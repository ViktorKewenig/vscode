/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { timeout } from '../../../../../base/common/async.js';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { Event } from '../../../../../base/common/event.js';
import { generateUuid } from '../../../../../base/common/uuid.js';
import { localize } from '../../../../../nls.js';
import { IChatQuestion } from '../../common/chatService/chatService.js';
import { getTextResponseFromStream, ChatMessageRole, IChatMessage, ILanguageModelsService } from '../../common/languageModels.js';
import { PlanningReadinessDimension } from '../../common/planning/chatPlanningReadiness.js';
import { getPlanningPhaseLabel, IPlanningRepositoryContext, IPlanningTarget, IPlanningTaskLens, IPlanningTransitionAnswer, isConcretePlanningArtifactReference, PlanningPhase, PlanningQuestionStage } from '../../common/planning/chatPlanningTransition.js';

export interface IPlanningQuestionGenerationContext {
	readonly userRequest: string;
	readonly modelId: string | undefined;
	readonly planningPhase: PlanningPhase;
	readonly questionStage: PlanningQuestionStage;
	readonly questionCount?: number;
	readonly missingDimensions?: readonly PlanningReadinessDimension[];
	readonly partialDimensions?: readonly PlanningReadinessDimension[];
	readonly shouldConfirmPlanningTarget?: boolean;
	readonly activeFilePath?: string;
	readonly selectedText?: string;
	readonly plannerNotes?: string;
	readonly recentConversation: readonly string[];
	readonly planningAnswers: readonly IPlanningTransitionAnswer[];
	readonly repositoryContext?: IPlanningRepositoryContext;
	readonly currentPlan?: string;
	readonly focusAreaLabel?: string;
	readonly focusHint?: string;
}

interface IGeneratedPlanningQuestion {
	readonly title?: string;
	readonly message?: string;
	readonly description?: string;
	readonly type?: 'text' | 'singleSelect' | 'multiSelect';
	readonly options?: ReadonlyArray<{
		readonly label?: string;
		readonly value?: string;
	}>;
	readonly required?: boolean;
	readonly allowFreeformInput?: boolean;
	readonly defaultValue?: string | string[];
}

interface IGeneratedPlanningQuestionEnvelope {
	readonly questions?: ReadonlyArray<IGeneratedPlanningQuestion>;
}

export interface IGeneratedPlanningQuestionsResult {
	readonly questions: IChatQuestion[];
	readonly modelId: string;
}

const preferredPlanningDefaultModelFamilies = ['gpt-4.1'];
const planningModelRegistrationPollMs = 1000;
const planningModelRegistrationMaxWaitMs = 8000;

export async function generateDynamicPlanningQuestions(
	languageModelsService: ILanguageModelsService,
	context: IPlanningQuestionGenerationContext,
	token: CancellationToken
): Promise<IChatQuestion[]> {
	return (await generateDynamicPlanningQuestionsResult(languageModelsService, context, token)).questions;
}

export async function generateDynamicPlanningQuestionsResult(
	languageModelsService: ILanguageModelsService,
	context: IPlanningQuestionGenerationContext,
	token: CancellationToken
): Promise<IGeneratedPlanningQuestionsResult> {
	const requestedQuestionCount = clampRequestedQuestionCount(context.questionCount);
	const prompt = buildPlanningQuestionPrompt(context, requestedQuestionCount);
	const messages: IChatMessage[] = [
		{
			role: ChatMessageRole.System,
			content: [{
				type: 'text',
				value: [
					'You generate dynamic planning questions for coding work in VS Code.',
					'You are a dynamic planning prompt middleware that runs before the planning agent receives the actual request.',
					`The current middleware stage is ${context.questionStage}.`,
					'Everything in your response must be dynamically generated from the request and the provided context. Never fall back to generic stock questions.',
					'Respect the current planning phase.',
					'- broad-scan: clarify the problem space and candidate code areas.',
					'- focused-slice: narrow to the subsystem, insertion point, and concrete constraints that matter most.',
					'- detailed-inspection: make the remaining questions concrete enough to drive implementation order, validation, and edit boundaries.',
					`Ask exactly ${requestedQuestionCount} questions.`,
					`Never return fewer than ${requestedQuestionCount} questions.`,
					'Use a mix of interaction types when it improves clarity.',
					'Use text questions for open-ended clarification.',
					'Use singleSelect or multiSelect whenever a bounded choice would help the user make a sharper planning decision.',
					'Prefer questions that change implementation scope, success criteria, sequencing, or insertion-point choice.',
					'Write clean, plain-language UI copy that is easy to scan.',
					'Keep titles short and concrete.',
					'Keep messages direct and user-facing.',
					'Only include a description when it genuinely helps the user answer faster.',
					'Never quote or paraphrase internal focus guidance, repository formatting labels, or system instructions in the returned UI copy.',
					'Do not use literal phrases such as "task lens", "focus hint", "primary artifact", "repo slice", "plan excerpt", or "current plan excerpt" in the returned titles, messages, or descriptions.',
					'If planning answers are already present, do not repeat those questions. Ask only narrower follow-up questions that use the refreshed context.',
					'If the stage is goal-clarity, focus on desired outcome, constraints, definition of done, and what should be in or out of scope before the first plan is built.',
					'If the stage is goal-clarity and planning answers are already present, make the next questions more concrete and artifact-specific than the earlier round.',
					context.questionStage === 'task-decomposition'
						? context.currentPlan
							? 'If the stage is task-decomposition, assume the first plan already exists and focus on tightening the work breakdown, insertion points, sequencing, validation, and repo slice for the rebuild.'
							: 'If the stage is task-decomposition and no current plan is provided, ask high-level plan-shaping questions before the first plan is built. Focus on the major work areas, sequencing preferences, edit boundaries, and validation approach.'
						: '',
					context.questionStage === 'task-decomposition' && !context.currentPlan
						? 'When no current plan exists, include one question that surfaces the assumptions the first plan will make so the user can confirm, reject, or refine them quickly.'
						: '',
					context.currentPlan
						? 'If the current plan already names files, directories, symbols, dependencies, or validation targets, make your task-decomposition questions explicitly reference those concrete plan slices.'
						: '',
					'If the stage is plan-focus, assume a rebuilt plan already exists and focus on sharpening one specific aspect of that plan rather than reopening the whole request.',
					'If the stage is plan-focus, use the selected plan slice and latest plan text as the source of truth. Your questions should feel like a zoom-in on that exact slice, not generic follow-up.',
					'If the stage is task-decomposition or plan-focus, treat prior goal-clarity answers as settled inputs, not new question topics.',
					'Never ask task-decomposition or plan-focus questions that re-open goal, scope, non-goals, or definition-of-done themes.',
					'When the current workspace is already clear, do not ask which repository or workspace to use. Ask about the file, directory, subsystem, or related artifacts inside it instead.',
					'If a primary artifact hint is provided, start from that artifact and only ask a repo-targeting question if the file, folder, or subsystem is still genuinely ambiguous.',
					!context.repositoryContext?.taskLens?.primaryArtifact && context.repositoryContext?.primaryArtifactHint
						? 'The exact primary artifact is still unresolved. Your first goal-clarity question should pin down the actual file, directory, symbol, or subsystem inside the current workspace.'
						: '',
					'If a task lens is provided, treat it as the best current summary of the work: the task kind, target artifact, adjacent artifacts, desired outcome, expected deliverable, validation targets, risks, and open decisions.',
					'Ask about unresolved decisions in that task lens rather than asking for generic project metadata.',
					'For bug-fix work, prefer symptom boundary, likely code path, preserved behavior, and validation questions.',
					'For feature work, prefer user outcome, surface area, acceptance, and dependency questions.',
					'For refactor work, prefer preserved behavior, edit boundaries, migration risk, and validation questions.',
					'For analysis work, prefer target artifact, related inputs, evidence, comparison set, and output questions.',
					'For test work, prefer failing path, expected assertions, harness, and coverage questions.',
					'For investigation work, prefer code path, evidence, uncertainty, and next-check questions.',
					'If the request intent is data analysis, lead with the data file, related schema, output, or directory that matters most.',
					'If the request intent is data analysis, prefer questions about the data file, related files, desired output, or result validation. Do not ask generic tooling or library questions unless the user or repo context suggests they matter.',
					'If the request intent is data analysis and the stage is goal-clarity, ask directly what kind of analysis the user wants. Prefer a multi-select question with concrete options such as summary statistics, group comparison, patterns or outliers, data quality checks, codebook creation, and a freeform option.',
					'If the request intent is data analysis and the stage is goal-clarity, ask what the user wants to learn, decide, or communicate with the analysis. Audience and helpful data context are strong follow-up topics when question budget allows.',
					'If the request intent is script work, lead with the script, entrypoint, or runtime context that matters most.',
					requestedQuestionCount > 1 ? 'For goal-clarity, prefer at least one structured choice question plus one open text question unless the context is genuinely too ambiguous.' : '',
					context.questionStage === 'task-decomposition' ? 'For task-decomposition, prefer a structured work-breakdown or insertion-point question and avoid drifting back into abstract scope questions.' : '',
					context.questionStage === 'plan-focus' ? 'For plan-focus, return a complementary set of follow-up controls that zoom in on the named focus area using the current plan and narrowed repo context. At least one question should lock the concrete repo slice, file, or subsystem, and at least one should sharpen risk, validation, sequencing, or dependency handling.' : '',
					'Descriptions should be concise and help the user understand why the question matters.',
					'Do not repeat the same question theme across both stages.',
					context.shouldConfirmPlanningTarget ? 'If the concrete file, folder, or subsystem is still ambiguous, use one sharply-targeted question to pin it down inside the current workspace.' : '',
					'Do not say or imply that the user has an existing workflow unless the user or repository context explicitly describes one.',
					'Return JSON only with the shape {"questions":[...]} and no markdown.'
				].join(' ')
			}]
		},
		{
			role: ChatMessageRole.User,
			content: [{ type: 'text', value: prompt }]
		}
	];

	let lastError: Error | undefined;
	let providerRetryAttempted = false;
	while (true) {
		const candidateModelIds = await getCandidateModelIds(languageModelsService, context.modelId);
		if (candidateModelIds.length === 0) {
			if (!providerRetryAttempted && shouldWaitForLanguageModelProvider(languageModelsService, context.modelId)) {
				providerRetryAttempted = true;
				const changed = await waitForLanguageModelRegistration(languageModelsService, context.modelId, token);
				if (changed) {
					continue;
				}

				throw new Error(localize(
					'chat.dynamicPlanning.modelProviderUnavailable',
					'No active language model is ready to generate planning questions yet. Try again in a moment.'
				));
			}

			throw new Error(localize(
				'chat.dynamicPlanning.noLanguageModel',
				'No language model is available to generate planning questions.'
			));
		}

		for (const modelId of candidateModelIds) {
			try {
				const normalized = requestedQuestionCount > 0
					? await requestModelPlanningQuestions(languageModelsService, modelId, messages, context, token)
					: [];
				const finalized = finalizeGeneratedQuestions(normalized, context);
				if (finalized.length > 0) {
					return { questions: finalized, modelId };
				}

				lastError = new Error(localize(
					'chat.dynamicPlanning.noUsableQuestions',
					'Language model "{0}" did not return enough usable planning questions.',
					modelId
				));
			} catch (error) {
				lastError = error instanceof Error
					? error
					: new Error(localize('chat.dynamicPlanning.unknownGenerationError', 'Planning question generation failed.'));
			}
		}

		if (!providerRetryAttempted && isMissingChatProviderError(lastError)) {
			providerRetryAttempted = true;
			const changed = await waitForLanguageModelRegistration(languageModelsService, context.modelId, token);
			if (changed) {
				continue;
			}
		}

		break;
	}

	if (isMissingChatProviderError(lastError)) {
		throw new Error(localize(
			'chat.dynamicPlanning.modelProviderUnavailable',
			'No active language model is ready to generate planning questions yet. Try again in a moment.'
		));
	}

	throw lastError ?? new Error(localize('chat.dynamicPlanning.unknownGenerationError', 'Planning question generation failed.'));
}

async function getCandidateModelIds(languageModelsService: ILanguageModelsService, preferredModelId: string | undefined): Promise<string[]> {
	const candidateModelIds: string[] = [];
	const seen = new Set<string>();
	const providerBackedModelIds = await getProviderBackedModelIds(languageModelsService, preferredModelId);
	const availableModelIds = providerBackedModelIds.length > 0 ? providerBackedModelIds : languageModelsService.getLanguageModelIds();
	const pushCandidate = (modelId: string | undefined) => {
		if (!modelId || seen.has(modelId)) {
			return;
		}
		if (!isExecutablePlanningModelId(modelId)) {
			return;
		}

		const metadata = languageModelsService.lookupLanguageModel(modelId);
		if (metadata?.targetChatSessionType) {
			return;
		}

		seen.add(modelId);
		candidateModelIds.push(modelId);
	};

	if (!providerBackedModelIds.length || providerBackedModelIds.includes(preferredModelId ?? '')) {
		pushCandidate(preferredModelId);
	}

	for (const modelId of availableModelIds) {
		const metadata = languageModelsService.lookupLanguageModel(modelId);
		if (metadata?.capabilities?.toolCalling && !metadata.targetChatSessionType) {
			pushCandidate(modelId);
		}
	}

	for (const modelId of availableModelIds) {
		pushCandidate(modelId);
	}

	return candidateModelIds;
}

async function getProviderBackedModelIds(languageModelsService: ILanguageModelsService, preferredModelId: string | undefined): Promise<string[]> {
	const result = new Set<string>();
	const preferredMetadata = preferredModelId ? languageModelsService.lookupLanguageModel(preferredModelId) : undefined;
	const preferredVendor = preferredMetadata?.vendor;
	const shouldPreferPlanningDefault = !preferredModelId || !isExecutablePlanningModelId(preferredModelId) || !preferredMetadata;

	if (shouldPreferPlanningDefault) {
		for (const modelId of await getPreferredPlanningDefaultModelIds(languageModelsService)) {
			result.add(modelId);
		}
	}

	if (preferredVendor && preferredMetadata?.id) {
		for (const modelId of await languageModelsService.selectLanguageModels({
			vendor: preferredVendor,
			id: preferredMetadata.id,
			family: preferredMetadata.family,
			version: preferredMetadata.version,
		})) {
			if (isExecutablePlanningModelId(modelId)) {
				result.add(modelId);
			}
		}
	}

	if (preferredVendor) {
		for (const modelId of await languageModelsService.selectLanguageModels({ vendor: preferredVendor })) {
			if (isExecutablePlanningModelId(modelId)) {
				result.add(modelId);
			}
		}
	}

	for (const modelId of await languageModelsService.selectLanguageModels({})) {
		if (isExecutablePlanningModelId(modelId)) {
			result.add(modelId);
		}
	}

	return [...result];
}

async function getPreferredPlanningDefaultModelIds(languageModelsService: ILanguageModelsService): Promise<string[]> {
	const result = new Set<string>();
	const pushModelId = (modelId: string) => {
		if (isExecutablePlanningModelId(modelId)) {
			result.add(modelId);
		}
	};

	for (const family of preferredPlanningDefaultModelFamilies) {
		for (const modelId of await languageModelsService.selectLanguageModels({ id: family })) {
			pushModelId(modelId);
		}
		for (const modelId of await languageModelsService.selectLanguageModels({ family })) {
			pushModelId(modelId);
		}
	}

	for (const modelId of languageModelsService.getLanguageModelIds()) {
		if (isPreferredPlanningDefaultModel(languageModelsService.lookupLanguageModel(modelId))) {
			pushModelId(modelId);
		}
	}

	return [...result];
}

function isPreferredPlanningDefaultModel(metadata: ReturnType<ILanguageModelsService['lookupLanguageModel']>): boolean {
	if (!metadata) {
		return false;
	}

	const normalizedId = metadata.id?.toLowerCase();
	const normalizedFamily = metadata.family?.toLowerCase();
	return preferredPlanningDefaultModelFamilies.some(family =>
		normalizedId === family
		|| normalizedId?.startsWith(`${family}-`) === true
		|| normalizedFamily === family
		|| normalizedFamily?.startsWith(`${family}-`) === true
	);
}

function isExecutablePlanningModelId(modelId: string): boolean {
	return modelId !== 'copilot/auto';
}

function shouldWaitForLanguageModelProvider(languageModelsService: ILanguageModelsService, preferredModelId: string | undefined): boolean {
	const hasObservableLanguageModelEvents = languageModelsService.onDidChangeLanguageModels !== Event.None
		|| languageModelsService.onDidChangeLanguageModelVendors !== Event.None;
	const hasRegisteredVendors = languageModelsService.getVendors().length > 0;
	const hasExecutableModelIds = languageModelsService.getLanguageModelIds().some(modelId => isExecutablePlanningModelId(modelId));

	if (!preferredModelId) {
		return hasObservableLanguageModelEvents || hasRegisteredVendors || hasExecutableModelIds;
	}

	if (!isExecutablePlanningModelId(preferredModelId)) {
		return hasObservableLanguageModelEvents || hasRegisteredVendors || hasExecutableModelIds;
	}

	return !!languageModelsService.lookupLanguageModel(preferredModelId)?.vendor
		|| hasObservableLanguageModelEvents
		|| hasRegisteredVendors;
}

function isMissingChatProviderError(error: Error | undefined): boolean {
	return !!error && /Chat provider for model .+ is not registered\./.test(error.message);
}

async function waitForLanguageModelRegistration(
	languageModelsService: ILanguageModelsService,
	preferredModelId: string | undefined,
	token: CancellationToken,
): Promise<boolean> {
	const preferredVendor = preferredModelId ? languageModelsService.lookupLanguageModel(preferredModelId)?.vendor : undefined;
	const deadline = Date.now() + planningModelRegistrationMaxWaitMs;
	while (!token.isCancellationRequested && Date.now() < deadline) {
		const changed = await waitForLanguageModelChange(languageModelsService, preferredVendor, token, Math.min(planningModelRegistrationPollMs, Math.max(deadline - Date.now(), 0)));
		if (changed) {
			return true;
		}
	}

	return false;
}

async function waitForLanguageModelChange(
	languageModelsService: ILanguageModelsService,
	preferredVendor: string | undefined,
	token: CancellationToken,
	timeoutMs: number,
): Promise<boolean> {
	const vendorChangeEvent = preferredVendor
		? Event.filter(languageModelsService.onDidChangeLanguageModelVendors, vendors => vendors.includes(preferredVendor))
		: languageModelsService.onDidChangeLanguageModelVendors;
	const modelChangeEvent = preferredVendor
		? Event.filter(languageModelsService.onDidChangeLanguageModels, vendor => vendor === preferredVendor)
		: languageModelsService.onDidChangeLanguageModels;
	const changeEvent = Event.any(vendorChangeEvent, modelChangeEvent);

	return new Promise<boolean>(resolve => {
		let done = false;
		const finish = (changed: boolean) => {
			if (done) {
				return;
			}
			done = true;
			listener.dispose();
			timer.cancel();
			cancellationListener.dispose();
			resolve(changed);
		};
		const listener = changeEvent(() => finish(true));
		const timer = timeout(timeoutMs);
		timer.then(() => finish(false));
		const cancellationListener = token.onCancellationRequested(() => finish(false));
	});
}

async function requestModelPlanningQuestions(
	languageModelsService: ILanguageModelsService,
	modelId: string,
	messages: IChatMessage[],
	context: IPlanningQuestionGenerationContext,
	token: CancellationToken,
): Promise<IChatQuestion[]> {
	const response = await languageModelsService.sendChatRequest(modelId, undefined, messages, {}, token);
	const responseText = await getTextResponseFromStream(response);
	const parsed = parseQuestionEnvelope(responseText);
	return normalizeGeneratedQuestions(parsed, context);
}

function buildPlanningQuestionPrompt(context: IPlanningQuestionGenerationContext, requestedQuestionCount: number): string {
	const sections = [
		`Planning phase:\n${context.planningPhase} (${getPlanningPhaseLabel(context.planningPhase)})`,
		`Question stage:\n${context.questionStage}`,
		`Requested question count:\n${requestedQuestionCount}`,
		`User request:\n${context.userRequest.trim()}`,
	];

	if (context.activeFilePath) {
		sections.push(`Active file:\n${context.activeFilePath}`);
	}

	if (context.selectedText) {
		sections.push(`Selected code or text:\n${truncate(context.selectedText.trim(), 1400)}`);
	}

	if (context.plannerNotes) {
		sections.push(`Planner notes:\n${context.plannerNotes}`);
	}

	if (context.currentPlan) {
		sections.push(`Current plan:\n${truncate(context.currentPlan, 1800)}`);
	}

	const planAnchors = extractPlanAnchors(context.currentPlan);
	if (planAnchors.length > 0) {
		sections.push(`Current plan anchors:\n${planAnchors.map(anchor => `- ${anchor}`).join('\n')}`);
	}

	if (context.planningAnswers.length > 0) {
		sections.push(`Existing planning answers:\n${context.planningAnswers.map(answer => `- ${answer.question}: ${answer.answer}`).join('\n')}`);
	}

	if (context.recentConversation.length > 0) {
		sections.push(`Recent planning conversation:\n${context.recentConversation.map(entry => `- ${entry}`).join('\n')}`);
	}

	if (context.repositoryContext) {
		sections.push(formatRepositoryContext(context.repositoryContext));
	}

	if (context.repositoryContext?.workspaceFolders?.length === 1 || context.repositoryContext?.workspaceRoot) {
		sections.push('Workspace grounding:\nThe current repository is already known. Ask about the file, directory, subsystem, or related artifacts inside it instead of asking which repo to use.');
	}

	if (!context.repositoryContext?.taskLens?.primaryArtifact && context.repositoryContext?.primaryArtifactHint) {
		sections.push(`Artifact targeting:\nThe request points at ${context.repositoryContext.primaryArtifactHint}, but the exact file, directory, symbol, or subsystem is not pinned down yet. Use the first goal-clarity question to lock that down inside the current workspace.`);
	}

	if (context.focusHint) {
		sections.push(`Internal focus guidance (do not quote verbatim):\n${context.focusHint}`);
	}

	if (context.focusAreaLabel) {
		sections.push(`User-selected focus area:\n${context.focusAreaLabel}`);
	}

	if (context.missingDimensions?.length) {
		sections.push(`Missing planning dimensions:\n${context.missingDimensions.join(', ')}`);
	}

	if (context.partialDimensions?.length) {
		sections.push(`Partially-covered dimensions:\n${context.partialDimensions.join(', ')}`);
	}

	sections.push([
		context.questionStage === 'goal-clarity'
			? `Return exactly ${requestedQuestionCount} questions that clarify the implementation goal, constraints, non-goals, and what success looks like before the first plan is built.`
			: context.questionStage === 'task-decomposition'
				? context.currentPlan
					? `Return exactly ${requestedQuestionCount} questions that tighten the first plan into a stronger work breakdown, insertion-point choice, repo slice, and validation path. At least two questions should hook into concrete files, steps, dependencies, or validation targets already named in the current plan or task lens.`
					: `Return exactly ${requestedQuestionCount} questions that help the user co-create the high-level plan before the first detailed plan is built. Cover the major work areas, ordering, scope boundaries, and validation approach without getting into low-level implementation minutiae.`
				: `Return exactly ${requestedQuestionCount} questions that zoom in on one specific aspect of the rebuilt plan using the named focus area, the latest plan text, and the narrowed repo context.`,
		context.questionStage === 'goal-clarity'
			? 'Prefer a light but engaging pre-planning UX: the questions should feel closer to ask-questions than a heavy middleware banner.'
			: context.questionStage === 'task-decomposition'
				? context.currentPlan
					? 'Prefer a concrete refinement UX: one question should usually lock in work breakdown, insertion point, file or repo slice, or validation.'
					: 'Prefer a collaborative plan-shaping UX: use checkable choices where possible so the user can quickly confirm, reject, or add work areas before seeing a full plan.'
				: 'Prefer a focused refinement UX: the questions should feel like a zoom-in on one part of the plan, not a restart of the whole plan. When possible, cover the exact repo slice, the key unresolved decision, and the evidence or validation needed for that focused change.',
		'Avoid generic project-management questions.',
		'Do not ask for information that is already clear from the repo context, current plan, or earlier answers.',
		'Use the richer repository context to narrow the work as the phase becomes more specific.',
		'In broad-scan, keep questions exploratory.',
		'In focused-slice, prefer subsystem, insertion-point, and constraints questions.',
		'In detailed-inspection, prefer implementation-order, validation, and edit-boundary questions.',
		context.missingDimensions?.length ? `Cover these missing dimensions first: ${context.missingDimensions.join(', ')}.` : '',
		context.partialDimensions?.length ? `Use the remaining questions to sharpen these partial dimensions: ${context.partialDimensions.join(', ')}.` : '',
		context.questionStage === 'goal-clarity'
			? 'Do not ask sequencing or implementation-order questions unless they are necessary to understand the goal.'
			: context.questionStage === 'task-decomposition'
				? context.currentPlan
					? 'Do not repeat goal-clarity questions that are already answered in the planning context.'
					: 'Do not ask the user to approve a plan that does not exist yet. Ask which high-level work areas and ordering choices should shape the first plan.'
				: 'Do not drift back into broad decomposition or restart the plan from scratch.',
		context.questionStage === 'task-decomposition' && !context.currentPlan ? 'Include one assumption-confirmation question so the user can quickly confirm, reject, or refine what the first plan will assume.' : '',
		context.shouldConfirmPlanningTarget ? 'Do not ask the user to confirm the primary repo target again; that is being collected separately.' : '',
		'For singleSelect and multiSelect questions, defaultValue must reference the option label.'
	].join(' '));

	sections.push([
		'JSON schema:',
		'{"questions":[{"title":"...","message":"...","type":"text|singleSelect|multiSelect","required":true|false,"allowFreeformInput":true|false,"options":[{"label":"...","value":"..."}],"defaultValue":"..."|["..."]}]}'
	].join('\n'));

	return sections.join('\n\n');
}

function formatRepositoryContext(repositoryContext: IPlanningRepositoryContext): string {
	return [
		'Repository context:',
		`Scope: ${repositoryContext.scope}`,
		`Workspace root: ${repositoryContext.workspaceRoot || 'Not provided'}`,
		`Planning target: ${formatPlanningTarget(repositoryContext.planningTarget)}`,
		`Request intent: ${repositoryContext.requestIntent || 'Not provided'}`,
		`Task lens:\n${formatTaskLens(repositoryContext.taskLens)}`,
		`Primary artifact hint: ${repositoryContext.primaryArtifactHint || 'Not provided'}`,
		`Related artifact hints: ${repositoryContext.relatedArtifactHints?.join(', ') || 'None'}`,
		`Focus summary: ${repositoryContext.focusSummary || 'Not provided'}`,
		`Focus queries: ${repositoryContext.focusQueries.join(', ') || 'None'}`,
		`Workspace folders: ${repositoryContext.workspaceFolders?.join(', ') || 'None'}`,
		`Workspace top-level entries: ${repositoryContext.workspaceTopLevelEntries?.join(', ') || 'None'}`,
		`Working set files: ${repositoryContext.workingSetFiles?.join(', ') || 'None'}`,
		`Active document symbols: ${formatSymbols(repositoryContext.activeDocumentSymbols)}`,
		`Workspace symbol matches: ${formatSymbols(repositoryContext.workspaceSymbolMatches)}`,
		`Nearby files: ${repositoryContext.nearbyFiles.join(', ') || 'None'}`,
		`Relevant file snippets:\n${formatSnippets(repositoryContext.relevantSnippets)}`,
	].join('\n');
}

function parseQuestionEnvelope(raw: string): IGeneratedPlanningQuestionEnvelope | undefined {
	const trimmed = raw.trim();
	if (!trimmed) {
		return undefined;
	}

	const withoutFences = trimmed.startsWith('```')
		? trimmed.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '')
		: trimmed;

	try {
		return JSON.parse(withoutFences) as IGeneratedPlanningQuestionEnvelope;
	} catch {
		return undefined;
	}
}

function normalizeGeneratedQuestions(parsed: IGeneratedPlanningQuestionEnvelope | undefined, context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	const questions = parsed?.questions;
	if (!questions?.length) {
		return [];
	}

	return finalizeGeneratedQuestions(questions
		.map(normalizeGeneratedQuestion)
		.filter((question): question is IChatQuestion => !!question), context);
}

function normalizeGeneratedQuestion(question: IGeneratedPlanningQuestion): IChatQuestion | undefined {
	const title = normalizeText(question.title);
	if (!title) {
		return undefined;
	}

	const normalizedType = question.type ?? 'text';
	const options = (question.options ?? [])
		.map(option => {
			const label = normalizeText(option.label);
			const value = normalizeText(option.value) ?? label;
			if (!label || !value) {
				return undefined;
			}

			return {
				id: label,
				label,
				value
			};
		})
		.filter((option): option is NonNullable<typeof option> => !!option);

	const type = options.length >= 2 ? normalizedType : 'text';
	const normalizedMessage = normalizeText(question.message);
	const normalizedDescription = normalizeText(question.description);

	return {
		id: generateUuid(),
		title,
		message: normalizedMessage ?? normalizedDescription,
		description: normalizedDescription,
		type,
		options: type === 'text' ? undefined : options,
		required: question.required === true,
		allowFreeformInput: question.allowFreeformInput ?? true,
		defaultValue: type === 'text' ? normalizeDefaultTextValue(question.defaultValue) : normalizeDefaultOptionValue(question.defaultValue, options)
	};
}

function normalizeDefaultTextValue(defaultValue: string | string[] | undefined): string | undefined {
	return typeof defaultValue === 'string' ? normalizeText(defaultValue) : undefined;
}

function normalizeDefaultOptionValue(defaultValue: string | string[] | undefined, options: ReadonlyArray<{ label: string }>): string | string[] | undefined {
	if (typeof defaultValue === 'string') {
		return options.some(option => option.label === defaultValue) ? defaultValue : undefined;
	}

	if (Array.isArray(defaultValue)) {
		const matchingValues = defaultValue.filter(value => options.some(option => option.label === value));
		return matchingValues.length > 0 ? matchingValues : undefined;
	}

	return undefined;
}

function normalizeText(value: string | undefined): string | undefined {
	return value?.replace(/\s+/g, ' ').trim() || undefined;
}

function truncate(value: string, maxLength: number): string {
	return value.length <= maxLength ? value : `${value.slice(0, maxLength)}...`;
}

function extractPlanAnchors(currentPlan: string | undefined): string[] {
	if (!currentPlan) {
		return [];
	}

	const anchors: string[] = [];
	const seen = new Set<string>();
	for (const rawLine of currentPlan.split(/\r?\n/g)) {
		const normalized = rawLine
			.replace(/^\s{0,3}(?:[-*+]|\d+[.)]|#{1,6})\s*/, '')
			.replace(/\s+/g, ' ')
			.trim();
		if (normalized.length < 12) {
			continue;
		}

		const key = normalized.toLowerCase();
		if (seen.has(key)) {
			continue;
		}

		seen.add(key);
		anchors.push(normalized);
		if (anchors.length >= 6) {
			break;
		}
	}

	return anchors;
}

function extractFileLikeMentions(value: string | undefined): string[] {
	if (!value) {
		return [];
	}

	return value.match(/(?:[A-Za-z0-9_.-]+[\\/])*[A-Za-z0-9_.-]+\.[A-Za-z0-9]+/g) ?? [];
}

function isDataAnalysisGoalClarityContext(context: IPlanningQuestionGenerationContext): boolean {
	return context.questionStage === 'goal-clarity' && context.repositoryContext?.requestIntent === 'data-analysis';
}

function isInitialTaskDecompositionContext(context: IPlanningQuestionGenerationContext): boolean {
	return context.questionStage === 'task-decomposition' && !context.currentPlan?.trim();
}

function addSupplementalPlanningQuestions(questions: readonly IChatQuestion[], context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	return addInitialPlanShapingQuestions(addDataAnalysisGoalQuestions(questions, context), context);
}

function addDataAnalysisGoalQuestions(questions: readonly IChatQuestion[], context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	if (!isDataAnalysisGoalClarityContext(context)) {
		return [...questions];
	}

	const enriched = [...questions];
	for (const candidate of createDataAnalysisGoalQuestions(context)) {
		if (enriched.some(existing => isSimilarDataAnalysisQuestion(existing, candidate))) {
			continue;
		}

		enriched.push(candidate);
	}

	return enriched;
}

function addInitialPlanShapingQuestions(questions: readonly IChatQuestion[], context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	if (!isInitialTaskDecompositionContext(context)) {
		return [...questions];
	}

	const assumptionQuestion = createInitialPlanAssumptionQuestion(context);
	if (questions.some(question => isPlanAssumptionQuestion(question) || computeOverlap(normalizeQuestionPrompt(question), normalizeQuestionPrompt(assumptionQuestion)) >= 0.42)) {
		return [...questions];
	}

	return [assumptionQuestion, ...questions];
}

function createInitialPlanAssumptionQuestion(context: IPlanningQuestionGenerationContext): IChatQuestion {
	const taskLens = context.repositoryContext?.taskLens;
	const options: { id: string; label: string; value: string }[] = [];
	const seen = new Set<string>();
	const pushOption = (label: string, value: string) => {
		const normalized = label.replace(/\s+/g, ' ').trim();
		const key = normalized.toLowerCase();
		if (!normalized || seen.has(key)) {
			return;
		}

		seen.add(key);
		options.push({
			id: `plan-assumption-${options.length}`,
			label: normalized,
			value,
		});
	};

	const primaryTarget = taskLens?.primaryArtifact ?? context.repositoryContext?.primaryArtifactHint ?? context.repositoryContext?.planningTarget?.label;
	if (primaryTarget && isConcretePlanningArtifactReference(primaryTarget)) {
		pushOption(localize('chat.dynamicPlanning.planAssumptionPrimaryTarget', 'Use {0} as the primary target', primaryTarget), `primary-target:${primaryTarget}`);
	}

	if (taskLens?.desiredOutcome) {
		pushOption(localize('chat.dynamicPlanning.planAssumptionOutcome', 'Optimize the plan for: {0}', taskLens.desiredOutcome), `outcome:${taskLens.desiredOutcome}`);
	}

	const relatedArtifacts = taskLens?.secondaryArtifacts ?? context.repositoryContext?.relatedArtifactHints;
	if (relatedArtifacts?.length) {
		pushOption(localize('chat.dynamicPlanning.planAssumptionRelatedArtifacts', 'Treat {0} as supporting context, not extra scope unless needed', relatedArtifacts.slice(0, 2).join(', ')), `supporting-context:${relatedArtifacts.slice(0, 2).join(', ')}`);
	}

	if (taskLens?.validationTargets?.length) {
		pushOption(localize('chat.dynamicPlanning.planAssumptionValidation', 'Reserve validation for {0}', taskLens.validationTargets.slice(0, 2).join(', ')), `validation:${taskLens.validationTargets.slice(0, 2).join(', ')}`);
	}

	pushOption(localize('chat.dynamicPlanning.planAssumptionStatedGoal', 'Keep the plan focused on the stated goal'), 'stated-goal');
	pushOption(localize('chat.dynamicPlanning.planAssumptionWorkspaceBoundary', 'Use the current workspace context as the boundary'), 'workspace-boundary');
	pushOption(localize('chat.dynamicPlanning.planAssumptionValidationFallback', 'Include a lightweight validation step'), 'lightweight-validation');

	const trimmedOptions = options.slice(0, 5);
	return {
		id: 'dynamic-planning-plan-assumptions',
		type: 'multiSelect',
		title: localize('chat.dynamicPlanning.planAssumptionTitle', 'Plan Assumptions'),
		message: localize('chat.dynamicPlanning.planAssumptionMessage', 'Which assumptions should shape the first plan?'),
		description: localize('chat.dynamicPlanning.planAssumptionDescription', 'Confirm the assumptions that fit, or add corrections.'),
		required: false,
		allowFreeformInput: true,
		options: trimmedOptions,
		defaultValue: trimmedOptions.map(option => option.id),
	};
}

function createDataAnalysisGoalQuestions(context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	const target = context.repositoryContext?.taskLens?.primaryArtifact ?? context.repositoryContext?.primaryArtifactHint;
	const targetDescription = target && isConcretePlanningArtifactReference(target)
		? localize('chat.dynamicPlanning.analysisKindDescriptionWithTarget', 'Select any that fit for {0}, or add your own.', target)
		: localize('chat.dynamicPlanning.analysisKindDescription', 'Select any that fit, or add your own.');

	return [
		{
			id: 'dynamic-planning-analysis-kind',
			type: 'multiSelect',
			title: localize('chat.dynamicPlanning.analysisKindTitle', 'Analysis Type'),
			message: localize('chat.dynamicPlanning.analysisKindMessage', 'What kind of analysis are you looking to do?'),
			description: targetDescription,
			required: true,
			allowFreeformInput: true,
			options: [
				{ id: 'summary-statistics', label: localize('chat.dynamicPlanning.analysisKindSummaryStatistics', 'Summary Statistics'), value: 'summary-statistics' },
				{ id: 'group-comparison', label: localize('chat.dynamicPlanning.analysisKindGroupComparison', 'Compare Groups or Segments'), value: 'group-comparison' },
				{ id: 'patterns-outliers', label: localize('chat.dynamicPlanning.analysisKindPatternsOutliers', 'Find Patterns or Outliers'), value: 'patterns-outliers' },
				{ id: 'data-quality', label: localize('chat.dynamicPlanning.analysisKindDataQuality', 'Check Data Quality'), value: 'data-quality' },
				{ id: 'codebook', label: localize('chat.dynamicPlanning.analysisKindCodebook', 'Create or Update a Codebook'), value: 'codebook' },
			],
		},
		{
			id: 'dynamic-planning-analysis-goal',
			type: 'text',
			title: localize('chat.dynamicPlanning.analysisGoalTitle', 'Analysis Goal'),
			message: localize('chat.dynamicPlanning.analysisGoalMessage', 'What do you want to learn, decide, or communicate with this analysis?'),
			description: localize('chat.dynamicPlanning.analysisGoalDescription', 'A short research question, decision, or desired takeaway is enough.'),
			required: true,
			allowFreeformInput: true,
		},
		{
			id: 'dynamic-planning-analysis-audience',
			type: 'singleSelect',
			title: localize('chat.dynamicPlanning.analysisAudienceTitle', 'Audience'),
			message: localize('chat.dynamicPlanning.analysisAudienceMessage', 'Who is this analysis for?'),
			required: false,
			allowFreeformInput: true,
			options: [
				{ id: 'just-me', label: localize('chat.dynamicPlanning.analysisAudienceJustMe', 'Just Me'), value: 'just-me' },
				{ id: 'technical-collaborators', label: localize('chat.dynamicPlanning.analysisAudienceTechnicalCollaborators', 'Technical Collaborators'), value: 'technical-collaborators' },
				{ id: 'nontechnical-stakeholders', label: localize('chat.dynamicPlanning.analysisAudienceNontechnicalStakeholders', 'Nontechnical Stakeholders'), value: 'nontechnical-stakeholders' },
			],
		},
		{
			id: 'dynamic-planning-analysis-context',
			type: 'text',
			title: localize('chat.dynamicPlanning.analysisContextTitle', 'Data Context'),
			message: localize('chat.dynamicPlanning.analysisContextMessage', 'What context about the data should shape the analysis?'),
			description: localize('chat.dynamicPlanning.analysisContextDescription', 'Mention caveats, definitions, filters, or known issues that would change the plan.'),
			required: false,
			allowFreeformInput: true,
		},
	];
}

function isSimilarDataAnalysisQuestion(existing: IChatQuestion, candidate: IChatQuestion): boolean {
	const existingPrompt = normalizeQuestionPrompt(existing);
	const candidatePrompt = normalizeQuestionPrompt(candidate);
	if (computeOverlap(existingPrompt, candidatePrompt) >= 0.42) {
		return true;
	}

	switch (candidate.id) {
		case 'dynamic-planning-analysis-kind':
			return isAnalysisKindQuestion(existing);
		case 'dynamic-planning-analysis-goal':
			return isAnalysisGoalQuestion(existing);
		case 'dynamic-planning-analysis-audience':
			return isAnalysisAudienceQuestion(existing);
		case 'dynamic-planning-analysis-context':
			return isAnalysisContextQuestion(existing);
		default:
			return false;
	}
}

function finalizeGeneratedQuestions(questions: readonly IChatQuestion[], context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	const requestedQuestionCount = clampRequestedQuestionCount(context.questionCount);
	const deduped = dedupeQuestionsByPrompt(questions);
	const stageCandidates = context.questionStage === 'goal-clarity'
		? deduped
		: deduped.filter(question => !isOverlappingGoalClarityQuestion(question, context.planningAnswers));
	const stageFiltered = dedupeQuestionsByPrompt(addSupplementalPlanningQuestions(stageCandidates, context));
	if (stageFiltered.length < requestedQuestionCount) {
		return [];
	}

	const ranked = rankQuestionsForContext(stageFiltered, context);
	const trimmed = selectQuestionsWithStageMix(ranked, requestedQuestionCount, context);
	const hasStructuredQuestion = trimmed.some(question => question.type !== 'text');
	const hasTextQuestion = trimmed.some(question => question.type === 'text');

	if (context.questionStage === 'goal-clarity') {
		if (requestedQuestionCount === 1) {
			return trimmed.slice(0, 1);
		}

		return hasStructuredQuestion && hasTextQuestion ? trimmed : [];
	}

	if (context.questionStage === 'plan-focus' && requestedQuestionCount > 2) {
		return hasStructuredQuestion && hasTextQuestion ? trimmed : [];
	}

	if (requestedQuestionCount === 1) {
		return trimmed.slice(0, 1);
	}

	return hasStructuredQuestion ? trimmed : [];
}

function rankQuestionsForContext(questions: readonly IChatQuestion[], context: IPlanningQuestionGenerationContext): IChatQuestion[] {
	return [...questions]
		.map((question, index) => ({
			question,
			index,
			score: scoreQuestionForContext(question, context),
		}))
		.sort((left, right) => right.score - left.score || left.index - right.index)
		.map(entry => entry.question);
}

function selectQuestionsWithStageMix(
	rankedQuestions: readonly IChatQuestion[],
	requestedQuestionCount: number,
	context: IPlanningQuestionGenerationContext,
): IChatQuestion[] {
	if (requestedQuestionCount <= 1) {
		return rankedQuestions.slice(0, 1);
	}

	if (isDataAnalysisGoalClarityContext(context)) {
		return selectDataAnalysisGoalQuestions(rankedQuestions, requestedQuestionCount, context);
	}

	if (isInitialTaskDecompositionContext(context)) {
		return selectInitialPlanShapingQuestions(rankedQuestions, requestedQuestionCount);
	}

	const requireTextAndStructured = context.questionStage === 'goal-clarity' || (context.questionStage === 'plan-focus' && requestedQuestionCount > 2);
	if (!requireTextAndStructured) {
		return rankedQuestions.slice(0, requestedQuestionCount);
	}

	const hasConcretePrimaryArtifact = isConcretePlanningArtifactReference(context.repositoryContext?.taskLens?.primaryArtifact)
		|| isConcretePlanningArtifactReference(context.repositoryContext?.primaryArtifactHint);
	const hasUnresolvedPrimaryArtifact = !hasConcretePrimaryArtifact && !!context.repositoryContext?.primaryArtifactHint;
	const textQuestion = rankedQuestions.find(question => question.type === 'text');
	const structuredQuestion = rankedQuestions.find(question => question.type !== 'text');
	if (!textQuestion || !structuredQuestion) {
		return rankedQuestions.slice(0, requestedQuestionCount);
	}

	const selected: IChatQuestion[] = [];
	const seen = new Set<string>();
	const prioritizedQuestions = hasUnresolvedPrimaryArtifact && context.questionStage === 'goal-clarity'
		? [rankedQuestions[0], textQuestion, structuredQuestion, ...rankedQuestions]
		: [structuredQuestion, textQuestion, ...rankedQuestions];
	for (const question of prioritizedQuestions) {
		if (seen.has(question.id)) {
			continue;
		}

		seen.add(question.id);
		selected.push(question);
		if (selected.length >= requestedQuestionCount) {
			break;
		}
	}

	return selected;
}

function selectDataAnalysisGoalQuestions(
	rankedQuestions: readonly IChatQuestion[],
	requestedQuestionCount: number,
	context: IPlanningQuestionGenerationContext,
): IChatQuestion[] {
	const taskLens = context.repositoryContext?.taskLens;
	const hasConcretePrimaryArtifact = isConcretePlanningArtifactReference(taskLens?.primaryArtifact)
		|| isConcretePlanningArtifactReference(context.repositoryContext?.primaryArtifactHint);
	const hasUnresolvedPrimaryArtifact = !hasConcretePrimaryArtifact && !!context.repositoryContext?.primaryArtifactHint;
	const selected: IChatQuestion[] = [];
	const seen = new Set<string>();
	const pushQuestion = (question: IChatQuestion | undefined) => {
		if (!question || seen.has(question.id)) {
			return;
		}

		seen.add(question.id);
		selected.push(question);
	};

	if (hasUnresolvedPrimaryArtifact) {
		pushQuestion(rankedQuestions.find(isArtifactTargetingQuestion));
	}

	pushQuestion(rankedQuestions.find(isAnalysisKindQuestion));
	pushQuestion(rankedQuestions.find(isAnalysisGoalQuestion));
	pushQuestion(rankedQuestions.find(isAnalysisAudienceQuestion));
	pushQuestion(rankedQuestions.find(isAnalysisContextQuestion));

	for (const question of rankedQuestions) {
		pushQuestion(question);
		if (selected.length >= requestedQuestionCount) {
			break;
		}
	}

	return selected.slice(0, requestedQuestionCount);
}

function selectInitialPlanShapingQuestions(rankedQuestions: readonly IChatQuestion[], requestedQuestionCount: number): IChatQuestion[] {
	const selected: IChatQuestion[] = [];
	const seen = new Set<string>();
	const pushQuestion = (question: IChatQuestion | undefined) => {
		if (!question || seen.has(question.id)) {
			return;
		}

		seen.add(question.id);
		selected.push(question);
	};

	pushQuestion(rankedQuestions.find(isPlanAssumptionQuestion));
	pushQuestion(rankedQuestions.find(question => question.type !== 'text' && !isPlanAssumptionQuestion(question)));
	pushQuestion(rankedQuestions.find(question => question.type === 'text'));
	for (const question of rankedQuestions) {
		pushQuestion(question);
		if (selected.length >= requestedQuestionCount) {
			break;
		}
	}

	return selected.slice(0, requestedQuestionCount);
}

function isPlanAssumptionQuestion(question: IChatQuestion): boolean {
	if (question.id === 'dynamic-planning-plan-assumptions') {
		return true;
	}

	return /\b(assumption|assumptions|confirm.*plan|shape.*first plan|first plan.*assume)\b/i.test(normalizeQuestionPrompt(question));
}

function isArtifactTargetingQuestion(question: IChatQuestion): boolean {
	const prompt = normalizeQuestionPrompt(question);
	return /\b(which|what)\b.{0,60}\b(file|folder|directory|dataset|csv|tsv|json|schema|table|notebook)\b/i.test(prompt)
		|| /\b(primary|exact|target|anchor)\b.{0,60}\b(file|folder|directory|dataset|csv|tsv|json|schema|table|notebook)\b/i.test(prompt);
}

function isAnalysisKindQuestion(question: IChatQuestion): boolean {
	if (question.id === 'dynamic-planning-analysis-kind') {
		return true;
	}

	const prompt = normalizeQuestionPrompt(question);
	return /\b(kind|type|approach)\b.{0,60}\banalys/i.test(prompt)
		|| /\b(summary statistics|group comparison|segments?|patterns?|outliers?|data quality|codebook)\b/i.test(prompt);
}

function isAnalysisGoalQuestion(question: IChatQuestion): boolean {
	if (question.id === 'dynamic-planning-analysis-goal') {
		return true;
	}

	const prompt = normalizeQuestionPrompt(question);
	return /\b(learn|decide|communicate|research question|question|takeaway|want to know|goal|outcome)\b/i.test(prompt)
		&& /\banalys/i.test(prompt);
}

function isAnalysisAudienceQuestion(question: IChatQuestion): boolean {
	if (question.id === 'dynamic-planning-analysis-audience') {
		return true;
	}

	const prompt = normalizeQuestionPrompt(question);
	return /\b(audience|stakeholder|reader|recipient|for whom|who is this)\b/i.test(prompt);
}

function isAnalysisContextQuestion(question: IChatQuestion): boolean {
	if (question.id === 'dynamic-planning-analysis-context') {
		return true;
	}

	const prompt = normalizeQuestionPrompt(question);
	return /\b(data context|caveat|definition|filter|known issue|quality issue|assumption|context about the data)\b/i.test(prompt);
}

function clampRequestedQuestionCount(questionCount: number | undefined): number {
	return Math.min(Math.max(questionCount ?? 3, 1), 4);
}

function dedupeQuestionsByPrompt(questions: readonly IChatQuestion[]): IChatQuestion[] {
	const deduped: IChatQuestion[] = [];

	for (const candidate of questions) {
		const candidateText = normalizeQuestionPrompt(candidate);
		if (!candidateText) {
			continue;
		}

		if (deduped.some(existing => computeOverlap(normalizeQuestionPrompt(existing), candidateText) >= 0.72)) {
			continue;
		}

		deduped.push(candidate);
	}

	return deduped;
}

function isOverlappingGoalClarityQuestion(question: IChatQuestion, planningAnswers: readonly IPlanningTransitionAnswer[]): boolean {
	const prompt = normalizeQuestionPrompt(question);
	if (!prompt) {
		return false;
	}

	if (looksLikeGoalClarityQuestion(prompt)) {
		return true;
	}

	return planningAnswers.some(answer => {
		const previousPrompt = normalizeWhitespace(answer.question);
		return previousPrompt ? computeOverlap(previousPrompt, prompt) >= 0.42 : false;
	});
}

function looksLikeGoalClarityQuestion(prompt: string): boolean {
	return /\b(goal|outcome|success|definition of done|done|scope|non-goal|non goal|constraint|boundary|in scope|out of scope)\b/i.test(prompt);
}

function normalizeQuestionPrompt(question: IChatQuestion): string {
	return normalizeWhitespace([question.title, typeof question.message === 'string' ? question.message : undefined, question.description].filter(Boolean).join(' '));
}

function normalizeWhitespace(value: string | undefined): string {
	return value?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '';
}

function computeOverlap(left: string, right: string): number {
	if (!left || !right) {
		return 0;
	}

	const leftTokens = new Set(left.match(/[a-z0-9]{4,}/gi) ?? []);
	const rightTokens = new Set(right.match(/[a-z0-9]{4,}/gi) ?? []);
	if (leftTokens.size === 0 || rightTokens.size === 0) {
		return 0;
	}

	let matches = 0;
	for (const token of leftTokens) {
		if (rightTokens.has(token)) {
			matches += 1;
		}
	}

	return matches / Math.max(leftTokens.size, rightTokens.size);
}

function scoreQuestionForContext(question: IChatQuestion, context: IPlanningQuestionGenerationContext): number {
	const prompt = normalizeQuestionPrompt(question);
	if (!prompt) {
		return -Infinity;
	}

	const taskLens = context.repositoryContext?.taskLens;
	const planAnchors = extractPlanAnchors(context.currentPlan);
	const planFileMentions = extractFileLikeMentions(context.currentPlan);
	const hasConcretePrimaryArtifact = isConcretePlanningArtifactReference(taskLens?.primaryArtifact)
		|| isConcretePlanningArtifactReference(context.repositoryContext?.primaryArtifactHint);
	const hasUnresolvedPrimaryArtifact = !hasConcretePrimaryArtifact && !!context.repositoryContext?.primaryArtifactHint;
	const overlap = (values: readonly string[] | undefined, weight: number) => {
		if (!values?.length) {
			return 0;
		}

		return values.reduce((total, value) => total + computeOverlap(prompt, normalizeWhitespace(value)) * weight, 0);
	};
	const keywordOverlap = (value: string | undefined, weight: number) => value ? computeOverlap(prompt, normalizeWhitespace(value)) * weight : 0;
	let score = overlap(context.missingDimensions, 1.8)
		+ overlap(context.partialDimensions, 1.1)
		+ keywordOverlap(taskLens?.taskSummary, 1.8)
		+ keywordOverlap(taskLens?.desiredOutcome, 2.2)
		+ keywordOverlap(taskLens?.primaryArtifact ?? context.repositoryContext?.primaryArtifactHint, 2.4)
		+ overlap(taskLens?.secondaryArtifacts ?? context.repositoryContext?.relatedArtifactHints, 1.5)
		+ overlap(taskLens?.planAreas, context.questionStage === 'goal-clarity' ? 0.7 : 2.4)
		+ overlap(taskLens?.validationTargets, 1.6)
		+ overlap(taskLens?.riskAreas, 1.5)
		+ overlap(taskLens?.unknowns, 2)
		+ keywordOverlap(context.focusAreaLabel, context.questionStage === 'plan-focus' ? 2.6 : 0.8)
		+ keywordOverlap(context.currentPlan, context.questionStage === 'goal-clarity' ? 0.4 : 1.8)
		+ overlap(planAnchors, context.questionStage === 'goal-clarity' ? 0.4 : context.questionStage === 'task-decomposition' ? 2.8 : 3)
		+ overlap(planFileMentions, context.questionStage === 'goal-clarity' ? 0.5 : 2.6);

	if (context.questionStage === 'task-decomposition' && question.type !== 'text') {
		score += 0.35;
	}

	if (context.questionStage === 'plan-focus' && question.type === 'text') {
		score += 0.25;
	}

	if (context.questionStage === 'task-decomposition' && planAnchors.length > 0) {
		if (computeOverlap(prompt, planAnchors.join(' ')) === 0 && computeOverlap(prompt, planFileMentions.join(' ')) === 0) {
			score -= 1.8;
		}

		if (/\b(anything else|other context|additional context|other preferences|anything to keep in mind)\b/i.test(prompt)) {
			score -= 2.4;
		}
	}

	if (context.questionStage === 'plan-focus' && context.focusAreaLabel) {
		if (computeOverlap(prompt, normalizeWhitespace(context.focusAreaLabel)) === 0 && computeOverlap(prompt, planAnchors.join(' ')) === 0) {
			score -= 2.2;
		}
	}

	if ((context.repositoryContext?.workspaceRoot || context.repositoryContext?.workspaceFolders?.length === 1)
		&& (taskLens?.primaryArtifact || context.repositoryContext?.primaryArtifactHint)
		&& /\b(repo|repository|workspace)\b/i.test(prompt)) {
		score -= 3.5;
	}

	if (hasUnresolvedPrimaryArtifact && context.questionStage === 'goal-clarity') {
		if (/\b(file|folder|directory|subsystem|symbol|module|entrypoint|script|csv|tsv|json|yaml|sql|notebook|test)\b/i.test(prompt)) {
			score += 2.4;
		}

		if (/\b(which|what)\b.{0,40}\b(file|folder|directory|subsystem|symbol|module|csv|script|test)\b/i.test(prompt)) {
			score += 1.2;
		}

		if (/\b(anchor|primary|exact|target)\b/i.test(prompt)) {
			score += 1.4;
		}

		if (/\brelated files?\b|\badjacent files?\b|\bother files?\b/i.test(prompt)) {
			score -= 1.6;
		}
	}

	if (taskLens?.artifactType === 'dataset' && /\b(language|framework|library|plotting|frontend)\b/i.test(prompt)) {
		score -= 2;
	}

	if (context.questionStage !== 'goal-clarity' && looksLikeGoalClarityQuestion(prompt)) {
		score -= 3;
	}

	if (/\bwhich repo\b|\bwhich workspace\b/i.test(prompt)) {
		score -= 4;
	}

	return score;
}

function formatSymbols(symbols: ReadonlyArray<{ name: string; kind: string; file?: string }>): string {
	if (!symbols.length) {
		return 'None';
	}

	return symbols
		.slice(0, 10)
		.map(symbol => `${symbol.name} (${symbol.kind}${symbol.file ? ` @ ${symbol.file}` : ''})`)
		.join(', ');
}

function formatSnippets(snippets: ReadonlyArray<{ path: string; preview: string; detailLevel?: string; reason?: string }>): string {
	if (!snippets.length) {
		return 'None';
	}

	return snippets
		.slice(0, 4)
		.map(snippet => [
			`FILE: ${snippet.path}`,
			snippet.detailLevel ? `DETAIL: ${snippet.detailLevel}` : '',
			snippet.reason ? `WHY: ${snippet.reason}` : '',
			snippet.preview,
		].filter(part => part.length > 0).join('\n'))
		.join('\n---\n');
}

function formatTaskLens(taskLens: IPlanningTaskLens | undefined): string {
	if (!taskLens) {
		return 'Not provided';
	}

	return [
		`Task kind: ${taskLens.taskKind}`,
		`Task summary: ${taskLens.taskSummary || 'Not provided'}`,
		`Primary artifact: ${taskLens.primaryArtifact || 'Not provided'}`,
		`Adjacent artifacts: ${taskLens.secondaryArtifacts?.join(', ') || 'None'}`,
		`Artifact type: ${taskLens.artifactType || 'Not provided'}`,
		`Desired outcome: ${taskLens.desiredOutcome || 'Not provided'}`,
		`Expected deliverable: ${taskLens.deliverableType || 'Not provided'}`,
		`Plan areas: ${taskLens.planAreas?.join(', ') || 'None'}`,
		`Validation targets: ${taskLens.validationTargets?.join(', ') || 'None'}`,
		`Risks or guardrails: ${taskLens.riskAreas?.join('; ') || 'None'}`,
		`Open decisions: ${taskLens.unknowns?.join('; ') || 'None'}`,
	].join('\n');
}

function formatPlanningTarget(target: IPlanningTarget | undefined): string {
	if (!target) {
		return 'Not provided';
	}

	return `${target.label} (${target.kind}${target.confidence ? `, ${target.confidence} confidence` : ''})`;
}
