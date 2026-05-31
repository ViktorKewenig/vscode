/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { CancellationToken } from '../../../../../../base/common/cancellation.js';
import { Event } from '../../../../../../base/common/event.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { generateDynamicPlanningPlanStepControlsResult, generateDynamicPlanningQuestions } from '../../../browser/planning/chatPlanningQuestionGenerator.js';
import { ChatMessageRole, IChatMessage, ILanguageModelsService } from '../../../common/languageModels.js';

suite('ChatPlanningQuestionGenerator', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('normalizes generated planning questions and includes coding context in the prompt', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{
										title: 'Primary implementation goal',
										message: 'What should be true when this coding task is complete?',
										type: 'text',
										required: true
									},
									{
										title: 'Execution strategy',
										description: 'Choose the best starting point for the work.',
										type: 'singleSelect',
										options: [
											{ label: 'Audit existing code path', value: 'Audit existing code path' },
											{ label: 'Implement directly', value: 'Implement directly' }
										],
										defaultValue: 'Audit existing code path'
									},
									{
										title: 'Validation signals',
										message: 'Which outcomes should this plan make explicit before coding starts?',
										type: 'multiSelect',
										options: [
											{ label: 'Changed files', value: 'Changed files' },
											{ label: 'Validation path', value: 'Validation path' }
										]
									}
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Add a planning scaffold before implementation starts',
			modelId: undefined,
			planningPhase: 'focused-slice',
			questionStage: 'goal-clarity',
			questionCount: 3,
			missingDimensions: ['constraints'],
			partialDimensions: ['repo-target'],
			activeFilePath: 'file:///workspace/src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
			selectedText: 'private async _acceptInput(...)',
			plannerNotes: 'Keep the changes inside the chat planning flow.',
			recentConversation: ['User: add richer planning context', 'Assistant: start with workspace symbols and snippets'],
			planningAnswers: [{ question: 'Implementation Goal', answer: 'Narrow toward the right insertion point.' }],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts', confidence: 'high' },
				requestIntent: 'script-work',
				taskLens: {
					taskKind: 'script-work',
					taskSummary: 'Refine the planning middleware around chatWidget.ts.',
					primaryArtifact: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
					secondaryArtifacts: ['src/vs/workbench/contrib/chat/browser/planning/chatPlanningQuestionGenerator.ts'],
					artifactType: 'file',
					desiredOutcome: 'Produce a tighter planning flow before the first plan runs.',
					deliverableType: 'code-change',
					riskAreas: ['Keep the changes inside the planning flow.'],
					unknowns: ['Validation path'],
					validationTargets: ['Planning tests'],
					planAreas: ['Question generation prompt', 'Planning context attachment'],
				},
				primaryArtifactHint: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
				relatedArtifactHints: ['src/vs/workbench/contrib/chat/browser/planning/chatPlanningQuestionGenerator.ts'],
				focusSummary: 'Focused slice around chat planning state.',
				focusQueries: ['planning', 'chatWidget'],
				workspaceFolders: ['vscode'],
				workspaceTopLevelEntries: ['src', 'extensions'],
				workingSetFiles: ['src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts'],
				activeDocumentSymbols: [{ name: '_acceptInput', kind: 'Method', file: 'file:///workspace/src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts' }],
				workspaceSymbolMatches: [{ name: 'refinePlan', kind: 'Method', file: 'file:///workspace/src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts' }],
				nearbyFiles: ['src/vs/workbench/contrib/chat/browser/actions/chatActions.ts'],
				relevantSnippets: [{
					path: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
					preview: 'private async _acceptInput(...)',
					detailLevel: 'focused',
					reason: 'Current implementation entry point.'
				}]
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 3);
		const goalQuestion = questions.find(question => question.message === 'What should be true when this coding task is complete?');
		const strategyQuestion = questions.find(question => question.message === 'Choose the best starting point for the work.');
		const validationQuestion = questions.find(question => question.title === 'Validation signals');
		assert.ok(goalQuestion);
		assert.ok(strategyQuestion);
		assert.ok(validationQuestion);
		assert.strictEqual(goalQuestion.type, 'text');
		assert.strictEqual(strategyQuestion.type, 'singleSelect');
		assert.deepStrictEqual(strategyQuestion.options, [
			{ id: 'Audit existing code path', label: 'Audit existing code path', value: 'Audit existing code path' },
			{ id: 'Implement directly', label: 'Implement directly', value: 'Implement directly' }
		]);
		assert.strictEqual(strategyQuestion.defaultValue, 'Audit existing code path');
		assert.strictEqual(validationQuestion.type, 'multiSelect');

		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		assert.strictEqual(capturedMessages[0].role, ChatMessageRole.System);
		const promptPart = capturedMessages[1].content[0];
		assert.strictEqual(promptPart.type, 'text');
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('User request:\nAdd a planning scaffold before implementation starts'));
		assert.ok(promptPart.value.includes('Planning phase:\nfocused-slice (Focused Slice)'));
		assert.ok(promptPart.value.includes('Question stage:\ngoal-clarity'));
		assert.ok(promptPart.value.includes('Requested question count:\n3'));
		assert.ok(promptPart.value.includes('Active file:\nfile:///workspace/src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts'));
		assert.ok(promptPart.value.includes('Selected code or text:\nprivate async _acceptInput(...)'));
		assert.ok(promptPart.value.includes('Planner notes:\nKeep the changes inside the chat planning flow.'));
		assert.ok(promptPart.value.includes('Recent planning conversation:\n- User: add richer planning context'));
		assert.ok(promptPart.value.includes('Repository context:'));
		assert.ok(promptPart.value.includes('Planning target: src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts (file, high confidence)'));
		assert.ok(promptPart.value.includes('Request intent: script-work'));
		assert.ok(promptPart.value.includes('Task lens:'));
		assert.ok(promptPart.value.includes('Desired outcome: Produce a tighter planning flow before the first plan runs.'));
		assert.ok(promptPart.value.includes('Primary artifact hint: src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts'));
		assert.ok(promptPart.value.includes('Missing planning dimensions:\nconstraints'));
	});

	test('uses task decomposition to generate inline questions for an existing plan', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{
										title: 'Plan Step Priority',
										message: 'Which current plan steps should be handled first?',
										type: 'multiSelect',
										options: [
											{ label: 'Update chatWidget transition', value: 'Update chatWidget transition' },
											{ label: 'Embed step questions', value: 'Embed step questions' },
											{ label: 'Run planning tests', value: 'Run planning tests' }
										]
									},
									{
										title: 'Validation Approach',
										message: 'What validation should this plan require after the UI changes?',
										type: 'text',
										required: true
									}
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a better planning middleware handoff.',
			modelId: undefined,
			planningPhase: 'focused-slice',
			questionStage: 'task-decomposition',
			questionCount: 2,
			currentPlan: [
				'## Plan',
				'1. Update chatWidget transition so goal clarity answers stream the first plan immediately.',
				'2. Embed generated refinement questions into the plan editor steps.',
				'3. Run ChatPlanning tests and smoke launch Code OSS.'
			].join('\n'),
			recentConversation: [],
			planningAnswers: [{ question: 'Clarifying Your Goals', answer: 'Make the plan easier to refine before implementation.' }],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts', confidence: 'high' },
				requestIntent: 'feature-work',
				taskLens: {
					taskKind: 'feature-work',
					taskSummary: 'Improve planning middleware handoff.',
					primaryArtifact: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
					artifactType: 'file',
					desiredOutcome: 'Shape the first plan with user input before implementation starts.',
					deliverableType: 'code-change',
					planAreas: ['Question flow', 'Plan handoff'],
					validationTargets: ['Planning tests'],
				},
				primaryArtifactHint: 'src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts',
				focusQueries: ['planning', 'chatWidget'],
				workspaceFolders: ['workspace'],
				workingSetFiles: ['src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: [],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 2);
		assert.strictEqual(questions[0].type, 'multiSelect');
		assert.strictEqual(questions[0].title, 'Plan Step Priority');
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const systemPart = capturedMessages[0].content[0];
		if (systemPart.type !== 'text') {
			throw new Error('Expected system prompt text part');
		}
		assert.ok(systemPart.value.includes('titles under five words'));
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Current plan:\n## Plan'));
		assert.ok(promptPart.value.includes('Current plan anchors:'));
		assert.ok(promptPart.value.includes('tighten the first plan into a stronger work breakdown'));
		assert.ok(!promptPart.value.includes('co-create the high-level plan before the first detailed plan is built'));
		assert.ok(!promptPart.value.includes('Include one assumption-confirmation question'));
	});

	test('generates distinct task controls keyed to each plan step', async () => {
		let capturedModelId: string | undefined;
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['copilot-fast', 'plan-model'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => ({ id: modelId, family: modelId, capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async (selector: { readonly id?: string; readonly family?: string }) => selector.id === 'copilot-fast' || selector.family === 'copilot-fast' ? ['copilot-fast'] : [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedModelId = modelId;
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								stepControls: [
									{
										stepIndex: 1,
										questions: [
											{
												title: 'Loader Boundary',
												message: 'How should this step limit the loader change?',
												type: 'singleSelect',
												allowFreeformInput: true,
												defaultValue: 'Parameter only',
												options: [
													{ label: 'Parameter only', value: 'parameter-only' },
													{ label: 'Trace first', value: 'trace-first' }
												]
											},
											{
												title: 'Loader Evidence',
												message: 'Which evidence should confirm the loader path?',
												type: 'singleSelect',
												options: [
													{ label: 'Call trace', value: 'call-trace' },
													{ label: 'Path print', value: 'path-print' }
												]
											}
										]
									},
									{
										stepIndex: 2,
										questions: [
											{
												title: 'Schema Checks',
												message: 'Which validation should happen before modeling?',
												type: 'singleSelect',
												options: [
													{ label: 'Required columns', value: 'required-columns' },
													{ label: 'Types and columns', value: 'types-and-columns' }
												]
											},
											{
												title: 'NA Reporting',
												message: 'How should missing values be reported?',
												type: 'multiSelect',
												options: [
													{ label: 'By column', value: 'by-column' },
													{ label: 'By condition', value: 'by-condition' }
												]
											}
										]
									},
									{
										stepIndex: 3,
										questions: [
											{
												title: 'Validation Signal',
												message: 'Which run should prove this step is done?',
												type: 'singleSelect',
												options: [
													{ label: 'Unit test', value: 'unit-test' },
													{ label: 'Script run', value: 'script-run' }
												]
											},
											{
												title: 'Output Check',
												message: 'Which output should be inspected first?',
												type: 'singleSelect',
												options: [
													{ label: 'CSV columns', value: 'csv-columns' },
													{ label: 'Console verdict', value: 'console-verdict' }
												]
											}
										]
									}
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const result = await generateDynamicPlanningPlanStepControlsResult(service, {
			userRequest: 'Tighten this CSV analysis plan.',
			modelId: undefined,
			planningPhase: 'focused-slice',
			questionStage: 'task-decomposition',
			currentPlan: [
				'Plan: CSV analysis',
				'**Steps**',
				'Confirm the input loader path.',
				'Validate required CSV columns.',
				'**Verification**',
				'Run the script and inspect output.'
			].join('\n'),
			recentConversation: [],
			planningAnswers: [{ question: 'Missing values', answer: 'Exclude rows with NA.' }],
		}, [
			{ index: 1, label: 'Confirm the input loader path.', text: 'Confirm the input loader path.', sectionTitle: 'Steps', kind: 'step' },
			{ index: 2, label: 'Validate required CSV columns.', text: 'Validate required CSV columns.', sectionTitle: 'Steps', kind: 'step' },
			{ index: 3, label: 'Run the script and inspect output.', text: 'Run the script and inspect output.', sectionTitle: 'Verification', kind: 'verification' },
		], CancellationToken.None);

		assert.strictEqual(capturedModelId, 'copilot-fast');
		assert.strictEqual(result.questionsByStep.size, 3);
		assert.deepStrictEqual(result.questionsByStep.get(1)?.map(question => question.title), ['Loader Boundary', 'Loader Evidence']);
		assert.deepStrictEqual(result.questionsByStep.get(2)?.map(question => question.title), ['Schema Checks', 'NA Reporting']);
		assert.deepStrictEqual(result.questionsByStep.get(3)?.map(question => question.title), ['Validation Signal', 'Output Check']);
		for (const questions of result.questionsByStep.values()) {
			assert.strictEqual(questions.length, 2);
			assert.ok(questions.every(question => question.type !== 'text'));
			assert.ok(questions.every(question => question.allowFreeformInput === false));
			assert.ok(questions.every(question => question.defaultValue === undefined));
		}
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Plan steps that need controls:'));
		assert.ok(promptPart.value.includes('stepIndex: 1'));
		assert.ok(promptPart.value.includes('stepIndex: 3'));
		assert.ok(promptPart.value.includes('Return at least two and at most four controls for every listed stepIndex.'));
		assert.ok(promptPart.value.includes('Use only singleSelect or multiSelect controls. Do not return text controls.'));
	});

	test('ranks task-lens-aligned questions ahead of generic repo questions', async () => {
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async () => ({
				stream: (async function* () {
					yield {
						type: 'text' as const,
						value: JSON.stringify({
							questions: [
								{ title: 'Repository choice', message: 'Which repository should I use for this task?', type: 'text' },
								{ title: 'Primary analysis target', message: 'Which file or directory around orders.csv should anchor the analysis?', type: 'text' },
								{
									title: 'Validation path',
									message: 'How should the plan validate the analysis against schema.json?',
									type: 'singleSelect',
									options: [
										{ label: 'Schema checks first', value: 'Schema checks first' },
										{ label: 'Manual review first', value: 'Manual review first' }
									]
								},
								{
									title: 'Output shape',
									message: 'What output should the plan optimize for?',
									type: 'singleSelect',
									options: [
										{ label: 'Plain-text summary', value: 'Plain-text summary' },
										{ label: 'Structured report', value: 'Structured report' }
									]
								}
							]
						})
					};
				})(),
				result: Promise.resolve({})
			}),
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan how to analyze orders.csv with schema.json in this repo.',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			questionCount: 3,
			recentConversation: [],
			planningAnswers: [],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'data/orders.csv', confidence: 'high' },
				requestIntent: 'data-analysis',
				taskLens: {
					taskKind: 'data-analysis',
					taskSummary: 'Analyze orders.csv against schema.json.',
					primaryArtifact: 'data/orders.csv',
					secondaryArtifacts: ['data/schema.json'],
					artifactType: 'dataset',
					desiredOutcome: 'Produce a useful analysis for data/orders.csv.',
					deliverableType: 'analysis',
					validationTargets: ['Schema checks', 'Analysis output'],
					unknowns: ['Desired output'],
					planAreas: ['Analysis target', 'Validation path'],
				},
				primaryArtifactHint: 'data/orders.csv',
				relatedArtifactHints: ['data/schema.json'],
				focusQueries: ['orders.csv', 'schema.json'],
				workspaceFolders: ['workspace'],
				workingSetFiles: ['data/orders.csv', 'data/schema.json'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: ['data/schema.json'],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 3);
		assert.ok(questions.every(question => !/repository/i.test(question.title)));
		assert.ok(questions.some(question => question.id === 'dynamic-planning-analysis-kind' && question.type === 'multiSelect'));
		assert.ok(questions.some(question => question.id === 'dynamic-planning-analysis-goal' && question.type === 'text'));
		assert.ok(questions.some(question => question.id === 'dynamic-planning-analysis-audience' && question.type === 'singleSelect'));
	});

	test('does not inject a static planning-target confirmation question', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Desired outcome', message: 'What should the user notice first?', type: 'text' },
									{
										title: 'Scope boundary',
										message: 'Which boundary matters most?',
										type: 'singleSelect',
										options: [
											{ label: 'Keep changes in chat planning', value: 'Keep changes in chat planning' },
											{ label: 'Allow follow-up UI work', value: 'Allow follow-up UI work' }
										]
									},
									{ title: 'Definition of done', message: 'What must be settled before the planner runs?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Refine planning mode around the right repo target.',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			questionCount: 3,
			shouldConfirmPlanningTarget: true,
			recentConversation: [],
			planningAnswers: [],
			repositoryContext: {
				scope: 'broad',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'workspace', label: 'vscode', confidence: 'low' },
				focusQueries: ['planning'],
				workspaceFolders: ['vscode'],
				workspaceTopLevelEntries: ['src', 'extensions'],
				workingSetFiles: ['src/vs/workbench/contrib/chat/browser/widget/chatWidget.ts'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: [],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 3);
		assert.ok(questions.every(question => question.id !== 'planning-target-confirmation'));
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Requested question count:\n3'));
	});

	test('throws when no language model is available', async () => {
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => [],
			getVendors: () => [],
			lookupLanguageModel: () => undefined,
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async () => {
				throw new Error('should not be called');
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await assert.rejects(() => generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'task-decomposition',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None), /No language model is available/);
	});

	test('throws when the model returns invalid JSON', async () => {
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async () => ({
				stream: (async function* () {
					yield { type: 'text' as const, value: 'not json' };
				})(),
				result: Promise.resolve({})
			}),
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await assert.rejects(() => generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: undefined,
			planningPhase: 'detailed-inspection',
			questionStage: 'task-decomposition',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None), /Planning question generation failed|did not return enough usable planning questions/);
	});

	test('throws when the model returns too few usable questions', async () => {
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async () => ({
				stream: (async function* () {
					yield {
						type: 'text' as const,
						value: JSON.stringify({
							questions: [
								{
									title: 'Approved implementation target',
									message: 'What must this plan make concrete before coding starts?',
									type: 'text',
									required: true
								},
								{
									title: 'Primary work split',
									message: 'Which work split best matches the request?',
									type: 'singleSelect',
									options: [
										{ label: 'Touch one subsystem first', value: 'Touch one subsystem first' },
										{ label: 'Stage the work across layers', value: 'Stage the work across layers' }
									]
								}
							]
						})
					};
				})(),
				result: Promise.resolve({})
			}),
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await assert.rejects(() => generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a scoped implementation change',
			modelId: undefined,
			planningPhase: 'focused-slice',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None), /did not return enough usable planning questions/);
	});

	test('uses the current plan and focus hint to request multiple plan-focus questions', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{
										title: 'Focused Repo Slice',
										message: 'Which file, folder, or subsystem should this zoom-in stay anchored on?',
										type: 'singleSelect',
										options: [
											{ label: 'Current working set', value: 'Current working set' },
											{ label: 'One specific file', value: 'One specific file' }
										],
										allowFreeformInput: true
									},
									{
										title: 'Unresolved Decision',
										message: 'What decision in this plan area still needs sharper guidance?',
										type: 'text',
										required: true
									},
									{
										title: 'Evidence To Add',
										message: 'Which evidence should the focused rebuild make explicit?',
										type: 'singleSelect',
										options: [
											{ label: 'Execution order', value: 'Execution order' },
											{ label: 'Validation path', value: 'Validation path' }
										]
									}
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Refine the planning middleware around CSV analysis.',
			modelId: undefined,
			planningPhase: 'detailed-inspection',
			questionStage: 'plan-focus',
			questionCount: 3,
			recentConversation: ['Assistant: The revised plan already covers the main flow.'],
			planningAnswers: [{ question: 'Related files', answer: 'Pay attention to orders.csv and schema.json.' }],
			currentPlan: [
				'1. Inspect orders.csv and nearby parsing utilities.',
				'2. Identify where schema.json informs validation.',
				'3. Rebuild the plan around the data flow and validation path.'
			].join('\n'),
			focusHint: 'Focus especially on the validation path around orders.csv and schema.json.',
			repositoryContext: {
				scope: 'detailed',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'data/orders.csv', confidence: 'high' },
				focusSummary: 'Plan focus around orders.csv validation.',
				focusQueries: ['orders.csv', 'schema.json', 'validation path'],
				workspaceFolders: ['workspace'],
				workspaceTopLevelEntries: ['data', 'src'],
				workingSetFiles: ['data/orders.csv', 'data/schema.json'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: ['data/schema.json'],
				relevantSnippets: []
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 3);
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Question stage:\nplan-focus'));
		assert.ok(promptPart.value.includes('Requested question count:\n3'));
		assert.ok(promptPart.value.includes('Current plan:\n1. Inspect orders.csv and nearby parsing utilities.'));
		assert.ok(promptPart.value.includes('Internal focus guidance (do not quote verbatim):\nFocus especially on the validation path around orders.csv and schema.json.'));
	});

	test('treats requested artifact hints as unresolved and prefers an artifact-targeting first question', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Repository choice', message: 'Which repository should I use for this analysis?', type: 'text' },
									{ title: 'CSV target', message: 'Which CSV file or directory should anchor the analysis?', type: 'text' },
									{
										title: 'Related inputs',
										message: 'Which related files should stay in view while planning the analysis?',
										type: 'singleSelect',
										options: [
											{ label: 'schema.json', value: 'schema.json' },
											{ label: 'No related file', value: 'No related file' }
										]
									}
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan how to analyze a csv file in this repo.',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			questionCount: 2,
			recentConversation: [],
			planningAnswers: [],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'workspace', label: 'workspace', confidence: 'low' },
				requestIntent: 'data-analysis',
				taskLens: {
					taskKind: 'data-analysis',
					artifactType: 'dataset',
					desiredOutcome: 'Plan how to analyze a csv file in this repo.',
					unknowns: ['Exact file, folder, or subsystem'],
				},
				primaryArtifactHint: 'Requested CSV file',
				relatedArtifactHints: ['data/schema.json'],
				focusQueries: ['csv', 'schema.json'],
				workspaceFolders: ['workspace'],
				workspaceTopLevelEntries: ['data', 'src'],
				workingSetFiles: ['data/schema.json'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: ['data/schema.json'],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 2);
		assert.strictEqual(questions[0].message, 'Which CSV file or directory should anchor the analysis?');
		assert.strictEqual(questions[1].id, 'dynamic-planning-analysis-kind');
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Artifact targeting:\nThe request points at Requested CSV file'));
	});

	test('uses follow-up goal clarity to surface editable assumptions without repeating answered analysis questions', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Analysis Type', message: 'What kind of analysis are you looking to do?', type: 'multiSelect', options: [{ label: 'Summary Statistics', value: 'Summary Statistics' }, { label: 'Compare Groups', value: 'Compare Groups' }] },
									{ title: 'Analysis Goal', message: 'What do you want to learn, decide, or communicate with this analysis?', type: 'text' },
									{ title: 'Data Caveats', message: 'What context about the data should shape the analysis?', type: 'text' },
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'want to analyse a csv file',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			questionCount: 2,
			missingDimensions: ['constraints'],
			partialDimensions: ['scope-boundaries'],
			focusHint: 'Review editable assumptions before drafting the first plan.',
			recentConversation: [],
			planningAnswers: [
				{ question: 'What kind of analysis are you looking to do?', answer: 'Summary statistics, Group comparisons' },
				{ question: 'What do you want to learn, decide, or communicate with this analysis?', answer: 'visualisation, result summary' },
				{ question: 'Who is this analysis for?', answer: 'Technical Collaborators' },
			],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'data/orders.csv', confidence: 'high' },
				requestIntent: 'data-analysis',
				taskLens: {
					taskKind: 'data-analysis',
					primaryArtifact: 'data/orders.csv',
					artifactType: 'dataset',
					desiredOutcome: 'visualisation, result summary',
					deliverableType: 'analysis',
					validationTargets: ['Analysis output'],
				},
				primaryArtifactHint: 'data/orders.csv',
				focusQueries: ['orders.csv'],
				workspaceFolders: ['workspace'],
				workingSetFiles: ['data/orders.csv'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: [],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 2);
		assert.strictEqual(questions[0].id, 'dynamic-planning-plan-assumptions');
		assert.strictEqual(questions[0].title, 'Working Assumptions');
		assert.ok(questions[0].description?.includes('replace goal-clarification questions'));
		assert.ok(!questions.some(question => question.id === 'dynamic-planning-analysis-kind'));
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const promptPart = capturedMessages[0].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('include an editable assumptions review'));
	});

	test('filters goal clarity and assumption questions out of task decomposition', async () => {
		let capturedMessages: IChatMessage[] | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['plan-model'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (_modelId: string, _from: unknown, messages: IChatMessage[]) => {
				capturedMessages = messages;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Goal', message: 'What outcome should this achieve?', type: 'text' },
									{ title: 'Scope', message: 'What is in or out of scope?', type: 'text' },
									{ title: 'Success Criteria', message: 'What should success look like?', type: 'text' },
									{ title: 'Working Assumptions', message: 'Which assumptions should shape the first plan?', type: 'multiSelect', options: [{ label: 'Keep scope narrow', value: 'Keep scope narrow' }, { label: 'Include validation', value: 'Include validation' }] },
									{ title: 'Work Areas', message: 'Which work areas should shape the first plan?', type: 'multiSelect', options: [{ label: 'Data loading', value: 'Data loading' }, { label: 'Visualization', value: 'Visualization' }] },
									{ title: 'Validation Path', message: 'Which validation path should the plan include?', type: 'singleSelect', options: [{ label: 'Smoke run', value: 'Smoke run' }, { label: 'Unit tests', value: 'Unit tests' }] },
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		const questions = await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan how to analyze orders.csv.',
			modelId: undefined,
			planningPhase: 'focused-slice',
			questionStage: 'task-decomposition',
			questionCount: 2,
			recentConversation: [],
			planningAnswers: [
				{ question: 'Analysis Goal', answer: 'visualisation and summary' },
				{ question: 'Working Assumptions', answer: 'Keep scope narrow' },
			],
			repositoryContext: {
				scope: 'focused',
				workspaceRoot: 'file:///workspace',
				planningTarget: { kind: 'file', label: 'data/orders.csv', confidence: 'high' },
				requestIntent: 'data-analysis',
				taskLens: {
					taskKind: 'data-analysis',
					primaryArtifact: 'data/orders.csv',
					artifactType: 'dataset',
					planAreas: ['Data loading', 'Visualization'],
					validationTargets: ['Smoke run'],
				},
				primaryArtifactHint: 'data/orders.csv',
				focusQueries: ['orders.csv'],
				workspaceFolders: ['workspace'],
				workingSetFiles: ['data/orders.csv'],
				activeDocumentSymbols: [],
				workspaceSymbolMatches: [],
				nearbyFiles: [],
				relevantSnippets: [],
			}
		}, CancellationToken.None);

		assert.strictEqual(questions.length, 2);
		assert.deepStrictEqual(questions.map(question => question.title), ['Work Areas', 'Validation Path']);
		assert.ok(!questions.some(question => /goal|scope|success|assumption/i.test(question.title)));
		assert.ok(capturedMessages, 'Expected a language-model request to be issued');
		const systemPart = capturedMessages[0].content[0];
		if (systemPart.type !== 'text') {
			throw new Error('Expected system prompt text part');
		}
		assert.ok(!systemPart.value.includes('success criteria'));
		const promptPart = capturedMessages[1].content[0];
		if (promptPart.type !== 'text') {
			throw new Error('Expected prompt text part');
		}
		assert.ok(promptPart.value.includes('Do not ask goal, scope, definition-of-done, or assumption-review questions here.'));
	});

	test('skips session-targeted models when choosing a fallback model', async () => {
		let capturedModelId: string | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['session-model', 'plan-model'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => modelId === 'session-model'
				? { capabilities: { toolCalling: true }, targetChatSessionType: 'agent-host' }
				: { capabilities: { toolCalling: true } },
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string) => {
				capturedModelId = modelId;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Goal', message: 'What should happen?', type: 'text' },
									{
										title: 'Constraint',
										message: 'What should the planner optimize for?',
										type: 'singleSelect',
										options: [
											{ label: 'Minimal surface area', value: 'Minimal surface area' },
											{ label: 'Fastest path', value: 'Fastest path' }
										]
									},
									{ title: 'Definition of done', message: 'What must be clarified?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: undefined,
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None);

		assert.strictEqual(capturedModelId, 'plan-model');
	});

	test('prefers provider-backed models over stale cached model ids', async () => {
		let capturedModelId: string | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['copilot/auto', 'plan-model'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => {
				if (modelId === 'copilot/auto') {
					return { vendor: 'copilot', id: 'auto', capabilities: { toolCalling: true } };
				}
				return { vendor: 'test-vendor', id: 'plan-model', capabilities: { toolCalling: true } };
			},
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async (selector: { vendor?: string }) => selector.vendor === 'copilot' ? [] : ['plan-model'],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string) => {
				capturedModelId = modelId;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Goal', message: 'What should happen?', type: 'text' },
									{
										title: 'Constraint',
										message: 'What should the planner optimize for?',
										type: 'singleSelect',
										options: [
											{ label: 'Minimal surface area', value: 'Minimal surface area' },
											{ label: 'Fastest path', value: 'Fastest path' }
										]
									},
									{ title: 'Definition of done', message: 'What must be clarified?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: 'copilot/auto',
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None);

		assert.strictEqual(capturedModelId, 'plan-model');
	});

	test('defaults planning question generation to GPT-4.1 when the picker is Auto', async () => {
		let capturedModelId: string | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['copilot/auto', 'gpt-4.1', 'plan-model'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => {
				if (modelId === 'copilot/auto') {
					return { vendor: 'copilot', id: 'auto', family: 'auto', capabilities: { toolCalling: true } };
				}
				if (modelId === 'gpt-4.1') {
					return { vendor: 'copilot', id: 'gpt-4.1', family: 'gpt-4.1', capabilities: { toolCalling: true } };
				}
				return { vendor: 'test-vendor', id: 'plan-model', family: 'plan-model', capabilities: { toolCalling: true } };
			},
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async (selector: { vendor?: string; id?: string; family?: string }) => {
				if (selector.id === 'gpt-4.1' || selector.family === 'gpt-4.1') {
					return ['gpt-4.1'];
				}
				if (selector.vendor === 'copilot') {
					return ['gpt-4.1'];
				}
				return ['gpt-4.1', 'plan-model'];
			},
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string) => {
				capturedModelId = modelId;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Goal', message: 'What should happen?', type: 'text' },
									{
										title: 'Constraint',
										message: 'What should the planner optimize for?',
										type: 'singleSelect',
										options: [
											{ label: 'Minimal surface area', value: 'Minimal surface area' },
											{ label: 'Fastest path', value: 'Fastest path' }
										]
									},
									{ title: 'Definition of done', message: 'What must be clarified?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: 'copilot/auto',
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None);

		assert.strictEqual(capturedModelId, 'gpt-4.1');
	});

	test('defaults task-decomposition control generation to the fast Copilot model', async () => {
		let capturedModelId: string | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['copilot/auto', 'copilot-fast', 'gpt-4.1', 'plan-model'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => {
				if (modelId === 'copilot/auto') {
					return { vendor: 'copilot', id: 'auto', family: 'auto', capabilities: { toolCalling: true } };
				}
				if (modelId === 'copilot-fast') {
					return { vendor: 'copilot', id: 'copilot-fast', family: 'copilot-fast', capabilities: { toolCalling: true } };
				}
				if (modelId === 'gpt-4.1') {
					return { vendor: 'copilot', id: 'gpt-4.1', family: 'gpt-4.1', capabilities: { toolCalling: true } };
				}
				return { vendor: 'test-vendor', id: 'plan-model', family: 'plan-model', capabilities: { toolCalling: true } };
			},
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async (selector: { vendor?: string; id?: string; family?: string }) => {
				if (selector.id === 'copilot-fast' || selector.family === 'copilot-fast') {
					return ['copilot-fast'];
				}
				if (selector.id === 'gpt-4.1' || selector.family === 'gpt-4.1') {
					return ['gpt-4.1'];
				}
				if (selector.vendor === 'copilot') {
					return ['copilot-fast', 'gpt-4.1'];
				}
				return ['copilot-fast', 'gpt-4.1', 'plan-model'];
			},
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string) => {
				capturedModelId = modelId;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{
										title: 'Step order',
										message: 'Which step should move first?',
										type: 'singleSelect',
										options: [
											{ label: 'Tests first', value: 'Tests first' },
											{ label: 'Code first', value: 'Code first' }
										]
									},
									{ title: 'Missing check', message: 'What check is missing?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: 'copilot/auto',
			planningPhase: 'focused-slice',
			questionStage: 'task-decomposition',
			questionCount: 2,
			currentPlan: '1. Add tests.\n2. Update code.',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None);

		assert.strictEqual(capturedModelId, 'copilot-fast');
	});

	test('defaults planning question generation to GPT-4.1 when the current concrete model is stale', async () => {
		let capturedModelId: string | undefined;
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['stale-model-id', 'gpt-4.1'],
			getVendors: () => [],
			lookupLanguageModel: (modelId: string) => {
				if (modelId === 'gpt-4.1') {
					return { vendor: 'copilot', id: 'gpt-4.1', family: 'gpt-4.1', capabilities: { toolCalling: true } };
				}
				return undefined;
			},
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async (selector: { vendor?: string; id?: string; family?: string }) => {
				if (selector.id === 'gpt-4.1' || selector.family === 'gpt-4.1') {
					return ['gpt-4.1'];
				}
				if (selector.vendor === 'copilot') {
					return ['gpt-4.1'];
				}
				return ['gpt-4.1'];
			},
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async (modelId: string) => {
				capturedModelId = modelId;
				return {
					stream: (async function* () {
						yield {
							type: 'text' as const,
							value: JSON.stringify({
								questions: [
									{ title: 'Goal', message: 'What should happen?', type: 'text' },
									{
										title: 'Constraint',
										message: 'What should the planner optimize for?',
										type: 'singleSelect',
										options: [
											{ label: 'Minimal surface area', value: 'Minimal surface area' },
											{ label: 'Fastest path', value: 'Fastest path' }
										]
									},
									{ title: 'Definition of done', message: 'What must be clarified?', type: 'text' }
								]
							})
						};
					})(),
					result: Promise.resolve({})
				};
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: 'stale-model-id',
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None);

		assert.strictEqual(capturedModelId, 'gpt-4.1');
	});

	test('surfaces a friendly error when no registered model provider is ready', async () => {
		const service = {
			_serviceBrand: undefined,
			onDidChangeLanguageModelVendors: Event.None,
			onDidChangeLanguageModels: Event.None,
			updateModelPickerPreference: () => { },
			getLanguageModelIds: () => ['copilot/auto'],
			getVendors: () => [],
			lookupLanguageModel: () => ({ vendor: 'copilot', id: 'auto', capabilities: { toolCalling: true } }),
			lookupLanguageModelByQualifiedName: () => undefined,
			getLanguageModelGroups: () => [],
			selectLanguageModels: async () => [],
			registerLanguageModelProvider: () => ({ dispose: () => { } }),
			deltaLanguageModelChatProviderDescriptors: () => { },
			sendChatRequest: async () => {
				throw new Error('Chat provider for model copilot/auto is not registered.');
			},
			computeTokenLength: async () => 0,
			getModelConfiguration: () => undefined,
			setModelConfiguration: async () => { },
			getModelConfigurationActions: () => [],
			addLanguageModelsProviderGroup: async () => { },
			removeLanguageModelsProviderGroup: async () => { },
			configureLanguageModelsProviderGroup: async () => { },
		} as unknown as ILanguageModelsService;

		await assert.rejects(() => generateDynamicPlanningQuestions(service, {
			userRequest: 'Plan a change',
			modelId: 'copilot/auto',
			planningPhase: 'broad-scan',
			questionStage: 'goal-clarity',
			recentConversation: [],
			planningAnswers: [],
		}, CancellationToken.None), /No language model is available to generate planning questions/);
	});
});
