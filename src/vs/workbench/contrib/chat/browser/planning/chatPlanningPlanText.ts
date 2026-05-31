/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IResponse } from '../../common/model/chatModel.js';

export interface IPlanningPlanChangeSummary {
	readonly added: readonly string[];
	readonly removed: readonly string[];
}

export type PlanningPlanStepKind = 'step' | 'verification' | 'decision' | 'guardrail' | 'other';

export interface IPlanningPlanStep {
	readonly index: number;
	readonly label: string;
	readonly text: string;
	readonly sectionTitle?: string;
	readonly kind: PlanningPlanStepKind;
}

const planHeadingPattern = /^(?:#{1,6}\s*plan\b.*|(?:\*\*)?plan\s*:\s*.+)$/gim;
const planSectionMarkerPattern = /^(?:\*\*(?:steps|relevant files|verification|decisions)\*\*|#{1,6}\s*(?:steps|verification|decisions)\b)/gim;
const planRenderStartPatterns = [
	/^(?:#{1,6}\s*)?(?:\*\*)?plan\b.*$/gim,
	/^(?:\*\*)?(?:steps|relevant files|verification|decisions)(?:\*\*)?\s*:?\s*$/gim,
];
const minimumPlanCandidateScore = 3;
const planningScaffoldingBlockHeadingPattern = /^(?:planning answers:|recent planning conversation:)\s*$/i;
const planningScaffoldingLinePattern = /^(?:planning context from the previous planning step:|planning phase:|planner notes:|planning handoff instruction:|use this planning context as the source of truth\b|do not ignore the planning answers\b|do not re-ask questions\b|do not include new goal-clarification questions\b|do not include unanswered goal-clarification questions\b|use the editable assumption answers\b|do not number phase or section headings\b|if you ask follow-up questions\b|treat a confirmed planning target\b|keep upfront verification concise\b|do not describe an existing workflow\b|the user already answered\b|the user submitted edits\b|the user selected a plan area\b|the previous planner response completed\b|this retry must end\b|this retry must output\b|do not use tools on this retry\b|do not announce that you will\b|create the first concrete plan now\b|revise the current plan now\b|ask a follow-up question only\b|do not ask more goal-clarity questions\b|while working, stream concise visible planning notes\b|visible planning notes must be followed\b|if you need to inspect files\b|if a tool needs approval\b|do not stop after a progress update\b|do not write a standalone future-tense status sentence\b|a response that only says\b|if no tool result is available\b|return markdown that starts\b|user request:|question stage:|requested question count:|active file:|selected text:|missing dimensions:|partial dimensions:|should confirm planning target:|repository context:|scope:|workspace root:|planning target:|request intent:|task lens:|primary artifact hint:|related artifact hints:|focus summary:|focus queries:|workspace folders:|workspace top-level entries:|working set files:|active document symbols:|workspace symbol matches:|nearby files:|relevant snippets:|task kind:|task summary:|primary artifact:|adjacent artifacts:|artifact type:|desired outcome:|expected deliverable:|plan areas:|validation targets:|risks or guardrails:|open decisions:|current plan excerpt:|selected plan slice:|internal focus guidance\b)/i;

export function extractPlanningPlanText(response: IResponse | undefined): string | undefined {
	if (!response) {
		return undefined;
	}

	const candidates = new Set<string>();
	const markdown = normalizePlanText(response.getMarkdown());
	if (markdown) {
		candidates.add(markdown);
	}

	for (const part of response.value) {
		const candidate = getMarkdownCandidate(part);
		if (candidate) {
			candidates.add(candidate);
		}
	}

	let bestCandidate: string | undefined;
	let bestScore = Number.NEGATIVE_INFINITY;
	for (const candidate of candidates) {
		for (const extracted of getPlanCandidateVariants(candidate)) {
			const score = scorePlanCandidate(extracted);
			if (score > bestScore) {
				bestScore = score;
				bestCandidate = extracted;
			}
		}
	}

	if (!bestCandidate || bestScore < minimumPlanCandidateScore || !isUsablePlanningPlanText(bestCandidate)) {
		return undefined;
	}

	return normalizePlanText(bestCandidate);
}

export function isUsablePlanningPlanText(planText: string | undefined): boolean {
	const normalized = normalizePlanText(planText);
	if (!normalized) {
		return false;
	}

	const lines = normalized.split(/\r?\n/g).map(line => line.trim()).filter(line => line.length > 0);
	if (lines.length === 0) {
		return false;
	}

	const hasPlanHeading = /^(?:#{1,6}\s*plan\b.*|(?:\*\*)?plan\s*:\s*.+)$/im.test(normalized);
	const hasPlanSection = planSectionMarkerPattern.test(normalized);
	planSectionMarkerPattern.lastIndex = 0;
	const topLevelListItems = normalized.match(/^\s{0,2}(?:[-*+]|\d+[.)])\s+/gm)?.length ?? 0;
	if (hasPlanHeading || hasPlanSection || topLevelListItems >= 2) {
		return true;
	}

	const extractedSteps = extractPlanningPlanSteps(normalized, 4);
	return extractedSteps.length >= 2;
}

export function findPlanningPlanStartOffset(value: string | undefined): number | undefined {
	if (!value) {
		return undefined;
	}

	for (const pattern of planRenderStartPatterns) {
		const matcher = new RegExp(pattern.source, pattern.flags);
		let match = matcher.exec(value);
		while (match) {
			if (isUsablePlanningPlanText(value.slice(match.index))) {
				return match.index;
			}
			match = matcher.exec(value);
		}
	}

	return undefined;
}

export function normalizePlanningPlanLine(line: string): string | undefined {
	const normalized = line
		.replace(/^\s{0,3}(?:(?:[-*+])\s+|\d+[.)]\s+|#{1,6}\s+)/, '')
		.replace(/\s+/g, ' ')
		.trim();
	return normalized.length >= 12 ? normalized : undefined;
}

export function extractPlanningPlanSteps(planText: string | undefined, maxSteps = 10): readonly IPlanningPlanStep[] {
	const normalized = normalizePlanText(planText);
	if (!normalized || maxSteps <= 0) {
		return [];
	}

	const steps: IPlanningPlanStep[] = [];
	const seen = new Set<string>();
	let currentSectionTitle: string | undefined;
	let currentSectionKind: PlanningPlanStepKind = 'step';
	let currentChunk: string[] = [];
	let currentChunkKind: PlanningPlanStepKind = currentSectionKind;
	let currentChunkSectionTitle: string | undefined = currentSectionTitle;
	let inCodeBlock = false;

	const flushChunk = () => {
		if (currentChunk.length === 0) {
			return;
		}

		const text = currentChunk.join('\n').trim();
		const label = normalizePlanningPlanStepLabel(currentChunk[0]);
		currentChunk = [];
		if (!label || !text) {
			return;
		}

		const key = `${currentChunkKind}:${label.toLowerCase()}:${text.toLowerCase()}`;
		if (seen.has(key) || shouldSkipPlanStepKind(currentChunkKind)) {
			return;
		}

		seen.add(key);
		steps.push({
			index: steps.length + 1,
			label,
			text,
			sectionTitle: currentChunkSectionTitle,
			kind: currentChunkKind,
		});
	};

	for (const rawLine of normalized.split(/\r?\n/g)) {
		const trimmed = rawLine.trim();
		if (/^```/.test(trimmed)) {
			inCodeBlock = !inCodeBlock;
			if (currentChunk.length > 0) {
				currentChunk.push(trimmed);
			}
			continue;
		}

		if (!trimmed) {
			continue;
		}

		if (!inCodeBlock) {
			const heading = parsePlanningPlanSectionHeading(trimmed);
			if (heading) {
				flushChunk();
				currentSectionTitle = heading.title;
				currentSectionKind = heading.kind;
				continue;
			}
		}

		const topLevelListItem = !inCodeBlock ? parseTopLevelPlanningListItem(rawLine) : undefined;
		if (topLevelListItem) {
			flushChunk();
			currentChunkKind = currentSectionKind;
			currentChunkSectionTitle = currentSectionTitle;
			currentChunk.push(topLevelListItem);
			if (steps.length >= maxSteps) {
				break;
			}
			continue;
		}

		if (!inCodeBlock && currentSectionTitle && !shouldSkipPlanStepKind(currentSectionKind)) {
			const isIndentedContinuation = /^\s{2,}\S/.test(rawLine);
			if (currentChunk.length > 0 && !isIndentedContinuation) {
				flushChunk();
			}

			currentChunkKind = currentSectionKind;
			currentChunkSectionTitle = currentSectionTitle;
			currentChunk.push(trimmed);
			if (steps.length >= maxSteps) {
				break;
			}
			continue;
		}

		if (currentChunk.length > 0) {
			currentChunk.push(trimmed);
		}
	}

	flushChunk();
	return steps.slice(0, maxSteps);
}

export function summarizePlanningPlanChanges(previousPlanText: string | undefined, currentPlanText: string | undefined): IPlanningPlanChangeSummary | undefined {
	if (!previousPlanText || !currentPlanText) {
		return undefined;
	}

	const previousLines = getPlanningPlanLines(previousPlanText);
	const currentLines = getPlanningPlanLines(currentPlanText);
	const previousKeys = new Set(previousLines.map(line => line.toLowerCase()));
	const currentKeys = new Set(currentLines.map(line => line.toLowerCase()));
	const added = currentLines.filter(line => !previousKeys.has(line.toLowerCase())).slice(0, 4);
	const removed = previousLines.filter(line => !currentKeys.has(line.toLowerCase())).slice(0, 3);
	return added.length > 0 || removed.length > 0 ? { added, removed } : undefined;
}

function getPlanningPlanLines(planText: string): string[] {
	const lines: string[] = [];
	const seen = new Set<string>();
	for (const rawLine of planText.split(/\r?\n/g)) {
		if (isNoisyPlanningLine(rawLine.trim())) {
			continue;
		}

		const normalized = normalizePlanningPlanLine(rawLine);
		if (!normalized) {
			continue;
		}

		const key = normalized.toLowerCase();
		if (seen.has(key)) {
			continue;
		}

		seen.add(key);
		lines.push(normalized);
	}

	return lines;
}

function normalizePlanningPlanStepLabel(line: string): string | undefined {
	const normalized = line
		.replace(/^\s{0,3}(?:(?:[-*+])\s+|\d+[.)]\s+)/, '')
		.replace(/\s+/g, ' ')
		.trim();
	return normalized.length >= 6 ? normalized : undefined;
}

function shouldSkipPlanStepKind(kind: PlanningPlanStepKind): boolean {
	return kind === 'other';
}

function parsePlanningPlanSectionHeading(line: string): { readonly title: string; readonly kind: PlanningPlanStepKind } | undefined {
	const headingMatch = /^(?:#{1,6}\s+(.+?)\s*#*|\*\*(.+?)\*\*\s*:?)$/.exec(line);
	const title = normalizePlanningPlanSectionTitle((headingMatch?.[1] ?? headingMatch?.[2])?.replace(/#+\s*$/, ''));
	if (!title) {
		return undefined;
	}

	if (/^plan\b/i.test(title)) {
		return undefined;
	}

	return {
		title,
		kind: getPlanningPlanSectionKind(title),
	};
}

function normalizePlanningPlanSectionTitle(title: string | undefined): string | undefined {
	const normalized = title?.replace(/\s+/g, ' ').trim();
	if (!normalized) {
		return undefined;
	}

	const withoutPhaseNumber = normalized.replace(/^phase\s+\d+\s*(?:[:.)-]\s*)?/i, '').trim();
	if (withoutPhaseNumber) {
		return withoutPhaseNumber;
	}

	const withoutBareNumber = normalized.replace(/^phase\s+\d+\b/i, 'Phase').trim();
	return withoutBareNumber || undefined;
}

function getPlanningPlanSectionKind(title: string): PlanningPlanStepKind {
	if (/\b(step|steps|implementation|approach|tasks?|work plan|execution)\b/i.test(title)) {
		return 'step';
	}

	if (/\b(verification|validate|validation|tests?|checks?|qa)\b/i.test(title)) {
		return 'verification';
	}

	if (/\b(decisions?|questions?|open items?|clarifications?)\b/i.test(title)) {
		return 'decision';
	}

	if (/\b(risks?|guardrails?|constraints?|assumptions?)\b/i.test(title)) {
		return 'guardrail';
	}

	if (/\b(relevant files?|files?|context|references?|dependencies?)\b/i.test(title)) {
		return 'other';
	}

	return 'step';
}

function parseTopLevelPlanningListItem(line: string): string | undefined {
	const match = /^\s{0,2}(?:[-*+]|\d+[.)])\s+(.+)$/.exec(line);
	return match?.[1]?.trim();
}

function getMarkdownCandidate(part: IResponse['value'][number]): string | undefined {
	switch (part.kind) {
		case 'markdownContent':
		case 'markdownVuln':
		case 'warning':
		case 'progressTask':
		case 'progressTaskSerialized':
			return normalizePlanText(part.content.value);
		case 'planningPlanEditor':
			return normalizePlanText(part.planText);
		default:
			return undefined;
	}
}

function getPlanCandidateVariants(candidate: string): string[] {
	const variants = new Set<string>();
	const normalized = normalizePlanText(candidate);
	if (!normalized) {
		return [];
	}

	variants.add(normalized);

	const lastPlanHeading = getLastMatchIndex(normalized, planHeadingPattern);
	if (lastPlanHeading !== undefined) {
		variants.add(normalizePlanText(normalized.slice(lastPlanHeading)) ?? normalized);
	}

	const firstPlanSection = getFirstMatchIndex(normalized, planSectionMarkerPattern);
	if (firstPlanSection !== undefined) {
		const sectionStart = findPreviousSectionBoundary(normalized, firstPlanSection);
		variants.add(normalizePlanText(normalized.slice(sectionStart)) ?? normalized);
	}

	return [...variants];
}

function scorePlanCandidate(candidate: string): number {
	const normalized = normalizePlanText(candidate);
	if (!normalized) {
		return Number.NEGATIVE_INFINITY;
	}

	const lines = normalized.split(/\r?\n/g).map(line => line.trim()).filter(line => line.length > 0);
	const firstPlanHeadingIndex = lines.findIndex(line => /^(?:#{1,6}\s*plan\b|(?:\*\*)?plan\s*:)/i.test(line));
	const startsWithPlanHeading = firstPlanHeadingIndex === 0;
	const startsWithPlanSection = /^(?:\*\*(?:steps|relevant files|verification|decisions)\*\*|#{1,6}\s*(?:steps|verification|decisions)\b)/i.test(lines[0] ?? '');
	const numberedSteps = normalized.match(/^\s*(?:[-*]|\d+[.)])\s+/gm)?.length ?? 0;
	const hasPlanHeading = /^(?:#{1,6}\s*plan\b.*|(?:\*\*)?plan\s*:\s*.+)$/im.test(normalized);
	const hasSteps = /^\*\*steps\*\*$/im.test(normalized) || /^#{1,6}\s*steps\b/im.test(normalized);
	const hasVerification = /^\*\*verification\*\*$/im.test(normalized) || /^#{1,6}\s*verification\b/im.test(normalized);
	const hasDecisions = /^\*\*decisions\*\*$/im.test(normalized) || /^#{1,6}\s*decisions\b/im.test(normalized);
	const hasRelevantFiles = /^\*\*relevant files\*\*$/im.test(normalized) || /^#{1,6}\s*relevant files\b/im.test(normalized);
	const noisyLines = lines.filter(isNoisyPlanningLine).length;

	let score = 0;
	if (hasPlanHeading) {
		score += 18;
	}
	if (startsWithPlanHeading) {
		score += 10;
	} else if (hasPlanHeading && firstPlanHeadingIndex > 0) {
		score -= Math.min(firstPlanHeadingIndex * 2, 8);
	}
	if (startsWithPlanSection) {
		score += 4;
	}
	if (hasSteps) {
		score += 10;
	}
	if (hasVerification) {
		score += 8;
	}
	if (hasDecisions) {
		score += 7;
	}
	if (hasRelevantFiles) {
		score += 7;
	}
	if (numberedSteps >= 3) {
		score += 6;
	} else if (numberedSteps > 0) {
		score += 3;
	}

	score += Math.min(lines.length, 12) * 0.25;
	score -= noisyLines * 5;

	return score;
}

function normalizePlanText(value: string | undefined): string | undefined {
	if (!value) {
		return undefined;
	}

	const normalized = stripPlanningScaffolding(value)
		.replace(/\r\n/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
	return normalized.length > 0 ? normalized : undefined;
}

function isNoisyPlanningLine(line: string): boolean {
	return /^(?:read\b|searched for\b|created memory file\b|read memory\b|wrote\b|updated\b)\b/i.test(line)
		|| /^(?:using your\b|i(?:'m| am)\b|i confirmed\b|i found\b|i checked\b|i reviewed\b|i will\b|i'll\b|next i(?:'ll| will)\b|now i(?:'ll| will)\b|plan persisted to\b|planner response\b|starting the planner response\b|estimated progress\b)/i.test(line)
		|| planningScaffoldingLinePattern.test(line)
		|| planningScaffoldingBlockHeadingPattern.test(line);
}

function stripPlanningScaffolding(value: string): string {
	const filteredLines: string[] = [];
	let skippingPromptList = false;

	for (const rawLine of value.replace(/\r\n/g, '\n').split('\n')) {
		const trimmed = rawLine.trim();
		if (skippingPromptList) {
			if (!trimmed) {
				skippingPromptList = false;
				continue;
			}

			if (/^(?:[-*+]\s+|\d+[.)]\s+)/.test(trimmed)) {
				continue;
			}

			skippingPromptList = false;
		}

		if (!trimmed) {
			filteredLines.push(rawLine);
			continue;
		}

		if (planningScaffoldingBlockHeadingPattern.test(trimmed)) {
			skippingPromptList = true;
			continue;
		}

		if (planningScaffoldingLinePattern.test(trimmed) || isToolTraceLine(trimmed)) {
			continue;
		}

		filteredLines.push(rawLine);
	}

	return filteredLines.join('\n');
}

function isToolTraceLine(line: string): boolean {
	return /^(?:read\b|searched for\b|created memory file\b|read memory\b|wrote\b|updated\b)\b/i.test(line);
}

function getLastMatchIndex(value: string, pattern: RegExp): number | undefined {
	const flags = pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`;
	const matcher = new RegExp(pattern.source, flags);
	let match = matcher.exec(value);
	let lastIndex: number | undefined;
	while (match) {
		lastIndex = match.index;
		match = matcher.exec(value);
	}
	return lastIndex;
}

function getFirstMatchIndex(value: string, pattern: RegExp): number | undefined {
	const flags = pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`;
	const matcher = new RegExp(pattern.source, flags);
	const match = matcher.exec(value);
	return match?.index;
}

function findPreviousSectionBoundary(value: string, index: number): number {
	const boundary = value.lastIndexOf('\n\n', index);
	return boundary >= 0 ? boundary + 2 : 0;
}
