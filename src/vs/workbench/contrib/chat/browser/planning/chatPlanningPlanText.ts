/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IResponse } from '../../common/model/chatModel.js';

export interface IPlanningPlanChangeSummary {
	readonly added: readonly string[];
	readonly removed: readonly string[];
}

const planHeadingPattern = /^#{1,6}\s*plan\b.*$/gim;
const planSectionMarkerPattern = /^(?:\*\*(?:steps|relevant files|verification|decisions)\*\*|#{1,6}\s*(?:steps|verification|decisions)\b)/gim;

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

	return bestCandidate ? normalizePlanText(bestCandidate) : undefined;
}

export function normalizePlanningPlanLine(line: string): string | undefined {
	const normalized = line
		.replace(/^\s{0,3}(?:[-*+]|\d+[.)]|#{1,6})\s*/, '')
		.replace(/\s+/g, ' ')
		.trim();
	return normalized.length >= 12 ? normalized : undefined;
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

function getMarkdownCandidate(part: IResponse['value'][number]): string | undefined {
	switch (part.kind) {
		case 'markdownContent':
		case 'markdownVuln':
		case 'warning':
		case 'progressTask':
		case 'progressTaskSerialized':
			return normalizePlanText(part.content.value);
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
	const numberedSteps = normalized.match(/^\s*(?:[-*]|\d+[.)])\s+/gm)?.length ?? 0;
	const hasPlanHeading = /^#{1,6}\s*plan\b.*$/im.test(normalized);
	const hasSteps = /^\*\*steps\*\*$/im.test(normalized) || /^#{1,6}\s*steps\b/im.test(normalized);
	const hasVerification = /^\*\*verification\*\*$/im.test(normalized) || /^#{1,6}\s*verification\b/im.test(normalized);
	const hasDecisions = /^\*\*decisions\*\*$/im.test(normalized) || /^#{1,6}\s*decisions\b/im.test(normalized);
	const hasRelevantFiles = /^\*\*relevant files\*\*$/im.test(normalized) || /^#{1,6}\s*relevant files\b/im.test(normalized);
	const noisyLines = lines.filter(isNoisyPlanningLine).length;

	let score = 0;
	if (hasPlanHeading) {
		score += 18;
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

	const normalized = value
		.replace(/\r\n/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
	return normalized.length > 0 ? normalized : undefined;
}

function isNoisyPlanningLine(line: string): boolean {
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
