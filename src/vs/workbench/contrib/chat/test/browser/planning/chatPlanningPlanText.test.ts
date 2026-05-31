/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { MarkdownString } from '../../../../../../base/common/htmlContent.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { extractPlanningPlanSteps, extractPlanningPlanText, findPlanningPlanStartOffset, summarizePlanningPlanChanges } from '../../../browser/planning/chatPlanningPlanText.js';
import { Response } from '../../../common/model/chatModel.js';

suite('ChatPlanningPlanText', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('extracts the actual plan body from mixed markdown content', () => {
		const response = new Response([
			{ kind: 'markdownContent', content: new MarkdownString('CSV and analysis scripts\n\nRead files and inspect the workspace.\n\n## Plan: Analyze Orders CSV\n\n**Steps**\n1. Load `orders.csv`.\n2. Compare it with `schema.json`.\n\n**Verification**\n1. Confirm expected columns.\n') }
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'## Plan: Analyze Orders CSV',
			'',
			'**Steps**',
			'1. Load `orders.csv`.',
			'2. Compare it with `schema.json`.',
			'',
			'**Verification**',
			'1. Confirm expected columns.'
		].join('\n'));
	});

	test('strips planning prompt scaffolding from extracted plan text', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString([
					'Planning context from the previous planning step:',
					'Planning phase: focused-slice',
					'Planning answers:',
					'- Which file matters most?: `orders.csv`',
					'- What should the analysis produce?: A concise markdown summary',
					'Use this planning context as the source of truth for implementation unless the codebase forces a concrete adjustment.',
					'',
					'## Plan: Analyze Orders CSV',
					'',
					'**Steps**',
					'1. Load `orders.csv` into pandas.',
					'2. Validate the columns against `schema.json`.',
					'3. Produce a concise markdown summary.',
					'',
					'**Verification**',
					'1. Confirm the expected columns are present.',
				].join('\n'))
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'## Plan: Analyze Orders CSV',
			'',
			'**Steps**',
			'1. Load `orders.csv` into pandas.',
			'2. Validate the columns against `schema.json`.',
			'3. Produce a concise markdown summary.',
			'',
			'**Verification**',
			'1. Confirm the expected columns are present.',
		].join('\n'));
	});

	test('strips planning handoff instructions from extracted plan text', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString([
					'Planning handoff instruction:',
					'User request: analyze a CSV file',
					'The previous planner response completed without extractable plan markdown.',
					'This retry must end with the plan markdown itself.',
					'Create the first concrete plan now.',
					'While working, stream concise visible planning notes before the final plan: what context you are using, which constraints matter, and which tradeoffs shape the plan.',
					'Visible planning notes must be followed by the final plan in this same response.',
					'Do not stop after a progress update such as "I will inspect..." or "I will present...".',
					'Do not number phase or section headings. Keep numbering only on the individual plan steps.',
					'Return markdown that starts with a plan heading or a **Steps** section and includes **Relevant files**, **Verification**, and **Decisions** when applicable.',
					'',
					'## Plan: Analyze CSV',
					'',
					'**Steps**',
					'1. Load the CSV and inspect columns.',
					'2. Compute summary statistics and group comparisons.',
					'',
					'**Verification**',
					'1. Confirm every reported value matches the CSV.',
				].join('\n'))
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'## Plan: Analyze CSV',
			'',
			'**Steps**',
			'1. Load the CSV and inspect columns.',
			'2. Compute summary statistics and group comparisons.',
			'',
			'**Verification**',
			'1. Confirm every reported value matches the CSV.',
		].join('\n'));
	});

	test('does not treat transient planning status text as a plan', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString('I will quickly gather context from memory and the workspace, then produce a concrete execution plan focused on analysing llm_judge_ratings.xlsx with summary stats, outlier detection, and visualizations across all columns.')
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), undefined);
	});

	test('finds the final plan start after visible planning notes', () => {
		const text = [
			'I confirmed the anchor script behavior and outputs.',
			'Next I will draft a concrete, execution-ready analysis plan.',
			'',
			'Plan: Intervention Effect CSV Analysis',
			'',
			'Steps',
			'1. Confirm the input path.',
			'2. Validate the schema.',
			'',
			'Verification',
			'1. Run the anchor script.',
		].join('\n');

		assert.strictEqual(findPlanningPlanStartOffset(text), text.indexOf('Plan: Intervention Effect CSV Analysis'));
	});

	test('extracts from a plain Plan colon heading after visible notes', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString([
					'Using your planning context as the source of truth, I confirmed the CSV target.',
					'',
					'Plan: Intervention Effect CSV Analysis',
					'',
					'**Steps**',
					'1. Confirm the input path.',
					'2. Validate the schema.',
					'',
					'**Verification**',
					'1. Run the anchor script.',
				].join('\n'))
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'Plan: Intervention Effect CSV Analysis',
			'',
			'**Steps**',
			'1. Confirm the input path.',
			'2. Validate the schema.',
			'',
			'**Verification**',
			'1. Run the anchor script.',
		].join('\n'));
	});

	test('extracts a minimal two-step plan without an explicit heading', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString([
					'1. Load `llm_judge_ratings.xlsx` and inspect the available columns.',
					'2. Compute summary statistics and validation checks before plotting.',
				].join('\n'))
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'1. Load `llm_judge_ratings.xlsx` and inspect the available columns.',
			'2. Compute summary statistics and validation checks before plotting.',
		].join('\n'));
	});

	test('extracts the canonical plan value from a planning plan editor part', () => {
		const response = new Response([
			{
				kind: 'markdownContent',
				content: new MarkdownString('Plan drafted. It is open in the canvas for task-breakdown controls and revisions.')
			},
			{
				kind: 'planningPlanEditor',
				planText: [
					'## Plan',
					'',
					'1. Inspect the planning response capture path.',
					'2. Wire the canvas to the captured plan value.',
				].join('\n'),
				steps: [],
			}
		]);

		assert.strictEqual(extractPlanningPlanText(response), [
			'## Plan',
			'',
			'1. Inspect the planning response capture path.',
			'2. Wire the canvas to the captured plan value.',
		].join('\n'));
	});

	test('summarizes added and removed plan lines', () => {
		const changeSummary = summarizePlanningPlanChanges(
			[
				'## Plan',
				'',
				'1. Load `orders.csv`.',
				'2. Produce a summary.',
			].join('\n'),
			[
				'## Plan',
				'',
				'1. Load `orders.csv`.',
				'2. Validate against `schema.json`.',
				'3. Produce a summary.',
			].join('\n')
		);

		assert.deepStrictEqual(changeSummary, {
			added: ['Validate against `schema.json`.'],
			removed: []
		});
	});

	test('keeps revision summaries focused on plan line changes', () => {
		const changeSummary = summarizePlanningPlanChanges(
			[
				'Plan: CSV analysis',
				'',
				'**Steps**',
				'1. Load `orders.csv`.',
				'2. Produce a summary.',
			].join('\n'),
			[
				'I confirmed the anchor script behavior and outputs.',
				'Plan persisted to /memories/session/plan.md',
				'',
				'Plan: CSV analysis',
				'',
				'**Steps**',
				'1. Load `orders.csv`.',
				'2. Validate against `schema.json`.',
				'3. Produce a summary.',
			].join('\n')
		);

		assert.deepStrictEqual(changeSummary, {
			added: ['Validate against `schema.json`.'],
			removed: []
		});
	});

	test('extracts editable plan steps from steps and verification sections', () => {
		const steps = extractPlanningPlanSteps([
			'## Plan: Analyze Orders CSV',
			'',
			'**Relevant Files**',
			'- `orders.csv`',
			'- `schema.json`',
			'',
			'**Steps**',
			'1. Load `orders.csv` into pandas.',
			'   - Preserve the original column names.',
			'2. Validate the columns against `schema.json`.',
			'3. Produce a concise markdown summary.',
			'',
			'**Verification**',
			'- Confirm the expected columns are present.',
		].join('\n'));

		assert.deepStrictEqual(steps.map(step => ({
			label: step.label,
			kind: step.kind,
			sectionTitle: step.sectionTitle,
			text: step.text,
		})), [
			{
				label: 'Load `orders.csv` into pandas.',
				kind: 'step',
				sectionTitle: 'Steps',
				text: [
					'Load `orders.csv` into pandas.',
					'- Preserve the original column names.',
				].join('\n'),
			},
			{
				label: 'Validate the columns against `schema.json`.',
				kind: 'step',
				sectionTitle: 'Steps',
				text: 'Validate the columns against `schema.json`.',
			},
			{
				label: 'Produce a concise markdown summary.',
				kind: 'step',
				sectionTitle: 'Steps',
				text: 'Produce a concise markdown summary.',
			},
			{
				label: 'Confirm the expected columns are present.',
				kind: 'verification',
				sectionTitle: 'Verification',
				text: 'Confirm the expected columns are present.',
			},
		]);
	});

	test('extracts unnumbered paragraph steps from a steps section', () => {
		const steps = extractPlanningPlanSteps([
			'Plan: Intervention Success Analysis',
			'',
			'**TL;DR:** Create a standalone script for the contrasts CSV.',
			'',
			'**Steps**',
			'Create `analyze_intervention_contrasts.py` at the workspace root.',
			'**Data quality checks** load the CSV and report NA values.',
			'**Summary statistics** compute mean, median, and std.',
			'Print all output to stdout.',
			'',
			'**Relevant files**',
			'contrast.csv input data',
			'',
			'**Verification**',
			'Run `python analyze_intervention_contrasts.py`.',
		].join('\n'));

		assert.deepStrictEqual(steps.map(step => ({
			label: step.label,
			kind: step.kind,
			sectionTitle: step.sectionTitle,
		})), [
			{
				label: 'Create `analyze_intervention_contrasts.py` at the workspace root.',
				kind: 'step',
				sectionTitle: 'Steps',
			},
			{
				label: '**Data quality checks** load the CSV and report NA values.',
				kind: 'step',
				sectionTitle: 'Steps',
			},
			{
				label: '**Summary statistics** compute mean, median, and std.',
				kind: 'step',
				sectionTitle: 'Steps',
			},
			{
				label: 'Print all output to stdout.',
				kind: 'step',
				sectionTitle: 'Steps',
			},
			{
				label: 'Run `python analyze_intervention_contrasts.py`.',
				kind: 'verification',
				sectionTitle: 'Verification',
			},
		]);
	});

	test('extracts top-level list items when a plan has no explicit steps heading', () => {
		const steps = extractPlanningPlanSteps([
			'## Plan',
			'',
			'1. Inspect the current middleware flow.',
			'2. Add a bounded plan review carousel.',
			'   1. Keep nested details with the parent step.',
			'3. Run the planning tests.',
		].join('\n'));

		assert.deepStrictEqual(steps.map(step => step.text), [
			'Inspect the current middleware flow.',
			[
				'Add a bounded plan review carousel.',
				'1. Keep nested details with the parent step.',
			].join('\n'),
			'Run the planning tests.',
		]);
	});

	test('drops phase numbers from section titles while preserving numbered steps', () => {
		const steps = extractPlanningPlanSteps([
			'## Plan',
			'',
			'### Phase 1: Data Preparation',
			'1. Load `orders.csv`.',
			'2. Inspect the columns.',
			'',
			'### Phase 2 - Visualization',
			'1. Build summary charts.',
		].join('\n'));

		assert.deepStrictEqual(steps.map(step => ({
			index: step.index,
			label: step.label,
			sectionTitle: step.sectionTitle,
		})), [
			{
				index: 1,
				label: 'Load `orders.csv`.',
				sectionTitle: 'Data Preparation',
			},
			{
				index: 2,
				label: 'Inspect the columns.',
				sectionTitle: 'Data Preparation',
			},
			{
				index: 3,
				label: 'Build summary charts.',
				sectionTitle: 'Visualization',
			},
		]);
	});

});
