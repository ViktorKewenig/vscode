/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { MarkdownString } from '../../../../../../base/common/htmlContent.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { extractPlanningPlanSteps, extractPlanningPlanText, summarizePlanningPlanChanges } from '../../../browser/planning/chatPlanningPlanText.js';
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

});
