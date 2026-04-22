/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { MarkdownString } from '../../../../../../base/common/htmlContent.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { extractPlanningPlanText, summarizePlanningPlanChanges } from '../../../browser/planning/chatPlanningPlanText.js';
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
});
