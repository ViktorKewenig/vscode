/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { mainWindow } from '../../../../../../../base/browser/window.js';
import { URI } from '../../../../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../../base/test/common/utils.js';
import { workbenchInstantiationService } from '../../../../../../test/browser/workbenchTestServices.js';
import { ChatPlanningPlanEditorPart } from '../../../../browser/widget/chatContentParts/chatPlanningPlanEditorPart.js';
import { IChatContentPartRenderContext } from '../../../../browser/widget/chatContentParts/chatContentParts.js';
import { IChatPlanningPlanEditor, IChatQuestionAnswers, IChatService } from '../../../../common/chatService/chatService.js';

const planningPlanRegenerateControlsAnswerKey = 'plan-editor-regenerate-controls';
const planningPlanRegenerateControlsAnswerValue = 'regenerate-controls';

function createMockContext(): IChatContentPartRenderContext {
	return {
		element: {
			requestId: 'request-1',
			sessionResource: URI.parse('chat://test/session'),
			setVote: () => { },
		},
	} as unknown as IChatContentPartRenderContext;
}

function getButtonByLabel(container: HTMLElement, label: string): HTMLElement {
	const button = [...container.querySelectorAll<HTMLElement>('.monaco-button')]
		.find(candidate => candidate.textContent === label);
	assert.ok(button, `Expected button "${label}"`);
	return button;
}

suite('ChatPlanningPlanEditorPart', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	let part: ChatPlanningPlanEditorPart | undefined;
	let submitted: { readonly requestId: string; readonly resolveId: string; readonly answers: IChatQuestionAnswers | undefined } | undefined;

	teardown(() => {
		if (part?.domNode.parentNode) {
			part.domNode.parentNode.removeChild(part.domNode);
		}
		part = undefined;
		submitted = undefined;
	});

	function createPart(): ChatPlanningPlanEditorPart {
		const editor: IChatPlanningPlanEditor = {
			kind: 'planningPlanEditor',
			planText: '## Plan\n\n1. Inspect the file.\n2. Add tests.',
			resolveId: 'plan-editor-1',
			steps: [
				{
					id: 'step-1',
					index: 1,
					label: 'Inspect the file.',
					text: 'Inspect the file.',
					kind: 'step',
				},
			],
		};
		const chatService = {
			_serviceBrand: undefined,
			notifyQuestionCarouselAnswer: (requestId: string, resolveId: string, answers: IChatQuestionAnswers | undefined) => {
				submitted = { requestId, resolveId, answers };
			},
		} as Partial<IChatService> as IChatService;
		const instantiationService = workbenchInstantiationService(undefined, store);
		instantiationService.stub(IChatService, chatService);
		part = store.add(instantiationService.createInstance(ChatPlanningPlanEditorPart, editor, createMockContext()));
		mainWindow.document.body.appendChild(part.domNode);
		return part;
	}

	test('submits a control-regeneration request from the plan canvas', () => {
		const editorPart = createPart();
		getButtonByLabel(editorPart.domNode, 'Refresh controls').click();

		assert.deepStrictEqual({
			requestId: submitted?.requestId,
			resolveId: submitted?.resolveId,
			regenerateControls: submitted?.answers?.[planningPlanRegenerateControlsAnswerKey],
		}, {
			requestId: 'request-1',
			resolveId: 'plan-editor-1',
			regenerateControls: planningPlanRegenerateControlsAnswerValue,
		});
		assert.ok(editorPart.domNode.textContent?.includes('Refreshing controls.'));
	});
});
