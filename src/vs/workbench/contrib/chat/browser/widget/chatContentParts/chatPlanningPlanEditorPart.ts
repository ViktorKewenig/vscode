/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import './media/chatPlanningPlanEditor.css';
import * as dom from '../../../../../../base/browser/dom.js';
import { Button } from '../../../../../../base/browser/ui/button/button.js';
import { Checkbox } from '../../../../../../base/browser/ui/toggle/toggle.js';
import { Emitter, Event } from '../../../../../../base/common/event.js';
import { Disposable, DisposableStore, IDisposable } from '../../../../../../base/common/lifecycle.js';
import { localize } from '../../../../../../nls.js';
import { defaultButtonStyles, defaultCheckboxStyles } from '../../../../../../platform/theme/browser/defaultStyles.js';
import { IChatMultiSelectAnswer, IChatPlanningPlanEditor, IChatQuestion, IChatQuestionAnswerValue, IChatQuestionAnswers, IChatService, IChatSingleSelectAnswer } from '../../../common/chatService/chatService.js';
import { IChatProgressRenderableResponseContent } from '../../../common/model/chatModel.js';
import { isResponseVM } from '../../../common/model/chatViewModel.js';
import { IChatContentPart, IChatContentPartRenderContext } from './chatContentParts.js';

const $ = dom.$;

interface IPlanningPlanEditorStepControls {
	readonly keep: Checkbox;
	readonly revise: Checkbox;
	readonly defer: Checkbox;
	readonly split: Checkbox;
	readonly note: HTMLTextAreaElement;
	readonly questions: readonly IPlanningPlanEditorQuestionControls[];
}

interface IPlanningPlanEditorQuestionControls {
	readonly question: IChatQuestion;
	readonly text?: HTMLTextAreaElement;
	readonly options?: readonly { readonly optionId: string; readonly value: string; readonly checkbox: Checkbox }[];
	readonly freeform?: HTMLTextAreaElement;
}

export class ChatPlanningPlanEditorPart extends Disposable implements IChatContentPart {
	public readonly domNode: HTMLElement;

	private readonly _onDidChangeHeight = this._register(new Emitter<void>());
	public readonly onDidChangeHeight: Event<void> = this._onDidChangeHeight.event;

	private readonly stepControls = new Map<string, IPlanningPlanEditorStepControls>();
	private additionalNotes: HTMLTextAreaElement | undefined;
	private applyButton: Button | undefined;

	constructor(
		private readonly editor: IChatPlanningPlanEditor,
		private readonly context: IChatContentPartRenderContext,
		@IChatService private readonly chatService: IChatService,
	) {
		super();

		this.domNode = $('.chat-planning-plan-editor');
		this.domNode.tabIndex = 0;
		this.render();
	}

	private render(): void {
		dom.clearNode(this.domNode);
		if (this.editor.isUsed) {
			this.renderSummary();
			return;
		}

		const header = dom.append(this.domNode, $('.chat-planning-plan-editor-header'));
		const title = dom.append(header, $('.chat-planning-plan-editor-title'));
		title.textContent = localize('chat.planningPlanEditor.title', 'Plan Editor');

		const subtitle = dom.append(header, $('.chat-planning-plan-editor-subtitle'));
		subtitle.textContent = localize('chat.planningPlanEditor.subtitle', 'Review the generated plan in place. Each step has its own refinement question.');

		const stepsContainer = dom.append(this.domNode, $('.chat-planning-plan-editor-steps'));
		for (const step of this.editor.steps) {
			this.renderStep(stepsContainer, step);
		}

		const additional = dom.append(this.domNode, $('.chat-planning-plan-editor-additional'));
		const additionalLabel = dom.append(additional, $('label.chat-planning-plan-editor-label'));
		additionalLabel.textContent = localize('chat.planningPlanEditor.otherEdits', 'Other edits');
		this.additionalNotes = dom.append(additional, $('textarea.chat-planning-plan-editor-textarea'));
		this.additionalNotes.rows = 2;
		this.additionalNotes.placeholder = localize('chat.planningPlanEditor.otherEditsPlaceholder', 'Add reordering, scope, validation, or missing-step notes.');
		this._register(dom.addDisposableListener(this.additionalNotes, dom.EventType.INPUT, () => this.handleInputChange()));

		const footer = dom.append(this.domNode, $('.chat-planning-plan-editor-footer'));
		const continueButton = this._register(new Button(footer, { ...defaultButtonStyles, secondary: true }));
		continueButton.label = localize('chat.planningPlanEditor.continue', 'Continue');
		this._register(continueButton.onDidClick(() => this.submit(false)));

		this.applyButton = this._register(new Button(footer, defaultButtonStyles));
		this.applyButton.label = localize('chat.planningPlanEditor.applyEdits', 'Apply Edits');
		this._register(this.applyButton.onDidClick(() => this.submit(true)));
		this.updateApplyButton();
	}

	private renderStep(container: HTMLElement, step: IChatPlanningPlanEditor['steps'][number]): void {
		const stepElement = dom.append(container, $('.chat-planning-plan-editor-step'));
		const stepHeader = dom.append(stepElement, $('.chat-planning-plan-editor-step-header'));
		const index = dom.append(stepHeader, $('.chat-planning-plan-editor-step-index'));
		index.textContent = String(step.index);
		const title = dom.append(stepHeader, $('.chat-planning-plan-editor-step-title'));
		title.textContent = step.sectionTitle && step.kind !== 'step'
			? `${step.sectionTitle}: ${step.label}`
			: step.label;

		if (step.text !== step.label) {
			const detail = dom.append(stepElement, $('.chat-planning-plan-editor-step-detail'));
			detail.textContent = step.text;
		}

		const store = this._register(new DisposableStore());
		const questionControls = this.renderStepQuestions(stepElement, step, store);

		const actions = dom.append(stepElement, $('.chat-planning-plan-editor-actions'));
		const keep = this.createActionCheckbox(actions, store, localize('chat.planningPlanEditor.keep', 'Keep'), true);
		const revise = this.createActionCheckbox(actions, store, localize('chat.planningPlanEditor.revise', 'Revise'));
		const defer = this.createActionCheckbox(actions, store, localize('chat.planningPlanEditor.defer', 'Defer or remove'));
		const split = this.createActionCheckbox(actions, store, localize('chat.planningPlanEditor.split', 'Split'));

		const note = dom.append(stepElement, $<HTMLTextAreaElement>('textarea.chat-planning-plan-editor-textarea'));
		note.rows = 1;
		note.placeholder = localize('chat.planningPlanEditor.stepNotePlaceholder', 'Answer or add a note for this step');
		this._register(dom.addDisposableListener(note, dom.EventType.INPUT, () => this.handleInputChange()));

		const controls = { keep, revise, defer, split, note, questions: questionControls };
		this.stepControls.set(step.id, controls);
		for (const checkbox of [keep, revise, defer, split]) {
			store.add(checkbox.onChange(() => {
				if (checkbox !== keep && checkbox.checked) {
					keep.checked = false;
				}
				if (checkbox === keep && keep.checked) {
					revise.checked = false;
					defer.checked = false;
					split.checked = false;
				}
				if (!keep.checked && !revise.checked && !defer.checked && !split.checked) {
					keep.checked = true;
				}
				this.handleInputChange();
			}));
		}
	}

	private renderStepQuestions(stepElement: HTMLElement, step: IChatPlanningPlanEditor['steps'][number], store: DisposableStore): readonly IPlanningPlanEditorQuestionControls[] {
		if (!step.questions?.length) {
			const question = dom.append(stepElement, $('.chat-planning-plan-editor-step-question'));
			question.textContent = this.getStepQuestion(step.kind);
			return [];
		}

		const questionsContainer = dom.append(stepElement, $('.chat-planning-plan-editor-step-questions'));
		return step.questions.map(question => this.renderStepQuestion(questionsContainer, question, store));
	}

	private renderStepQuestion(container: HTMLElement, question: IChatQuestion, store: DisposableStore): IPlanningPlanEditorQuestionControls {
		const questionElement = dom.append(container, $('.chat-planning-plan-editor-step-question-card'));
		const title = dom.append(questionElement, $('.chat-planning-plan-editor-step-question-title'));
		title.textContent = question.title;

		const message = typeof question.message === 'string' ? question.message : question.message?.value;
		if (message && message !== question.title) {
			const messageElement = dom.append(questionElement, $('.chat-planning-plan-editor-step-question-message'));
			messageElement.textContent = message;
		}

		if (question.description) {
			const description = dom.append(questionElement, $('.chat-planning-plan-editor-step-question-description'));
			description.textContent = question.description;
		}

		if (question.type === 'text') {
			const text = dom.append(questionElement, $<HTMLTextAreaElement>('textarea.chat-planning-plan-editor-textarea'));
			text.rows = 1;
			text.placeholder = localize('chat.planningPlanEditor.questionTextPlaceholder', 'Answer this question for the step');
			this._register(dom.addDisposableListener(text, dom.EventType.INPUT, () => this.handleInputChange()));
			return { question, text };
		}

		const optionControls = this.renderQuestionOptions(questionElement, question, store);
		const freeform = question.allowFreeformInput === false
			? undefined
			: dom.append(questionElement, $<HTMLTextAreaElement>('textarea.chat-planning-plan-editor-textarea.chat-planning-plan-editor-question-freeform'));
		if (freeform) {
			freeform.rows = 1;
			freeform.placeholder = localize('chat.planningPlanEditor.questionFreeformPlaceholder', 'Add a different answer or note');
			this._register(dom.addDisposableListener(freeform, dom.EventType.INPUT, () => this.handleInputChange()));
		}

		return {
			question,
			options: optionControls,
			...(freeform ? { freeform } : {}),
		};
	}

	private renderQuestionOptions(container: HTMLElement, question: IChatQuestion, store: DisposableStore): readonly { readonly optionId: string; readonly value: string; readonly checkbox: Checkbox }[] {
		const optionsContainer = dom.append(container, $('.chat-planning-plan-editor-question-options'));
		const controls = (question.options ?? []).map(option => {
			const checkbox = this.createActionCheckbox(optionsContainer, store, option.label);
			return {
				optionId: option.id,
				value: option.value,
				checkbox,
			};
		});

		for (const control of controls) {
			store.add(control.checkbox.onChange(() => {
				if (question.type === 'singleSelect' && control.checkbox.checked) {
					for (const otherControl of controls) {
						if (otherControl !== control) {
							otherControl.checkbox.checked = false;
						}
					}
				}
				this.handleInputChange();
			}));
		}

		return controls;
	}

	private createActionCheckbox(container: HTMLElement, store: DisposableStore, label: string, checked = false): Checkbox {
		const checkbox = store.add(new Checkbox(label, checked, defaultCheckboxStyles));
		checkbox.domNode.classList.add('chat-planning-plan-editor-action');
		container.appendChild(checkbox.domNode);
		return checkbox;
	}

	private getStepQuestion(kind: IChatPlanningPlanEditor['steps'][number]['kind']): string {
		switch (kind) {
			case 'verification':
				return localize('chat.planningPlanEditor.verificationQuestion', 'Question: What evidence or check would make this verification step sufficient?');
			case 'decision':
				return localize('chat.planningPlanEditor.decisionQuestion', 'Question: What decision needs to be resolved before this step is useful?');
			case 'guardrail':
				return localize('chat.planningPlanEditor.guardrailQuestion', 'Question: What constraint or assumption should this guardrail capture?');
			default:
				return localize('chat.planningPlanEditor.stepQuestion', 'Question: What should change before this step is implemented?');
		}
	}

	private handleInputChange(): void {
		this.updateApplyButton();
		this._onDidChangeHeight.fire();
	}

	private updateApplyButton(): void {
		if (this.applyButton) {
			this.applyButton.enabled = this.hasEdits();
		}
	}

	private hasEdits(): boolean {
		if (this.additionalNotes?.value.trim()) {
			return true;
		}

		for (const controls of this.stepControls.values()) {
			if (controls.revise.checked || controls.defer.checked || controls.split.checked || controls.note.value.trim()) {
				return true;
			}
			if (controls.questions.some(question => this.getQuestionAnswer(question) !== undefined)) {
				return true;
			}
		}

		return false;
	}

	private submit(requireEdits: boolean): void {
		if (requireEdits && !this.hasEdits()) {
			return;
		}

		const answers = this.collectAnswers();
		this.editor.data = answers;
		this.editor.isUsed = true;
		this.renderSummary();
		this._onDidChangeHeight.fire();

		if (isResponseVM(this.context.element) && this.editor.resolveId) {
			this.chatService.notifyQuestionCarouselAnswer(this.context.element.requestId, this.editor.resolveId, answers);
		}
	}

	private collectAnswers(): IChatQuestionAnswers {
		const answers: IChatQuestionAnswers = {};
		for (const step of this.editor.steps) {
			const controls = this.stepControls.get(step.id);
			if (!controls) {
				continue;
			}

			const selectedValues: string[] = [];
			if (controls.keep.checked) {
				selectedValues.push('keep');
			}
			if (controls.revise.checked) {
				selectedValues.push('revise');
			}
			if (controls.defer.checked) {
				selectedValues.push('defer');
			}
			if (controls.split.checked) {
				selectedValues.push('split');
			}

			answers[`plan-editor-step-${step.id}`] = {
				selectedValues,
				freeformValue: controls.note.value.trim() || undefined,
			} satisfies IChatMultiSelectAnswer;

			for (const questionControl of controls.questions) {
				const answer = this.getQuestionAnswer(questionControl);
				if (answer !== undefined) {
					answers[`plan-editor-question-${step.id}-${questionControl.question.id}`] = answer;
				}
			}
		}

		const additional = this.additionalNotes?.value.trim();
		if (additional) {
			answers['plan-editor-additional'] = additional;
		}

		return answers;
	}

	private getQuestionAnswer(control: IPlanningPlanEditorQuestionControls): IChatQuestionAnswerValue | undefined {
		if (control.question.type === 'text') {
			const value = control.text?.value.trim();
			return value ? value : undefined;
		}

		const selectedValues = control.options
			?.filter(option => option.checkbox.checked)
			.map(option => option.value) ?? [];
		const freeformValue = control.freeform?.value.trim();
		if (control.question.type === 'singleSelect') {
			if (selectedValues.length === 0 && !freeformValue) {
				return undefined;
			}

			return {
				selectedValue: selectedValues[0],
				...(freeformValue ? { freeformValue } : {}),
			} satisfies IChatSingleSelectAnswer;
		}

		if (selectedValues.length === 0 && !freeformValue) {
			return undefined;
		}

		return {
			selectedValues,
			...(freeformValue ? { freeformValue } : {}),
		} satisfies IChatMultiSelectAnswer;
	}

	private renderSummary(): void {
		dom.clearNode(this.domNode);
		this.domNode.classList.add('chat-planning-plan-editor-used');
		const summary = dom.append(this.domNode, $('.chat-planning-plan-editor-summary'));
		const title = dom.append(summary, $('.chat-planning-plan-editor-title'));
		title.textContent = localize('chat.planningPlanEditor.summaryTitle', 'Plan Editor');

		const edits = this.countSubmittedEdits();
		const detail = dom.append(summary, $('.chat-planning-plan-editor-subtitle'));
		detail.textContent = edits > 0
			? localize('chat.planningPlanEditor.summaryWithEdits', '{0} plan edit(s) submitted.', edits)
			: localize('chat.planningPlanEditor.summaryNoEdits', 'No step edits submitted.');
	}

	private countSubmittedEdits(): number {
		const data = this.editor.data;
		if (!data) {
			return 0;
		}

		let edits = typeof data['plan-editor-additional'] === 'string' && data['plan-editor-additional'].trim() ? 1 : 0;
		for (const [key, value] of Object.entries(data)) {
			if (key.startsWith('plan-editor-question-') && this.hasSubmittedQuestionAnswer(value)) {
				edits++;
				continue;
			}
			if (typeof value !== 'object' || value === null || !('selectedValues' in value)) {
				continue;
			}
			const selectedValues = Array.isArray(value.selectedValues) ? value.selectedValues : [];
			const freeformValue = typeof value.freeformValue === 'string' ? value.freeformValue.trim() : '';
			if (selectedValues.some(selected => selected !== 'keep') || freeformValue) {
				edits++;
			}
		}
		return edits;
	}

	private hasSubmittedQuestionAnswer(value: IChatQuestionAnswerValue): boolean {
		if (typeof value === 'string') {
			return !!value.trim();
		}

		if ('selectedValues' in value) {
			return value.selectedValues.length > 0 || !!value.freeformValue?.trim();
		}

		return !!value.selectedValue || !!value.freeformValue?.trim();
	}

	hasSameContent(other: IChatProgressRenderableResponseContent): boolean {
		return other.kind === 'planningPlanEditor'
			&& other.resolveId === this.editor.resolveId
			&& other.isUsed === this.editor.isUsed;
	}

	addDisposable(disposable: IDisposable): void {
		this._register(disposable);
	}
}
