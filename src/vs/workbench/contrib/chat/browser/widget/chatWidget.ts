/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import './media/chat.css';
import './media/chatAgentHover.css';
import './media/chatViewWelcome.css';
import * as dom from '../../../../../base/browser/dom.js';
import { status } from '../../../../../base/browser/ui/aria/aria.js';
import { IMouseWheelEvent } from '../../../../../base/browser/mouseEvent.js';
import { disposableTimeout, raceTimeout, timeout } from '../../../../../base/common/async.js';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { Codicon } from '../../../../../base/common/codicons.js';
import { toErrorMessage } from '../../../../../base/common/errorMessage.js';
import { Emitter, Event } from '../../../../../base/common/event.js';
import { hash } from '../../../../../base/common/hash.js';
import { IMarkdownString, MarkdownString } from '../../../../../base/common/htmlContent.js';
import { Iterable } from '../../../../../base/common/iterator.js';
import { Disposable, DisposableStore, IDisposable, MutableDisposable, thenIfNotDisposed } from '../../../../../base/common/lifecycle.js';
import { ResourceSet } from '../../../../../base/common/map.js';
import { Schemas } from '../../../../../base/common/network.js';
import { IsSessionsWindowContext } from '../../../../common/contextkeys.js';
import { filter } from '../../../../../base/common/objects.js';
import { autorun, constObservable, derived, observableFromEvent, observableValue } from '../../../../../base/common/observable.js';
import { extUri, isEqual } from '../../../../../base/common/resources.js';
import { MicrotaskDelay } from '../../../../../base/common/symbols.js';
import { hasKey, isDefined } from '../../../../../base/common/types.js';
import { URI } from '../../../../../base/common/uri.js';
import { ICodeEditor } from '../../../../../editor/browser/editorBrowser.js';
import { ICodeEditorService } from '../../../../../editor/browser/services/codeEditorService.js';
import { OffsetRange } from '../../../../../editor/common/core/ranges/offsetRange.js';
import { Range } from '../../../../../editor/common/core/range.js';
import { ILanguageFeaturesService } from '../../../../../editor/common/services/languageFeatures.js';
import { ITextModelService } from '../../../../../editor/common/services/resolverService.js';
import { localize } from '../../../../../nls.js';
import { MenuId } from '../../../../../platform/actions/common/actions.js';
import { ICommandService } from '../../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { IContextKey, IContextKeyService } from '../../../../../platform/contextkey/common/contextkey.js';
import { IDialogService } from '../../../../../platform/dialogs/common/dialogs.js';

import { ITextResourceEditorInput } from '../../../../../platform/editor/common/editor.js';
import { IFileService } from '../../../../../platform/files/common/files.js';
import { IInstantiationService } from '../../../../../platform/instantiation/common/instantiation.js';
import { ServiceCollection } from '../../../../../platform/instantiation/common/serviceCollection.js';
import { ILogService } from '../../../../../platform/log/common/log.js';
import { bindContextKey } from '../../../../../platform/observable/common/platformObservableUtils.js';
import product from '../../../../../platform/product/common/product.js';
import { ITelemetryService } from '../../../../../platform/telemetry/common/telemetry.js';
import { IThemeService } from '../../../../../platform/theme/common/themeService.js';
import { IWorkspaceContextService } from '../../../../../platform/workspace/common/workspace.js';
import { IChatEntitlementService } from '../../../../services/chat/common/chatEntitlementService.js';
import { ILifecycleService } from '../../../../services/lifecycle/common/lifecycle.js';
import { checkModeOption } from '../../common/chat.js';
import { IChatAgentAttachmentCapabilities, IChatAgentCommand, IChatAgentData, IChatAgentService, UserSelectedTools } from '../../common/participants/chatAgents.js';
import { ChatContextKeys } from '../../common/actions/chatContextKeys.js';
import { applyingChatEditsFailedContextKey, decidedChatEditingResourceContextKey, hasAppliedChatEditsContextKey, hasUndecidedChatEditingResourceContextKey, IChatEditingService, IChatEditingSession, inChatEditingSessionContextKey, ModifiedFileEntryState } from '../../common/editing/chatEditingService.js';
import { IChatLayoutService } from '../../common/widget/chatLayoutService.js';
import { IChatModel, IChatModelInputState, IChatResponseModel } from '../../common/model/chatModel.js';
import { ChatMode, getModeNameForTelemetry, IChatMode, IChatModeService } from '../../common/chatModes.js';
import { chatAgentLeader, ChatRequestAgentPart, ChatRequestDynamicVariablePart, ChatRequestSlashPromptPart, ChatRequestToolPart, ChatRequestToolSetPart, chatSubcommandLeader, formatChatQuestion, IParsedChatRequest } from '../../common/requestParser/chatParserTypes.js';
import { ChatRequestParser } from '../../common/requestParser/chatRequestParser.js';
import { getDynamicVariablesForWidget, getSelectedToolAndToolSetsForWidget } from '../attachments/chatVariables.js';
import { ChatRequestQueueKind, ChatSendResult, IChatLocationData, IChatPlanningPlanEditor, IChatPlanningPlanEditorStep, IChatProgress, IChatProgressMessage, IChatQuestion, IChatQuestionAnswerValue, IChatQuestionAnswers, IChatQuestionCarousel, IChatSendRequestOptions, IChatService, IChatToolInvocation } from '../../common/chatService/chatService.js';
import { IChatSessionsService } from '../../common/chatSessionsService.js';
import { IChatSlashCommandService } from '../../common/participants/chatSlashCommands.js';
import { IChatArtifactsService } from '../../common/tools/chatArtifactsService.js';
import { IChatTodoListService } from '../../common/tools/chatTodoListService.js';
import { ChatRequestVariableSet, IChatRequestVariableEntry, isPromptFileVariableEntry, isPromptTextVariableEntry, isWorkspaceVariableEntry, PromptFileVariableKind, toPromptFileVariableEntry, toPromptTextVariableEntry } from '../../common/attachments/chatVariableEntries.js';
import { ChatViewModel, IChatRequestViewModel, IChatResponseViewModel, isRequestVM, isResponseVM } from '../../common/model/chatViewModel.js';
import { CodeBlockModelCollection } from '../../common/widget/codeBlockModelCollection.js';
import { ChatAgentLocation, ChatConfiguration, ChatModeKind, ChatPermissionLevel, ThinkingDisplayMode } from '../../common/constants.js';
import { ILanguageModelToolsService, isToolSet } from '../../common/tools/languageModelToolsService.js';
import { ILanguageModelsService } from '../../common/languageModels.js';
import { ComputeAutomaticInstructions } from '../../common/promptSyntax/computeAutomaticInstructions.js';
import { IHandOff, PromptHeader } from '../../common/promptSyntax/promptFileParser.js';
import { assessPlanningReadiness } from '../../common/planning/chatPlanningReadiness.js';
import { shouldRegeneratePlanningQuestions } from '../../common/planning/chatPlanningQuestionHeuristics.js';
import { augmentPromptWithPlanningContext, buildPlanningTransitionContext, getNextPlanningPhase, getPreviousPlanningPhase, IPlanningTransitionContext, isPlanningMiddlewareQuestionCarousel, isPlanningModeName, mergePlanningTransitionContexts, PlanningPhase, PlanningQuestionStage, planningMiddlewareQuestionCarouselResolveIdPrefix } from '../../common/planning/chatPlanningTransition.js';
import { IPromptsService, PromptsStorage } from '../../common/promptSyntax/service/promptsService.js';
import { GENERATE_AGENT_INSTRUCTIONS_COMMAND_ID, handleModeSwitch, OPEN_PLANNING_PLAN_ACTION_ID, UPDATE_PLANNING_PLAN_ACTION_ID } from '../actions/chatActions.js';
import { ChatTreeItem, IChatAcceptInputOptions, IChatAccessibilityService, IChatCodeBlockInfo, IChatFileTreeInfo, IChatListItemRendererOptions, IChatWidget, IChatWidgetService, IChatWidgetViewContext, IChatWidgetViewModelChangeEvent, IChatWidgetViewOptions, isIChatResourceViewContext, isIChatViewViewContext } from '../chat.js';
import { ChatAttachmentModel } from '../attachments/chatAttachmentModel.js';
import { IChatAttachmentResolveService } from '../attachments/chatAttachmentResolveService.js';
import { ChatDynamicVariableModel } from '../attachments/chatDynamicVariables.js';
import { ChatSuggestNextWidget } from './chatContentParts/chatSuggestNextWidget.js';
import { ChatInputPart, IChatInputPartOptions, IChatInputStyles } from './input/chatInputPart.js';
import { IChatListItemTemplate } from './chatListRenderer.js';
import { ChatListWidget } from './chatListWidget.js';
import { ChatEditorOptions } from './chatOptions.js';
import { ChatViewWelcomePart, IChatViewWelcomeContent } from '../viewsWelcome/chatViewWelcomeController.js';
import { IChatTipService } from '../chatTipService.js';
import { ChatTipContentPart } from './chatContentParts/chatTipContentPart.js';
import { ChatContentMarkdownRenderer } from './chatContentMarkdownRenderer.js';
import { IAgentSessionsService } from '../agentSessions/agentSessionsService.js';
import { IChatDebugService } from '../../common/chatDebugService.js';
import { ChatQuestionCarouselData } from '../../common/model/chatProgressTypes/chatQuestionCarouselData.js';
import { collectPlanningRepositoryContext } from '../planning/chatPlanningContextCollector.js';
import { extractPlanningPlanSteps, extractPlanningPlanText, IPlanningPlanStep, isUsablePlanningPlanText } from '../planning/chatPlanningPlanText.js';
import { generateDynamicPlanningPlanStepControlsResult, generateDynamicPlanningQuestionsResult, IGeneratedPlanningPlanStepControlsResult, IGeneratedPlanningQuestionsResult, IPlanningQuestionGenerationContext } from '../planning/chatPlanningQuestionGenerator.js';

const $ = dom.$;
const planningQuestionGenerationTimeoutMs = 45000;
const planningPlanProgressUpdateIntervalMs = 1000;
const planningPlanProgressBarSegments = 24;
const planningPlanProgressMessageId = 'planning-plan-progress';
const minGoalClarityQuestionRoundsBeforeFirstPlan = 2;
const goalClarityRoundToOfferPlanProceed = 3;
const goalClarityProceedQuestionId = 'dynamic-planning-proceed-to-plan';
const goalClarityProceedNowValue = 'proceed-to-plan';
const goalClarityContinueClarifyingValue = 'continue-goal-clarity';
const planningPlanRegenerateControlsAnswerKey = 'plan-editor-regenerate-controls';
const planningPlanRegenerateControlsAnswerValue = 'regenerate-controls';
const planningPlanControlFocusAnswerKey = 'plan-editor-control-focus';

export interface IChatWidgetStyles extends IChatInputStyles {
	readonly inputEditorBackground: string;
	readonly resultEditorBackground: string;
}

export interface IChatWidgetContrib extends IDisposable {

	readonly id: string;

	/**
	 * A piece of state which is related to the input editor of the chat widget.
	 * Takes in the `contrib` object that will be saved in the {@link IChatModelInputState}.
	 */
	getInputState?(contrib: Record<string, unknown>): void;

	/**
	 * Called with the result of getInputState when navigating input history.
	 */
	setInputState?(contrib: Readonly<Record<string, unknown>>): void;
}

interface IChatRequestInputOptions {
	input: string;
	attachedContext: ChatRequestVariableSet;
}

interface IPlanningPlanSnapshot {
	readonly requestId: string;
	readonly planText: string;
	readonly previousRequestId?: string;
	readonly previousPlanText?: string;
}

type PlanningPlanProgressSource = 'goal-clarity' | 'task-decomposition' | 'plan-review' | 'plan-focus';

interface IPlanningPlanRequestResult {
	readonly response: IChatResponseModel;
	readonly planSnapshot?: IPlanningPlanSnapshot;
}

interface IPlanningPlanProgressSnapshot {
	readonly percent: number;
	readonly elapsedLabel: string;
	readonly status: string;
	readonly progressBar: string;
}

interface IPlanningPlanProgressTracker extends IDisposable {
	setResponse(response: IChatResponseModel): void;
	complete(status: string): Promise<void>;
}

interface IGoalClarityContinuationDecision {
	readonly shouldContinue: boolean;
	readonly focusHint?: string;
}

interface IPlanningPlanStepReviewOutcome {
	readonly hasEdits: boolean;
	readonly plannerNotes?: string;
}

export interface IChatWidgetLocationOptions {
	location: ChatAgentLocation;

	resolveData?(): IChatLocationData | undefined;
}

export function isQuickChat(widget: IChatWidget): boolean {
	return isIChatResourceViewContext(widget.viewContext) && Boolean(widget.viewContext.isQuickChat);
}

function isInlineChat(widget: IChatWidget): boolean {
	return isIChatResourceViewContext(widget.viewContext) && Boolean(widget.viewContext.isInlineChat);
}

type ChatHandoffClickEvent = {
	fromAgent: string;
	toAgent: string;
	hasPrompt: boolean;
	autoSend: boolean;
};

type ChatHandoffClickClassification = {
	owner: 'digitarald';
	comment: 'Event fired when a user clicks on a handoff prompt in the chat suggest-next widget';
	fromAgent: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The agent/mode the user was in before clicking the handoff' };
	toAgent: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The agent/mode specified in the handoff' };
	hasPrompt: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; isMeasurement: true; comment: 'Whether the handoff includes a prompt' };
	autoSend: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; isMeasurement: true; comment: 'Whether the handoff automatically submits the request' };
};

type ChatHandoffWidgetShownEvent = {
	agent: string;
	handoffCount: number;
};

type ChatHandoffWidgetShownClassification = {
	owner: 'digitarald';
	comment: 'Event fired when the suggest-next widget is shown with handoff prompts';
	agent: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The current agent/mode that has handoffs defined' };
	handoffCount: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; isMeasurement: true; comment: 'Number of handoff options shown to the user' };
};

type ChatPromptRunEvent = {
	storage: PromptsStorage;
	extensionId?: string;
	promptName?: string;
	promptNameHash?: string;
};

type ChatPromptRunClassification = {
	owner: 'digitarald';
	comment: 'Event fired when a prompt slash command is resolved into a follow instructions request';
	storage: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Where the prompt is stored (local, user, extension).' };
	extensionId?: { classification: 'PublicNonPersonalData'; purpose: 'FeatureInsight'; comment: 'Identifier of the extension that contributed the prompt, when applicable.' };
	promptName?: { classification: 'PublicNonPersonalData'; purpose: 'FeatureInsight'; comment: 'Name of the core or extension-contributed prompt.' };
	promptNameHash?: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Hashed name of local or user prompt for privacy.' };
};

type ChatThinkingStyleUsageEvent = {
	thinkingStyle: ThinkingDisplayMode;
	location: ChatAgentLocation;
	requestKind: 'submit' | 'rerun';
};

type ChatThinkingStyleUsageClassification = {
	owner: 'justschen';
	comment: 'Event fired when a chat request uses the configured thinking style rendering mode.';
	thinkingStyle: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The configured rendering mode for thinking content.' };
	location: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The location where the request was made.' };
	requestKind: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Whether the request was a new submit or a rerun.' };
};

const supportsAllAttachments: Required<IChatAgentAttachmentCapabilities> = {
	supportsFileAttachments: true,
	supportsToolAttachments: true,
	supportsMCPAttachments: true,
	supportsImageAttachments: true,
	supportsSearchResultAttachments: true,
	supportsInstructionAttachments: true,
	supportsSourceControlAttachments: true,
	supportsProblemAttachments: true,
	supportsSymbolAttachments: true,
	supportsTerminalAttachments: true,
	supportsPromptAttachments: true,
	supportsHandOffs: true,
};

const DISCLAIMER = localize('chatDisclaimer', "AI responses may be inaccurate");

export class ChatWidget extends Disposable implements IChatWidget {

	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	static readonly CONTRIBS: { new(...args: [IChatWidget, ...any]): IChatWidgetContrib }[] = [];

	private readonly _onDidSubmitAgent = this._register(new Emitter<{ agent: IChatAgentData; slashCommand?: IChatAgentCommand }>());
	readonly onDidSubmitAgent = this._onDidSubmitAgent.event;

	private _onDidChangeAgent = this._register(new Emitter<{ agent: IChatAgentData; slashCommand?: IChatAgentCommand }>());
	readonly onDidChangeAgent = this._onDidChangeAgent.event;

	private _onDidFocus = this._register(new Emitter<void>());
	readonly onDidFocus = this._onDidFocus.event;

	private _onDidChangeViewModel = this._register(new Emitter<IChatWidgetViewModelChangeEvent>());
	readonly onDidChangeViewModel = this._onDidChangeViewModel.event;

	private _onDidScroll = this._register(new Emitter<void>());
	readonly onDidScroll = this._onDidScroll.event;

	private _onDidAcceptInput = this._register(new Emitter<void>());
	readonly onDidAcceptInput = this._onDidAcceptInput.event;

	private _onDidHide = this._register(new Emitter<void>());
	readonly onDidHide = this._onDidHide.event;

	private _onDidShow = this._register(new Emitter<void>());
	readonly onDidShow = this._onDidShow.event;

	private _onDidChangeParsedInput = this._register(new Emitter<void>());
	readonly onDidChangeParsedInput = this._onDidChangeParsedInput.event;

	private _onDidChangeActiveInputEditor = this._register(new Emitter<void>());
	readonly onDidChangeActiveInputEditor = this._onDidChangeActiveInputEditor.event;

	private readonly _onWillMaybeChangeHeight = this._register(new Emitter<void>());
	readonly onWillMaybeChangeHeight: Event<void> = this._onWillMaybeChangeHeight.event;

	private _onDidChangeHeight = this._register(new Emitter<number>());
	readonly onDidChangeHeight = this._onDidChangeHeight.event;

	private readonly _onDidChangeContentHeight = this._register(new Emitter<void>());
	readonly onDidChangeContentHeight: Event<void> = this._onDidChangeContentHeight.event;

	private _onDidChangeEmptyState = this._register(new Emitter<void>());
	readonly onDidChangeEmptyState = this._onDidChangeEmptyState.event;

	contribs: ReadonlyArray<IChatWidgetContrib> = [];

	private listContainer!: HTMLElement;
	private container!: HTMLElement;

	get domNode() { return this.container; }

	private listWidget!: ChatListWidget;
	private readonly _codeBlockModelCollection: CodeBlockModelCollection;
	private inputPartMaxHeightOverride: number | undefined;

	private readonly visibilityTimeoutDisposable: MutableDisposable<IDisposable> = this._register(new MutableDisposable());
	private readonly visibilityAnimationFrameDisposable: MutableDisposable<IDisposable> = this._register(new MutableDisposable());

	private readonly inputPartDisposable: MutableDisposable<ChatInputPart> = this._register(new MutableDisposable());
	private readonly inlineInputPartDisposable: MutableDisposable<ChatInputPart> = this._register(new MutableDisposable());
	private inputContainer!: HTMLElement;
	private focusedInputDOM!: HTMLElement;
	private editorOptions!: ChatEditorOptions;

	private recentlyRestoredCheckpoint: boolean = false;
	private _planningPhase: PlanningPhase = 'broad-scan';
	private _planningTransitionContext: IPlanningTransitionContext | undefined;
	private _lastPlanningQuestionModelId: string | undefined;
	private _lastPlanningQuestionSourceInput: string | undefined;
	private _goalClarityQuestionRounds = 0;
	private _pendingPlanningQuestionResolveId: string | undefined;
	private _pendingPlanningPlaceholderRequestId: string | undefined;
	private _planningTranscriptScrollPreservationDepth = 0;
	private _currentPlanningPlanRequestId: string | undefined;
	private _previousPlanningPlanRequestId: string | undefined;
	private readonly _planningPlanTextByRequestId = new Map<string, string>();
	private readonly _pendingPlanningQuestionAnswersListener = this._register(new MutableDisposable<IDisposable>());
	private readonly _pendingPlanningResponseListener = this._register(new MutableDisposable<IDisposable>());
	private _skipDynamicPlanningQuestionsOnce = false;

	private welcomeMessageContainer!: HTMLElement;
	private readonly welcomePart: MutableDisposable<ChatViewWelcomePart> = this._register(new MutableDisposable());

	private readonly _gettingStartedTipPart = this._register(new MutableDisposable<DisposableStore>());
	private _gettingStartedTipPartRef: ChatTipContentPart | undefined;

	private readonly chatSuggestNextWidget: ChatSuggestNextWidget;

	private bodyDimension: dom.Dimension | undefined;
	private visibleChangeCount = 0;
	private requestInProgress: IContextKey<boolean>;
	private agentInInput: IContextKey<boolean>;

	private _visible = false;
	get visible() { return this._visible; }

	private _instructionFilesCheckPromise: Promise<boolean> | undefined;
	private _instructionFilesExist: boolean | undefined;

	private _isRenderingWelcome = false;

	// Coding agent locking state
	private _lockedAgent?: {
		id: string;
		name: string;
		prefix: string;
		displayName: string;
	};
	private readonly _lockedToCodingAgentContextKey: IContextKey<boolean>;
	private readonly _lockedCodingAgentIdContextKey: IContextKey<string>;
	private readonly _chatSessionSupportsForkContextKey: IContextKey<boolean>;
	private readonly _agentSupportsAttachmentsContextKey: IContextKey<boolean>;
	private readonly _sessionIsEmptyContextKey: IContextKey<boolean>;
	private readonly _hasPendingRequestsContextKey: IContextKey<boolean>;
	private readonly _sessionHasDebugDataContextKey: IContextKey<boolean>;
	private _attachmentCapabilities: IChatAgentAttachmentCapabilities = supportsAllAttachments;

	private readonly viewModelDisposables = this._register(new DisposableStore());
	private _viewModel: ChatViewModel | undefined;

	private set viewModel(viewModel: ChatViewModel | undefined) {
		if (this._viewModel === viewModel) {
			return;
		}

		const previousSessionResource = this._viewModel?.sessionResource;
		if (!isEqual(previousSessionResource, viewModel?.sessionResource)) {
			this._planningPhase = 'broad-scan';
			this._planningTransitionContext = undefined;
			this._lastPlanningQuestionModelId = undefined;
			this._lastPlanningQuestionSourceInput = undefined;
			this._goalClarityQuestionRounds = 0;
			this._pendingPlanningQuestionResolveId = undefined;
			this._pendingPlanningPlaceholderRequestId = undefined;
			this._currentPlanningPlanRequestId = undefined;
			this._previousPlanningPlanRequestId = undefined;
			this._planningPlanTextByRequestId.clear();
			this._pendingPlanningQuestionAnswersListener.clear();
			this._pendingPlanningResponseListener.clear();
			this._skipDynamicPlanningQuestionsOnce = false;
		}
		this.viewModelDisposables.clear();

		this._viewModel = viewModel;
		if (viewModel) {
			this.viewModelDisposables.add(viewModel);
			this.logService.debug('ChatWidget#setViewModel: have viewModel');

			// If switching to a model with a request in progress, play progress sound
			if (viewModel.model.requestInProgress.get()) {
				this.chatAccessibilityService.acceptRequest(viewModel.sessionResource, true);
			}
		} else {
			this.logService.debug('ChatWidget#setViewModel: no viewModel');
		}

		this._onDidChangeViewModel.fire({ previousSessionResource, currentSessionResource: this._viewModel?.sessionResource });
	}

	get viewModel() {
		return this._viewModel;
	}

	private readonly _editingSession = observableValue<IChatEditingSession | undefined>(this, undefined);
	private readonly _viewModelObs = observableFromEvent(this, this.onDidChangeViewModel, () => this.viewModel);

	private parsedChatRequest: IParsedChatRequest | undefined;
	get parsedInput() {
		if (this.parsedChatRequest === undefined) {
			if (!this.viewModel) {
				return { text: '', parts: [] };
			}

			this.parsedChatRequest = this.instantiationService.createInstance(ChatRequestParser)
				.parseChatRequestWithReferences(getDynamicVariablesForWidget(this), getSelectedToolAndToolSetsForWidget(this), this.getInput(), this.location, {
					selectedAgent: this._lastSelectedAgent,
					mode: this.input.currentModeKind,
					attachmentCapabilities: this.attachmentCapabilities,
					forcedAgent: this._lockedAgent?.id ? this.chatAgentService.getAgent(this._lockedAgent.id) : undefined
				});
			this._onDidChangeParsedInput.fire();
		}

		return this.parsedChatRequest;
	}

	get scopedContextKeyService(): IContextKeyService {
		return this.contextKeyService;
	}

	private readonly _location: IChatWidgetLocationOptions;
	get location() {
		return this._location.location;
	}

	readonly viewContext: IChatWidgetViewContext;

	get supportsChangingModes(): boolean {
		return !!this.viewOptions.supportsChangingModes;
	}

	get locationData() {
		return this._location.resolveData?.();
	}

	constructor(
		location: ChatAgentLocation | IChatWidgetLocationOptions,
		viewContext: IChatWidgetViewContext | undefined,
		private readonly viewOptions: IChatWidgetViewOptions,
		private readonly styles: IChatWidgetStyles,
		@ICodeEditorService private readonly codeEditorService: ICodeEditorService,
		@IConfigurationService private readonly configurationService: IConfigurationService,
		@IDialogService private readonly dialogService: IDialogService,
		@ICommandService private readonly commandService: ICommandService,
		@IContextKeyService private readonly contextKeyService: IContextKeyService,
		@IInstantiationService private readonly instantiationService: IInstantiationService,
		@IChatService private readonly chatService: IChatService,
		@IChatAgentService private readonly chatAgentService: IChatAgentService,
		@IChatWidgetService private readonly chatWidgetService: IChatWidgetService,
		@IChatAccessibilityService private readonly chatAccessibilityService: IChatAccessibilityService,
		@ILogService private readonly logService: ILogService,
		@IThemeService private readonly themeService: IThemeService,
		@IChatSlashCommandService private readonly chatSlashCommandService: IChatSlashCommandService,
		@IChatEditingService chatEditingService: IChatEditingService,
		@ITelemetryService private readonly telemetryService: ITelemetryService,
		@IPromptsService private readonly promptsService: IPromptsService,
		@ILanguageModelToolsService private readonly toolsService: ILanguageModelToolsService,
		@IChatModeService private readonly chatModeService: IChatModeService,
		@IChatLayoutService private readonly chatLayoutService: IChatLayoutService,
		@IChatEntitlementService private readonly chatEntitlementService: IChatEntitlementService,
		@IChatSessionsService private readonly chatSessionsService: IChatSessionsService,
		@IAgentSessionsService private readonly agentSessionsService: IAgentSessionsService,
		@IChatTodoListService private readonly chatTodoListService: IChatTodoListService,
		@IChatArtifactsService private readonly chatArtifactsService: IChatArtifactsService,
		@ILifecycleService private readonly lifecycleService: ILifecycleService,
		@IChatAttachmentResolveService private readonly chatAttachmentResolveService: IChatAttachmentResolveService,
		@IChatTipService private readonly chatTipService: IChatTipService,
		@IChatDebugService private readonly chatDebugService: IChatDebugService,
		@ILanguageModelsService private readonly languageModelsService: ILanguageModelsService,
		@IFileService private readonly fileService: IFileService,
		@ITextModelService private readonly textModelService: ITextModelService,
		@IWorkspaceContextService private readonly workspaceContextService: IWorkspaceContextService,
		@ILanguageFeaturesService private readonly languageFeaturesService: ILanguageFeaturesService,
	) {
		super();

		this._lockedToCodingAgentContextKey = ChatContextKeys.lockedToCodingAgent.bindTo(this.contextKeyService);
		this._lockedCodingAgentIdContextKey = ChatContextKeys.lockedCodingAgentId.bindTo(this.contextKeyService);
		this._chatSessionSupportsForkContextKey = ChatContextKeys.chatSessionSupportsFork.bindTo(this.contextKeyService);
		this._agentSupportsAttachmentsContextKey = ChatContextKeys.agentSupportsAttachments.bindTo(this.contextKeyService);
		this._sessionIsEmptyContextKey = ChatContextKeys.chatSessionIsEmpty.bindTo(this.contextKeyService);
		this._hasPendingRequestsContextKey = ChatContextKeys.hasPendingRequests.bindTo(this.contextKeyService);
		this._sessionHasDebugDataContextKey = ChatContextKeys.chatSessionHasDebugData.bindTo(this.contextKeyService);

		this._register(this.chatDebugService.onDidAddEvent(e => {
			const sessionResource = this.viewModel?.sessionResource;
			if (sessionResource && e.sessionResource.toString() === sessionResource.toString()) {
				this._sessionHasDebugDataContextKey.set(true);
			}
		}));

		this.viewContext = viewContext ?? {};

		const viewModelObs = this._viewModelObs;

		if (typeof location === 'object') {
			this._location = location;
		} else {
			this._location = { location };
		}

		ChatContextKeys.inChatSession.bindTo(contextKeyService).set(true);
		ChatContextKeys.location.bindTo(contextKeyService).set(this._location.location);
		ChatContextKeys.inQuickChat.bindTo(contextKeyService).set(isQuickChat(this));
		this.agentInInput = ChatContextKeys.inputHasAgent.bindTo(contextKeyService);
		this.requestInProgress = ChatContextKeys.requestInProgress.bindTo(contextKeyService);

		this._register(this.chatEntitlementService.onDidChangeAnonymous(() => this.renderWelcomeViewContentIfNeeded()));

		this._register(this.configurationService.onDidChangeConfiguration(e => {
			if (e.affectsConfiguration('chat.tips.enabled')) {
				if (!this.configurationService.getValue<boolean>('chat.tips.enabled')) {
					// Clear the existing tip so it doesn't linger
					if (this.inputPart) {
						this._gettingStartedTipPartRef = undefined;
						this._gettingStartedTipPart.clear();
						const tipContainer = this.inputPart.gettingStartedTipContainerElement;
						dom.clearNode(tipContainer);
						dom.setVisibility(false, tipContainer);
					}
				} else {
					this.updateChatViewVisibility();
				}
			}
		}));

		this._register(bindContextKey(decidedChatEditingResourceContextKey, contextKeyService, (reader) => {
			const currentSession = this._editingSession.read(reader);
			if (!currentSession) {
				return;
			}
			const entries = currentSession.entries.read(reader);
			const decidedEntries = entries.filter(entry => entry.state.read(reader) !== ModifiedFileEntryState.Modified);
			return decidedEntries.map(entry => entry.entryId);
		}));
		this._register(bindContextKey(hasUndecidedChatEditingResourceContextKey, contextKeyService, (reader) => {
			const currentSession = this._editingSession.read(reader);
			const entries = currentSession?.entries.read(reader) ?? []; // using currentSession here
			const decidedEntries = entries.filter(entry => entry.state.read(reader) === ModifiedFileEntryState.Modified);
			return decidedEntries.length > 0;
		}));
		this._register(bindContextKey(hasAppliedChatEditsContextKey, contextKeyService, (reader) => {
			const currentSession = this._editingSession.read(reader);
			if (!currentSession) {
				return false;
			}
			const entries = currentSession.entries.read(reader);
			return entries.length > 0;
		}));
		this._register(bindContextKey(inChatEditingSessionContextKey, contextKeyService, (reader) => {
			return this._editingSession.read(reader) !== null;
		}));
		this._register(bindContextKey(ChatContextKeys.chatEditingCanUndo, contextKeyService, (r) => {
			return this._editingSession.read(r)?.canUndo.read(r) || false;
		}));
		this._register(bindContextKey(ChatContextKeys.chatEditingCanRedo, contextKeyService, (r) => {
			return this._editingSession.read(r)?.canRedo.read(r) || false;
		}));
		this._register(bindContextKey(applyingChatEditsFailedContextKey, contextKeyService, (r) => {
			const chatModel = viewModelObs.read(r)?.model;
			const editingSession = this._editingSession.read(r);
			if (!editingSession || !chatModel) {
				return false;
			}
			const lastResponse = observableFromEvent(this, chatModel.onDidChange, () => chatModel.getRequests().at(-1)?.response).read(r);
			return lastResponse?.result?.errorDetails && !lastResponse?.result?.errorDetails.responseIsIncomplete;
		}));

		this._codeBlockModelCollection = this._register(instantiationService.createInstance(CodeBlockModelCollection, undefined));
		this.chatSuggestNextWidget = this._register(this.instantiationService.createInstance(ChatSuggestNextWidget));

		this._register(autorun(r => {
			const viewModel = viewModelObs.read(r);
			const sessions = chatEditingService.editingSessionsObs.read(r);

			const session = sessions.find(candidate => isEqual(candidate.chatSessionResource, viewModel?.sessionResource));
			this._editingSession.set(undefined, undefined);
			this.renderChatEditingSessionState(); // this is necessary to make sure we dispose previous buttons, etc.

			if (!session) {
				// none or for a different chat widget
				return;
			}

			const entries = session.entries.read(r);
			for (const entry of entries) {
				entry.state.read(r); // SIGNAL
			}

			this._editingSession.set(session, undefined);

			r.store.add(session.onDidDispose(() => {
				this._editingSession.set(undefined, undefined);
				this.renderChatEditingSessionState();
			}));
			r.store.add(this.inputEditor.onDidChangeModelContent(() => {
				if (this.getInput() === '') {
					this.refreshParsedInput();
				}
			}));
			this.renderChatEditingSessionState();
		}));

		this._register(this.codeEditorService.registerCodeEditorOpenHandler(async (input: ITextResourceEditorInput, _source: ICodeEditor | null, _sideBySide?: boolean): Promise<ICodeEditor | null> => {
			const resource = input.resource;
			if (resource.scheme !== Schemas.vscodeChatCodeBlock) {
				return null;
			}

			const responseId = resource.path.split('/').at(1);
			if (!responseId) {
				return null;
			}

			const item = this.viewModel?.getItems().find(item => item.id === responseId);
			if (!item) {
				return null;
			}

			// TODO: needs to reveal the chat view

			this.reveal(item);

			await timeout(0); // wait for list to actually render

			for (const codeBlockPart of this.listWidget.editorsInUse()) {
				if (extUri.isEqual(codeBlockPart.uri, resource, true)) {
					const editor = codeBlockPart.editor;

					let relativeTop = 0;
					const editorDomNode = editor.getDomNode();
					if (editorDomNode) {
						const row = dom.findParentWithClass(editorDomNode, 'monaco-list-row');
						if (row) {
							relativeTop = dom.getTopLeftOffset(editorDomNode).top - dom.getTopLeftOffset(row).top;
						}
					}

					if (input.options?.selection) {
						const editorSelectionTopOffset = editor.getTopForPosition(input.options.selection.startLineNumber, input.options.selection.startColumn);
						relativeTop += editorSelectionTopOffset;

						editor.focus();
						editor.setSelection({
							startLineNumber: input.options.selection.startLineNumber,
							startColumn: input.options.selection.startColumn,
							endLineNumber: input.options.selection.endLineNumber ?? input.options.selection.startLineNumber,
							endColumn: input.options.selection.endColumn ?? input.options.selection.startColumn
						});
					}

					this.reveal(item, relativeTop);

					return editor;
				}
			}
			return null;
		}));

		this._register(this.onDidChangeParsedInput(() => this.updateChatInputContext()));

		this._register(this.chatTodoListService.onDidUpdateTodos((sessionResource) => {
			if (isEqual(this.viewModel?.sessionResource, sessionResource)) {
				this.inputPart.renderChatTodoListWidget(sessionResource);
			}
		}));

		this._register(this.chatArtifactsService.onDidUpdateArtifacts((sessionResource) => {
			if (isEqual(this.viewModel?.sessionResource, sessionResource)) {
				this.inputPart.renderArtifactsWidget(sessionResource);
			}
		}));
	}

	private _lastSelectedAgent: IChatAgentData | undefined;
	set lastSelectedAgent(agent: IChatAgentData | undefined) {
		this.parsedChatRequest = undefined;
		this._lastSelectedAgent = agent;
		this._updateAgentCapabilitiesContextKeys(agent);
		this._onDidChangeParsedInput.fire();
	}

	get lastSelectedAgent(): IChatAgentData | undefined {
		return this._lastSelectedAgent;
	}

	private _updateAgentCapabilitiesContextKeys(agent: IChatAgentData | undefined): void {
		// Check if the agent has capabilities defined directly
		const capabilities = agent?.capabilities ?? (this._lockedAgent ? this.chatSessionsService.getCapabilitiesForSessionType(this._lockedAgent.id) : undefined);
		this._attachmentCapabilities = capabilities ?? supportsAllAttachments;

		const supportsAttachments = Object.keys(filter(this._attachmentCapabilities, (key, value) => value === true)).length > 0;
		this._agentSupportsAttachmentsContextKey.set(supportsAttachments);
	}

	get supportsFileReferences(): boolean {
		return !!this.viewOptions.supportsFileReferences;
	}

	get attachmentCapabilities(): IChatAgentAttachmentCapabilities {
		return this._attachmentCapabilities;
	}

	/**
	 * Either the inline input (when editing) or the main input part
	 */
	get input(): ChatInputPart {
		return this.viewModel?.editing && this.configurationService.getValue<string>('chat.editRequests') !== 'input' ? this.inlineInputPart : this.inputPart;
	}

	/**
	 * The main input part at the buttom of the chat widget. Use `input` to get the active input (main or inline editing part).
	 */
	get inputPart(): ChatInputPart {
		return this.inputPartDisposable.value!;
	}

	private get inlineInputPart(): ChatInputPart {
		return this.inlineInputPartDisposable.value!;
	}

	get inputEditor(): ICodeEditor {
		return this.input.inputEditor;
	}

	get contentHeight(): number {
		return this.input.height.get() + this.listWidget.contentHeight + this.chatSuggestNextWidget.height;
	}

	get scrollTop(): number {
		return this.listWidget.scrollTop;
	}

	set scrollTop(value: number) {
		this.listWidget.scrollTop = value;
	}

	get attachmentModel(): ChatAttachmentModel {
		return this.input.attachmentModel;
	}

	render(parent: HTMLElement): void {
		const viewId = isIChatViewViewContext(this.viewContext) ? this.viewContext.viewId : undefined;
		this.editorOptions = this._register(this.instantiationService.createInstance(ChatEditorOptions, viewId, this.styles.listForeground, this.styles.inputEditorBackground, this.styles.resultEditorBackground));
		const renderInputOnTop = this.viewOptions.renderInputOnTop ?? false;
		const renderFollowups = this.viewOptions.renderFollowups ?? !renderInputOnTop;
		const renderStyle = this.viewOptions.renderStyle;
		const renderInputToolbarBelowInput = this.viewOptions.renderInputToolbarBelowInput ?? false;

		this.container = dom.append(parent, $('.interactive-session'));
		this.welcomeMessageContainer = dom.append(this.container, $('.chat-welcome-view-container', { style: 'display: none' }));
		this._register(dom.addStandardDisposableListener(this.welcomeMessageContainer, dom.EventType.CLICK, () => this.focusInput()));

		this._register(this.chatSuggestNextWidget.onDidChangeHeight(() => {
			if (this.bodyDimension) {
				this.layout(this.bodyDimension.height, this.bodyDimension.width);
			}
		}));
		this._register(this.chatSuggestNextWidget.onDidSelectPrompt(({ handoff, agentId, withAutopilot }) => {
			this.handleNextPromptSelection(handoff, agentId, withAutopilot);
		}));

		if (renderInputOnTop) {
			this.createInput(this.container, { renderFollowups, renderStyle, renderInputToolbarBelowInput });
			this.listContainer = dom.append(this.container, $(`.interactive-list`));
		} else {
			this.listContainer = dom.append(this.container, $(`.interactive-list`));
			dom.append(this.container, this.chatSuggestNextWidget.domNode);
			this.createInput(this.container, { renderFollowups, renderStyle, renderInputToolbarBelowInput });
		}

		this.renderWelcomeViewContentIfNeeded();
		this.createList(this.listContainer, { editable: !isInlineChat(this) && !isQuickChat(this), ...this.viewOptions.rendererOptions, renderStyle });

		// Forward scroll events from the parent container margins (outside the max-width area) to the chat list
		this._register(dom.addDisposableListener(parent, dom.EventType.MOUSE_WHEEL, (e: IMouseWheelEvent) => {
			if (e.defaultPrevented) {
				return;
			}

			if (dom.isAncestor(e.target as Node | null, this.container)) {
				return;
			}

			this.listWidget.delegateScrollFromMouseWheelEvent(e);
		}));

		// Update the font family and size
		this._register(autorun(reader => {
			const fontFamily = this.chatLayoutService.fontFamily.read(reader);
			const fontSize = this.chatLayoutService.fontSize.read(reader);

			this.container.style.setProperty('--vscode-chat-font-family', fontFamily);
			this.container.style.fontSize = `${fontSize}px`;

			if (this.visible) {
				this.listWidget.rerender();
			}
		}));

		this._register(Event.runAndSubscribe(this.editorOptions.onDidChange, () => this.onDidStyleChange()));

		// Do initial render
		if (this.viewModel) {
			this.onDidChangeItems();
			this.listWidget.scrollToEnd();
		}

		this.contribs = ChatWidget.CONTRIBS.map(contrib => {
			try {
				return this._register(this.instantiationService.createInstance(contrib, this));
			} catch (err) {
				this.logService.error('Failed to instantiate chat widget contrib', toErrorMessage(err));
				return undefined;
			}
		}).filter(isDefined);

		this._register(this.chatWidgetService.register(this));

		const parsedInput = observableFromEvent(this.onDidChangeParsedInput, () => this.parsedInput);
		this._register(autorun(r => {
			const input = parsedInput.read(r);

			const newPromptAttachments = new Map<string, IChatRequestVariableEntry>();
			const oldPromptAttachments = new Set<string>();

			// get all attachments, know those that are prompt-referenced
			for (const attachment of this.attachmentModel.attachments) {
				if (attachment.range) {
					oldPromptAttachments.add(attachment.id);
				}
			}

			// update/insert prompt-referenced attachments
			for (const part of input.parts) {
				if (part instanceof ChatRequestToolPart || part instanceof ChatRequestToolSetPart || part instanceof ChatRequestDynamicVariablePart) {
					const entry = part.toVariableEntry();
					newPromptAttachments.set(entry.id, entry);
					oldPromptAttachments.delete(entry.id);
				}
			}

			this.attachmentModel.updateContext(oldPromptAttachments, newPromptAttachments.values());
		}));

		if (!this.focusedInputDOM) {
			this.focusedInputDOM = this.container.appendChild(dom.$('.focused-input-dom'));
		}
	}

	focusInput(): void {
		this.input.focus();

		// Sometimes focusing the input part is not possible,
		// but we'd like to be the last focused chat widget,
		// so we emit an optimistic onDidFocus event nonetheless.
		this._onDidFocus.fire();
	}

	focusTodosView(): boolean {
		if (!this.input.hasVisibleTodos()) {
			return false;
		}

		return this.input.focusTodoList();
	}

	toggleTodosViewFocus(): boolean {
		if (!this.input.hasVisibleTodos()) {
			return false;
		}

		if (this.input.isTodoListFocused()) {
			this.focusInput();
			return true;
		}

		return this.input.focusTodoList();
	}

	focusQuestionCarousel(): boolean {
		if (!this.input.questionCarousel) {
			return false;
		}

		return this.input.focusQuestionCarousel();
	}

	toggleQuestionCarouselFocus(): boolean {
		if (!this.input.questionCarousel) {
			return false;
		}

		if (this.input.isQuestionCarouselFocused()) {
			this.focusInput();
			return true;
		}

		return this.input.focusQuestionCarousel();
	}

	navigateToPreviousQuestion(): boolean {
		if (!this.input.questionCarousel) {
			return false;
		}

		return this.input.navigateToPreviousQuestion();
	}

	navigateToNextQuestion(): boolean {
		if (!this.input.questionCarousel) {
			return false;
		}

		return this.input.navigateToNextQuestion();
	}

	async refinePlan(): Promise<boolean> {
		const questionStage: PlanningQuestionStage = this.getCurrentPlanningResponseText() ? 'task-decomposition' : 'goal-clarity';
		return this.triggerDynamicPlanningQuestions(this.getCurrentPlanningInput(), { storeToHistory: false }, { forceRegenerate: true, phase: this._planningPhase, questionStage });
	}

	async retreatPlanPhase(): Promise<boolean> {
		const previousPhase = getPreviousPlanningPhase(this._planningPhase);
		if (!previousPhase) {
			return false;
		}

		const questionStage: PlanningQuestionStage = this.getCurrentPlanningResponseText() ? 'task-decomposition' : 'goal-clarity';
		return this.triggerDynamicPlanningQuestions(this.getCurrentPlanningInput(), { storeToHistory: false }, { forceRegenerate: true, phase: previousPhase, questionStage });
	}

	async advancePlanPhase(): Promise<boolean> {
		const nextPhase = getNextPlanningPhase(this._planningPhase);
		if (!nextPhase) {
			return false;
		}

		const questionStage: PlanningQuestionStage = this.getCurrentPlanningResponseText() ? 'task-decomposition' : 'goal-clarity';
		return this.triggerDynamicPlanningQuestions(this.getCurrentPlanningInput(), { storeToHistory: false }, { forceRegenerate: true, phase: nextPhase, questionStage });
	}

	toggleTipFocus(): boolean {
		if (this._gettingStartedTipPartRef?.hasFocus()) {
			this.focusInput();
			return true;
		}

		if (!this._gettingStartedTipPartRef) {
			return false;
		}
		this._gettingStartedTipPartRef.focus();
		return true;
	}

	hasInputFocus(): boolean {
		return this.input.hasFocus();
	}

	refreshParsedInput() {
		if (!this.viewModel) {
			return;
		}

		const previous = this.parsedChatRequest;
		this.parsedChatRequest = this.instantiationService.createInstance(ChatRequestParser).parseChatRequestWithReferences(getDynamicVariablesForWidget(this), getSelectedToolAndToolSetsForWidget(this), this.getInput(), this.location, { selectedAgent: this._lastSelectedAgent, mode: this.input.currentModeKind, attachmentCapabilities: this.attachmentCapabilities });
		if (!previous || !IParsedChatRequest.equals(previous, this.parsedChatRequest)) {
			this._onDidChangeParsedInput.fire();
		}
	}

	getSibling(item: ChatTreeItem, type: 'next' | 'previous'): ChatTreeItem | undefined {
		if (!isResponseVM(item)) {
			return;
		}
		const items = this.viewModel?.getItems();
		if (!items) {
			return;
		}
		const responseItems = items.filter(i => isResponseVM(i));
		const targetIndex = responseItems.indexOf(item);
		if (targetIndex === undefined) {
			return;
		}
		const indexToFocus = type === 'next' ? targetIndex + 1 : targetIndex - 1;
		if (indexToFocus < 0 || indexToFocus > responseItems.length - 1) {
			return;
		}
		return responseItems[indexToFocus];
	}

	async clear(): Promise<void> {
		this.logService.debug('ChatWidget#clear');
		if (this._dynamicMessageLayoutData) {
			this._dynamicMessageLayoutData.enabled = true;
		}

		if (this.viewModel?.editing) {
			this.finishedEditing();
		}

		if (this.viewModel) {
			this.viewModel.resetInputPlaceholder();
		}
		if (this._lockedAgent) {
			this.lockToCodingAgent(this._lockedAgent.name, this._lockedAgent.displayName, this._lockedAgent.id);
		} else {
			this.unlockFromCodingAgent();
		}

		this.inputPart.clearTodoListWidget(this.viewModel?.sessionResource, true);
		this.inputPart.clearArtifactsWidget();
		this.chatSuggestNextWidget.hide();
		await this.viewOptions.clear?.();
	}

	private onDidChangeItems(skipDynamicLayout?: boolean) {
		if (this._visible || !this.viewModel) {
			const items = this.viewModel?.getItems() ?? [];

			if (items.length > 0) {
				this.updateChatViewVisibility();
			} else {
				this.renderWelcomeViewContentIfNeeded();
			}

			this._onWillMaybeChangeHeight.fire();

			// Update list widget state and refresh
			this.listWidget.setVisibleChangeCount(this.visibleChangeCount);
			this.listWidget.refresh();

			if (!skipDynamicLayout && this._dynamicMessageLayoutData) {
				this.layoutDynamicChatTreeItemMode();
			}

			this.renderFollowups();
		}
	}

	/**
	 * Updates the DOM visibility of welcome view and chat list immediately
	 */
	private updateChatViewVisibility(): void {
		if (this.viewModel) {
			const isStandardLayout = this.viewOptions.renderStyle !== 'compact' && this.viewOptions.renderStyle !== 'minimal';
			const numItems = this.viewModel.getItems().length;
			dom.setVisibility(numItems === 0, this.welcomeMessageContainer);
			dom.setVisibility(numItems !== 0, this.listContainer);

			// Show/hide the getting-started tip container based on empty state.
			// Only use this in the standard chat layout where the welcome view is shown.
			if (isStandardLayout && this.inputPart) {
				const tipContainer = this.inputPart.gettingStartedTipContainerElement;
				if (numItems === 0) {
					this.renderGettingStartedTipIfNeeded();
				} else {
					// Dispose the cached tip part so the next empty state picks a
					// fresh (rotated) tip instead of re-showing the stale one.
					this._gettingStartedTipPartRef = undefined;
					this._gettingStartedTipPart.clear();
					dom.clearNode(tipContainer);
					dom.setVisibility(false, tipContainer);
				}
			}
		}

		// Only show welcome getting started until extension is installed
		this.container.classList.toggle('chat-view-getting-started-disabled', this.chatEntitlementService.sentiment.installed);

		this._onDidChangeEmptyState.fire();
	}

	isEmpty(): boolean {
		return (this.viewModel?.getItems().length ?? 0) === 0;
	}

	/**
	 * Renders the welcome view content when needed.
	 */
	private renderWelcomeViewContentIfNeeded() {
		if (this._isRenderingWelcome) {
			return;
		}

		this._isRenderingWelcome = true;
		try {
			if (this.viewOptions.renderStyle === 'compact' || this.viewOptions.renderStyle === 'minimal' || this.lifecycleService.willShutdown) {
				return;
			}

			const numItems = this.viewModel?.getItems().length ?? 0;
			if (!numItems) {
				const defaultAgent = this.chatAgentService.getDefaultAgent(this.location, this.input.currentModeKind);
				let additionalMessage: string | IMarkdownString | undefined;
				if (this.chatEntitlementService.anonymous && !this.chatEntitlementService.sentiment.installed) {
					const providers = product.defaultChatAgent.provider;
					additionalMessage = new MarkdownString(localize({ key: 'settings', comment: ['{Locked="]({2})"}', '{Locked="]({3})"}'] }, "By continuing with {0} Copilot, you agree to {1}'s [Terms]({2}) and [Privacy Statement]({3}).", providers.default.name, providers.default.name, product.defaultChatAgent.termsStatementUrl, product.defaultChatAgent.privacyStatementUrl), { isTrusted: true });
				} else {
					additionalMessage = defaultAgent?.metadata.additionalWelcomeMessage;
				}
				if (!additionalMessage && !this._lockedAgent) {
					additionalMessage = this._getGenerateInstructionsMessage();
				}
				const welcomeContent = this.getWelcomeViewContent(additionalMessage);
				if (!this.welcomePart.value || this.welcomePart.value.needsRerender(welcomeContent)) {
					dom.clearNode(this.welcomeMessageContainer);

					this.welcomePart.value = this.instantiationService.createInstance(
						ChatViewWelcomePart,
						welcomeContent,
						{
							location: this.location,
							isWidgetAgentWelcomeViewContent: this.input?.currentModeKind === ChatModeKind.Agent
						}
					);
					dom.append(this.welcomeMessageContainer, this.welcomePart.value.element);
				}
			}

			this.updateChatViewVisibility();
		} finally {
			this._isRenderingWelcome = false;
		}
	}

	private renderGettingStartedTipIfNeeded(): void {
		if (!this.inputPart) {
			return;
		}

		const tipContainer = this.inputPart.gettingStartedTipContainerElement;

		const tip = this.chatTipService.getWelcomeTip(this.contextKeyService);
		if (!tip) {
			if (this._gettingStartedTipPart.value) {
				this._gettingStartedTipPartRef = undefined;
				this._gettingStartedTipPart.clear();
				dom.clearNode(tipContainer);
			}
			dom.setVisibility(false, tipContainer);
			return;
		}

		// Already showing an eligible tip
		if (this._gettingStartedTipPart.value) {
			dom.setVisibility(true, tipContainer);
			return;
		}

		const store = new DisposableStore();
		const renderer = this.instantiationService.createInstance(ChatContentMarkdownRenderer);
		const tipPart = store.add(this.instantiationService.createInstance(ChatTipContentPart,
			tip,
			renderer,
		));
		this._gettingStartedTipPartRef = tipPart;

		store.add(tipPart.onDidHide(() => {
			tipPart.domNode.remove();
			this._gettingStartedTipPartRef = undefined;
			this._gettingStartedTipPart.clear();
			dom.setVisibility(false, tipContainer);
			this.focusInput();
		}));

		// Set the guard before appending to DOM so that any re-entrant calls
		// triggered by context-key changes during construction see the guard
		// and return early without adding a duplicate tip node.
		this._gettingStartedTipPart.value = store;
		// Clear any stale nodes left from a previous tip that was not properly
		// removed (e.g. if re-entrancy bypassed the guard above).
		dom.clearNode(tipContainer);
		tipContainer.appendChild(tipPart.domNode);
		dom.setVisibility(true, tipContainer);
	}

	private _getGenerateInstructionsMessage(): IMarkdownString {
		// Start checking for instruction files immediately if not already done
		if (!this._instructionFilesCheckPromise) {
			this._instructionFilesCheckPromise = this._checkForAgentInstructionFiles();
			// Use VS Code's idiomatic pattern for disposal-safe promise callbacks
			this._register(thenIfNotDisposed(this._instructionFilesCheckPromise, hasFiles => {
				this._instructionFilesExist = hasFiles;
				// Only re-render if the current view still doesn't have items and we're showing the welcome message
				const hasViewModelItems = this.viewModel?.getItems().length ?? 0;
				if (hasViewModelItems === 0) {
					this.renderWelcomeViewContentIfNeeded();
				}
			}));
		}

		// If we already know the result, use it
		if (this._instructionFilesExist === true) {
			// Don't show generate instructions message if files exist
			return new MarkdownString('');
		} else if (this._instructionFilesExist === false) {
			// Show generate instructions message if no files exist
			return new MarkdownString(localize(
				'chatWidget.instructions',
				"[Generate Agent Instructions]({0}) to onboard AI onto your codebase.",
				`command:${GENERATE_AGENT_INSTRUCTIONS_COMMAND_ID}`
			), { isTrusted: { enabledCommands: [GENERATE_AGENT_INSTRUCTIONS_COMMAND_ID] } });
		}

		// While checking, don't show the generate instructions message
		return new MarkdownString('');
	}

	/**
	 * Checks if any agent instruction files (.github/copilot-instructions.md or AGENTS.md) exist in the workspace.
	 * Used to determine whether to show the "Generate Agent Instructions" hint.
	 *
	 * @returns true if instruction files exist OR if instruction features are disabled (to hide the hint)
	 */
	private async _checkForAgentInstructionFiles(): Promise<boolean> {
		try {
			return (await this.promptsService.listAgentInstructions(CancellationToken.None)).length > 0;
		} catch (error) {
			// On error, assume no instruction files exist to be safe
			this.logService.warn('[ChatWidget] Error checking for instruction files:', error);
			return false;
		}
	}

	private getWelcomeViewContent(additionalMessage: string | IMarkdownString | undefined): IChatViewWelcomeContent {
		if (this.isLockedToCodingAgent) {
			// Check for provider-specific customizations from chat sessions service
			const contribution = this._lockedAgent ? this.chatSessionsService.getChatSessionContribution(this._lockedAgent.id) : undefined;
			const providerIcon = contribution?.icon;
			const providerTitle = contribution?.welcomeTitle;
			const providerMessage = contribution?.welcomeMessage;

			// Fallback to default messages if provider doesn't specify
			const message = providerMessage
				? new MarkdownString(providerMessage)
				: (this._lockedAgent?.prefix === '@copilot '
					? new MarkdownString(localize('copilotCodingAgentMessage', "This chat session will be forwarded to the {0} [coding agent]({1}) where work is completed in the background. ", this._lockedAgent.prefix, 'https://aka.ms/coding-agent-docs') + DISCLAIMER, { isTrusted: true })
					: new MarkdownString(localize('genericCodingAgentMessage', "This chat session will be forwarded to the {0} coding agent where work is completed in the background. ", this._lockedAgent?.prefix) + DISCLAIMER));

			return {
				title: providerTitle ?? localize('codingAgentTitle', "Delegate to {0}", this._lockedAgent?.prefix),
				message,
				icon: providerIcon ?? Codicon.sendToRemoteAgent,
				additionalMessage,
				useLargeIcon: !!providerIcon,
			};
		}

		let title: string;
		if (this.input.currentModeKind === ChatModeKind.Ask) {
			title = localize('chatDescription', "Ask about your code");
		} else if (this.input.currentModeKind === ChatModeKind.Edit) {
			title = localize('editsTitle', "Edit in context");
		} else {
			title = localize('agentTitle', "Build with Agent");
		}

		return {
			title,
			message: new MarkdownString(DISCLAIMER),
			icon: Codicon.chatSparkle,
			additionalMessage,
		};
	}

	private async renderChatEditingSessionState() {
		if (!this.input) {
			return;
		}
		this.input.renderChatEditingSessionState(this._editingSession.get() ?? null);
	}

	private async renderFollowups(): Promise<void> {
		const lastItem = this.listWidget.lastItem;
		if (lastItem && isResponseVM(lastItem) && lastItem.isComplete) {
			this.input.renderFollowups(lastItem.replyFollowups, lastItem);
		} else {
			this.input.renderFollowups(undefined, undefined);
		}
	}

	private renderChatSuggestNextWidget(): void {
		if (this.lifecycleService.willShutdown) {
			return;
		}

		// Skip rendering in coding agent sessions unless the agent supports hand-offs
		if (this.isLockedToCodingAgent && !this._attachmentCapabilities.supportsHandOffs) {
			this.chatSuggestNextWidget.hide();
			return;
		}

		const items = this.viewModel?.getItems() ?? [];
		if (!items.length) {
			return;
		}

		const lastItem = items[items.length - 1];
		const lastResponseComplete = lastItem && isResponseVM(lastItem) && lastItem.isComplete;
		if (!lastResponseComplete || lastItem.isCanceled) {
			this.chatSuggestNextWidget.hide();
			return;
		}

		// Derive handoffs from the mode that generated the last response, not the current UI selection.
		// This ensures handoffs reflect what the response agent offers, regardless of mode picker state.
		// Fall back to the current mode picker for old sessions where modeInfo was not persisted.
		const modeInfo = lastItem.model.request?.modeInfo;
		let responseMode: IChatMode | undefined;
		if (modeInfo?.modeInstructions?.name) {
			responseMode = this.chatModeService.findModeByName(modeInfo.modeInstructions.name);
		} else if (modeInfo?.modeId) {
			responseMode = this.chatModeService.findModeById(modeInfo.modeId);
		} else {
			responseMode = this.input.currentModeObs.get();
		}

		const handoffs = responseMode?.handOffs?.get();

		if (responseMode && handoffs && handoffs.length > 0) {
			// In Autopilot mode, automatically trigger the first auto-send handoff
			// so the plan flows seamlessly into implementation without user interaction.
			const permissionLevel = this.inputPart.currentModeInfo.permissionLevel;
			if (permissionLevel === ChatPermissionLevel.Autopilot) {
				const autoSendHandoff = handoffs.find(h => h.send);
				if (autoSendHandoff) {
					this.handleNextPromptSelection(autoSendHandoff);
					return;
				}
			}

			// Log telemetry only when widget transitions from hidden to visible
			const wasHidden = this.chatSuggestNextWidget.domNode.style.display === 'none';
			this.chatSuggestNextWidget.render(responseMode);

			if (wasHidden) {
				this.telemetryService.publicLog2<ChatHandoffWidgetShownEvent, ChatHandoffWidgetShownClassification>('chat.handoffWidgetShown', {
					agent: getModeNameForTelemetry(responseMode),
					handoffCount: handoffs.length
				});
			}
		} else {
			this.chatSuggestNextWidget.hide();
		}

		// Trigger layout update
		if (this.bodyDimension) {
			this.layout(this.bodyDimension.height, this.bodyDimension.width);
		}
	}

	private handleNextPromptSelection(handoff: IHandOff, agentId?: string, withAutopilot?: boolean): void {
		// Hide the widget after selection
		this.chatSuggestNextWidget.hide();

		// If starting with Autopilot, set permission level before submitting
		if (withAutopilot) {
			this.inputPart.setPermissionLevel(ChatPermissionLevel.Autopilot);
		}

		const promptToUse = handoff.prompt;

		// Log telemetry
		const currentMode = this.input.currentModeObs.get();
		const toMode = handoff.agent ? this.chatModeService.findModeByName(handoff.agent) : undefined;
		this.telemetryService.publicLog2<ChatHandoffClickEvent, ChatHandoffClickClassification>('chat.handoffClicked', {
			fromAgent: getModeNameForTelemetry(currentMode),
			toAgent: agentId || (toMode ? getModeNameForTelemetry(toMode) : ''),
			hasPrompt: Boolean(promptToUse),
			autoSend: Boolean(handoff.send)
		});

		this.executeHandoff(handoff, agentId).catch(e => {
			const target = agentId ?? handoff.agent ?? 'unknown';
			this.logService.error(`[Handoff] Failed to execute handoff '${handoff.label}' to '${target}'`, e);
		});
	}

	async executeHandoff(handoff: IHandOff, agentId?: string): Promise<void> {
		this.chatSuggestNextWidget.hide();

		const promptToUse = this.getHandoffPrompt(handoff.prompt);

		// If agentId is provided (from chevron dropdown), delegate to that chat session
		// Otherwise, switch to the handoff agent
		if (agentId) {
			// Delegate to chat session (e.g., @background or @cloud)
			this.input.setValue(`@${agentId} ${promptToUse}`, false);
			this.input.focus();
			// Auto-submit for delegated chat sessions
			this.acceptInput().catch(e => this.logService.error(`[Handoff] Failed to submit delegated handoff to '@${agentId}'`, e));
		} else if (handoff.agent) {
			// Regular handoff to specified agent
			this._switchToAgentByName(handoff.agent);
			// Switch to the specified model if provided
			if (handoff.model) {
				this.input.switchModelByQualifiedName([handoff.model]);
			}
			// Insert the handoff prompt into the input
			this.input.setValue(promptToUse, false);
			this.input.focus();

			// Auto-submit if send flag is true
			if (handoff.send) {
				this.acceptInput();
			}
		}
	}

	private getHandoffPrompt(prompt: string): string {
		return augmentPromptWithPlanningContext(prompt, this.getPlanningTransitionContextForCurrentResponse());
	}

	private getLatestPlanningResponseModel(): IChatResponseModel | undefined {
		return [...(this.viewModel?.model.getRequests() ?? [])]
			.reverse()
			.find(request =>
				this.isPlanningModeInfo(request.modeInfo)
				&& request.id !== this._pendingPlanningPlaceholderRequestId
				&& request.response?.isComplete
				&& !request.response.isCompleteAddedRequest
			)
			?.response;
	}

	private getPlanningTransitionContextForCurrentResponse() {
		const latestPlanningResponse = this.getLatestPlanningResponseModel();
		if (!latestPlanningResponse) {
			return undefined;
		}

		const planningCarousel = [...latestPlanningResponse.response.value].reverse().find(isUsedQuestionCarousel);
		return mergePlanningTransitionContexts(planningCarousel ? buildPlanningTransitionContext(planningCarousel) : undefined, this._planningTransitionContext);
	}

	async handleDelegationExitIfNeeded(sourceAgent: Pick<IChatAgentData, 'id' | 'name'> | undefined, targetAgent: IChatAgentData | undefined): Promise<void> {
		if (!this._shouldExitAfterDelegation(sourceAgent, targetAgent)) {
			return;
		}

		this.logService.debug(`[Delegation] Will exit after delegation: sourceAgent=${sourceAgent?.id}, targetAgent=${targetAgent?.id}`);
		try {
			await this._handleDelegationExit();
		} catch (e) {
			this.logService.error('[Delegation] Failed to handle delegation exit', e);
		}
	}

	private _shouldExitAfterDelegation(sourceAgent: Pick<IChatAgentData, 'id' | 'name'> | undefined, targetAgent: IChatAgentData | undefined): boolean {
		if (!targetAgent) {
			this.logService.debug('[Delegation] _shouldExitAfterDelegation: false (no targetAgent)');
			return false;
		}

		if (!this.configurationService.getValue<boolean>(ChatConfiguration.ExitAfterDelegation)) {
			this.logService.debug('[Delegation] _shouldExitAfterDelegation: false (ExitAfterDelegation config disabled)');
			return false;
		}

		// Never exit if the source and target are the same (that means that you're providing a follow up, etc.)
		// NOTE: sourceAgent would be the chatWidget's 'lockedAgent'
		if (sourceAgent && sourceAgent.id === targetAgent.id) {
			this.logService.debug('[Delegation] _shouldExitAfterDelegation: false (source and target agents are the same)');
			return false;
		}

		if (!isIChatViewViewContext(this.viewContext)) {
			this.logService.debug('[Delegation] _shouldExitAfterDelegation: false (not in chat view context)');
			return false;
		}

		const contribution = this.chatSessionsService.getChatSessionContribution(targetAgent.id);
		if (!contribution) {
			this.logService.debug(`[Delegation] _shouldExitAfterDelegation: false (no contribution found for targetAgent.id=${targetAgent.id})`);
			return false;
		}

		if (contribution.canDelegate !== true) {
			this.logService.debug(`[Delegation] _shouldExitAfterDelegation: false (contribution.canDelegate=${contribution.canDelegate}, expected true)`);
			return false;
		}

		this.logService.debug('[Delegation] _shouldExitAfterDelegation: true');
		return true;
	}

	/**
	 * Handles the exit of the panel chat when a delegation to another session occurs.
	 * Waits for the response to complete and any pending confirmations to be resolved,
	 * then clears the widget unless the final message is an error.
	 */
	private async _handleDelegationExit(): Promise<void> {
		const viewModel = this.viewModel;
		if (!viewModel) {
			this.logService.debug('[Delegation] _handleDelegationExit: no viewModel, returning');
			return;
		}

		const parentSessionResource = viewModel.sessionResource;
		this.logService.debug(`[Delegation] _handleDelegationExit: parentSessionResource=${parentSessionResource.toString()}`);

		// Check if response is complete, not pending confirmation, and has no error
		const checkIfShouldClear = (): boolean => {
			const items = viewModel.getItems();
			const lastItem = items[items.length - 1];
			if (lastItem && isResponseVM(lastItem) && lastItem.model && lastItem.isComplete && !lastItem.model.isPendingConfirmation.get()) {
				const hasError = Boolean(lastItem.result?.errorDetails);
				return !hasError;
			}
			return false;
		};

		if (checkIfShouldClear()) {
			this.logService.debug('[Delegation] Response complete, archiving session before clearing');
			// Archive BEFORE clearing to ensure session still exists in agentSessionsService
			await this.archiveLocalParentSession(parentSessionResource);
			await this.clear();
			return;
		}

		this.logService.debug('[Delegation] Waiting for response to complete...');
		const shouldClear = await new Promise<boolean>(resolve => {
			const disposable = viewModel.onDidChange(() => {
				const result = checkIfShouldClear();
				if (result) {
					cleanup();
					resolve(true);
				}
			});
			const timeout = setTimeout(() => {
				this.logService.debug('[Delegation] Timeout waiting for response to complete');
				cleanup();
				resolve(false);
			}, 30_000); // 30 second timeout
			const cleanup = () => {
				clearTimeout(timeout);
				disposable.dispose();
			};
		});

		if (shouldClear) {
			this.logService.debug('[Delegation] Response completed, archiving session before clearing');
			await this.archiveLocalParentSession(parentSessionResource);
			await this.clear();
		} else {
			this.logService.debug('[Delegation] Not clearing (timeout or error)');
		}
	}

	private async archiveLocalParentSession(sessionResource: URI): Promise<void> {
		// In the regular workbench, only archive local chat sessions.
		// In the sessions window, allow archiving any session type after delegation.
		if (sessionResource.scheme !== Schemas.vscodeLocalChatSession && !IsSessionsWindowContext.getValue(this.contextKeyService)) {
			return;
		}

		this.logService.debug(`[Delegation] archiveLocalParentSession: archiving session ${sessionResource.toString()}`);

		// Implicitly keep parent session's changes as they've now been delegated to the new agent.
		await this.chatService.getSession(sessionResource)?.editingSession?.accept();

		const session = this.agentSessionsService.getSession(sessionResource);
		if (session) {
			session.setArchived(true);
			this.logService.debug('[Delegation] archiveLocalParentSession: session archived successfully');
		} else {
			this.logService.warn(`[Delegation] archiveLocalParentSession: session not found in agentSessionsService for ${sessionResource.toString()}`);
		}
	}

	setVisible(visible: boolean): void {
		const wasVisible = this._visible;
		this._visible = visible;
		this.visibleChangeCount++;
		this.listWidget.setVisible(visible);
		this.input.setVisible(visible);

		if (visible) {
			if (!wasVisible) {
				this.visibilityTimeoutDisposable.value = disposableTimeout(() => {
					// Progressive rendering paused while hidden, so start it up again.
					// Do it after a timeout because the container is not visible yet (it should be but offsetHeight returns 0 here)
					if (this._visible) {
						this.onDidChangeItems(true);
					}
				}, 0);

				this.visibilityAnimationFrameDisposable.value = dom.scheduleAtNextAnimationFrame(dom.getWindow(this.listContainer), () => {
					this._onDidShow.fire();
				});
			}
		} else if (wasVisible) {
			this._onDidHide.fire();
		}
	}

	private createList(listContainer: HTMLElement, options: IChatListItemRendererOptions): void {
		// Create a dom element to hold UI from editor widgets embedded in chat messages
		const overflowWidgetsContainer = document.createElement('div');
		overflowWidgetsContainer.classList.add('chat-overflow-widget-container', 'monaco-editor');
		listContainer.append(overflowWidgetsContainer);

		// Create chat list widget
		this.listWidget = this._register(this.instantiationService.createInstance(
			ChatListWidget,
			listContainer,
			{
				rendererOptions: options,
				renderStyle: this.viewOptions.renderStyle,
				defaultElementHeight: this.viewOptions.defaultElementHeight ?? 200,
				overflowWidgetsDomNode: overflowWidgetsContainer,
				styles: {
					listForeground: this.styles.listForeground,
					listBackground: this.styles.listBackground,
				},
				currentChatMode: () => this.input.currentModeKind,
				filter: this.viewOptions.filter ? { filter: this.viewOptions.filter.bind(this.viewOptions) } : undefined,
				codeBlockModelCollection: this._codeBlockModelCollection,
				viewModel: this.viewModel,
				editorOptions: this.editorOptions,
				location: this.location,
				getCurrentLanguageModelId: () => this.input.currentLanguageModel,
				getCurrentModeInfo: () => this.input.currentModeInfo,
			}
		));

		// Wire up ChatWidget-specific list widget events
		this._register(this.listWidget.onDidClickRequest(async item => {
			this.clickedRequest(item);
		}));

		this._register(this.listWidget.onDidRerender(item => {
			if (isRequestVM(item.currentElement) && this.configurationService.getValue<string>('chat.editRequests') !== 'input') {
				if (!item.rowContainer.contains(this.inputContainer)) {
					item.rowContainer.appendChild(this.inputContainer);
				}
				this.input.focus();
			}
		}));

		this._register(this.listWidget.onDidDispose(() => {
			this.focusedInputDOM.appendChild(this.inputContainer);
			this.input.focus();
		}));

		this._register(this.listWidget.onDidFocusOutside(() => {
			this.finishedEditing();
		}));

		this._register(this.listWidget.onDidClickFollowup(item => {
			// is this used anymore?
			this.acceptInput(item.message);
		}));

		this._register(this.listWidget.onDidChangeContentHeight(() => {
			this._onDidChangeContentHeight.fire();
		}));

		this._register(this.listWidget.onDidFocus(() => {
			this._onDidFocus.fire();
		}));
		this._register(this.listWidget.onDidScroll(() => {
			this._onDidScroll.fire();
		}));
	}

	startEditing(requestId: string): void {
		const editedRequest = this.listWidget.getTemplateDataForRequestId(requestId);
		if (editedRequest) {
			this.clickedRequest(editedRequest);
		}
	}

	private clickedRequest(item: IChatListItemTemplate) {

		const currentElement = item.currentElement;
		if (isRequestVM(currentElement) && !this.viewModel?.editing) {

			const requests = this.viewModel?.model.getRequests();
			if (!requests || !this.viewModel?.sessionResource) {
				return;
			}

			// this will only ever be true if we restored a checkpoint
			if (this.viewModel?.model.checkpoint) {
				this.recentlyRestoredCheckpoint = true;
			}

			this.viewModel?.model.setCheckpoint(currentElement.id);

			// set contexts and request to false
			const currentContext: IChatRequestVariableEntry[] = [];
			const addedContextIds = new Set<string>();
			const addToContext = (entry: IChatRequestVariableEntry) => {
				const dedupKey = entry.range ? `${entry.id}:${entry.range.start}-${entry.range.endExclusive}` : entry.id;
				if (addedContextIds.has(dedupKey) || isWorkspaceVariableEntry(entry)) {
					return;
				}
				if ((isPromptFileVariableEntry(entry) || isPromptTextVariableEntry(entry)) && entry.automaticallyAdded) {
					return;
				}
				addedContextIds.add(dedupKey);
				currentContext.push(entry);
			};
			for (let i = requests.length - 1; i >= 0; i -= 1) {
				const request = requests[i];
				if (request.id === currentElement.id) {
					request.setShouldBeBlocked(false); // unblocking just this request.
					request.attachedContext?.forEach(addToContext);
				}
			}
			currentElement.variables.forEach(addToContext);

			// set states
			this.viewModel?.setEditing(currentElement);
			if (item?.contextKeyService) {
				ChatContextKeys.currentlyEditing.bindTo(item.contextKeyService).set(true);
			}

			const isEditingSentRequest = currentElement.pendingKind === undefined
				? ChatContextKeys.EditingRequestType.Sent
				: currentElement.pendingKind === ChatRequestQueueKind.Queued
					? ChatContextKeys.EditingRequestType.Queue
					: ChatContextKeys.EditingRequestType.Steer;
			const isInput = this.configurationService.getValue<string>('chat.editRequests') === 'input';
			this.inputPart?.setEditing(!!this.viewModel?.editing && isInput, isEditingSentRequest);

			if (!isInput) {
				const rowContainer = item.rowContainer;
				this.inputContainer = dom.$('.chat-edit-input-container');
				rowContainer.appendChild(this.inputContainer);
				this.createInput(this.inputContainer);
				this.input.setChatMode(this.inputPart.currentModeObs.get().id);
				this.input.setPermissionLevel(this.inputPart.currentModeInfo.permissionLevel ?? ChatPermissionLevel.Default);
				this.input.setEditing(true, isEditingSentRequest);
				this._onDidChangeActiveInputEditor.fire();
			} else {
				this.inputPart.element.classList.add('editing');
			}

			this.inputPart.toggleChatInputOverlay(!isInput);
			if (currentContext.length > 0) {
				this.input.attachmentModel.addContext(...currentContext);
			}

			// rerenders
			this.inputPart.dnd.setDisabledOverlay(!isInput);
			this.input.renderAttachedContext();
			this.input.setValue(currentElement.messageText, false);

			// restore dynamic variables in the model so decorations and parsing work
			const dynamicVariableModel = this.getContrib<ChatDynamicVariableModel>(ChatDynamicVariableModel.ID);
			const editorModel = this.input.inputEditor.getModel();
			if (dynamicVariableModel && editorModel) {
				const modelTextLength = editorModel.getValueLength();
				for (const entry of currentContext) {
					if (entry.range) {
						if (entry.range.start >= entry.range.endExclusive) {
							continue;
						}

						if (entry.range.start < 0 || entry.range.endExclusive > modelTextLength) {
							continue;
						}

						const startPos = editorModel.getPositionAt(entry.range.start);
						const endPos = editorModel.getPositionAt(entry.range.endExclusive);
						dynamicVariableModel.addReference({
							id: entry.id,
							range: new Range(startPos.lineNumber, startPos.column, endPos.lineNumber, endPos.column),
							data: entry.value,
							fullName: entry.fullName,
							icon: entry.icon,
							modelDescription: entry.modelDescription,
							isFile: entry.kind === 'file',
							isDirectory: entry.kind === 'directory',
						});
					}
				}
			}

			this.listWidget.suppressAutoScroll = true;
			this.onDidChangeItems();
			this.input.inputEditor.focus();

			this._register(this.inputPart.onDidClickOverlay(() => {
				if (this.viewModel?.editing && this.configurationService.getValue<string>('chat.editRequests') !== 'input') {
					this.finishedEditing();
				}
			}));

			// listeners
			if (!isInput) {
				this._register(this.inlineInputPart.inputEditor.onDidChangeModelContent(() => {
					this.listWidget.scrollToCurrentItem(currentElement);
				}));

				this._register(this.inlineInputPart.inputEditor.onDidChangeCursorSelection((e) => {
					this.listWidget.scrollToCurrentItem(currentElement);
				}));
			}
		}

		type StartRequestEvent = { editRequestType: string };

		type StartRequestEventClassification = {
			owner: 'justschen';
			comment: 'Event used to gain insights into when edits are being pressed.';
			editRequestType: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Current entry point for editing a request.' };
		};

		this.telemetryService.publicLog2<StartRequestEvent, StartRequestEventClassification>('chat.startEditingRequests', {
			editRequestType: this.configurationService.getValue<string>('chat.editRequests'),
		});
	}

	finishedEditing(completedEdit?: boolean): void {
		// reset states
		this.listWidget.suppressAutoScroll = false;
		const editedRequest = this.listWidget.getTemplateDataForRequestId(this.viewModel?.editing?.id);
		if (this.recentlyRestoredCheckpoint) {
			this.recentlyRestoredCheckpoint = false;
		} else {
			this.viewModel?.model.setCheckpoint(undefined);
		}
		this.inputPart.dnd.setDisabledOverlay(false);
		if (editedRequest?.contextKeyService) {
			ChatContextKeys.currentlyEditing.bindTo(editedRequest.contextKeyService).set(false);
		}

		const isInput = this.configurationService.getValue<string>('chat.editRequests') === 'input';

		if (!isInput) {
			this.inputPart.setChatMode(this.input.currentModeObs.get().id);
			this.inputPart.setPermissionLevel(this.input.currentModeInfo.permissionLevel ?? ChatPermissionLevel.Default);
			const currentModel = this.input.selectedLanguageModel.get();
			if (currentModel) {
				this.inputPart.switchModel(currentModel.metadata);
			}

			this.inputPart?.toggleChatInputOverlay(false);
			try {
				if (editedRequest?.rowContainer?.contains(this.inputContainer)) {
					editedRequest.rowContainer.removeChild(this.inputContainer);
				} else if (this.inputContainer.parentElement) {
					this.inputContainer.parentElement.removeChild(this.inputContainer);
				}
			} catch (e) {
				this.logService.error('Error occurred while finishing editing:', e);
			}
			this.inputContainer = dom.$('.empty-chat-state');

			// only dispose if we know the input is not the bottom input object.
			this.input.dispose();
		}

		if (isInput) {
			this.inputPart.element.classList.remove('editing');
		}
		this.viewModel?.setEditing(undefined);
		this.inputPart?.setEditing(false, undefined);

		if (!isInput) {
			this._onDidChangeActiveInputEditor.fire();
		}

		this.onDidChangeItems();

		type CancelRequestEditEvent = {
			editRequestType: string;
			editCanceled: boolean;
		};

		type CancelRequestEventEditClassification = {
			owner: 'justschen';
			editRequestType: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Current entry point for editing a request.' };
			editCanceled: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'Indicates whether the edit was canceled.' };
			comment: 'Event used to gain insights into when edits are being canceled.';
		};

		this.telemetryService.publicLog2<CancelRequestEditEvent, CancelRequestEventEditClassification>('chat.editRequestsFinished', {
			editRequestType: this.configurationService.getValue<string>('chat.editRequests'),
			editCanceled: !completedEdit
		});

		this.inputPart.focus();
	}

	private getWidgetViewKindTag(): string {
		if (!this.viewContext) {
			return 'editor';
		} else if (isIChatViewViewContext(this.viewContext)) {
			return 'view';
		} else {
			return 'quick';
		}
	}

	private createInput(container: HTMLElement, options?: { renderFollowups: boolean; renderStyle?: 'compact' | 'minimal'; renderInputToolbarBelowInput?: boolean }): void {
		const commonConfig: IChatInputPartOptions = {
			renderFollowups: options?.renderFollowups ?? true,
			renderStyle: options?.renderStyle === 'minimal' ? 'compact' : options?.renderStyle,
			renderInputToolbarBelowInput: options?.renderInputToolbarBelowInput ?? false,
			menus: {
				executeToolbar: MenuId.ChatExecute,
				telemetrySource: 'chatWidget',
				...this.viewOptions.menus
			},
			editorOverflowWidgetsDomNode: this.viewOptions.editorOverflowWidgetsDomNode,
			enableImplicitContext: this.viewOptions.enableImplicitContext,
			renderWorkingSet: this.viewOptions.enableWorkingSet === 'explicit',
			supportsChangingModes: this.viewOptions.supportsChangingModes,
			dndContainer: this.viewOptions.dndContainer,
			inputEditorMinLines: this.viewOptions.inputEditorMinLines,
			widgetViewKindTag: this.getWidgetViewKindTag(),
			defaultMode: this.viewOptions.defaultMode,
			sessionTypePickerDelegate: this.viewOptions.sessionTypePickerDelegate,
			workspacePickerDelegate: this.viewOptions.workspacePickerDelegate,
			isSessionsWindow: this.viewOptions.isSessionsWindow,
		};

		if (this.viewModel?.editing) {
			const editedRequest = this.listWidget.getTemplateDataForRequestId(this.viewModel?.editing?.id);
			const scopedInstantiationService = this._register(this.instantiationService.createChild(new ServiceCollection([IContextKeyService, editedRequest?.contextKeyService])));
			this.inlineInputPartDisposable.value = scopedInstantiationService.createInstance(ChatInputPart,
				this.location,
				commonConfig,
				this.styles,
				true
			);
		} else {
			this.inputPartDisposable.value = this.instantiationService.createInstance(ChatInputPart,
				this.location,
				commonConfig,
				this.styles,
				false
			);
			this._register(autorun(reader => {
				this.inputPart.height.read(reader);
				if (!this.listWidget) {
					// This is set up before the list/renderer are created
					return;
				}

				if (this.bodyDimension) {
					// Only re-layout the list/containers to match the new input
					// height. Do NOT re-call this.layout() here: the input part
					// has already laid itself out and re-entering inputPart.layout
					// creates a layout loop when the viewPane also reacts.
					this._layoutListForInputHeight();
				}

				this._onDidChangeContentHeight.fire();
			}));
		}

		this.input.render(container, '', this);
		if (this.bodyDimension?.width) {
			this.input.layout(this.bodyDimension.width);
		}

		this._register(this.input.onDidLoadInputState(() => {
			this.refreshParsedInput();
		}));
		this._register(this.input.onDidFocus(() => this._onDidFocus.fire()));
		this._register(this.input.onDidAcceptFollowup(e => {
			if (!this.viewModel) {
				return;
			}

			let msg = '';
			if (e.followup.agentId && e.followup.agentId !== this.chatAgentService.getDefaultAgent(this.location, this.input.currentModeKind)?.id) {
				const agent = this.chatAgentService.getAgent(e.followup.agentId);
				if (!agent) {
					return;
				}

				this.lastSelectedAgent = agent;
				msg = `${chatAgentLeader}${agent.name} `;
				if (e.followup.subCommand) {
					msg += `${chatSubcommandLeader}${e.followup.subCommand} `;
				}
			} else if (!e.followup.agentId && e.followup.subCommand && this.chatSlashCommandService.hasCommand(e.followup.subCommand)) {
				msg = `${chatSubcommandLeader}${e.followup.subCommand} `;
			}

			msg += e.followup.message;
			this.acceptInput(msg);

			if (!e.response) {
				// Followups can be shown by the welcome message, then there is no response associated.
				// At some point we probably want telemetry for these too.
				return;
			}

			this.chatService.notifyUserAction({
				sessionResource: this.viewModel.sessionResource,
				requestId: e.response.requestId,
				agentId: e.response.agent?.id,
				command: e.response.slashCommand?.name,
				result: e.response.result,
				action: {
					kind: 'followUp',
					followup: e.followup
				},
			});
		}));
		this._register(this.inputEditor.onDidChangeModelContent(() => {
			this.parsedChatRequest = undefined;
			this.updateChatInputContext();
		}));
		this._register(this.chatAgentService.onDidChangeAgents(() => {
			this.parsedChatRequest = undefined;
			// Tools agent loads -> welcome content changes
			this.renderWelcomeViewContentIfNeeded();
		}));
		let wasPlanningMode = this.isInPlanningMode();
		this._register(this.input.onDidChangeCurrentChatMode(() => {
			const isPlanningMode = this.isInPlanningMode();
			if (isPlanningMode !== wasPlanningMode) {
				this.resetPlanningMiddlewareState();
			}
			wasPlanningMode = isPlanningMode;
			this.renderWelcomeViewContentIfNeeded();
			this.refreshParsedInput();
			this.renderFollowups();
			this.renderChatSuggestNextWidget();
		}));
		const foregroundSessionCountContextKeys = new Set([ChatContextKeys.foregroundSessionCount.key]);
		this._register(this.contextKeyService.onDidChangeContext(e => {
			if (e.affectsSome(foregroundSessionCountContextKeys) && this.isEmpty()) {
				this.renderGettingStartedTipIfNeeded();
			}
		}));
		let previousModelIdentifier: string | undefined;
		this._register(autorun(reader => {
			const modelIdentifier = this.inputPart.selectedLanguageModel.read(reader)?.identifier;
			if (previousModelIdentifier === undefined) {
				previousModelIdentifier = modelIdentifier;
				return;
			}

			if (previousModelIdentifier === modelIdentifier) {
				return;
			}

			previousModelIdentifier = modelIdentifier;
			if (!this._gettingStartedTipPartRef) {
				return;
			}

			this.chatTipService.getWelcomeTip(this.contextKeyService);
		}));

		this._register(autorun(r => {
			const toolSetIds = new Set<string>();
			const toolIds = new Set<string>();
			for (const [entry, enabled] of this.input.selectedToolsModel.entriesMap.read(r)) {
				if (enabled) {
					if (isToolSet(entry)) {
						toolSetIds.add(entry.id);
					} else {
						toolIds.add(entry.id);
					}
				}
			}
			const disabledTools = this.input.attachmentModel.attachments
				.filter(a => a.kind === 'tool' && !toolIds.has(a.id) || a.kind === 'toolset' && !toolSetIds.has(a.id))
				.map(a => a.id);

			this.input.attachmentModel.updateContext(disabledTools, Iterable.empty());
			this.refreshParsedInput();
		}));
	}

	private onDidStyleChange(): void {
		this.container.style.setProperty('--vscode-interactive-result-editor-background-color', this.editorOptions.configuration.resultEditor.backgroundColor?.toString() ?? '');
		this.container.style.setProperty('--vscode-interactive-session-foreground', this.editorOptions.configuration.foreground?.toString() ?? '');
		this.container.style.setProperty('--vscode-chat-list-background', this.themeService.getColorTheme().getColor(this.styles.listBackground)?.toString() ?? '');
	}


	setModel(model: IChatModel | undefined): void {
		if (!this.container) {
			throw new Error('Call render() before setModel()');
		}

		if (!model) {
			if (this.viewModel?.editing) {
				this.finishedEditing();
			}
			this.viewModel = undefined;
			this.onDidChangeItems();
			this._hasPendingRequestsContextKey.set(false);
			return;
		}

		if (isEqual(model.sessionResource, this.viewModel?.sessionResource)) {
			return;
		}

		if (this.viewModel?.editing) {
			this.finishedEditing();
		}
		this.inputPart.clearTodoListWidget(model.sessionResource, false);
		this.inputPart.clearArtifactsWidget();
		this.chatSuggestNextWidget.hide();
		this.chatTipService.resetSession();

		// Switching sessions resets tip service state; clear any rendered tip so
		// empty-state rendering picks a fresh, context-appropriate tip.
		this._gettingStartedTipPartRef = undefined;
		this._gettingStartedTipPart.clear();
		const tipContainer = this.inputPart.gettingStartedTipContainerElement;
		dom.clearNode(tipContainer);
		dom.setVisibility(false, tipContainer);

		this._codeBlockModelCollection.clear();

		// Set the input model on the inputPart before assigning this.viewModel. Assigning this.viewModel
		// fires onDidChangeViewModel, which ChatInputPart listens to and expects the input model to be initialized.
		// Pass input model reference to input part for state syncing
		this.inputPart.setInputModel(model.inputModel, model.getRequests().length === 0);

		this.viewModel = this.instantiationService.createInstance(ChatViewModel, model, this._codeBlockModelCollection, undefined);

		this.listWidget.setViewModel(this.viewModel);

		if (this._lockedAgent) {
			let placeholder = this.chatSessionsService.getChatSessionContribution(this._lockedAgent.id)?.inputPlaceholder;
			if (!placeholder) {
				placeholder = localize('chat.input.placeholder.lockedToAgent', "Chat with {0}", this._lockedAgent.displayName || this._lockedAgent.name);
			}
			this.viewModel.setInputPlaceholder(placeholder);
			this.inputEditor.updateOptions({ placeholder });
		} else if (this.viewModel.inputPlaceholder) {
			this.inputEditor.updateOptions({ placeholder: this.viewModel.inputPlaceholder });
		}

		const renderImmediately = this.configurationService.getValue<boolean>('chat.experimental.renderMarkdownImmediately');
		const delay = renderImmediately ? MicrotaskDelay : 0;
		this.viewModelDisposables.add(Event.runAndSubscribe(Event.accumulate(this.viewModel.onDidChange, delay), (events => {
			if (!this.viewModel || this._store.isDisposed) {
				// See https://github.com/microsoft/vscode/issues/278969
				return;
			}

			this.requestInProgress.set(this.viewModel.model.requestInProgress.get());

			// Update the editor's placeholder text when it changes in the view model
			if (events?.some(e => e?.kind === 'changePlaceholder')) {
				this.inputEditor.updateOptions({ placeholder: this.viewModel.inputPlaceholder });
			}

			this.onDidChangeItems();
			if (events?.some(e => e?.kind === 'addRequest') && this.visible && this._planningTranscriptScrollPreservationDepth === 0) {
				this.listWidget.scrollToEnd();
			}
		})));
		this.viewModelDisposables.add(this.viewModel.onDidDisposeModel(() => {
			// Ensure that view state is saved here, because we will load it again when a new model is assigned
			if (this.viewModel?.editing) {
				this.finishedEditing();
			}
			// Disposes the viewmodel and listeners
			this.viewModel = undefined;
			this.onDidChangeItems();
		}));
		this._sessionIsEmptyContextKey.set(model.getRequests().length === 0);
		const supportsFork = this.chatSessionsService.sessionSupportsFork(model.sessionResource);
		this._chatSessionSupportsForkContextKey.set(supportsFork);
		this.listWidget?.updateRendererOptions({ supportsFork });
		this._sessionHasDebugDataContextKey.set(this.chatDebugService.getEvents(model.sessionResource).length > 0);
		let lastSteeringCount = 0;
		const updatePendingRequestKeys = (announceSteering: boolean) => {
			const pendingRequests = model.getPendingRequests();
			const pendingCount = pendingRequests.length;
			this._hasPendingRequestsContextKey.set(pendingCount > 0);
			const steeringCount = pendingRequests.filter(pending => pending.kind === ChatRequestQueueKind.Steering).length;
			if (announceSteering && steeringCount > 0 && lastSteeringCount === 0) {
				status(localize('chat.pendingRequests.steeringQueued', "Steering"));
			}
			lastSteeringCount = steeringCount;
		};
		updatePendingRequestKeys(false);
		this.viewModelDisposables.add(model.onDidChangePendingRequests(() => updatePendingRequestKeys(true)));

		this.refreshParsedInput();
		this.viewModelDisposables.add(model.onDidChange((e) => {
			if (e.kind === 'setAgent') {
				this._onDidChangeAgent.fire({ agent: e.agent, slashCommand: e.command });
				// Update capabilities context keys when agent changes
				this._updateAgentCapabilitiesContextKeys(e.agent);
			}
			if (e.kind === 'addRequest') {
				this.inputPart.clearTodoListWidget(this.viewModel?.sessionResource, false);
				this._sessionIsEmptyContextKey.set(false);
				this.chatSuggestNextWidget.hide();
			}
			// Hide widget on request removal
			if (e.kind === 'removeRequest') {
				this.inputPart.clearTodoListWidget(this.viewModel?.sessionResource, true);
				this.chatSuggestNextWidget.hide();
				this._sessionIsEmptyContextKey.set((this.viewModel?.model.getRequests().length ?? 0) === 0);
			}
			// Show next steps widget when response completes (not when request starts)
			if (e.kind === 'completedRequest') {
				const lastRequest = this.viewModel?.model.getRequests().at(-1);
				const wasCancelled = lastRequest?.response?.isCanceled ?? false;
				if (wasCancelled) {
					// Clear todo list when request is cancelled
					this.inputPart.clearTodoListWidget(this.viewModel?.sessionResource, true);
				}
				// Only show if response wasn't canceled
				this.renderChatSuggestNextWidget();

				// Mark the session as read when the request completes and the widget is visible
				if (this.visible && this.viewModel?.sessionResource) {
					this.agentSessionsService.getSession(this.viewModel.sessionResource)?.setRead(true);
				}
			}
		}));

		if (this.listWidget && this.visible) {
			this.onDidChangeItems();
			this.listWidget.scrollToEnd();
		}

		this.renderChatSuggestNextWidget();
		this.updateChatInputContext();
		this.input.renderChatTodoListWidget(this.viewModel.sessionResource);
		this.input.renderArtifactsWidget(this.viewModel.sessionResource);
	}

	getFocus(): ChatTreeItem | undefined {
		return this.listWidget.getFocus()[0] ?? undefined;
	}

	reveal(item: ChatTreeItem, relativeTop?: number): void {
		this.listWidget.reveal(item, relativeTop);
	}

	focus(item: ChatTreeItem): void {
		if (!this.listWidget.hasElement(item)) {
			return;
		}

		this.listWidget.focusItem(item);
	}

	setInputPlaceholder(placeholder: string): void {
		this.viewModel?.setInputPlaceholder(placeholder);
	}

	resetInputPlaceholder(): void {
		this.viewModel?.resetInputPlaceholder();
	}

	setInput(value = ''): void {
		this.input.setValue(value, false);
		this.refreshParsedInput();
	}

	getInput(): string {
		return this.input.inputEditor.getValue();
	}

	getContrib<T extends IChatWidgetContrib>(id: string): T | undefined {
		return this.contribs.find(c => c.id === id) as T | undefined;
	}

	// Coding agent locking methods
	lockToCodingAgent(name: string, displayName: string, agentId: string): void {
		this._lockedAgent = {
			id: agentId,
			name,
			prefix: `@${name} `,
			displayName
		};
		this._lockedToCodingAgentContextKey.set(true);
		this._lockedCodingAgentIdContextKey.set(agentId);
		this.renderWelcomeViewContentIfNeeded();
		// Update capabilities for the locked agent
		const agent = this.chatAgentService.getAgent(agentId);
		this._updateAgentCapabilitiesContextKeys(agent);
		this.listWidget?.updateRendererOptions({ restorable: false, editable: false, noFooter: true, progressMessageAtBottomOfResponse: true });
		if (this.visible) {
			this.listWidget?.rerender();
		}
	}

	unlockFromCodingAgent(): void {
		// Clear all state related to locking
		this._lockedAgent = undefined;
		this._lockedToCodingAgentContextKey.set(false);
		this._lockedCodingAgentIdContextKey.set('');
		this._chatSessionSupportsForkContextKey.set(false);
		this._updateAgentCapabilitiesContextKeys(undefined);

		// Explicitly update the DOM to reflect unlocked state
		this.renderWelcomeViewContentIfNeeded();

		// Reset to default placeholder
		if (this.viewModel) {
			this.viewModel.resetInputPlaceholder();
		}
		this.inputEditor?.updateOptions({ placeholder: undefined });
		this.listWidget?.updateRendererOptions({ restorable: true, editable: true, noFooter: false, progressMessageAtBottomOfResponse: mode => mode !== ChatModeKind.Ask });
		if (this.visible) {
			this.listWidget?.rerender();
		}
	}

	get isLockedToCodingAgent(): boolean {
		return !!this._lockedAgent;
	}

	get lockedAgentId(): string | undefined {
		return this._lockedAgent?.id;
	}

	logInputHistory(): void {
		this.input.logInputHistory();
	}

	async acceptInput(query?: string, options?: IChatAcceptInputOptions): Promise<IChatResponseModel | undefined> {
		return this._acceptInput(query ? { query } : undefined, options);
	}

	async rerunLastRequest(): Promise<void> {
		if (!this.viewModel) {
			return;
		}

		const sessionResource = this.viewModel.sessionResource;
		const lastRequest = this.chatService.getSession(sessionResource)?.getRequests().at(-1);
		if (!lastRequest) {
			return;
		}

		const options: IChatSendRequestOptions = {
			attempt: lastRequest.attempt + 1,
			location: this.location,
			userSelectedModelId: this.input.currentLanguageModel,
			modeInfo: this.input.currentModeInfo,
		};
		const result = await this.chatService.resendRequest(lastRequest, options);
		this.logThinkingStyleUsage('rerun');
		return result;
	}

	private getConfiguredThinkingStyle(): ThinkingDisplayMode {
		const thinkingStyle = this.configurationService.getValue<ThinkingDisplayMode>(ChatConfiguration.ThinkingStyle);
		switch (thinkingStyle) {
			case ThinkingDisplayMode.Collapsed:
			case ThinkingDisplayMode.CollapsedPreview:
			case ThinkingDisplayMode.FixedScrolling:
				return thinkingStyle;
			default:
				return ThinkingDisplayMode.FixedScrolling;
		}
	}

	private logThinkingStyleUsage(requestKind: ChatThinkingStyleUsageEvent['requestKind']): void {
		this.telemetryService.publicLog2<ChatThinkingStyleUsageEvent, ChatThinkingStyleUsageClassification>('chat.thinkingStyleUsage', {
			thinkingStyle: this.getConfiguredThinkingStyle(),
			location: this.location,
			requestKind,
		});
	}

	private async _applyPromptFileIfSet(requestInput: IChatRequestInputOptions): Promise<void> {
		// first check if the input has a prompt slash command
		const agentSlashPromptPart = this.parsedInput.parts.find((r): r is ChatRequestSlashPromptPart => r instanceof ChatRequestSlashPromptPart);
		if (!agentSlashPromptPart) {
			return;
		}

		// Prompt slash commands are transformed out of the input before sendRequest.
		// Track them now so tip exclusions still update for commands like /init.
		this.chatTipService.recordSlashCommandUsage(agentSlashPromptPart.name);

		// need to resolve the slash command to get the prompt file
		const slashCommand = await this.promptsService.resolvePromptSlashCommand(agentSlashPromptPart.name, CancellationToken.None);
		if (!slashCommand) {
			return;
		}
		const parseResult = slashCommand.parsedPromptFile;
		// add the prompt file to the context
		const refs = parseResult.body?.variableReferences.map(({ name, offset }) => ({ name, range: new OffsetRange(offset, offset + name.length + 1) })) ?? [];
		const toolReferences = this.toolsService.toToolReferences(refs);
		requestInput.attachedContext.insertFirst(toPromptFileVariableEntry(parseResult.uri, PromptFileVariableKind.PromptFile, undefined, true, toolReferences));

		const promptPath = slashCommand.promptPath;
		const promptRunEvent: ChatPromptRunEvent = {
			storage: promptPath.storage,
		};
		if (promptPath.storage === PromptsStorage.extension) {
			promptRunEvent.extensionId = promptPath.extension.identifier.value;
			promptRunEvent.promptName = slashCommand.name;
		} else {
			promptRunEvent.promptNameHash = hash(slashCommand.name).toString(16);
		}
		this.telemetryService.publicLog2<ChatPromptRunEvent, ChatPromptRunClassification>('chat.promptRun', promptRunEvent);

		if (parseResult.header) {
			await this._applyPromptMetadata(parseResult.header, requestInput);
		}
	}

	private async maybeRunDynamicPlanningQuestions(input: string, options: IChatAcceptInputOptions): Promise<boolean> {
		if (this._skipDynamicPlanningQuestionsOnce) {
			this._skipDynamicPlanningQuestionsOnce = false;
			return false;
		}

		if (!this.shouldRunDynamicPlanningQuestions(input)) {
			return false;
		}

		return this.triggerDynamicPlanningQuestions(input, options);
	}

	private resetPlanningMiddlewareState(): void {
		if (isPlanningMiddlewareQuestionCarousel(this._pendingPlanningQuestionResolveId)) {
			this.input.clearQuestionCarousel(undefined, this._pendingPlanningQuestionResolveId);
		}
		void this.clearPendingPlanningPlaceholder({ preserveUsedPlanningCarousel: false });

		this._planningPhase = 'broad-scan';
		this._planningTransitionContext = undefined;
		this._lastPlanningQuestionModelId = undefined;
		this._lastPlanningQuestionSourceInput = undefined;
		this._goalClarityQuestionRounds = 0;
		this._pendingPlanningQuestionResolveId = undefined;
		this._pendingPlanningPlaceholderRequestId = undefined;
		this._currentPlanningPlanRequestId = undefined;
		this._previousPlanningPlanRequestId = undefined;
		this._planningPlanTextByRequestId.clear();
		this._pendingPlanningQuestionAnswersListener.clear();
		this._pendingPlanningResponseListener.clear();
		this._skipDynamicPlanningQuestionsOnce = false;
	}

	private async triggerDynamicPlanningQuestions(
		input: string,
		options: IChatAcceptInputOptions,
		planningOptions?: { readonly forceRegenerate?: boolean; readonly phase?: PlanningPhase; readonly questionStage?: PlanningQuestionStage; readonly useSubmittedPlaceholder?: boolean }
	): Promise<boolean> {
		if (!input.trim()) {
			return false;
		}

		if (!this.isInPlanningMode()) {
			return false;
		}

		if (this._pendingPlanningQuestionResolveId) {
			this.input.focusQuestionCarousel();
			return true;
		}

		if (this.input.questionCarousel) {
			this.input.focusQuestionCarousel();
			return true;
		}

		const planningPhase = planningOptions?.phase ?? this._planningPhase;
		const questionStage = planningOptions?.questionStage ?? 'goal-clarity';
		const useSubmittedPlaceholder = planningOptions?.useSubmittedPlaceholder ?? (questionStage === 'goal-clarity' && (options.storeToHistory ?? true));
		const recentConversation = this.getRecentPlanningConversation();

		try {
			if (useSubmittedPlaceholder && this.canShowSubmittedPlanningPlaceholder()) {
				await this.showSubmittedPlanningPlaceholder(input, options, questionStage);
			}

			const generationContext = await this.getPlanningQuestionGenerationContext(
				input,
				planningPhase,
				questionStage,
				recentConversation
			);
			if (questionStage === 'task-decomposition' && generationContext.currentPlan?.trim()) {
				const planSteps = extractPlanningPlanSteps(generationContext.currentPlan, 24);
				const generationResult = await this.generatePlanningPlanStepControlsWithTimeout(generationContext, planSteps);
				this._lastPlanningQuestionModelId = generationResult.modelId;
				this._planningPhase = generationContext.planningPhase;
				this.showPlanningPlanEditor(
					input,
					options,
					generationContext.planningPhase,
					this.getLatestPlanningPlanSnapshot(),
					generationContext.focusAreaLabel,
					undefined,
					generationResult.questionsByStep
				);
				return true;
			}

			const generationResult = questionStage === 'goal-clarity'
				? await this.generateNonEmptyGoalClarityQuestionsWithRetry(generationContext)
				: await this.generatePlanningQuestionsWithTimeout(generationContext);
			const questions = generationResult.questions;
			this._lastPlanningQuestionModelId = generationResult.modelId;
			if (!questions.length) {
				await this.showPlanningQuestionGenerationError(questionStage, new Error(localize(
					'chat.dynamicPlanning.emptyQuestions',
					'Planning question generation did not return any usable questions.'
				)));
				return true;
			}

			this._planningPhase = generationContext.planningPhase;
			this.showDynamicPlanningQuestionCarousel(
				input,
				questions,
				options,
				generationContext,
				useSubmittedPlaceholder
			);
			return true;
		} catch (error) {
			await this.showPlanningQuestionGenerationError(questionStage, error);
			return true;
		}
	}

	private shouldRunDynamicPlanningQuestions(input: string): boolean {
		if (!input.trim()) {
			return false;
		}

		if (!this.isInPlanningMode()) {
			return false;
		}

		if (this.input.questionCarousel) {
			return false;
		}

		const itemCount = this.viewModel?.getItems().length ?? 0;
		if (itemCount === 0) {
			return true;
		}

		const hasPlanningContext = Boolean(this.getPlanningTransitionContextForCurrentResponse() ?? this._planningTransitionContext);
		const previousPlanningInput = this.getLastPlanningQuestionSourceInput();
		if (!previousPlanningInput) {
			return !hasPlanningContext;
		}

		return shouldRegeneratePlanningQuestions(input, previousPlanningInput, hasPlanningContext);
	}

	private isInPlanningMode(): boolean {
		const mode = this.input.currentModeObs.get();
		return isPlanningModeName(mode.id) || isPlanningModeName(mode.name.get());
	}

	private getLastPlanningQuestionSourceInput(): string | undefined {
		if (this._lastPlanningQuestionSourceInput?.trim()) {
			return this._lastPlanningQuestionSourceInput;
		}

		const lastPlanningRequestId = [...(this.viewModel?.model.getRequests() ?? [])]
			.reverse()
			.find(request => this.isPlanningModeInfo(request.modeInfo))
			?.id;
		if (!lastPlanningRequestId) {
			return undefined;
		}

		return [...(this.viewModel?.getItems() ?? [])]
			.reverse()
			.find((item): item is IChatRequestViewModel => isRequestVM(item) && item.id === lastPlanningRequestId)
			?.messageText;
	}

	private isPlanningModeInfo(modeInfo: { readonly modeInstructions?: { readonly name?: string }; readonly modeId?: string } | undefined): boolean {
		return isPlanningModeName(modeInfo?.modeInstructions?.name) || isPlanningModeName(modeInfo?.modeId);
	}

	private async getPlanningQuestionGenerationContext(
		input: string,
		planningPhase: PlanningPhase,
		questionStage: PlanningQuestionStage,
		recentConversationOverride?: readonly string[],
		overrides?: {
			readonly planningContext?: IPlanningTransitionContext;
			readonly currentPlan?: string;
			readonly focusAreaLabel?: string;
			readonly focusHint?: string;
		}
	): Promise<IPlanningQuestionGenerationContext> {
		const contextEditors = this.getPlanningContextEditors();
		const editor = contextEditors[0];
		const model = editor?.getModel();
		const selection = editor?.getSelection();
		const selectedText = model && selection && !selection.isEmpty() ? model.getValueInRange(selection) : undefined;
		const mergedPlanningContext = overrides?.planningContext ?? this.getPlanningTransitionContextForCurrentResponse() ?? this._planningTransitionContext;
		const recentConversation = [...(recentConversationOverride ?? this.getRecentPlanningConversation())];
		const currentPlan = overrides?.currentPlan ?? this.getCurrentPlanningResponseText();
		const repositoryContext = await collectPlanningRepositoryContext({
			phase: planningPhase,
			questionStage,
			userRequest: input,
			plannerNotes: mergedPlanningContext?.plannerNotes,
			recentConversation,
			planningAnswers: mergedPlanningContext?.answers.map(answer => answer.answer) ?? [],
			currentPlan,
			focusHint: overrides?.focusHint,
			previousRepositoryContext: mergedPlanningContext?.repositoryContext,
			confirmedPlanningTarget: mergedPlanningContext?.repositoryContext?.planningTarget,
			activeEditor: editor,
			contextEditors,
		}, {
			fileService: this.fileService,
			textModelService: this.textModelService,
			workspaceContextService: this.workspaceContextService,
			languageFeaturesService: this.languageFeaturesService,
		});
		const readiness = assessPlanningReadiness({
			userRequest: input,
			plannerNotes: mergedPlanningContext?.plannerNotes,
			planningAnswers: mergedPlanningContext?.answers ?? [],
			recentConversation,
			repositoryContext,
			currentPlan,
		});
		const stageReadiness = this.getPlanningStageReadiness(questionStage, readiness, overrides?.focusHint, currentPlan);

		return {
			userRequest: input,
			modelId: this.resolvePlanningQuestionModelId(questionStage),
			planningPhase,
			questionStage,
			questionCount: stageReadiness.questionCount,
			missingDimensions: stageReadiness.missingDimensions,
			partialDimensions: stageReadiness.partialDimensions,
			shouldConfirmPlanningTarget: stageReadiness.shouldConfirmPlanningTarget,
			activeFilePath: model?.uri.toString(),
			selectedText,
			plannerNotes: mergedPlanningContext?.plannerNotes,
			recentConversation,
			planningAnswers: mergedPlanningContext?.answers ?? [],
			repositoryContext,
			currentPlan,
			focusAreaLabel: overrides?.focusAreaLabel?.trim(),
			focusHint: this.buildPlanningFocusHint(questionStage, repositoryContext, currentPlan, overrides?.focusAreaLabel, overrides?.focusHint),
		};
	}

	private async generatePlanningQuestionsWithTimeout(generationContext: IPlanningQuestionGenerationContext, timeoutMs = planningQuestionGenerationTimeoutMs): Promise<IGeneratedPlanningQuestionsResult> {
		const result = await raceTimeout(
			generateDynamicPlanningQuestionsResult(this.languageModelsService, generationContext, CancellationToken.None),
			timeoutMs,
			() => this.logService.warn(`[Planning] Timed out generating ${generationContext.questionStage} questions.`)
		);
		if (!result) {
			throw new Error(localize(
				'chat.dynamicPlanning.questionGenerationTimedOut',
				'Planning question generation took too long.'
			));
		}

		return result;
	}

	private async generatePlanningPlanStepControlsWithTimeout(generationContext: IPlanningQuestionGenerationContext, planSteps: readonly IPlanningPlanStep[], timeoutMs = planningQuestionGenerationTimeoutMs): Promise<IGeneratedPlanningPlanStepControlsResult> {
		const result = await raceTimeout(
			generateDynamicPlanningPlanStepControlsResult(this.languageModelsService, generationContext, planSteps, CancellationToken.None),
			timeoutMs,
			() => this.logService.warn(`[Planning] Timed out generating ${generationContext.questionStage} step controls.`)
		);
		if (!result) {
			throw new Error(localize(
				'chat.dynamicPlanning.stepControlGenerationTimedOut',
				'Planning control generation took too long.'
			));
		}

		return result;
	}

	private resolvePlanningQuestionModelId(questionStage: PlanningQuestionStage): string | undefined {
		const currentModelId = this.input.currentLanguageModel;
		if (this.isExecutablePlanningQuestionModelId(currentModelId)) {
			return currentModelId;
		}

		if (questionStage !== 'goal-clarity' && this.isExecutablePlanningQuestionModelId(this._lastPlanningQuestionModelId)) {
			return this._lastPlanningQuestionModelId;
		}

		return currentModelId;
	}

	private isExecutablePlanningQuestionModelId(modelId: string | undefined): modelId is string {
		return !!modelId && modelId !== 'copilot/auto';
	}

	private getCurrentPlanningInput(): string {
		return this.getInput().trim() || this._lastPlanningQuestionSourceInput?.trim() || '';
	}

	private getRecentPlanningConversation(): string[] {
		const items = this.viewModel?.getItems() ?? [];
		const entries: string[] = [];

		for (const item of items.slice(-6)) {
			if (isRequestVM(item)) {
				if (item.id === this._pendingPlanningPlaceholderRequestId || item.isCompleteAddedRequest) {
					continue;
				}
				if (item.messageText.trim()) {
					entries.push(`User: ${item.messageText.trim()}`);
				}
				continue;
			}

			if (isResponseVM(item)) {
				if (item.requestId === this._pendingPlanningPlaceholderRequestId || item.isCompleteAddedRequest) {
					continue;
				}
				const responseText = (extractPlanningPlanText(item.response) ?? item.response.getMarkdown() ?? item.response.toString()).replace(/\s+/g, ' ').trim();
				if (responseText) {
					entries.push(`Assistant: ${responseText.slice(0, 400)}`);
				}
			}
		}

		return entries.slice(-4);
	}

	private getPlanningResponseTextForRequest(requestId: string | undefined): string | undefined {
		if (!requestId || !this.viewModel) {
			return undefined;
		}

		const cachedPlanText = this._planningPlanTextByRequestId.get(requestId);
		if (cachedPlanText) {
			return cachedPlanText;
		}

		const request = this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().find(candidate => candidate.id === requestId);
		const planText = this.getCanonicalPlanningPlanText(request?.response);
		if (planText) {
			this._planningPlanTextByRequestId.set(requestId, planText);
		}
		return planText;
	}

	private getCanonicalPlanningPlanText(response: IChatResponseModel | undefined): string | undefined {
		if (!response || response.isCompleteAddedRequest) {
			return undefined;
		}

		return extractPlanningPlanText(response.entireResponse)
			?? extractPlanningPlanText(response.response);
	}

	private capturePlanningPlanSnapshot(response: IChatResponseModel | undefined): IPlanningPlanSnapshot | undefined {
		if (!response || response.isCompleteAddedRequest) {
			return undefined;
		}

		const planText = this.getCanonicalPlanningPlanText(response);
		if (!planText) {
			return undefined;
		}

		this._planningPlanTextByRequestId.set(response.requestId, planText);
		const previousRequestId = this._currentPlanningPlanRequestId;
		const previousPlanText = this.getPlanningResponseTextForRequest(previousRequestId);
		this._previousPlanningPlanRequestId = previousRequestId;
		this._currentPlanningPlanRequestId = response.requestId;

		return {
			requestId: response.requestId,
			planText,
			...(previousRequestId ? { previousRequestId } : {}),
			...(previousPlanText ? { previousPlanText } : {}),
		};
	}

	private getLatestPlanningPlanSnapshot(): IPlanningPlanSnapshot | undefined {
		const currentRequestId = this._currentPlanningPlanRequestId;
		const currentPlanText = this.getPlanningResponseTextForRequest(currentRequestId);
		if (currentRequestId && currentPlanText) {
			const previousPlanText = this.getPlanningResponseTextForRequest(this._previousPlanningPlanRequestId);
			return {
				requestId: currentRequestId,
				planText: currentPlanText,
				...(this._previousPlanningPlanRequestId ? { previousRequestId: this._previousPlanningPlanRequestId } : {}),
				...(previousPlanText ? { previousPlanText } : {}),
			};
		}

		const latestPlanningResponse = this.getLatestPlanningResponseModel();
		return latestPlanningResponse ? this.capturePlanningPlanSnapshot(latestPlanningResponse) : undefined;
	}

	private getCurrentPlanningResponseText(): string | undefined {
		return this.getLatestPlanningPlanSnapshot()?.planText
			?? this.getCanonicalPlanningPlanText(this.getLatestPlanningResponseModel());
	}

	private getPlanningPlanEditorStepsFromText(planText: string | undefined): IChatPlanningPlanEditorStep[] {
		return extractPlanningPlanSteps(planText, 24).map(step => ({
			id: `step-${step.index}`,
			index: step.index,
			label: step.label,
			text: step.text,
			sectionTitle: step.sectionTitle,
			kind: step.kind,
		}));
	}

	private updatePlanningPlanCanvas(requestId: string, planText: string, previousPlanText: string | undefined, isComplete: boolean): void {
		if (!this.viewModel || !planText.trim()) {
			return;
		}

		void this.commandService.executeCommand(UPDATE_PLANNING_PLAN_ACTION_ID, {
			sessionResource: this.viewModel.sessionResource.toString(),
			requestId,
			planText,
			previousPlanText,
			planSteps: this.getPlanningPlanEditorStepsFromText(planText),
			isComplete,
		}).catch(error => this.logService.warn('[Planning] Failed to update plan canvas.', error));
	}

	private getPlanningContextEditors(): ICodeEditor[] {
		const seen = new Set<string>();
		const editors: ICodeEditor[] = [];
		for (const editor of [
			this.codeEditorService.getFocusedCodeEditor(),
			this.codeEditorService.getActiveCodeEditor(),
			...this.codeEditorService.listCodeEditors()
		]) {
			const resource = editor?.getModel()?.uri;
			if (!editor || !resource || this.isChatInternalEditor(resource)) {
				continue;
			}

			const key = resource.toString();
			if (seen.has(key)) {
				continue;
			}

			seen.add(key);
			editors.push(editor);
		}

		return editors;
	}

	private getPlanningStageReadiness(
		questionStage: PlanningQuestionStage,
		readiness: ReturnType<typeof assessPlanningReadiness>,
		focusHint: string | undefined,
		currentPlan: string | undefined,
	) {
		if (questionStage === 'goal-clarity') {
			return readiness.goalClarity;
		}

		if (questionStage === 'plan-focus') {
			return {
				...readiness.taskDecomposition,
				questionCount: focusHint?.trim() ? Math.min(Math.max(readiness.taskDecomposition.questionCount + 2, 3), 4) : 1,
				shouldConfirmPlanningTarget: false,
			};
		}

		const hasCurrentPlan = !!currentPlan?.trim();
		return {
			...readiness.taskDecomposition,
			questionCount: hasCurrentPlan
				? Math.max(readiness.taskDecomposition.questionCount, 3)
				: Math.min(Math.max(readiness.taskDecomposition.questionCount, 2), 3),
		};
	}

	private getGoalClarityContinuationDecisionBeforeFirstPlan(originalQuery: string, planningContext: IPlanningTransitionContext, forceProceedToPlan = false): IGoalClarityContinuationDecision {
		if (this._currentPlanningPlanRequestId && this.getLatestPlanningPlanSnapshot()?.planText.trim()) {
			return { shouldContinue: false };
		}

		if (forceProceedToPlan) {
			return { shouldContinue: false };
		}

		if (planningContext.answers.length === 0) {
			return { shouldContinue: false };
		}

		const readiness = assessPlanningReadiness({
			userRequest: originalQuery,
			plannerNotes: planningContext.plannerNotes,
			planningAnswers: planningContext.answers,
			recentConversation: planningContext.recentConversation ?? [],
			repositoryContext: planningContext.repositoryContext,
		}).goalClarity;
		const hasAssumptionReview = this.hasGoalClarityAssumptionAnswer(planningContext);
		const requiresMinimumFollowup = this._goalClarityQuestionRounds < minGoalClarityQuestionRoundsBeforeFirstPlan;
		const hasUnsettledGoalDimension = readiness.missingDimensions.length > 0
			|| readiness.partialDimensions.some(dimension => dimension !== 'repo-target');
		const shouldContinue = requiresMinimumFollowup || !hasAssumptionReview || hasUnsettledGoalDimension;

		return {
			shouldContinue,
			...(shouldContinue ? { focusHint: this.buildGoalClarityFollowupFocusHint(readiness, hasAssumptionReview, requiresMinimumFollowup) } : {}),
		};
	}

	private hasGoalClarityAssumptionAnswer(planningContext: IPlanningTransitionContext): boolean {
		return planningContext.answers.some(answer => /\b(assumption|assumptions|before.*plan|shape.*first plan)\b/i.test(`${answer.question} ${answer.answer}`));
	}

	private buildGoalClarityFollowupFocusHint(
		readiness: ReturnType<typeof assessPlanningReadiness>['goalClarity'],
		hasAssumptionReview: boolean,
		requiresMinimumFollowup: boolean
	): string {
		const dimensions = [...readiness.missingDimensions, ...readiness.partialDimensions.filter(dimension => dimension !== 'repo-target')];
		const hints: string[] = [];
		if (requiresMinimumFollowup) {
			hints.push('Ask another goal-clarity round before drafting the first plan. The plan should not be built until at least two rounds of goal context have been captured.');
		}
		if (dimensions.length) {
			hints.push(`Focus on plan-critical context still missing or partial: ${dimensions.join(', ')}.`);
		}
		if (!hasAssumptionReview) {
			hints.push('Include an editable assumptions review if assumptions would materially shape the first plan.');
		}
		hints.push('Ask only goal-context questions that the planning agent would need before it can build a concrete, accurate first plan. Do not ask task-breakdown, sequencing, or implementation-order questions in this round.');
		return hints.join(' ');
	}

	private async continueGoalClarityBeforeFirstPlan(
		originalQuery: string,
		options: IChatAcceptInputOptions,
		planningPhase: PlanningPhase,
		planningContext: IPlanningTransitionContext,
		focusHint: string | undefined
	): Promise<boolean> {
		await this.showPlanningGenerationPlaceholder(originalQuery, 'goal-clarity');

		try {
			const generationContext = await this.getPlanningQuestionGenerationContext(
				originalQuery,
				planningPhase,
				'goal-clarity',
				undefined,
				{
					planningContext,
					focusHint,
				}
			);
			const generationResult = await this.generateNonEmptyGoalClarityQuestionsWithRetry(generationContext);
			const questions = generationResult.questions;
			this._lastPlanningQuestionModelId = generationResult.modelId;
			await this.showGoalClarityFollowupQuestions(originalQuery, questions, options, generationContext);
			return true;
		} catch (error) {
			await this.showPlanningQuestionGenerationError('goal-clarity', error);
			return true;
		}
	}

	private async generateNonEmptyGoalClarityQuestionsWithRetry(generationContext: IPlanningQuestionGenerationContext): Promise<IGeneratedPlanningQuestionsResult> {
		const generationResult = await this.generatePlanningQuestionsWithTimeout(generationContext);
		if (generationResult.questions.length) {
			return generationResult;
		}

		this.logService.warn('[Planning] Goal clarity follow-up generation returned no questions. Retrying with stricter dynamic instructions.');
		const retryGenerationContext: IPlanningQuestionGenerationContext = {
			...generationContext,
			questionCount: Math.max(generationContext.questionCount ?? 2, 2),
			focusHint: [
				generationContext.focusHint,
				'The previous dynamic goal-clarity generation returned no usable questions. Generate at least two concrete, context-aware goal-clarity questions from the current repo context, prior answers, and recent chat. Do not return task-decomposition or implementation-order controls.'
			].filter(isDefined).join(' '),
		};
		const retryResult = await this.generatePlanningQuestionsWithTimeout(retryGenerationContext);
		if (retryResult.questions.length) {
			return retryResult;
		}

		throw new Error(localize(
			'chat.dynamicPlanning.emptyGoalClarityFollowupQuestions',
			'The planning agent did not return any usable goal-clarity questions.'
		));
	}

	private async showGoalClarityFollowupQuestions(
		originalQuery: string,
		questions: readonly IChatQuestion[],
		options: IChatAcceptInputOptions,
		generationContext: IPlanningQuestionGenerationContext
	): Promise<void> {
		this._planningPhase = generationContext.planningPhase;
		const carousel = this.createDynamicPlanningQuestionCarousel(this.withGoalClarityProceedChoice([...questions]), generationContext);
		this._pendingPlanningQuestionResolveId = carousel.resolveId;

		if (this.canShowSubmittedPlanningPlaceholder()) {
			try {
				if (await this.setSubmittedPlanningPlaceholderCarousel(originalQuery, carousel, options, generationContext)) {
					return;
				}
				throw this.createGoalClarityTranscriptCarouselError();
			} catch (error) {
				this.logService.warn('[Planning] Failed to show goal clarity follow-up questions in the transcript.', error);
				await this.showPlanningQuestionGenerationError('goal-clarity', error);
				return;
			}
		}

		this.renderGoalClarityQuestionCarouselInInput(originalQuery, carousel, options, generationContext);
	}

	private createGoalClarityTranscriptCarouselError(): Error {
		return new Error(localize(
			'chat.dynamicPlanning.goalClarityTranscriptCarouselFailed',
			'Generated goal-clarity questions could not be attached to the chat transcript.'
		));
	}

	private withGoalClarityProceedChoice(questions: IChatQuestion[]): IChatQuestion[] {
		if (this._goalClarityQuestionRounds + 1 < goalClarityRoundToOfferPlanProceed) {
			return questions;
		}

		if (questions.some(question => question.id === goalClarityProceedQuestionId)) {
			return questions;
		}

		return [
			...questions,
			{
				id: goalClarityProceedQuestionId,
				type: 'singleSelect',
				title: localize('chat.dynamicPlanning.proceedToPlanTitle', 'Proceed to Plan'),
				message: localize('chat.dynamicPlanning.proceedToPlanMessage', 'Do you want to draft the first plan now, or keep clarifying?'),
				description: localize('chat.dynamicPlanning.proceedToPlanDescription', 'You can proceed after three goal-clarity rounds even if some non-blocking details are still assumptions.'),
				required: false,
				allowFreeformInput: false,
				options: [
					{ id: 'proceed', label: localize('chat.dynamicPlanning.proceedToPlanNow', 'Draft the plan now'), value: goalClarityProceedNowValue },
					{ id: 'continue', label: localize('chat.dynamicPlanning.continueGoalClarity', 'Keep clarifying'), value: goalClarityContinueClarifyingValue },
				],
			}
		];
	}

	private hasGoalClarityProceedAnswer(answersRecord: IChatQuestionAnswers | undefined): boolean {
		const answer = answersRecord?.[goalClarityProceedQuestionId];
		if (typeof answer === 'string') {
			return answer === goalClarityProceedNowValue;
		}

		return typeof answer === 'object'
			&& answer !== null
			&& hasKey(answer, { selectedValue: true })
			&& answer.selectedValue === goalClarityProceedNowValue;
	}

	private buildPlanningFocusHint(
		questionStage: PlanningQuestionStage,
		repositoryContext: IPlanningQuestionGenerationContext['repositoryContext'],
		currentPlan: string | undefined,
		focusAreaLabel?: string,
		focusHint?: string,
	): string | undefined {
		const hints: string[] = [];
		const normalizedFocusAreaLabel = focusAreaLabel?.trim();
		const normalizedFocusHint = focusHint?.trim();
		const taskLens = repositoryContext?.taskLens;

		if (normalizedFocusAreaLabel) {
			hints.push(`User-selected focus area: ${normalizedFocusAreaLabel}`);
		}

		if (normalizedFocusHint) {
			hints.push(normalizedFocusHint);
		}

		if (taskLens?.taskSummary) {
			hints.push(`Task summary: ${taskLens.taskSummary}`);
		}

		if (repositoryContext?.planningTarget) {
			hints.push(`Repo focus: ${repositoryContext.planningTarget.label}`);
		}

		if (taskLens?.primaryArtifact ?? repositoryContext?.primaryArtifactHint) {
			hints.push(`Primary artifact: ${taskLens?.primaryArtifact ?? repositoryContext?.primaryArtifactHint}`);
		}

		const relatedArtifacts = taskLens?.secondaryArtifacts ?? repositoryContext?.relatedArtifactHints;
		if (relatedArtifacts?.length) {
			hints.push(`Adjacent artifacts: ${relatedArtifacts.slice(0, 3).join(', ')}`);
		}

		if (taskLens?.validationTargets?.length) {
			hints.push(`Validation targets: ${taskLens.validationTargets.slice(0, 3).join(', ')}`);
		}

		if (taskLens?.riskAreas?.length) {
			hints.push(`Guardrails: ${taskLens.riskAreas.slice(0, 3).join('; ')}`);
		}

		if (taskLens?.unknowns?.length) {
			hints.push(`Open decisions: ${taskLens.unknowns.slice(0, 3).join('; ')}`);
		}

		if (currentPlan && questionStage === 'task-decomposition') {
			hints.push('Refinement mode: tighten the existing plan');
		}

		if (currentPlan && questionStage === 'plan-focus') {
			hints.push('Refinement mode: sharpen one part of the existing plan');
		}

		return hints.length > 0 ? hints.join(' | ') : undefined;
	}

	private async refreshPlanningTransitionContextForStage(
		originalQuery: string,
		planningPhase: PlanningPhase,
		questionStage: PlanningQuestionStage,
		planningContext: IPlanningTransitionContext,
		options?: {
			readonly currentPlan?: string;
			readonly focusAreaLabel?: string;
			readonly focusHint?: string;
		}
	): Promise<{ readonly planningContext: IPlanningTransitionContext; readonly generationContext: IPlanningQuestionGenerationContext }> {
		const generationContext = await this.getPlanningQuestionGenerationContext(
			originalQuery,
			planningPhase,
			questionStage,
			undefined,
			{
				planningContext,
				currentPlan: options?.currentPlan,
				focusAreaLabel: options?.focusAreaLabel,
				focusHint: options?.focusHint,
			}
		);
		const refreshedPlanningContext = mergePlanningTransitionContexts(planningContext, {
			phase: generationContext.planningPhase,
			answers: [],
			plannerNotes: planningContext.plannerNotes,
			recentConversation: generationContext.recentConversation,
			repositoryContext: generationContext.repositoryContext,
		}) ?? planningContext;

		return {
			planningContext: refreshedPlanningContext,
			generationContext,
		};
	}

	private schedulePlanningFollowupAfterResponse(response: IChatResponseModel | undefined, callback: (response: IChatResponseModel) => Promise<void>): void {
		this._pendingPlanningResponseListener.clear();
		if (!response) {
			return;
		}

		const runCallback = () => {
			this._pendingPlanningResponseListener.clear();
			void callback(response).catch(error => {
				this.logService.error('[Planning] Failed to render planning follow-up UI', error);
			});
		};

		if (response.isComplete) {
			runCallback();
			return;
		}

		this._pendingPlanningResponseListener.value = Event.once(Event.filter(
			response.onDidChange,
			() => response.isComplete
		))(runCallback);
	}

	private async waitForPlanningResponseComplete(response: IChatResponseModel): Promise<IChatResponseModel> {
		if (!response.isComplete) {
			await Event.toPromise(Event.once(Event.filter(
				response.onDidChange,
				() => response.isComplete
			)));
		}

		return response;
	}

	private getPendingPlanningPlaceholderRequest() {
		if (!this._pendingPlanningPlaceholderRequestId || !this.viewModel) {
			return undefined;
		}

		return this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().find(candidate => candidate.id === this._pendingPlanningPlaceholderRequestId);
	}

	private pendingPlanningPlaceholderHasUsedPlanningCarousel(): boolean {
		return this.getPendingPlanningPlaceholderRequest()?.response?.response.value.some(content => {
			return isUsedQuestionCarousel(content) && isPlanningMiddlewareQuestionCarousel(content.resolveId);
		}) ?? false;
	}

	private async clearPendingPlanningPlaceholder(options: { readonly preserveUsedPlanningCarousel?: boolean } = {}): Promise<void> {
		if (!this._pendingPlanningPlaceholderRequestId || !this.viewModel) {
			return;
		}

		if (options.preserveUsedPlanningCarousel !== false && this.pendingPlanningPlaceholderHasUsedPlanningCarousel()) {
			this._pendingPlanningPlaceholderRequestId = undefined;
			return;
		}

		await this.chatService.removeRequest(this.viewModel.sessionResource, this._pendingPlanningPlaceholderRequestId);
		this._pendingPlanningPlaceholderRequestId = undefined;
	}

	private hidePlanningRequestPromptFromTranscript(requestId: string | undefined): void {
		if (!requestId || !this.viewModel) {
			return;
		}

		const request = this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().find(candidate => candidate.id === requestId);
		if (request) {
			request.shouldBeRemovedOnSend = { requestId };
		}
	}

	private addCompletePlanningTranscriptRequest(message: IChatProgress[], options: { readonly preserveScroll?: boolean } = {}): string | undefined {
		if (!this.viewModel) {
			return undefined;
		}

		if (options.preserveScroll === false) {
			this.chatService.addCompleteRequest(this.viewModel.sessionResource, '', undefined, 0, {
				message,
				result: {},
			});
			return this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().at(-1)?.id;
		}

		const scrollTop = this.listWidget.scrollTop;
		this._planningTranscriptScrollPreservationDepth++;
		this.listWidget.suppressAutoScroll = true;
		try {
			this.chatService.addCompleteRequest(this.viewModel.sessionResource, '', undefined, 0, {
				message,
				result: {},
			});
		} finally {
			this.listWidget.suppressAutoScroll = false;
			this._planningTranscriptScrollPreservationDepth--;
			this.listWidget.scrollTop = scrollTop;
		}

		return this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().at(-1)?.id;
	}

	private async replacePlanningProgressPlaceholder(message: IChatProgress[]): Promise<void> {
		if (!this.viewModel) {
			return;
		}

		const scrollTop = this.listWidget.scrollTop;
		this._planningTranscriptScrollPreservationDepth++;
		this.listWidget.suppressAutoScroll = true;
		try {
			if (this._pendingPlanningPlaceholderRequestId) {
				if (this.pendingPlanningPlaceholderHasUsedPlanningCarousel()) {
					this._pendingPlanningPlaceholderRequestId = undefined;
				} else {
					await this.chatService.removeRequest(this.viewModel.sessionResource, this._pendingPlanningPlaceholderRequestId);
					this._pendingPlanningPlaceholderRequestId = undefined;
				}
			}

			this.chatService.addCompleteRequest(this.viewModel.sessionResource, '', undefined, 0, {
				message,
				result: {},
			});
			this._pendingPlanningPlaceholderRequestId = this.chatService.getSession(this.viewModel.sessionResource)?.getRequests().at(-1)?.id;
		} finally {
			this.listWidget.suppressAutoScroll = false;
			this._planningTranscriptScrollPreservationDepth--;
			this.listWidget.scrollTop = scrollTop;
		}
	}

	private createPlanningProgressContent(
		kind: 'first-plan' | 'updated-plan',
		source: PlanningPlanProgressSource,
		progress?: IPlanningPlanProgressSnapshot,
	): MarkdownString {
		const content = new MarkdownString(undefined, { supportThemeIcons: true, supportHtml: true, isTrusted: true });
		content.appendMarkdown('$(sync~spin) ');
		content.appendMarkdown(kind === 'first-plan'
			? localize('chat.dynamicPlanning.generatingFirstPlan', '**Building Plan**\n\nStarting the planner response. Reasoning, tool progress, and confirmations will appear below as soon as the agent begins.')
			: source === 'task-decomposition'
				? localize('chat.dynamicPlanning.generatingUpdatedPlanFromTaskDecomposition', '**Updating Plan**\n\nStarting the planner response from your plan-shaping decisions.')
				: source === 'plan-review'
					? localize('chat.dynamicPlanning.generatingUpdatedPlanFromPlanReview', '**Updating Plan**\n\nStarting the planner response from your step edits.')
					: source === 'plan-focus'
						? localize('chat.dynamicPlanning.generatingUpdatedPlanFromFocus', '**Updating Plan**\n\nStarting the planner response for the selected plan area.')
						: localize('chat.dynamicPlanning.generatingUpdatedPlan', '**Updating Plan**\n\nStarting the planner response from your latest decisions.'));
		if (progress) {
			content.appendMarkdown('\n\n');
			content.appendMarkdown(localize('chat.dynamicPlanning.planProgressEstimate', 'Estimated progress: {0}% - elapsed {1}', Math.floor(progress.percent), progress.elapsedLabel));
			content.appendMarkdown(`\n\n${progress.progressBar}`);
			content.appendMarkdown('\n\n');
			content.appendText(progress.status);
		}
		return content;
	}

	private async showPlanningProgressPlaceholder(
		kind: 'first-plan' | 'updated-plan',
		source: PlanningPlanProgressSource = 'goal-clarity',
		progress?: IPlanningPlanProgressSnapshot,
	): Promise<void> {
		if (!this.viewModel || !this.isInPlanningMode()) {
			return;
		}

		await this.replacePlanningProgressPlaceholder([{
			kind: 'markdownContent',
			content: this.createPlanningProgressContent(kind, source, progress),
		}]);
	}

	private startPlanningProgressTracker(
		kind: 'first-plan' | 'updated-plan',
		source: PlanningPlanProgressSource
	): IPlanningPlanProgressTracker {
		const startedAt = Date.now();
		const targetWindow = dom.getWindow(this.container);
		let response: IChatResponseModel | undefined;
		let interval: number | undefined;
		let isDisposed = false;
		let isRefreshing = false;
		const stopInterval = () => {
			if (interval !== undefined) {
				targetWindow.clearInterval(interval);
				interval = undefined;
			}
		};

		const refresh = () => {
			if (isDisposed || isRefreshing) {
				return;
			}

			isRefreshing = true;
			const progress = this.getPlanningProgressSnapshot(startedAt, response);
			void this.updatePlanningProgress(progress, response, kind, source).finally(() => {
				isRefreshing = false;
			});
		};

		refresh();
		interval = targetWindow.setInterval(refresh, planningPlanProgressUpdateIntervalMs);
		return {
			setResponse: value => {
				response = value;
				void this.clearPendingPlanningPlaceholder().finally(refresh);
			},
			complete: async status => {
				isDisposed = true;
				stopInterval();
				await this.updatePlanningProgress(this.getPlanningProgressSnapshot(startedAt, response, status), response, kind, source);
			},
			dispose: () => {
				isDisposed = true;
				stopInterval();
			},
		};
	}

	private async updatePlanningProgress(
		progress: IPlanningPlanProgressSnapshot,
		response: IChatResponseModel | undefined,
		kind: 'first-plan' | 'updated-plan',
		source: PlanningPlanProgressSource
	): Promise<void> {
		const progressMessage = this.createPlanningProgressMessage(kind, source, progress);
		const request = response?.request;
		if (request) {
			this.chatService.appendProgress(request, progressMessage);
			return;
		}

		await this.showPlanningProgressPlaceholder(kind, source, progress);
	}

	private createPlanningProgressMessage(
		kind: 'first-plan' | 'updated-plan',
		source: PlanningPlanProgressSource,
		progress: IPlanningPlanProgressSnapshot
	): IChatProgressMessage {
		return {
			kind: 'progressMessage',
			id: planningPlanProgressMessageId,
			isSticky: true,
			shimmer: true,
			content: this.createPlanningProgressContent(kind, source, progress),
		};
	}

	private getPlanningProgressSnapshot(startedAt: number, response: IChatResponseModel | undefined, completedStatus?: string): IPlanningPlanProgressSnapshot {
		const elapsedMs = Date.now() - startedAt;
		const percent = completedStatus ? 100 : this.getEstimatedPlanningProgressPercent(elapsedMs, response);
		return {
			percent,
			elapsedLabel: this.formatPlanningProgressElapsed(elapsedMs),
			status: completedStatus ?? this.getPlanningProgressStatus(response),
			progressBar: this.createPlanningProgressBar(percent),
		};
	}

	private getEstimatedPlanningProgressPercent(elapsedMs: number, response: IChatResponseModel | undefined): number {
		const elapsedSeconds = Math.max(0, elapsedMs / 1000);
		const timeEstimate = 99 * (1 - Math.exp(-elapsedSeconds / 60));
		if (!response) {
			return Math.min(25, Math.max(0, timeEstimate));
		}

		if (this.getCanonicalPlanningPlanText(response)?.trim()) {
			const planAvailableEstimate = 90 + (9 * (1 - Math.exp(-Math.max(0, elapsedSeconds - 5) / 25)));
			return Math.min(99, Math.max(90, planAvailableEstimate));
		}

		if (this.planningResponseHasPendingConfirmation(response)) {
			return Math.min(85, Math.max(40, timeEstimate));
		}

		if (this.planningResponseHasVisibleActivity(response)) {
			return Math.min(88, Math.max(25, timeEstimate));
		}

		return Math.min(55, Math.max(0, timeEstimate));
	}

	private getPlanningProgressStatus(response: IChatResponseModel | undefined): string {
		if (!response) {
			return localize('chat.dynamicPlanning.planProgressWaitingForResponse', 'Waiting for the planner response to start.');
		}

		if (this.planningResponseHasPendingConfirmation(response)) {
			return localize('chat.dynamicPlanning.planProgressWaitingForConfirmation', 'Waiting for your confirmation in the planner response below.');
		}

		if (this.getCanonicalPlanningPlanText(response)?.trim()) {
			return localize('chat.dynamicPlanning.planProgressPlanAvailable', 'Plan text is available in the canvas; finishing the planner response.');
		}

		if (this.planningResponseHasVisibleActivity(response)) {
			return localize('chat.dynamicPlanning.planProgressActivityVisible', 'Planner reasoning, tool progress, or approvals should be visible in the response below.');
		}

		return localize('chat.dynamicPlanning.planProgressWaitingForFirstUpdate', 'Planner response is live; waiting for the first visible reasoning or tool update.');
	}

	private planningResponseHasPendingConfirmation(response: IChatResponseModel): boolean {
		return response.response.value.some(part => {
			if (part.kind !== 'toolInvocation') {
				return part.kind === 'confirmation' && !part.isUsed;
			}

			const state = part.state.get();
			return state.type === IChatToolInvocation.StateKind.WaitingForConfirmation
				|| state.type === IChatToolInvocation.StateKind.WaitingForPostApproval;
		});
	}

	private planningResponseHasVisibleActivity(response: IChatResponseModel): boolean {
		return response.response.value.some(part => {
			if (part.kind === 'progressMessage') {
				return part.id !== planningPlanProgressMessageId && !!part.content.value.trim();
			}

			if (part.kind === 'markdownContent' || part.kind === 'warning') {
				return !!part.content.value.trim();
			}

			if (part.kind === 'thinking') {
				const value = Array.isArray(part.value) ? part.value.join('') : part.value;
				return !!value?.trim();
			}

			return part.kind !== 'undoStop';
		});
	}

	private createPlanningProgressBar(percent: number): string {
		const clampedPercent = Math.min(100, Math.max(0, percent));
		const totalEighths = Math.floor((clampedPercent / 100) * planningPlanProgressBarSegments * 8);
		const filledSegments = Math.min(planningPlanProgressBarSegments, Math.floor(totalEighths / 8));
		const partialEighths = filledSegments < planningPlanProgressBarSegments ? totalEighths % 8 : 0;
		const partialSegments = ['&#9615;', '&#9614;', '&#9613;', '&#9612;', '&#9611;', '&#9610;', '&#9609;'];
		const partial = partialEighths > 0 ? partialSegments[partialEighths - 1] : '';
		const emptySegments = Math.max(0, planningPlanProgressBarSegments - filledSegments - (partial ? 1 : 0));
		const filled = '&#9608;'.repeat(filledSegments) + partial;
		const empty = '&#9608;'.repeat(emptySegments);
		const filledBar = filled
			? `<span style="color:var(--vscode-charts-blue);">${filled}</span>`
			: '';
		const emptyBar = empty
			? `<span style="color:var(--vscode-editorWidget-border);">${empty}</span>`
			: '';
		return `${filledBar}${emptyBar}`;
	}

	private formatPlanningProgressElapsed(elapsedMs: number): string {
		const totalSeconds = Math.max(0, Math.floor(elapsedMs / 1000));
		const minutes = Math.floor(totalSeconds / 60);
		const seconds = totalSeconds % 60;
		return minutes > 0
			? localize('chat.dynamicPlanning.elapsedMinutesSeconds', '{0}m {1}s', minutes, seconds)
			: localize('chat.dynamicPlanning.elapsedSeconds', '{0}s', seconds);
	}

	private async showPlanningGenerationPlaceholder(
		_originalQuery: string,
		questionStage: PlanningQuestionStage,
	): Promise<void> {
		if (!this.viewModel || !this.isInPlanningMode()) {
			return;
		}

		this.addCompletePlanningTranscriptRequest([this.createPlanningMiddlewareIntroContent(questionStage, 'generating')]);
	}

	private async showPlanningQuestionGenerationError(questionStage: PlanningQuestionStage, error: unknown): Promise<void> {
		this.logService.error('[Planning] Failed to generate planning questions', error);
		this._pendingPlanningQuestionAnswersListener.clear();
		this._pendingPlanningQuestionResolveId = undefined;
		await this.clearPendingPlanningPlaceholder();
		const stageLabel = questionStage === 'goal-clarity'
			? localize('chat.dynamicPlanning.goalClarityStageLabel', 'Clarifying Your Goals')
			: questionStage === 'task-decomposition'
				? localize('chat.dynamicPlanning.taskDecompositionStageLabel', 'Shaping the Plan')
				: localize('chat.dynamicPlanning.planFocusStageLabel', 'Refining a Plan Area');
		await this.dialogService.error(
			localize(
				'chat.dynamicPlanning.questionGenerationErrorTitle',
				'Unable to Generate {0} Questions',
				stageLabel
			),
			toErrorMessage(error)
		);
	}

	private async submitPlanningRequestWithContext(
		originalQuery: string,
		options: IChatAcceptInputOptions,
		planningContext: IPlanningTransitionContext | undefined,
		allowPlannerFollowupQuestions: boolean,
		onResponseComplete?: (response: IChatResponseModel, planSnapshot?: IPlanningPlanSnapshot) => Promise<void>,
		progressKind?: 'first-plan' | 'updated-plan',
		progressSource: PlanningPlanProgressSource = 'goal-clarity',
	): Promise<void> {
		this._skipDynamicPlanningQuestionsOnce = true;
		this._pendingPlanningQuestionAnswersListener.clear();

		const extraAttachedContext = planningContext ? [this.createPlanningContextAttachment(planningContext, {
			originalQuery,
			progressKind,
			progressSource,
			allowPlannerFollowupQuestions,
		})] : undefined;
		const userSelectedToolsOverride = this.getPlanningSubmissionToolOverrides(options.userSelectedToolsOverride, allowPlannerFollowupQuestions);
		let canvasUpdateListener: IDisposable | undefined;
		let progressTracker: IPlanningPlanProgressTracker | undefined;

		try {
			if (progressKind) {
				progressTracker = this.startPlanningProgressTracker(progressKind, progressSource);
			} else {
				await this.clearPendingPlanningPlaceholder();
			}

			const response = await this._acceptInput(
				{ query: originalQuery },
				{ ...options, extraAttachedContext, storeToHistory: false, userSelectedToolsOverride }
			);
			if (!response) {
				progressTracker?.dispose();
				canvasUpdateListener?.dispose();
				if (progressKind) {
					await this.clearPendingPlanningPlaceholder();
				}
				return;
			}
			if (progressKind) {
				progressTracker?.setResponse(response);
			}
			const canvasPreviousPlanText = progressKind ? this.getLatestPlanningPlanSnapshot()?.planText : undefined;
			let lastCanvasPlanText: string | undefined;
			if (progressKind) {
				canvasUpdateListener = response.onDidChange(() => {
					const planText = this.getCanonicalPlanningPlanText(response);
					if (!planText || planText === lastCanvasPlanText) {
						return;
					}

					lastCanvasPlanText = planText;
					this.updatePlanningPlanCanvas(response.requestId, planText, canvasPreviousPlanText, false);
				});
			}
			if (onResponseComplete || progressKind) {
				this.schedulePlanningFollowupAfterResponse(response, async completedResponse => {
					canvasUpdateListener?.dispose();
					canvasUpdateListener = undefined;
					let responseForCompletion = completedResponse;
					let planSnapshot = progressKind ? this.capturePlanningPlanSnapshot(completedResponse) : undefined;
					if (progressKind && !planSnapshot) {
						let retryResult: IPlanningPlanRequestResult | undefined;
						try {
							retryResult = await this.retryPlanningRequestForMissingPlan(
								originalQuery,
								options,
								planningContext,
								allowPlannerFollowupQuestions,
								progressKind,
								progressSource,
								progressTracker,
							);
						} catch (error) {
							this.logService.error('[Planning] Markdown output retry failed.', error);
						}
						if (retryResult) {
							responseForCompletion = retryResult.response;
							planSnapshot = retryResult.planSnapshot;
						}
					}
					if (planSnapshot) {
						this.updatePlanningPlanCanvas(planSnapshot.requestId, planSnapshot.planText, planSnapshot.previousPlanText, true);
					}
					if (progressKind) {
						if (planSnapshot) {
							await progressTracker?.complete(localize('chat.dynamicPlanning.planProgressComplete', 'Plan ready; the canvas is up to date.'));
							await this.clearPendingPlanningPlaceholder();
						} else {
							this.logService.warn('[Planning] Planner response completed without text output after retry. Keeping the response visible.');
							await progressTracker?.complete(localize('chat.dynamicPlanning.planProgressNoPlan', 'Planner stopped without a usable plan. The response remains visible for review.'));
							await this.clearPendingPlanningPlaceholder();
						}
					}
					progressTracker?.dispose();
					if (!progressKind || planSnapshot) {
						await onResponseComplete?.(responseForCompletion, planSnapshot);
					}
				});
			}
		} catch (error) {
			progressTracker?.dispose();
			canvasUpdateListener?.dispose();
			if (progressKind) {
				await this.clearPendingPlanningPlaceholder();
			}
			this.logService.error('[Planning] Failed to submit dynamic planning questions', error);
		}
	}

	private async retryPlanningRequestForMissingPlan(
		originalQuery: string,
		options: IChatAcceptInputOptions,
		planningContext: IPlanningTransitionContext | undefined,
		allowPlannerFollowupQuestions: boolean,
		progressKind: 'first-plan' | 'updated-plan',
		progressSource: PlanningPlanProgressSource,
		progressTracker?: IPlanningPlanProgressTracker,
	): Promise<IPlanningPlanRequestResult | undefined> {
		this.logService.warn('[Planning] Planner response completed without text output. Retrying once for markdown output.');
		this._skipDynamicPlanningQuestionsOnce = true;

		const extraAttachedContext = planningContext ? [this.createPlanningContextAttachment(planningContext, {
			originalQuery,
			progressKind,
			progressSource,
			allowPlannerFollowupQuestions,
			isRetry: true,
		})] : undefined;
		const userSelectedToolsOverride = this.getPlanningSubmissionToolOverrides(options.userSelectedToolsOverride, allowPlannerFollowupQuestions);
		const response = await this._acceptInput(
			{ query: originalQuery },
			{ ...options, extraAttachedContext, storeToHistory: false, userSelectedToolsOverride }
		);
		if (!response) {
			return undefined;
		}
		progressTracker?.setResponse(response);
		const completedResponse = await this.waitForPlanningResponseComplete(response);
		return {
			response: completedResponse,
			planSnapshot: this.capturePlanningPlanSnapshot(completedResponse),
		};
	}

	private async showPlanningPlanEditorWithInlineQuestions(
		originalQuery: string,
		options: IChatAcceptInputOptions,
		planningPhase: PlanningPhase,
		planSnapshot?: IPlanningPlanSnapshot,
		lastFocusAreaLabel?: string,
		focusHint?: string,
	): Promise<void> {
		const currentSnapshot = planSnapshot ?? this.getLatestPlanningPlanSnapshot();
		const planText = currentSnapshot?.planText ?? this.getCurrentPlanningResponseText();
		const planningContext = this.getPlanningTransitionContextForCurrentResponse() ?? this._planningTransitionContext;
		if (!planText?.trim() || !planningContext) {
			this.showPlanningPlanEditor(originalQuery, options, planningPhase, currentSnapshot, lastFocusAreaLabel);
			return;
		}

		this.showPlanningPlanEditor(originalQuery, options, planningPhase, currentSnapshot, lastFocusAreaLabel);
		try {
			const { planningContext: refreshedPlanningContext, generationContext } = await this.refreshPlanningTransitionContextForStage(
				originalQuery,
				planningPhase,
				'task-decomposition',
				planningContext,
				{
					currentPlan: planText,
					focusAreaLabel: lastFocusAreaLabel,
					focusHint,
				}
			);
			this._planningTransitionContext = refreshedPlanningContext;

			const planSteps = extractPlanningPlanSteps(planText, 24);
			const generationResult = await this.generatePlanningPlanStepControlsWithTimeout(generationContext, planSteps);
			this._lastPlanningQuestionModelId = generationResult.modelId;
			this.showPlanningPlanEditor(
				originalQuery,
				options,
				planningPhase,
				currentSnapshot,
				lastFocusAreaLabel,
				undefined,
				generationResult.questionsByStep
			);
		} catch (error) {
			this.logService.warn('[Planning] Failed to prepare inline plan questions. Showing the plan editor without generated questions.', error);
			this.showPlanningPlanEditor(originalQuery, options, planningPhase, currentSnapshot, lastFocusAreaLabel);
		}
	}

	private showPlanningPlanEditor(
		originalQuery: string,
		options: IChatAcceptInputOptions,
		planningPhase: PlanningPhase,
		planSnapshot?: IPlanningPlanSnapshot,
		lastFocusAreaLabel?: string,
		stepQuestions?: readonly IChatQuestion[],
		stepQuestionsByStep?: ReadonlyMap<number, readonly IChatQuestion[]>,
	): void {
		const currentSnapshot = planSnapshot ?? this.getLatestPlanningPlanSnapshot();
		const planText = currentSnapshot?.planText ?? this.getCurrentPlanningResponseText();
		const visiblePlanText = planText && isUsablePlanningPlanText(planText) ? planText : undefined;
		const extractedPlanSteps = extractPlanningPlanSteps(visiblePlanText, 24);
		const planSteps = extractedPlanSteps.length > 0
			? extractedPlanSteps
			: visiblePlanText?.trim()
				? [{
					index: 1,
					label: localize('chat.planEditor.wholePlanFallbackLabel', 'Review the whole plan'),
					text: visiblePlanText.trim(),
					kind: 'step' as const,
				}]
				: [];
		if (planSteps.length === 0) {
			return;
		}

		const resolveId = `${planningMiddlewareQuestionCarouselResolveIdPrefix}plan-editor-${Date.now()}`;
		const questionsByStep = stepQuestionsByStep ?? this.assignPlanningQuestionsToPlanSteps(planSteps, stepQuestions);
		const editor: IChatPlanningPlanEditor = {
			kind: 'planningPlanEditor',
			planText: visiblePlanText ?? '',
			steps: planSteps.map(step => ({
				id: `step-${step.index}`,
				index: step.index,
				label: step.label,
				text: step.text,
				sectionTitle: step.sectionTitle,
				kind: step.kind,
				...(questionsByStep.get(step.index)?.length ? { questions: [...(questionsByStep.get(step.index) ?? [])] } : {}),
			})),
			resolveId,
		};

		void this.showPlanningPlanEditorRequest(originalQuery, editor, async answersRecord => {
			await this.submitPlanningPlanEditorAnswers(
				originalQuery,
				editor,
				answersRecord,
				options,
				planningPhase,
				currentSnapshot,
				lastFocusAreaLabel
			);
		}, currentSnapshot, lastFocusAreaLabel);
	}

	private assignPlanningQuestionsToPlanSteps(planSteps: readonly IPlanningPlanStep[], questions: readonly IChatQuestion[] | undefined): Map<number, IChatQuestion[]> {
		const questionsByStep = new Map<number, IChatQuestion[]>();
		if (!questions?.length || planSteps.length === 0) {
			return questionsByStep;
		}

		let fallbackIndex = 0;
		for (const question of questions) {
			const scoredSteps = planSteps
				.map((step, index) => ({
					step,
					index,
					score: this.scorePlanningQuestionStepMatch(question, step),
				}))
				.sort((left, right) => right.score - left.score || left.index - right.index);
			const selectedStep = scoredSteps[0]?.score > 0
				? scoredSteps[0].step
				: planSteps[fallbackIndex++ % planSteps.length];
			const existing = questionsByStep.get(selectedStep.index) ?? [];
			existing.push(question);
			questionsByStep.set(selectedStep.index, existing);
		}

		return questionsByStep;
	}

	private scorePlanningQuestionStepMatch(question: IChatQuestion, step: IPlanningPlanStep): number {
		const questionText = this.getPlanningQuestionSearchText(question);
		const stepText = [step.sectionTitle, step.label, step.text].filter((value): value is string => !!value).join(' ');
		let score = this.computeTokenOverlap(questionText, stepText);
		if (step.kind === 'verification' && /\b(test|tests|verify|verification|validate|validation|check|evidence|assert|coverage)\b/i.test(questionText)) {
			score += 4;
		}
		if (step.kind === 'decision' && /\b(decide|decision|choice|choose|clarify|open question|resolve)\b/i.test(questionText)) {
			score += 4;
		}
		if (step.kind === 'guardrail' && /\b(risk|guardrail|constraint|assumption|boundary|preserve|avoid)\b/i.test(questionText)) {
			score += 4;
		}
		return score;
	}

	private getPlanningQuestionSearchText(question: IChatQuestion): string {
		const message = typeof question.message === 'string' ? question.message : question.message?.value;
		return [
			question.title,
			message,
			question.description,
			...(question.options?.map(option => `${option.label} ${option.value}`) ?? []),
		].filter((value): value is string => !!value).join(' ');
	}

	private computeTokenOverlap(left: string | undefined, right: string | undefined): number {
		if (!left || !right) {
			return 0;
		}

		const leftTokens = new Set(left.toLowerCase().match(/[a-z0-9]{4,}/g) ?? []);
		const rightTokens = new Set(right.toLowerCase().match(/[a-z0-9]{4,}/g) ?? []);
		if (leftTokens.size === 0 || rightTokens.size === 0) {
			return 0;
		}

		let score = 0;
		for (const token of leftTokens) {
			if (rightTokens.has(token)) {
				score++;
			}
		}
		return score;
	}

	private getPlanStepReviewAnswerParts(answer: unknown): { readonly selectedValues: readonly string[]; readonly freeformValue?: string } {
		if (typeof answer === 'string') {
			return { selectedValues: [answer] };
		}

		if (typeof answer !== 'object' || answer === null || !hasKey(answer, { selectedValues: true })) {
			return { selectedValues: [] };
		}

		const multiSelectAnswer = answer as { readonly selectedValues?: unknown; readonly freeformValue?: unknown };
		const selectedValues = Array.isArray(multiSelectAnswer.selectedValues)
			? multiSelectAnswer.selectedValues.filter((value): value is string => typeof value === 'string')
			: [];
		const freeformValue = typeof multiSelectAnswer.freeformValue === 'string'
			? multiSelectAnswer.freeformValue.trim()
			: undefined;
		return {
			selectedValues,
			...(freeformValue ? { freeformValue } : {}),
		};
	}

	private getPlanStepReviewOutcome(editor: IChatPlanningPlanEditor, answersRecord: IChatQuestionAnswers | undefined): IPlanningPlanStepReviewOutcome {
		if (!editor.steps.length || !answersRecord) {
			return { hasEdits: false };
		}

		const notes: string[] = [];
		for (const step of editor.steps) {
			const selectedPartAnswer = answersRecord[`plan-editor-selected-part-${step.id}`];
			const selectedPart = typeof selectedPartAnswer === 'string' && selectedPartAnswer.trim() && selectedPartAnswer.trim() !== step.label
				? selectedPartAnswer.trim()
				: undefined;
			const stepTarget = selectedPart
				? `${step.index}: ${step.label} (selected text: ${selectedPart})`
				: `${step.index}: ${step.label}`;
			if (selectedPart && answersRecord[planningPlanControlFocusAnswerKey]) {
				notes.push(`- Selected plan text for controls: Step ${stepTarget}`);
			}
			const stepCustomAnswer = answersRecord[`plan-editor-step-custom-${step.id}`];
			const stepCustomEdit = typeof stepCustomAnswer === 'string'
				? stepCustomAnswer.trim()
				: undefined;
			if (stepCustomEdit) {
				notes.push(`- Step ${stepTarget} -> ${stepCustomEdit}`);
			}

			const answer = answersRecord[`plan-editor-step-${step.id}`];
			const { selectedValues, freeformValue } = this.getPlanStepReviewAnswerParts(answer);
			const actionValues = selectedValues.filter(value => value !== 'keep');
			if (actionValues.length > 0 || freeformValue) {
				const actions = actionValues.map(value => {
					switch (value) {
						case 'revise':
							return localize('chat.planStepReview.actionRevise', 'revise');
						case 'defer':
							return localize('chat.planStepReview.actionDefer', 'defer or remove');
						case 'split':
							return localize('chat.planStepReview.actionSplit', 'split into smaller steps');
						default:
							return value;
					}
				});
				const noteParts = [
					actions.length > 0 ? actions.join(', ') : localize('chat.planStepReview.actionEdit', 'edit'),
					freeformValue ? localize('chat.planStepReview.freeformNote', 'note: {0}', freeformValue) : undefined,
				].filter((value): value is string => !!value);
				notes.push(`- ${step.kind === 'verification' ? 'Verification' : 'Step'} ${stepTarget} -> ${noteParts.join('; ')}`);
			}

			for (const question of step.questions ?? []) {
				const questionAnswer = answersRecord[`plan-editor-question-${step.id}-${question.id}`];
				const formattedAnswer = this.formatPlanEditorQuestionAnswer(question, questionAnswer);
				if (formattedAnswer) {
					notes.push(`- Step ${stepTarget} question "${question.title}" -> ${formattedAnswer}`);
				}
			}
		}

		const additionalEdits = typeof answersRecord['plan-editor-additional'] === 'string'
			? answersRecord['plan-editor-additional'].trim()
			: undefined;
		if (additionalEdits) {
			notes.push(`- Additional edits: ${additionalEdits}`);
		}

		const controlFocus = typeof answersRecord[planningPlanControlFocusAnswerKey] === 'string'
			? answersRecord[planningPlanControlFocusAnswerKey].trim()
			: undefined;
		if (controlFocus) {
			notes.push(`- Requested control focus: ${controlFocus}`);
		}

		return notes.length > 0
			? {
				hasEdits: true,
				plannerNotes: ['Plan step review edits:', ...notes].join('\n'),
			}
			: { hasEdits: false };
	}

	private isPlanningPlanControlRegenerationRequest(answersRecord: IChatQuestionAnswers | undefined): boolean {
		return answersRecord?.[planningPlanRegenerateControlsAnswerKey] === planningPlanRegenerateControlsAnswerValue;
	}

	private formatPlanEditorQuestionAnswer(question: IChatQuestion, answer: IChatQuestionAnswerValue | undefined): string | undefined {
		if (answer === undefined) {
			return undefined;
		}

		if (typeof answer === 'string') {
			return answer.trim() || undefined;
		}

		const parts = hasKey(answer, { selectedValues: true })
			? [
				...answer.selectedValues.map((value: string) => this.getPlanEditorQuestionOptionLabel(question, value)).filter((value: string) => !!value),
				...(answer.freeformValue?.trim() ? [answer.freeformValue.trim()] : []),
			]
			: [
				answer.selectedValue ? this.getPlanEditorQuestionOptionLabel(question, answer.selectedValue) : undefined,
				answer.freeformValue?.trim(),
			].filter((value): value is string => !!value);
		return parts.length > 0 ? parts.join(', ') : undefined;
	}

	private getPlanEditorQuestionOptionLabel(question: IChatQuestion, value: string): string {
		return question.options?.find(option => option.value === value || option.id === value)?.label ?? value;
	}

	private mergePlanningPlannerNotes(existingNotes: string | undefined, newNotes: string | undefined): string | undefined {
		const notes = [existingNotes?.trim(), newNotes?.trim()].filter((value): value is string => !!value);
		return notes.length > 0 ? notes.join('\n\n') : undefined;
	}

	private async showPlanningPlanReviewSubmissionProgress(hasEdits: boolean, isControlRegenerationRequest: boolean): Promise<void> {
		if (!this.viewModel) {
			return;
		}

		if (isControlRegenerationRequest) {
			return;
		}

		if (hasEdits) {
			return;
		}

		const content = new MarkdownString(undefined, { supportThemeIcons: true });
		content.appendMarkdown('$(info) ');
		content.appendMarkdown(localize('chat.planReview.planEditorNoEdits', 'No section edits were selected, so the current plan is unchanged.'));
		const requestId = this.addCompletePlanningTranscriptRequest([{
			kind: 'markdownContent',
			content,
		}]);
		this.hidePlanningRequestPromptFromTranscript(requestId);
	}

	private async showPlanningPlanEditorRequest(
		_originalQuery: string,
		editor: IChatPlanningPlanEditor,
		onSubmit: (answersRecord: IChatQuestionAnswers | undefined) => Promise<void>,
		planSnapshot?: IPlanningPlanSnapshot,
		_lastFocusAreaLabel?: string,
	): Promise<void> {
		if (!this.viewModel) {
			return;
		}
		if (!editor.resolveId) {
			return;
		}

		this._pendingPlanningQuestionAnswersListener.clear();
		const resolveId = editor.resolveId;

		let didSubmit = false;
		const isPlanEditorSubmission = (event: { readonly resolveId: string }) => {
			return event.resolveId === resolveId;
		};

		this._pendingPlanningQuestionAnswersListener.value = Event.filter(
			this.chatService.onDidReceiveQuestionCarouselAnswer,
			event => isPlanEditorSubmission(event)
		)(event => {
			if (didSubmit) {
				return;
			}

			didSubmit = true;
			this._pendingPlanningQuestionAnswersListener.clear();
			const reviewOutcome = this.getPlanStepReviewOutcome(editor, event.answers);
			const isControlRegenerationRequest = this.isPlanningPlanControlRegenerationRequest(event.answers);
			void (async () => {
				await this.showPlanningPlanReviewSubmissionProgress(reviewOutcome.hasEdits, isControlRegenerationRequest);
				await onSubmit(event.answers);
			})();
		});

		await this.commandService.executeCommand(OPEN_PLANNING_PLAN_ACTION_ID, {
			sessionResource: this.viewModel.sessionResource.toString(),
			requestId: planSnapshot?.requestId ?? resolveId,
			previousRequestId: planSnapshot?.previousRequestId,
			planText: editor.planText,
			previousPlanText: planSnapshot?.previousPlanText,
			planSteps: editor.steps,
			planEditorResolveId: resolveId,
		});
	}

	private async submitPlanningPlanEditorAnswers(
		originalQuery: string,
		editor: IChatPlanningPlanEditor,
		answersRecord: IChatQuestionAnswers | undefined,
		options: IChatAcceptInputOptions,
		planningPhase: PlanningPhase,
		planSnapshot: IPlanningPlanSnapshot | undefined,
		lastFocusAreaLabel: string | undefined,
	): Promise<void> {
		editor.data = answersRecord ?? {};
		editor.isUsed = true;

		const reviewOutcome = this.getPlanStepReviewOutcome(editor, answersRecord);
		if (this.isPlanningPlanControlRegenerationRequest(answersRecord)) {
			const stageContext: IPlanningTransitionContext = {
				phase: planningPhase,
				answers: [],
				plannerNotes: this.mergePlanningPlannerNotes(this._planningTransitionContext?.plannerNotes, reviewOutcome.plannerNotes),
				recentConversation: this._planningTransitionContext?.recentConversation,
				repositoryContext: this._planningTransitionContext?.repositoryContext,
			};
			this._planningTransitionContext = mergePlanningTransitionContexts(this._planningTransitionContext, stageContext) ?? this._planningTransitionContext;
			await this.showPlanningPlanEditorWithInlineQuestions(
				originalQuery,
				options,
				planningPhase,
				planSnapshot,
				lastFocusAreaLabel,
				localize('chat.planningPlanEditor.regenerateControlsFocus', 'Regenerate task-decomposition controls for the current plan canvas. Use any draft step notes as context, but do not revise the plan yet.')
			);
			return;
		}

		if (!reviewOutcome.hasEdits) {
			return;
		}

		const stageContext: IPlanningTransitionContext = {
			phase: planningPhase,
			answers: [],
			plannerNotes: this.mergePlanningPlannerNotes(this._planningTransitionContext?.plannerNotes, reviewOutcome.plannerNotes),
			recentConversation: this._planningTransitionContext?.recentConversation,
			repositoryContext: this._planningTransitionContext?.repositoryContext,
		};
		const planningContext = mergePlanningTransitionContexts(this._planningTransitionContext, stageContext);
		this._planningTransitionContext = planningContext;
		if (!planningContext) {
			return;
		}

		await this.submitPlanningRequestWithContext(
			originalQuery,
			options,
			planningContext,
			false,
			async (_response, updatedPlanSnapshot) => this.showPlanningPlanEditorWithInlineQuestions(
				originalQuery,
				options,
				planningPhase,
				updatedPlanSnapshot,
				lastFocusAreaLabel
			),
			'updated-plan',
			'plan-review'
		);
	}

	private isChatInternalEditor(uri: URI | undefined): boolean {
		if (!uri) {
			return true;
		}

		return uri.scheme === 'gutter-input'
			|| uri.scheme === Schemas.untitled
			|| uri.scheme === Schemas.vscodeChatInput
			|| uri.scheme === Schemas.vscodeChatCodeBlock
			|| uri.scheme === Schemas.vscodeChatCodeCompareBlock
			|| uri.scheme === Schemas.vscodeChatEditor;
	}

	private createDynamicPlanningQuestionCarousel(
		questions: IChatQuestion[],
		generationContext: IPlanningQuestionGenerationContext
	): IChatQuestionCarousel {
		return new ChatQuestionCarouselData(
			questions,
			true,
			`${planningMiddlewareQuestionCarouselResolveIdPrefix}${generationContext.questionStage}-${Date.now()}`,
			undefined,
			undefined,
			this.buildPlanningMiddlewareCarouselMessage(generationContext)
		);
	}

	private renderGoalClarityQuestionCarouselInInput(
		originalQuery: string,
		carousel: IChatQuestionCarousel,
		options: IChatAcceptInputOptions,
		generationContext: IPlanningQuestionGenerationContext
	): void {
		if (!this.viewModel) {
			return;
		}

		const context = {
			element: {
				id: 'dynamic-planning-questions',
				sessionResource: this.viewModel.sessionResource,
				dataId: 'dynamic-planning-questions',
				username: '',
				message: { text: originalQuery, parts: [] },
				messageText: originalQuery,
				attempt: 0,
				variables: [],
				currentRenderedHeight: undefined,
				shouldBeRemovedOnSend: undefined,
				isComplete: true,
				isCompleteAddedRequest: false,
				slashCommand: undefined,
				agentOrSlashCommandDetected: false,
				shouldBeBlocked: constObservable(false),
				timestamp: Date.now(),
			},
			elementIndex: Math.max(this.viewModel.getItems().length ?? 0, 0),
			container: this.container,
			content: [carousel],
			contentIndex: 0,
			editorPool: undefined as never,
			codeBlockStartIndex: 0,
			treeStartIndex: 0,
			diffEditorPool: undefined as never,
			codeBlockModelCollection: this._codeBlockModelCollection,
			currentWidth: constObservable(this.container.clientWidth),
			onDidChangeVisibility: Event.None,
			inlineTextModels: undefined as never,
		};
		this.input.renderQuestionCarousel(carousel, context, {
			shouldAutoFocus: true,
			onSubmit: async answers => {
				await this.submitDynamicPlanningQuestionAnswers(
					originalQuery,
					carousel,
					answers ? Object.fromEntries(answers) : undefined,
					options,
					generationContext,
					false
				);
			}
		});
		this.input.focusQuestionCarousel();
	}

	private showDynamicPlanningQuestionCarousel(
		originalQuery: string,
		questions: IChatQuestion[],
		options: IChatAcceptInputOptions,
		generationContext: IPlanningQuestionGenerationContext,
		useSubmittedPlaceholder: boolean,
		planSnapshot?: IPlanningPlanSnapshot,
	): void {
		if (generationContext.questionStage !== 'goal-clarity' && generationContext.currentPlan?.trim()) {
			this.showPlanningPlanEditor(
				originalQuery,
				options,
				generationContext.planningPhase,
				planSnapshot ?? this.getLatestPlanningPlanSnapshot(),
				generationContext.focusAreaLabel,
				questions
			);
			return;
		}
		if (generationContext.questionStage === 'task-decomposition') {
			this.showPlanningPlanEditor(
				originalQuery,
				options,
				generationContext.planningPhase,
				planSnapshot ?? this.getLatestPlanningPlanSnapshot(),
				generationContext.focusAreaLabel,
				questions
			);
			return;
		}
		if (generationContext.questionStage === 'plan-focus') {
			this.logService.warn('[Planning] Ignoring plan-focus questions because no current plan is available for inline review.');
			return;
		}

		const carousel = this.createDynamicPlanningQuestionCarousel(questions, generationContext);
		this._pendingPlanningQuestionResolveId = carousel.resolveId;

		if (useSubmittedPlaceholder && this.canShowSubmittedPlanningPlaceholder()) {
			void this.setSubmittedPlanningPlaceholderCarousel(originalQuery, carousel, options, generationContext).then(wasShown => {
				if (!wasShown && generationContext.questionStage === 'goal-clarity') {
					void this.showPlanningQuestionGenerationError('goal-clarity', this.createGoalClarityTranscriptCarouselError());
				}
			}, error => {
				this.logService.warn('[Planning] Failed to show planning questions in the transcript.', error);
				if (generationContext.questionStage === 'goal-clarity') {
					void this.showPlanningQuestionGenerationError('goal-clarity', error);
				}
			});
			return;
		}

		if (generationContext.questionStage === 'goal-clarity') {
			this.renderGoalClarityQuestionCarouselInInput(originalQuery, carousel, options, generationContext);
			return;
		}

	}

	private canShowSubmittedPlanningPlaceholder(): boolean {
		return !!this.viewModel?.sessionResource && this.isInPlanningMode();
	}

	private async showSubmittedPlanningPlaceholder(
		_originalQuery: string,
		options: IChatAcceptInputOptions,
		questionStage: PlanningQuestionStage
	): Promise<void> {
		if (!this.viewModel) {
			return;
		}
		if (!this.isInPlanningMode()) {
			return;
		}

		this._pendingPlanningQuestionAnswersListener.clear();
		this._onDidAcceptInput.fire();
		this.input.acceptInput(options.storeToHistory ?? true);

		this.addCompletePlanningTranscriptRequest([this.createPlanningMiddlewareIntroContent(questionStage, 'generating')]);
	}

	private async setSubmittedPlanningPlaceholderCarousel(
		originalQuery: string,
		carousel: IChatQuestionCarousel,
		options: IChatAcceptInputOptions,
		generationContext: IPlanningQuestionGenerationContext
	): Promise<boolean> {
		if (!this.viewModel) {
			return false;
		}
		if (!this.isInPlanningMode()) {
			return false;
		}

		this._pendingPlanningQuestionAnswersListener.clear();

		const requestId = this.addCompletePlanningTranscriptRequest(
			[carousel],
			{ preserveScroll: false }
		);
		this._pendingPlanningPlaceholderRequestId = requestId;
		if (!requestId || !carousel.resolveId) {
			return false;
		}

		this._pendingPlanningQuestionAnswersListener.value = Event.once(Event.filter(
			this.chatService.onDidReceiveQuestionCarouselAnswer,
			event => event.requestId === requestId && event.resolveId === carousel.resolveId
		))(event => {
			void this.submitDynamicPlanningQuestionAnswers(
				originalQuery,
				carousel,
				event.answers,
				options,
				generationContext,
				true
			);
		});

		queueMicrotask(() => {
			this.input.focusQuestionCarousel();
		});

		return true;
	}

	private async submitDynamicPlanningQuestionAnswers(
		originalQuery: string,
		carousel: IChatQuestionCarousel,
		answersRecord: IChatQuestionAnswers | undefined,
		options: IChatAcceptInputOptions,
		generationContext: IPlanningQuestionGenerationContext,
		fromSubmittedPlaceholder: boolean
	): Promise<void> {
		if (!this.isInPlanningMode()) {
			this._pendingPlanningQuestionAnswersListener.clear();
			await this.clearPendingPlanningPlaceholder({ preserveUsedPlanningCarousel: false });
			if (isPlanningMiddlewareQuestionCarousel(carousel.resolveId)) {
				this.input.clearQuestionCarousel(undefined, carousel.resolveId);
			}
			this._pendingPlanningQuestionResolveId = undefined;
			return;
		}

		carousel.data = answersRecord ?? {};
		carousel.isUsed = true;
		if (!fromSubmittedPlaceholder) {
			this.input.clearQuestionCarousel(undefined, carousel.resolveId);
		}

		this._pendingPlanningQuestionResolveId = undefined;
		this._lastPlanningQuestionSourceInput = originalQuery;
		if (generationContext.questionStage === 'goal-clarity') {
			this._goalClarityQuestionRounds += 1;
		}

		const stageContext = buildPlanningTransitionContext(carousel, answersRecord, {
			phase: generationContext.planningPhase,
			plannerNotes: generationContext.plannerNotes,
			recentConversation: generationContext.recentConversation,
			repositoryContext: generationContext.repositoryContext,
		});
		const planningContext = mergePlanningTransitionContexts(this._planningTransitionContext, stageContext);
		this._planningTransitionContext = planningContext;
		if (!planningContext) {
			return;
		}

		try {
			const { planningContext: refreshedPlanningContext } = await this.refreshPlanningTransitionContextForStage(
				originalQuery,
				generationContext.planningPhase,
				generationContext.questionStage,
				planningContext,
				{ currentPlan: generationContext.currentPlan }
			);
			this._planningTransitionContext = refreshedPlanningContext;

			if (generationContext.questionStage === 'goal-clarity') {
				const continuationDecision = this.getGoalClarityContinuationDecisionBeforeFirstPlan(
					originalQuery,
					refreshedPlanningContext,
					this.hasGoalClarityProceedAnswer(answersRecord)
				);
				if (continuationDecision.shouldContinue
					&& await this.continueGoalClarityBeforeFirstPlan(originalQuery, options, generationContext.planningPhase, refreshedPlanningContext, continuationDecision.focusHint)) {
					return;
				}

				await this.submitPlanningRequestWithContext(
					originalQuery,
					options,
					refreshedPlanningContext,
					false,
					async (_response, planSnapshot) => this.showPlanningPlanEditorWithInlineQuestions(
						originalQuery,
						options,
						generationContext.planningPhase,
						planSnapshot
					),
					'first-plan',
					'goal-clarity',
				);
				return;
			}

			if (generationContext.questionStage === 'task-decomposition') {
				const isInitialTaskDecomposition = !generationContext.currentPlan?.trim();
				await this.submitPlanningRequestWithContext(
					originalQuery,
					options,
					refreshedPlanningContext,
					false,
					async (_response, planSnapshot) => this.showPlanningPlanEditorWithInlineQuestions(
						originalQuery,
						options,
						generationContext.planningPhase,
						planSnapshot
					),
					isInitialTaskDecomposition ? 'first-plan' : 'updated-plan',
					'task-decomposition',
				);
				return;
			}

			await this.submitPlanningRequestWithContext(
				originalQuery,
				options,
				refreshedPlanningContext,
				false,
				async (_response, planSnapshot) => this.showPlanningPlanEditor(
					originalQuery,
					options,
					generationContext.planningPhase,
					planSnapshot,
					generationContext.focusAreaLabel
				),
				'updated-plan',
				'plan-focus'
			);
		} catch (error) {
			await this.clearPendingPlanningPlaceholder();
			this.logService.error('[Planning] Failed to update the plan after planning questions.', error);
		}
	}

	private createPlanningContextAttachment(
		context: IPlanningTransitionContext,
		submission?: {
			readonly originalQuery: string;
			readonly progressKind?: 'first-plan' | 'updated-plan';
			readonly progressSource: PlanningPlanProgressSource;
			readonly allowPlannerFollowupQuestions: boolean;
			readonly isRetry?: boolean;
		}
	): IChatRequestVariableEntry {
		const contextPrompt = augmentPromptWithPlanningContext('', context);
		const submissionPrompt = submission?.progressKind
			? this.buildPlanningSubmissionInstruction(
				submission.originalQuery,
				submission.progressKind,
				submission.progressSource,
				submission.allowPlannerFollowupQuestions,
				!!submission.isRetry
			)
			: undefined;
		return {
			...toPromptTextVariableEntry(submissionPrompt ? `${contextPrompt}\n\n${submissionPrompt}` : contextPrompt, true),
			id: 'vscode.planning.context',
			name: 'prompt:planningContext',
			modelDescription: 'Planning context',
		};
	}

	private buildPlanningSubmissionInstruction(
		originalQuery: string,
		progressKind: 'first-plan' | 'updated-plan',
		progressSource: PlanningPlanProgressSource,
		allowPlannerFollowupQuestions: boolean,
		isRetry: boolean
	): string {
		const request = originalQuery.trim() || localize('chat.dynamicPlanning.emptyOriginalRequest', '(empty request)');
		const action = progressKind === 'first-plan'
			? 'Create the first concrete plan now.'
			: 'Revise the current plan now.';
		const source = progressSource === 'goal-clarity'
			? 'The user already answered goal-clarity questions.'
			: progressSource === 'task-decomposition'
				? 'The user already answered plan-shaping questions.'
				: progressSource === 'plan-review'
					? 'The user submitted edits to the current plan.'
					: 'The user selected a plan area to refine.';
		const followupInstruction = allowPlannerFollowupQuestions
			? 'Ask a follow-up question only if a truly blocking decision is still missing.'
			: 'Do not ask more goal-clarity questions. If a minor detail is unknown, state the assumption in the plan.';
		const retryInstruction = isRetry
			? [
				'The previous planner response completed without usable plan markdown.',
				'This retry must output the plan markdown now.',
				'Do not use tools on this retry unless a required tool call is already approved and immediately available.',
				'Do not announce that you will inspect, gather, read, check, or present something later.',
			]
			: [];

		return [
			'Planning handoff instruction:',
			`User request: ${request}`,
			...retryInstruction,
			source,
			action,
			followupInstruction,
			'Do not include unanswered goal-clarification questions, open-ended goal prompts, or a goal-clarity "Further Considerations" section in the plan.',
			'Use the editable assumption answers as settled input. Put any remaining non-blocking assumptions under **Decisions** or **Assumptions** instead of asking the user another goal question.',
			'Do not number phase or section headings. Keep numbering only on the individual plan steps.',
			'While working, stream concise visible planning notes before the final plan: what context you are using, which constraints matter, and which tradeoffs shape the plan.',
			'Visible planning notes must be followed by the final plan in this same response.',
			'If you need to inspect files or use tools, do that in this same response and still finish with plan markdown.',
			'If a tool needs approval, emit the tool call so the chat response can show the approval UI; do not replace the tool call with prose or a notification-only request.',
			'Do not stop after a progress update such as "I will inspect...", "I will gather...", or "I will present...".',
			'Do not write a standalone future-tense status sentence. A response that only says what you will do next is invalid.',
			'If no tool result is available, draft the plan from the provided planning context and clearly mark non-blocking assumptions.',
			'Return markdown that starts with a plan heading or a **Steps** section and includes **Relevant files**, **Verification**, and **Decisions** when applicable.',
		].join('\n');
	}

	private buildPlanningMiddlewareCarouselMessage(generationContext: IPlanningQuestionGenerationContext | PlanningQuestionStage): MarkdownString {
		const markdown = new MarkdownString(undefined, { supportThemeIcons: true });
		const questionStage = typeof generationContext === 'string' ? generationContext : generationContext.questionStage;

		if (questionStage === 'goal-clarity') {
			markdown.appendMarkdown('$(sparkle) ');
			markdown.appendMarkdown(localize('chat.dynamicPlanning.goalClarityMessage', '**Clarifying Your Goals**'));
		} else if (questionStage === 'task-decomposition') {
			markdown.appendMarkdown('$(list-unordered) ');
			markdown.appendMarkdown(localize('chat.dynamicPlanning.taskDecompositionMessage', '**Shaping Breakdown Controls**'));
		} else {
			markdown.appendMarkdown('$(target) ');
			markdown.appendMarkdown(localize('chat.dynamicPlanning.planFocusMessage', '**Refining a Plan Area**'));
		}

		return markdown;
	}

	private getPlanningSubmissionToolOverrides(baseTools?: UserSelectedTools, allowPlannerFollowupQuestions = false): UserSelectedTools {
		return {
			...(baseTools ?? this.input.selectedToolsModel.userSelectedTools.get()),
			vscode_askQuestions: allowPlannerFollowupQuestions,
			copilot_askQuestions: allowPlannerFollowupQuestions,
		};
	}

	private createPlanningMiddlewareIntroContent(questionStage: PlanningQuestionStage, mode: 'captured' | 'generating' = 'captured') {
		const content = new MarkdownString(undefined, { supportThemeIcons: true });
		if (mode === 'captured') {
			content.appendMarkdown(questionStage === 'goal-clarity'
				? '$(sparkle) '
				: questionStage === 'task-decomposition'
					? '$(list-unordered) '
					: '$(target) ');
			content.appendMarkdown(questionStage === 'goal-clarity'
				? localize('chat.dynamicPlanning.prePlanningIntro', '**Clarifying Your Goals**\n\nQuestions and answers stay here for reference.')
				: questionStage === 'task-decomposition'
					? localize('chat.dynamicPlanning.prePlanningIntroDecomposition', '**Shaping Breakdown Controls**\n\nPlan-shaping choices stay here for reference.')
					: localize('chat.dynamicPlanning.prePlanningIntroFocus', '**Refining a Plan Area**\n\nRefinement choices stay here for reference.'));
		} else {
			content.appendMarkdown('$(sync~spin) ');
			content.appendMarkdown(questionStage === 'goal-clarity'
				? localize('chat.dynamicPlanning.generatingGoalClarityIntro', '**Clarifying Your Goals**\n\nReading your answers, checking repo context, and deciding what context is still needed before a plan can be drafted.')
				: questionStage === 'task-decomposition'
					? localize('chat.dynamicPlanning.generatingTaskDecompositionIntro', '**Shaping Breakdown Controls**\n\nReading the plan, goal-clarity answers, repo context, and recent chat to generate task-breakdown controls for the canvas.')
					: localize('chat.dynamicPlanning.generatingPlanFocusIntro', '**Refining a Plan Area**\n\nReading the current plan slice, repo context, and recent chat to prepare targeted follow-up controls.'));
		}
		return {
			kind: 'markdownContent' as const,
			content,
		};
	}

	private async _acceptInput(query: { query: string } | undefined, options: IChatAcceptInputOptions = {}): Promise<IChatResponseModel | undefined> {
		if (!query && this.input.generating) {
			// if the user submits the input and generation finishes quickly, just submit it for them
			const generatingAutoSubmitWindow = 500;
			const start = Date.now();
			await this.input.generating;
			if (Date.now() - start > generatingAutoSubmitWindow) {
				return;
			}
		}

		while (!this._viewModel && !this._store.isDisposed) {
			await Event.toPromise(this.onDidChangeViewModel, this._store);
		}

		if (!this.viewModel) {
			return;
		}

		// Check if a custom submit handler wants to handle this submission
		if (this.viewOptions.submitHandler) {
			const inputValue = !query ? this.getInput() : query.query;
			const handled = await this.viewOptions.submitHandler(inputValue, this.input.currentModeKind);
			if (handled) {
				return;
			}
		}

		const inputValue = !query ? this.getInput() : query.query;
		if (await this.maybeRunDynamicPlanningQuestions(inputValue, options)) {
			return;
		}

		this._onDidAcceptInput.fire();
		this.listWidget.setScrollLock(this.isLockedToCodingAgent || !!checkModeOption(this.input.currentModeKind, this.viewOptions.autoScroll));

		const requestInputs: IChatRequestInputOptions = {
			input: inputValue,
			attachedContext: options?.enableImplicitContext === false ? this.input.getAttachedContext() : this.input.getAttachedAndImplicitContext(),
		};
		if (options.extraAttachedContext?.length) {
			requestInputs.attachedContext.add(...options.extraAttachedContext);
		}

		const isUserQuery = !query;
		const isEditing = this.viewModel?.editing;
		if (isEditing) {
			const editingPendingRequest = this.viewModel.editing!.pendingKind;
			if (editingPendingRequest !== undefined) {
				const editingRequestId = this.viewModel.editing!.id;
				this.chatService.removePendingRequest(this.viewModel.sessionResource, editingRequestId);
				options.queue ??= editingPendingRequest;
			} else {
				await this.chatService.cancelCurrentRequestForSession(this.viewModel.sessionResource, 'acceptInput-editing');
				options.queue = undefined;
			}

			this.finishedEditing(true);
			this.viewModel.model?.setCheckpoint(undefined);
		}

		const model = this.viewModel.model;
		const requestInProgress = model.requestInProgress.get();
		// Cancel the request if the user chooses to take a different path.
		// This is a bit of a heuristic for the common case of tool confirmation+reroute.
		// But we don't do this if there are queued messages, because we would either
		// discard them or need a prompt (as in `confirmPendingRequestsBeforeSend`)
		// which could be a surprising behavior if the user finishes typing a steering
		// request just as confirmation is triggered.
		if (model.requestNeedsInput.get() && !model.getPendingRequests().length) {
			await this.chatService.cancelCurrentRequestForSession(this.viewModel.sessionResource, 'acceptInput-needsInput');
			options.queue ??= ChatRequestQueueKind.Queued;
		}
		if (requestInProgress) {
			options.queue ??= ChatRequestQueueKind.Queued;
		}
		if (!requestInProgress && !isEditing && !(await this.confirmPendingRequestsBeforeSend(model, options))) {
			return;
		}

		// process the prompt command
		await this._applyPromptFileIfSet(requestInputs);
		await this._autoAttachInstructions(requestInputs);

		if (this.viewOptions.enableWorkingSet !== undefined && this.input.currentModeKind === ChatModeKind.Edit) {
			const uniqueWorkingSetEntries = new ResourceSet(); // NOTE: this is used for bookkeeping so the UI can avoid rendering references in the UI that are already shown in the working set
			const editingSessionAttachedContext: ChatRequestVariableSet = requestInputs.attachedContext;

			// Collect file variables from previous requests before sending the request
			const previousRequests = this.viewModel.model.getRequests();
			for (const request of previousRequests) {
				for (const variable of request.variableData.variables) {
					if (URI.isUri(variable.value) && variable.kind === 'file') {
						const uri = variable.value;
						if (!uniqueWorkingSetEntries.has(uri)) {
							editingSessionAttachedContext.add(variable);
							uniqueWorkingSetEntries.add(variable.value);
						}
					}
				}
			}
			requestInputs.attachedContext = editingSessionAttachedContext;

			type ChatEditingWorkingSetClassification = {
				owner: 'joyceerhl';
				comment: 'Information about the working set size in a chat editing request';
				originalSize: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The number of files that the user tried to attach in their editing request.' };
				actualSize: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The number of files that were actually sent in their editing request.' };
			};
			type ChatEditingWorkingSetEvent = {
				originalSize: number;
				actualSize: number;
			};
			this.telemetryService.publicLog2<ChatEditingWorkingSetEvent, ChatEditingWorkingSetClassification>('chatEditing/workingSetSize', { originalSize: uniqueWorkingSetEntries.size, actualSize: uniqueWorkingSetEntries.size });
		}

		this.input.validateAgentMode();

		if (this.viewModel.model.checkpoint) {
			const requests = this.viewModel.model.getRequests();
			for (let i = requests.length - 1; i >= 0; i -= 1) {
				const request = requests[i];
				if (request.shouldBeBlocked) {
					this.chatService.removeRequest(this.viewModel.sessionResource, request.id);
				}
			}
		}
		// Expand directory attachments: extract images as binary entries
		const resolvedImageVariables = await this._resolveDirectoryImageAttachments(requestInputs.attachedContext.asArray());
		const submittedSessionResource = this.viewModel.sessionResource;

		const modeRequestOptions = this.getModeRequestOptions();
		const result = await this.chatService.sendRequest(this.viewModel.sessionResource, requestInputs.input, {
			userSelectedModelId: this.input.currentLanguageModel,
			location: this.location,
			locationData: this._location.resolveData?.(),
			parserContext: { selectedAgent: this._lastSelectedAgent, mode: this.input.currentModeKind, attachmentCapabilities: this._lastSelectedAgent?.capabilities ?? this.attachmentCapabilities },
			attachedContext: requestInputs.attachedContext.asArray(),
			resolvedVariables: resolvedImageVariables,
			noCommandDetection: options?.noCommandDetection,
			...modeRequestOptions,
			userSelectedTools: options?.userSelectedToolsOverride ? constObservable(options.userSelectedToolsOverride) : modeRequestOptions.userSelectedTools,
			modeInfo: this.input.currentModeInfo,
			agentIdSilent: this._lockedAgent?.id,
			queue: options?.queue,

		});

		if (ChatSendResult.isRejected(result)) {
			return;
		}

		this.logThinkingStyleUsage('submit');

		// visibility sync before firing events to hide the welcome view
		this.updateChatViewVisibility();
		this.input.acceptInput(options?.storeToHistory ?? isUserQuery);

		const sent = ChatSendResult.isQueued(result) ? await result.deferred : result;
		if (!ChatSendResult.isSent(sent)) {
			return;
		}

		this._onDidSubmitAgent.fire({ agent: sent.data.agent, slashCommand: sent.data.slashCommand });
		this.handleDelegationExitIfNeeded(this._lockedAgent, sent.data.agent);

		// If the session was replaced (untitled -> real contributed session), swap the widget's model
		if (sent.newSessionResource) {
			const newModel = this.chatService.getSession(sent.newSessionResource);
			if (newModel) {
				this.setModel(newModel);
			}
		}

		sent.data.responseCreatedPromise.then(() => {
			// Only start accessibility progress once a real request/response model exists.
			this.chatAccessibilityService.acceptRequest(submittedSessionResource);
			sent.data.responseCompletePromise.then(() => {
				const responses = this.viewModel?.getItems().filter(isResponseVM);
				const lastResponse = responses?.[responses.length - 1];
				this.chatAccessibilityService.acceptResponse(this, this.container, lastResponse, submittedSessionResource, options?.isVoiceInput);
				if (lastResponse?.result?.nextQuestion) {
					const { prompt, participant, command } = lastResponse.result.nextQuestion;
					const question = formatChatQuestion(this.chatAgentService, this.location, prompt, participant, command);
					if (question) {
						this.input.setValue(question, false);
					}
				}
			});
		});

		return sent.data.responseCreatedPromise;
	}

	// Resolve images from directory attachments to send as additional variables.
	private async _resolveDirectoryImageAttachments(attachments: IChatRequestVariableEntry[]): Promise<IChatRequestVariableEntry[]> {
		const imagePromises: Promise<IChatRequestVariableEntry[]>[] = [];

		for (const attachment of attachments) {
			if (attachment.kind === 'directory' && URI.isUri(attachment.value)) {
				imagePromises.push(
					this.chatAttachmentResolveService.resolveDirectoryImages(attachment.value)
				);
			}
		}

		if (imagePromises.length === 0) {
			return [];
		}

		const resolved = await Promise.all(imagePromises);
		return resolved.flat();
	}

	private async confirmPendingRequestsBeforeSend(model: IChatModel, options: IChatAcceptInputOptions): Promise<boolean> {
		if (options.queue) {
			return true;
		}

		const hasPendingRequests = model.getPendingRequests().length > 0;
		if (!hasPendingRequests) {
			return true;
		}

		const promptResult = await this.dialogService.prompt({
			type: 'question',
			message: localize('chat.pendingRequests.prompt.message', "You already have pending requests."),
			detail: localize('chat.pendingRequests.prompt.detail', "Do you want to keep them in the queue or remove them before sending this message?"),
			buttons: [
				{
					label: localize('chat.pendingRequests.prompt.keep', "Keep Pending Requests"),
					run: () => 'keep'
				},
				{
					label: localize('chat.pendingRequests.prompt.remove', "Remove Pending Requests"),
					run: () => 'remove'
				}
			],
			cancelButton: true
		});

		if (!promptResult.result) {
			return false;
		}

		if (promptResult.result === 'remove') {
			for (const pendingRequest of [...model.getPendingRequests()]) {
				this.chatService.removePendingRequest(model.sessionResource, pendingRequest.request.id);
			}
		}

		return true;
	}

	getModeRequestOptions(): Partial<IChatSendRequestOptions> {
		const sessionResource = this.viewModel?.sessionResource;
		const capturedModeId = this.input.currentModeObs.get().id;
		const userSelectedTools = this.input.selectedToolsModel.userSelectedTools;

		let lastToolsSnapshot = userSelectedTools.get();

		// When the widget has loaded a new session, return a snapshot of the tools for this session.
		// Only sync with the tools model when this session is shown with the same mode.
		const scopedTools = derived(reader => {
			const activeSession = this._viewModelObs.read(reader)?.sessionResource;
			const currentModeId = this.input.currentModeObs.read(reader).id;
			if (isEqual(activeSession, sessionResource) && currentModeId === capturedModeId) {
				const tools = userSelectedTools.read(reader);
				lastToolsSnapshot = tools;
				return tools;
			}
			return lastToolsSnapshot;
		});

		return {
			modeInfo: this.input.currentModeInfo,
			userSelectedTools: scopedTools,
		};
	}

	getCodeBlockInfosForResponse(response: IChatResponseViewModel): IChatCodeBlockInfo[] {
		return this.listWidget.getCodeBlockInfosForResponse(response);
	}

	getCodeBlockInfoForEditor(uri: URI): IChatCodeBlockInfo | undefined {
		return this.listWidget.getCodeBlockInfoForEditor(uri);
	}

	getFileTreeInfosForResponse(response: IChatResponseViewModel): IChatFileTreeInfo[] {
		return this.listWidget.getFileTreeInfosForResponse(response);
	}

	getLastFocusedFileTreeForResponse(response: IChatResponseViewModel): IChatFileTreeInfo | undefined {
		return this.listWidget.getLastFocusedFileTreeForResponse(response);
	}

	focusResponseItem(lastFocused?: boolean): void {
		this.listWidget.focusLastItem(lastFocused);
	}

	setInputPartMaxHeightOverride(maxHeight: number | undefined): void {
		this.inputPartMaxHeightOverride = maxHeight;
	}

	layout(height: number, width: number): void {
		width = Math.min(width, this.viewOptions.renderStyle === 'minimal' ? width : 950); // no min width of inline chat

		this.bodyDimension = new dom.Dimension(width, height);

		if (this.viewModel?.editing) {
			this.inlineInputPart?.layout(width);
		}

		const chatSuggestNextWidgetHeight = this.chatSuggestNextWidget.height;
		const inputMaxHeight = this._dynamicMessageLayoutData || this.location !== ChatAgentLocation.Chat
			? undefined
			: this.inputPartMaxHeightOverride !== undefined
				? Math.max(0, this.inputPartMaxHeightOverride - chatSuggestNextWidgetHeight - MIN_LIST_HEIGHT)
				: Math.max(0, height - chatSuggestNextWidgetHeight - MIN_LIST_HEIGHT);
		this.inputPart.setMaxHeight(inputMaxHeight);
		this.inputPart.layout(width);

		this._layoutListForInputHeight();
	}

	/**
	 * Re-layout just the list, welcome container, and list container to match
	 * the current input-part height. Called both from {@link layout} and from
	 * the inputPart.height autorun so we never re-enter inputPart.layout when
	 * only the input height changed.
	 */
	private _layoutListForInputHeight(): void {
		if (!this.bodyDimension) {
			return;
		}

		const { height, width } = this.bodyDimension;
		const chatSuggestNextWidgetHeight = this.chatSuggestNextWidget.height;

		const inputHeight = this.inputPart.height.get();
		const lastElementVisible = this.listWidget.isScrolledToBottom;
		const lastItem = this.listWidget.lastItem;

		const contentHeight = Math.max(0, height - inputHeight - chatSuggestNextWidgetHeight);
		this.listWidget.layout(contentHeight, width);

		this.welcomeMessageContainer.style.height = `${contentHeight}px`;

		const lastResponseIsRendering = isResponseVM(lastItem) && lastItem.renderData;
		if (lastElementVisible && (!lastResponseIsRendering || checkModeOption(this.input.currentModeKind, this.viewOptions.autoScroll))) {
			this.listWidget.scrollToEnd();
		}
		this.listContainer.style.height = `${contentHeight}px`;

		this._onDidChangeHeight.fire(height);
	}

	private _dynamicMessageLayoutData?: { numOfMessages: number; maxHeight: number; enabled: boolean };

	// An alternative to layout, this allows you to specify the number of ChatTreeItems
	// you want to show, and the max height of the container. It will then layout the
	// tree to show that many items.
	// TODO@TylerLeonhardt: This could use some refactoring to make it clear which layout strategy is being used
	setDynamicChatTreeItemLayout(numOfChatTreeItems: number, maxHeight: number) {
		this._dynamicMessageLayoutData = { numOfMessages: numOfChatTreeItems, maxHeight, enabled: true };
		this._register(this.listWidget.onDidChangeItemHeight(() => this.layoutDynamicChatTreeItemMode()));

		const mutableDisposable = this._register(new MutableDisposable());
		this._register(this.listWidget.onDidScroll((e) => {
			// TODO@TylerLeonhardt this should probably just be disposed when this is disabled
			// and then set up again when it is enabled again
			if (!this._dynamicMessageLayoutData?.enabled) {
				return;
			}
			mutableDisposable.value = dom.scheduleAtNextAnimationFrame(dom.getWindow(this.listContainer), () => {
				if (!e.scrollTopChanged || e.heightChanged || e.scrollHeightChanged) {
					return;
				}
				const renderHeight = e.height;
				const diff = e.scrollHeight - renderHeight - e.scrollTop;
				if (diff === 0) {
					return;
				}

				const possibleMaxHeight = (this._dynamicMessageLayoutData?.maxHeight ?? maxHeight);
				const width = this.bodyDimension?.width ?? this.container.offsetWidth;
				this.input.layout(width);
				const inputPartHeight = this.input.height.get();
				const chatSuggestNextWidgetHeight = this.chatSuggestNextWidget.height;
				const newHeight = Math.min(renderHeight + diff, possibleMaxHeight - inputPartHeight - chatSuggestNextWidgetHeight);
				this.layout(newHeight + inputPartHeight + chatSuggestNextWidgetHeight, width);
			});
		}));
	}

	updateDynamicChatTreeItemLayout(numOfChatTreeItems: number, maxHeight: number) {
		this._dynamicMessageLayoutData = { numOfMessages: numOfChatTreeItems, maxHeight, enabled: true };
		let hasChanged = false;
		let height = this.bodyDimension!.height;
		let width = this.bodyDimension!.width;
		if (maxHeight < this.bodyDimension!.height) {
			height = maxHeight;
			hasChanged = true;
		}
		const containerWidth = this.container.offsetWidth;
		if (this.bodyDimension?.width !== containerWidth) {
			width = containerWidth;
			hasChanged = true;
		}
		if (hasChanged) {
			this.layout(height, width);
		}
	}

	get isDynamicChatTreeItemLayoutEnabled(): boolean {
		return this._dynamicMessageLayoutData?.enabled ?? false;
	}

	set isDynamicChatTreeItemLayoutEnabled(value: boolean) {
		if (!this._dynamicMessageLayoutData) {
			return;
		}
		this._dynamicMessageLayoutData.enabled = value;
	}

	layoutDynamicChatTreeItemMode(): void {
		if (!this.viewModel || !this._dynamicMessageLayoutData?.enabled) {
			return;
		}

		const width = this.bodyDimension?.width ?? this.container.offsetWidth;
		this.input.layout(width);
		const inputHeight = this.input.height.get();
		const chatSuggestNextWidgetHeight = this.chatSuggestNextWidget.height;

		const totalMessages = this.viewModel.getItems();
		// grab the last N messages
		const messages = totalMessages.slice(-this._dynamicMessageLayoutData.numOfMessages);

		const needsRerender = messages.some(m => m.currentRenderedHeight === undefined);
		const listHeight = needsRerender
			? this._dynamicMessageLayoutData.maxHeight
			: messages.reduce((acc, message) => acc + message.currentRenderedHeight!, 0);

		this.layout(
			Math.min(
				// we add an additional 18px in order to show that there is scrollable content
				inputHeight + chatSuggestNextWidgetHeight + listHeight + (totalMessages.length > 2 ? 18 : 0),
				this._dynamicMessageLayoutData.maxHeight
			),
			width
		);

		if (needsRerender || !listHeight) {
			this.listWidget.scrollToEnd();
		}
	}

	saveState(): void {
		// no-op
	}

	getViewState(): IChatModelInputState | undefined {
		return this.input.getCurrentInputState();
	}

	private updateChatInputContext() {
		const currentAgent = this.parsedInput.parts.find(part => part instanceof ChatRequestAgentPart);
		this.agentInInput.set(!!currentAgent);
	}

	private async _switchToAgentByName(agentName: string): Promise<void> {
		const currentAgent = this.input.currentModeObs.get();

		// switch to appropriate agent if needed
		if (agentName !== currentAgent.name.get()) {
			// Find the mode object to get its kind
			const agent = this.chatModeService.findModeByName(agentName);
			if (agent) {
				if (currentAgent.kind !== agent.kind) {
					const chatModeCheck = await this.instantiationService.invokeFunction(handleModeSwitch, currentAgent.kind, agent.kind, this.viewModel?.model.getRequests().length ?? 0, this.viewModel?.model);
					if (!chatModeCheck) {
						return;
					}

					if (chatModeCheck.needToClearSession) {
						await this.clear();
					}
				}
				this.input.setChatMode(agent.id);
			}
		}
	}

	private async _applyPromptMetadata({ agent, tools, model }: PromptHeader, requestInput: IChatRequestInputOptions): Promise<void> {

		if (tools !== undefined && !agent && this.input.currentModeKind !== ChatModeKind.Agent) {
			agent = ChatMode.Agent.name.get();
		}
		// switch to appropriate agent if needed
		if (agent) {
			this._switchToAgentByName(agent);
		}

		// if not tools to enable are present, we are done
		if (tools !== undefined && this.input.currentModeKind === ChatModeKind.Agent) {
			const enablementMap = this.toolsService.toToolAndToolSetEnablementMap(tools, this.input.selectedLanguageModel.get()?.metadata);
			this.input.selectedToolsModel.set(enablementMap, true);
		}

		if (model !== undefined) {
			this.input.switchModelByQualifiedName(model);
		}
	}

	/**
	 * Adds additional instructions to the context
	 * - instructions that have a 'applyTo' pattern that matches the current input
	 * - instructions referenced in the copilot settings 'copilot-instructions'
	 * - instructions referenced in an already included instruction file
	 */
	private async _autoAttachInstructions({ attachedContext }: IChatRequestInputOptions): Promise<void> {
		const contribution = this._lockedAgent ? this.chatSessionsService.getChatSessionContribution(this._lockedAgent.id) : undefined;

		// For contributed session types, default to false for autoAttachReferences.
		const isContributedSession = !!contribution;
		const autoAttachEnabled = isContributedSession ?
			contribution.autoAttachReferences === true : true;

		if (!autoAttachEnabled) {
			this.logService.debug(`ChatWidget#_autoAttachInstructions: skipped, autoAttachReferences is disabled`);
			return;
		}

		this.logService.debug(`ChatWidget#_autoAttachInstructions: prompt files are enabled`);
		const enabledTools = this.input.currentModeKind === ChatModeKind.Agent ? this.input.selectedToolsModel.userSelectedTools.get() : undefined;
		const enabledSubAgents = this.input.currentModeKind === ChatModeKind.Agent ? this.input.currentModeObs.get().agents?.get() : undefined;
		const sessionResource = this._viewModel?.model.sessionResource;
		const computer = this.instantiationService.createInstance(ComputeAutomaticInstructions, this.input.currentModeKind, enabledTools, enabledSubAgents, sessionResource);
		await computer.collect(attachedContext, CancellationToken.None);
	}

	delegateScrollFromMouseWheelEvent(browserEvent: IMouseWheelEvent): void {
		this.listWidget.delegateScrollFromMouseWheelEvent(browserEvent);
	}
}

function isUsedQuestionCarousel(content: unknown): content is IChatQuestionCarousel {
	if (!content || typeof content !== 'object' || !('kind' in content) || content.kind !== 'questionCarousel') {
		return false;
	}

	const carousel = content as IChatQuestionCarousel;
	return Boolean(carousel.isUsed && carousel.data && Object.keys(carousel.data).length > 0);
}

const MIN_LIST_HEIGHT = 50;
