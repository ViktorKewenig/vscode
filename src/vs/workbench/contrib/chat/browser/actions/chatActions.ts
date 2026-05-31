/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { isAncestorOfActiveElement } from '../../../../../base/browser/dom.js';
import { alert } from '../../../../../base/browser/ui/aria/aria.js';
import { WorkbenchActionExecutedClassification, WorkbenchActionExecutedEvent } from '../../../../../base/common/actions.js';
import { coalesce } from '../../../../../base/common/arrays.js';
import { timeout } from '../../../../../base/common/async.js';
import { Codicon } from '../../../../../base/common/codicons.js';
import { safeIntl } from '../../../../../base/common/date.js';
import { Event } from '../../../../../base/common/event.js';
import { MarkdownString } from '../../../../../base/common/htmlContent.js';
import { KeyCode, KeyMod } from '../../../../../base/common/keyCodes.js';
import { Schemas } from '../../../../../base/common/network.js';
import { language } from '../../../../../base/common/platform.js';
import { basename } from '../../../../../base/common/resources.js';
import { ThemeIcon } from '../../../../../base/common/themables.js';
import { URI } from '../../../../../base/common/uri.js';
import { ICodeEditor } from '../../../../../editor/browser/editorBrowser.js';
import { EditorAction2 } from '../../../../../editor/browser/editorExtensions.js';
import { IRange } from '../../../../../editor/common/core/range.js';
import { localize, localize2 } from '../../../../../nls.js';
import { Action2, ICommandPaletteOptions, MenuId, MenuRegistry, registerAction2 } from '../../../../../platform/actions/common/actions.js';
import { ICommandService } from '../../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { ContextKeyExpr } from '../../../../../platform/contextkey/common/contextkey.js';
import { IsLinuxContext, IsWindowsContext } from '../../../../../platform/contextkey/common/contextkeys.js';
import { IDialogService } from '../../../../../platform/dialogs/common/dialogs.js';
import { IFileService } from '../../../../../platform/files/common/files.js';
import { IInstantiationService, ServicesAccessor } from '../../../../../platform/instantiation/common/instantiation.js';
import { KeybindingWeight } from '../../../../../platform/keybinding/common/keybindingsRegistry.js';
import { ILogService } from '../../../../../platform/log/common/log.js';
import { INotificationService } from '../../../../../platform/notification/common/notification.js';
import { IOpenerService } from '../../../../../platform/opener/common/opener.js';
import product from '../../../../../platform/product/common/product.js';
import { ITelemetryService } from '../../../../../platform/telemetry/common/telemetry.js';
import { ActiveEditorContext } from '../../../../common/contextkeys.js';
import { IResourceDiffEditorInput, IUntitledTextResourceEditorInput } from '../../../../common/editor.js';
import { IViewDescriptorService, ViewContainerLocation } from '../../../../common/views.js';
import { ChatEntitlement, IChatEntitlementService } from '../../../../services/chat/common/chatEntitlementService.js';
import { ACTIVE_GROUP, AUX_WINDOW_GROUP, IEditorService, SIDE_GROUP } from '../../../../services/editor/common/editorService.js';
import { IHostService } from '../../../../services/host/browser/host.js';
import { IWorkbenchLayoutService, Parts } from '../../../../services/layout/browser/layoutService.js';
import { IPreferencesService } from '../../../../services/preferences/common/preferences.js';
import { IViewsService } from '../../../../services/views/common/viewsService.js';
import { EXTENSIONS_CATEGORY, IExtensionsWorkbenchService } from '../../../extensions/common/extensions.js';
import { SCMHistoryItemChangeRangeContentProvider, ScmHistoryItemChangeRangeUriFields } from '../../../scm/browser/scmHistoryChatContext.js';
import { ISCMService } from '../../../scm/common/scm.js';
import { IChatAgentResult, IChatAgentService } from '../../common/participants/chatAgents.js';
import { ChatContextKeys } from '../../common/actions/chatContextKeys.js';
import { ModifiedFileEntryState } from '../../common/editing/chatEditingService.js';
import { IChatModel, IChatResponseModel } from '../../common/model/chatModel.js';
import { ChatMode, IChatMode, IChatModeService } from '../../common/chatModes.js';
import { ElicitationState, IChatPlanningPlanEditorStep, IChatQuestion, IChatQuestionAnswers, IChatService, IChatToolInvocation } from '../../common/chatService/chatService.js';
import { ISCMHistoryItemChangeRangeVariableEntry, ISCMHistoryItemChangeVariableEntry } from '../../common/attachments/chatVariableEntries.js';
import { IChatRequestViewModel, IChatResponseViewModel, isRequestVM } from '../../common/model/chatViewModel.js';
import { IChatWidgetHistoryService } from '../../common/widget/chatWidgetHistoryService.js';
import { ChatAgentLocation, ChatConfiguration, ChatModeKind } from '../../common/constants.js';
import { AICustomizationManagementCommands } from '../aiCustomization/aiCustomizationManagement.js';
import { ILanguageModelChatSelector, ILanguageModelsService } from '../../common/languageModels.js';
import { CopilotUsageExtensionFeatureId } from '../../common/languageModelStats.js';
import { ILanguageModelToolsConfirmationService } from '../../common/tools/languageModelToolsConfirmationService.js';
import { ILanguageModelToolsService, IToolData, IToolSet, isToolSet } from '../../common/tools/languageModelToolsService.js';
import { ChatViewId, IChatWidget, IChatWidgetService, isIChatViewViewContext } from '../chat.js';
import { IChatEditorOptions } from '../widgetHosts/editor/chatEditor.js';
import { ChatEditorInput, showClearEditingSessionConfirmation } from '../widgetHosts/editor/chatEditorInput.js';
import { convertBufferToScreenshotVariable } from '../attachments/chatScreenshotContext.js';
import { getChatSessionType, LocalChatSessionUri } from '../../common/model/chatUri.js';
import { localChatSessionType } from '../../common/chatSessionsService.js';
import { extractPlanningPlanSteps, extractPlanningPlanText, isUsablePlanningPlanText, summarizePlanningPlanChanges } from '../planning/chatPlanningPlanText.js';
import { generateUuid } from '../../../../../base/common/uuid.js';
import { ChatViewPane } from '../widgetHosts/viewPane/chatViewPane.js';
import { IWebviewWorkbenchService } from '../../../webviewPanel/browser/webviewWorkbenchService.js';
import { WebviewContentPurpose } from '../../../webview/browser/webview.js';

export const CHAT_CATEGORY = localize2('chat.category', 'Chat');

export const ACTION_ID_NEW_CHAT = `workbench.action.chat.newChat`;
export const ACTION_ID_NEW_EDIT_SESSION = `workbench.action.chat.newEditSession`;
export const ACTION_ID_OPEN_CHAT = 'workbench.action.openChat';
export const CHAT_OPEN_ACTION_ID = 'workbench.action.chat.open';
export const CHAT_SETUP_ACTION_ID = 'workbench.action.chat.triggerSetup';
export const CHAT_SETUP_SUPPORT_ANONYMOUS_ACTION_ID = 'workbench.action.chat.triggerSetupSupportAnonymousAction';
const TOGGLE_CHAT_ACTION_ID = 'workbench.action.chat.toggle';

export const GENERATE_AGENT_INSTRUCTIONS_COMMAND_ID = 'workbench.action.chat.generateAgentInstructions';
export const GENERATE_ON_DEMAND_INSTRUCTIONS_COMMAND_ID = 'workbench.action.chat.generateOnDemandInstructions';
export const GENERATE_PROMPT_COMMAND_ID = 'workbench.action.chat.generatePrompt';
export const GENERATE_SKILL_COMMAND_ID = 'workbench.action.chat.generateSkill';
export const GENERATE_AGENT_COMMAND_ID = 'workbench.action.chat.generateAgent';
export const GENERATE_HOOK_COMMAND_ID = 'workbench.action.chat.generateHook';
export const INSERT_FORK_CONVERSATION_COMMAND_ID = 'workbench.action.chat.insertForkConversationCommand';
export const INSERT_TROUBLESHOOT_COMMAND_ID = 'workbench.action.chat.insertTroubleshootCommand';
export const OPEN_PLANNING_PLAN_ACTION_ID = 'workbench.action.chat.openPlanningPlan';
export const OPEN_PLANNING_PLAN_TO_SIDE_ACTION_ID = 'workbench.action.chat.openPlanningPlanToSide';
export const OPEN_PLANNING_PLAN_DIFF_ACTION_ID = 'workbench.action.chat.openPlanningPlanDiff';
export const UPDATE_PLANNING_PLAN_ACTION_ID = 'workbench.action.chat.updatePlanningPlan';
const PLANNING_PLAN_WEBVIEW_VIEW_TYPE = 'workbench.chat.planningPlanEditor';
const planningPlanRegenerateControlsAnswerKey = 'plan-editor-regenerate-controls';
const planningPlanRegenerateControlsAnswerValue = 'regenerate-controls';

type PlanningPlanWebviewInput = ReturnType<IWebviewWorkbenchService['openWebview']>;
const planningPlanWebviews = new Map<string, { readonly input: PlanningPlanWebviewInput; listener: { dispose(): void } }>();

interface IPlanningPlanCommandArgs {
	readonly sessionResource: URI | string;
	readonly requestId: string;
	readonly previousRequestId?: string;
	readonly planText?: string;
	readonly previousPlanText?: string;
	readonly planSteps?: readonly IChatPlanningPlanEditorStep[];
	readonly planEditorResolveId?: string;
}

interface IPlanningPlanUpdateCommandArgs {
	readonly sessionResource: URI | string;
	readonly requestId?: string;
	readonly planText?: string;
	readonly previousPlanText?: string;
	readonly planSteps?: readonly IChatPlanningPlanEditorStep[];
	readonly isComplete?: boolean;
}

const defaultChat = {
	manageSettingsUrl: product.defaultChatAgent?.manageSettingsUrl ?? '',
	provider: product.defaultChatAgent?.provider ?? { enterprise: { id: '' } },
	completionsAdvancedSetting: product.defaultChatAgent?.completionsAdvancedSetting ?? '',
	completionsMenuCommand: product.defaultChatAgent?.completionsMenuCommand ?? '',
};

export interface IChatViewOpenOptions {
	/**
	 * The query for chat.
	 */
	query: string;
	/**
	 * Whether the query is partial and will await more input from the user.
	 */
	isPartialQuery?: boolean;
	/**
	 * A list of tools IDs with `canBeReferencedInPrompt` that will be resolved and attached if they exist.
	 */
	toolIds?: string[];
	/**
	 * Any previous chat requests and responses that should be shown in the chat view.
	 */
	previousRequests?: IChatViewOpenRequestEntry[];
	/**
	 * Whether a screenshot of the focused window should be taken and attached
	 */
	attachScreenshot?: boolean;
	/**
	 * A list of file URIs to attach to the chat as context.
	 */
	attachFiles?: (URI | { uri: URI; range: IRange })[];
	/**
	 * A list of source control history item changes to attach to the chat as context.
	 */
	attachHistoryItemChanges?: { uri: URI; historyItemId: string }[];
	/**
	 * A list of source control history item change ranges to attach to the chat as context.
	 */
	attachHistoryItemChangeRanges?: {
		start: { uri: URI; historyItemId: string };
		end: { uri: URI; historyItemId: string };
	}[];
	/**
	 * The mode ID or name to open the chat in.
	 */
	mode?: ChatModeKind | string;

	/**
	 * The language model selector to use for the chat.
	 * An Error will be thrown if there's no match. If there are multiple
	 * matches, the first match will be used.
	 *
	 * Examples:
	 *
	 * ```
	 * {
	 *   id: 'claude-sonnet-4',
	 *   vendor: 'copilot'
	 * }
	 * ```
	 *
	 * Use `claude-sonnet-4` from any vendor:
	 *
	 * ```
	 * {
	 *   id: 'claude-sonnet-4',
	 * }
	 * ```
	 */
	modelSelector?: ILanguageModelChatSelector;

	/**
	 * Wait to resolve the command until the chat response reaches a terminal state (complete, error, or pending user confirmation, etc.).
	 */
	blockOnResponse?: boolean;

	/**
	 * A list of tool identifiers to include. When specified alone, only these tools will be enabled.
	 * Identifiers can be tool IDs, tool reference names (`toolReferenceName`),
	 * toolset IDs, or toolset reference names (`referenceName`).
	 * When a toolset identifier matches, all tools in that toolset are included.
	 * Can be combined with `toolsExclude` for fine-grained control.
	 */
	toolsInclude?: string[];

	/**
	 * A list of tool identifiers to exclude. When specified alone, all tools except these will be enabled.
	 * Identifiers can be tool IDs, tool reference names (`toolReferenceName`),
	 * toolset IDs, or toolset reference names (`referenceName`).
	 * When a toolset identifier matches, all tools in that toolset are excluded.
	 * Can be combined with `toolsInclude` - exclusions are applied after inclusions.
	 * Explicit tool references in `toolsInclude` override toolset exclusions,
	 * but explicit tool exclusions always win.
	 */
	toolsExclude?: string[];
}

export interface IChatViewOpenRequestEntry {
	request: string;
	response: string;
}

export const CHAT_CONFIG_MENU_ID = new MenuId('workbench.chat.menu.config');

const OPEN_CHAT_QUOTA_EXCEEDED_DIALOG = 'workbench.action.chat.openQuotaExceededDialog';

abstract class OpenChatGlobalAction extends Action2 {
	constructor(overrides: Pick<ICommandPaletteOptions, 'keybinding' | 'title' | 'id' | 'menu'>, private readonly mode?: IChatMode) {
		super({
			...overrides,
			icon: Codicon.chatSparkle,
			f1: true,
			category: CHAT_CATEGORY,
			precondition: ContextKeyExpr.and(
				ChatContextKeys.Setup.hidden.negate(),
				ChatContextKeys.Setup.disabled.negate()
			)
		});
	}

	override async run(accessor: ServicesAccessor, opts?: string | IChatViewOpenOptions): Promise<IChatAgentResult & { type?: 'confirmation' } | undefined> {
		opts = typeof opts === 'string' ? { query: opts } : opts;

		const chatService = accessor.get(IChatService);
		const widgetService = accessor.get(IChatWidgetService);
		const toolsService = accessor.get(ILanguageModelToolsService);
		const hostService = accessor.get(IHostService);
		const chatAgentService = accessor.get(IChatAgentService);
		const instaService = accessor.get(IInstantiationService);
		const commandService = accessor.get(ICommandService);
		const chatModeService = accessor.get(IChatModeService);
		const fileService = accessor.get(IFileService);
		const languageModelService = accessor.get(ILanguageModelsService);
		const scmService = accessor.get(ISCMService);
		const logService = accessor.get(ILogService);
		const configurationService = accessor.get(IConfigurationService);

		let chatWidget = widgetService.lastFocusedWidget;
		// When this was invoked to switch to a mode via keybinding, and some chat widget is focused, use that one.
		// Otherwise, open the view.
		if (!this.mode || !chatWidget || !isAncestorOfActiveElement(chatWidget.domNode)) {
			chatWidget = await widgetService.revealWidget();
		}

		if (!chatWidget) {
			return;
		}

		const switchToMode = (opts?.mode ? chatModeService.findModeByName(opts?.mode) : undefined) ?? this.mode;
		if (switchToMode) {
			await this.handleSwitchToMode(switchToMode, chatWidget, instaService, commandService);
		}

		if (opts?.modelSelector) {
			const ids = await languageModelService.selectLanguageModels(opts.modelSelector);
			const id = ids.sort().at(0);
			if (!id) {
				throw new Error(`No language models found matching selector: ${JSON.stringify(opts.modelSelector)}.`);
			}

			const model = languageModelService.lookupLanguageModel(id);
			if (!model) {
				throw new Error(`Language model not loaded: ${id}.`);
			}

			chatWidget.input.setCurrentLanguageModel({ metadata: model, identifier: id });
		}

		if (opts?.toolsInclude || opts?.toolsExclude) {
			const model = chatWidget.input.selectedLanguageModel.get()?.metadata;
			const allTools = Array.from(toolsService.getTools(model));
			const allToolSets = Array.from(toolsService.getToolSetsForModel(model));

			const result = computeToolEnablementMap({
				allTools,
				allToolSets,
				toolsInclude: opts.toolsInclude,
				toolsExclude: opts.toolsExclude,
			});

			for (const identifier of result.unknownIdentifiers) {
				logService.warn(`Tool filtering: Unknown identifier '${identifier}' - no matching tool or toolset found.`);
			}

			chatWidget.input.selectedToolsModel.set(result.enablementMap, true);
		}

		if (opts?.previousRequests?.length && chatWidget.viewModel) {
			for (const { request, response } of opts.previousRequests) {
				chatService.addCompleteRequest(chatWidget.viewModel.sessionResource, request, undefined, 0, { message: response });
			}
		}
		if (opts?.attachScreenshot) {
			const screenshot = await hostService.getScreenshot();
			if (screenshot) {
				chatWidget.attachmentModel.addContext(convertBufferToScreenshotVariable(screenshot));
			}
		}
		if (opts?.attachFiles) {
			for (const file of opts.attachFiles) {
				const uri = file instanceof URI ? file : file.uri;
				const range = file instanceof URI ? undefined : file.range;

				if (await fileService.exists(uri)) {
					chatWidget.attachmentModel.addFile(uri, range);
				}
			}
		}
		if (opts?.attachHistoryItemChanges) {
			for (const historyItemChange of opts.attachHistoryItemChanges) {
				const repository = scmService.getRepository(URI.file(historyItemChange.uri.path));
				const historyProvider = repository?.provider.historyProvider.get();
				if (!historyProvider) {
					continue;
				}

				const historyItem = await historyProvider.resolveHistoryItem(historyItemChange.historyItemId);
				if (!historyItem) {
					continue;
				}

				chatWidget.attachmentModel.addContext({
					id: historyItemChange.uri.toString(),
					name: `${basename(historyItemChange.uri)}`,
					value: historyItemChange.uri,
					historyItem: historyItem,
					kind: 'scmHistoryItemChange'
				} satisfies ISCMHistoryItemChangeVariableEntry);
			}
		}
		if (opts?.attachHistoryItemChangeRanges) {
			for (const historyItemChangeRange of opts.attachHistoryItemChangeRanges) {
				const repository = scmService.getRepository(URI.file(historyItemChangeRange.end.uri.path));
				const historyProvider = repository?.provider.historyProvider.get();
				if (!repository || !historyProvider) {
					continue;
				}

				const [historyItemStart, historyItemEnd] = await Promise.all([
					historyProvider.resolveHistoryItem(historyItemChangeRange.start.historyItemId),
					historyProvider.resolveHistoryItem(historyItemChangeRange.end.historyItemId),
				]);
				if (!historyItemStart || !historyItemEnd) {
					continue;
				}

				const uri = historyItemChangeRange.end.uri.with({
					scheme: SCMHistoryItemChangeRangeContentProvider.scheme,
					query: JSON.stringify({
						repositoryId: repository.id,
						start: historyItemStart.id,
						end: historyItemChangeRange.end.historyItemId
					} satisfies ScmHistoryItemChangeRangeUriFields)
				});

				chatWidget.attachmentModel.addContext({
					id: uri.toString(),
					name: `${basename(uri)}`,
					value: uri,
					historyItemChangeStart: {
						uri: historyItemChangeRange.start.uri,
						historyItem: historyItemStart
					},
					historyItemChangeEnd: {
						uri: historyItemChangeRange.end.uri,
						historyItem: {
							...historyItemEnd,
							displayId: historyItemChangeRange.end.historyItemId
						}
					},
					kind: 'scmHistoryItemChangeRange'
				} satisfies ISCMHistoryItemChangeRangeVariableEntry);
			}
		}

		let resp: Promise<IChatResponseModel | undefined> | undefined;

		if (opts?.query) {

			if (opts.isPartialQuery) {
				chatWidget.input.showScrollbarUntilAccept();
				chatWidget.setInput(opts.query);
			} else {
				if (!chatWidget.viewModel) {
					await Event.toPromise(chatWidget.onDidChangeViewModel);
				}
				await waitForDefaultAgent(chatAgentService, chatWidget.input.currentModeKind);
				chatWidget.setInput(opts.query); // wait until the model is restored before setting the input, or it will be cleared when the model is restored
				resp = chatWidget.acceptInput();
			}
		}

		if (opts?.toolIds && opts.toolIds.length > 0) {
			for (const toolId of opts.toolIds) {
				const tool = toolsService.getTool(toolId);
				if (tool) {
					chatWidget.attachmentModel.addContext({
						id: tool.id,
						name: tool.displayName,
						fullName: tool.displayName,
						value: undefined,
						icon: ThemeIcon.isThemeIcon(tool.icon) ? tool.icon : undefined,
						kind: 'tool'
					});
				}
			}
		}

		chatWidget.focusInput();

		if (opts?.blockOnResponse) {
			const response = await resp;
			if (response) {
				const autoReplyEnabled = configurationService.getValue<boolean>(ChatConfiguration.AutoReply);
				await new Promise<void>(resolve => {
					const d = response.onDidChange(async () => {
						if (response.isComplete) {
							d.dispose();
							resolve();
							return;
						}

						const pendingConfirmation = response.isPendingConfirmation.get();
						if (pendingConfirmation) {
							// Check if the pending confirmation is a question carousel that will be auto-replied.
							// Only question carousels are auto-replied; other confirmation types (tool approvals,
							// elicitations, etc.) should cause us to resolve immediately.
							const hasPendingQuestionCarousel = response.response.value.some(
								part => part.kind === 'questionCarousel' && !part.isUsed
							);
							if (autoReplyEnabled && hasPendingQuestionCarousel) {
								// Auto-reply will handle this question carousel, keep waiting
								return;
							}
							d.dispose();
							resolve();
						}
					});
				});

				const confirmationInfo = getPendingConfirmationInfo(response);
				if (confirmationInfo) {
					return { ...response.result, ...confirmationInfo };
				}
				return { ...response.result };
			}
		}

		return undefined;
	}

	private async handleSwitchToMode(switchToMode: IChatMode, chatWidget: IChatWidget, instaService: IInstantiationService, commandService: ICommandService): Promise<void> {
		const currentMode = chatWidget.input.currentModeKind;

		if (switchToMode) {
			const model = chatWidget.viewModel?.model;
			const chatModeCheck = model ? await instaService.invokeFunction(handleModeSwitch, currentMode, switchToMode.kind, model.getRequests().length, model) : { needToClearSession: false };
			if (!chatModeCheck) {
				return;
			}
			chatWidget.input.setChatMode(switchToMode.id);

			if (chatModeCheck.needToClearSession) {
				await commandService.executeCommand(ACTION_ID_NEW_CHAT);
			}
		}
	}
}

async function waitForDefaultAgent(chatAgentService: IChatAgentService, mode: ChatModeKind): Promise<void> {
	const defaultAgent = chatAgentService.getDefaultAgent(ChatAgentLocation.Chat, mode);
	if (defaultAgent) {
		return;
	}

	await Promise.race([
		Event.toPromise(Event.filter(chatAgentService.onDidChangeAgents, () => {
			const defaultAgent = chatAgentService.getDefaultAgent(ChatAgentLocation.Chat, mode);
			return Boolean(defaultAgent);
		})),
		timeout(60_000).then(() => { throw new Error('Timed out waiting for default agent'); })
	]);
}

/**
 * Information about a pending confirmation in a chat response.
 */
export type IChatPendingConfirmationInfo =
	| { type: 'confirmation'; kind: 'toolInvocation'; toolId: string }
	| { type: 'confirmation'; kind: 'toolPostApproval'; toolId: string }
	| { type: 'confirmation'; kind: 'confirmation'; title: string; data: unknown }
	| { type: 'confirmation'; kind: 'questionCarousel'; questions: unknown[] }
	| { type: 'confirmation'; kind: 'elicitation'; title: string };

/**
 * Extracts detailed information about the pending confirmation from a chat response.
 * Returns undefined if there is no pending confirmation.
 */
function getPendingConfirmationInfo(response: IChatResponseModel): IChatPendingConfirmationInfo | undefined {
	for (const part of response.response.value) {
		if (part.kind === 'toolInvocation') {
			const state = part.state.get();
			if (state.type === IChatToolInvocation.StateKind.WaitingForConfirmation) {
				return {
					type: 'confirmation',
					kind: 'toolInvocation',
					toolId: part.toolId,
				};
			}
			if (state.type === IChatToolInvocation.StateKind.WaitingForPostApproval) {
				return {
					type: 'confirmation',
					kind: 'toolPostApproval',
					toolId: part.toolId,
				};
			}
		}
		if (part.kind === 'confirmation' && !part.isUsed) {
			return {
				type: 'confirmation',
				kind: 'confirmation',
				title: part.title,
				data: part.data,
			};
		}
		if (part.kind === 'questionCarousel' && !part.isUsed) {
			return {
				type: 'confirmation',
				kind: 'questionCarousel',
				questions: part.questions,
			};
		}
		if (part.kind === 'elicitation2' && part.state.get() === ElicitationState.Pending) {
			const title = part.title;
			return {
				type: 'confirmation',
				kind: 'elicitation',
				title: typeof title === 'string' ? title : title.value,
			};
		}
	}
	return undefined;
}

class PrimaryOpenChatGlobalAction extends OpenChatGlobalAction {
	constructor() {
		super({
			id: CHAT_OPEN_ACTION_ID,
			title: localize2('openChat', "Open Chat"),
			keybinding: {
				weight: KeybindingWeight.WorkbenchContrib,
				primary: KeyMod.CtrlCmd | KeyMod.Alt | KeyCode.KeyI,
				mac: {
					primary: KeyMod.CtrlCmd | KeyMod.WinCtrl | KeyCode.KeyI
				}
			},
			menu: [{
				id: MenuId.ChatTitleBarMenu,
				group: 'a_open',
				order: 1
			}]
		});
	}
}

export function getOpenChatActionIdForMode(mode: IChatMode): string {
	return `workbench.action.chat.open${mode.name.get()}`;
}

export abstract class ModeOpenChatGlobalAction extends OpenChatGlobalAction {
	constructor(mode: IChatMode, keybinding?: ICommandPaletteOptions['keybinding']) {
		super({
			id: getOpenChatActionIdForMode(mode),
			title: localize2('openChatMode', "Open Chat ({0})", mode.label.get()),
			keybinding
		}, mode);
	}
}

export function registerChatActions() {
	registerAction2(PrimaryOpenChatGlobalAction);
	registerAction2(class extends ModeOpenChatGlobalAction {
		constructor() { super(ChatMode.Ask); }
	});
	registerAction2(class extends ModeOpenChatGlobalAction {
		constructor() {
			super(ChatMode.Agent, {
				when: ContextKeyExpr.has(`config.${ChatConfiguration.AgentEnabled}`),
				weight: KeybindingWeight.WorkbenchContrib,
				primary: KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.KeyI,
				linux: {
					primary: KeyMod.CtrlCmd | KeyMod.Alt | KeyMod.Shift | KeyCode.KeyI
				}
			},);
		}
	});
	registerAction2(class extends ModeOpenChatGlobalAction {
		constructor() { super(ChatMode.Edit); }
	});

	registerAction2(class ToggleChatAction extends Action2 {
		constructor() {
			super({
				id: TOGGLE_CHAT_ACTION_ID,
				title: localize2('toggleChat', "Toggle Chat"),
				category: CHAT_CATEGORY
			});
		}

		async run(accessor: ServicesAccessor) {
			const layoutService = accessor.get(IWorkbenchLayoutService);
			const viewsService = accessor.get(IViewsService);
			const viewDescriptorService = accessor.get(IViewDescriptorService);
			const widgetService = accessor.get(IChatWidgetService);

			const chatLocation = viewDescriptorService.getViewLocationById(ChatViewId);
			const chatVisible = viewsService.isViewVisible(ChatViewId);
			if (chatVisible) {
				this.updatePartVisibility(layoutService, chatLocation, false);
			} else {
				this.updatePartVisibility(layoutService, chatLocation, true);
				(await widgetService.revealWidget())?.focusInput();
			}
		}

		private updatePartVisibility(layoutService: IWorkbenchLayoutService, location: ViewContainerLocation | null, visible: boolean): void {
			let part: Parts.PANEL_PART | Parts.SIDEBAR_PART | Parts.AUXILIARYBAR_PART | undefined;
			switch (location) {
				case ViewContainerLocation.Panel:
					part = Parts.PANEL_PART;
					break;
				case ViewContainerLocation.Sidebar:
					part = Parts.SIDEBAR_PART;
					break;
				case ViewContainerLocation.AuxiliaryBar:
					part = Parts.AUXILIARYBAR_PART;
					break;
			}

			if (part) {
				layoutService.setPartHidden(!visible, part);
			}
		}
	});


	registerAction2(class NewChatEditorAction extends Action2 {
		constructor() {
			super({
				id: ACTION_ID_OPEN_CHAT,
				title: localize2('interactiveSession.open', "New Chat Editor"),
				icon: Codicon.plus,
				f1: true,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
				keybinding: {
					weight: KeybindingWeight.WorkbenchContrib,
					primary: KeyMod.CtrlCmd | KeyCode.KeyN,
					when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.inChatEditor)
				},
				menu: [{
					id: MenuId.ChatTitleBarMenu,
					group: 'b_new',
					order: 0
				}, {
					id: MenuId.ChatNewMenu,
					group: '2_new',
					order: 2
				}, {
					id: MenuId.EditorTitle,
					group: 'navigation',
					when: ContextKeyExpr.and(ActiveEditorContext.isEqualTo(ChatEditorInput.EditorID), ChatContextKeys.newChatButtonExperimentIcon.notEqualsTo('copilot'), ChatContextKeys.newChatButtonExperimentIcon.notEqualsTo('new-session'), ChatContextKeys.newChatButtonExperimentIcon.notEqualsTo('comment')),
					order: 1
				}],
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), ACTIVE_GROUP, { pinned: true } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class NewChatEditorCopilotIconAction extends Action2 {
		constructor() {
			super({
				id: ACTION_ID_OPEN_CHAT + '.copilotIcon',
				title: localize2('interactiveSession.open', "New Chat Editor"),
				icon: Codicon.copilot,
				f1: false,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
				menu: [{
					id: MenuId.EditorTitle,
					group: 'navigation',
					when: ContextKeyExpr.and(ActiveEditorContext.isEqualTo(ChatEditorInput.EditorID), ChatContextKeys.newChatButtonExperimentIcon.isEqualTo('copilot')),
					order: 1
				}],
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), ACTIVE_GROUP, { pinned: true } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class NewChatEditorNewSessionIconAction extends Action2 {
		constructor() {
			super({
				id: ACTION_ID_OPEN_CHAT + '.newSessionIcon',
				title: localize2('interactiveSession.open', "New Chat Editor"),
				icon: Codicon.newSession,
				f1: false,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
				menu: [{
					id: MenuId.EditorTitle,
					group: 'navigation',
					when: ContextKeyExpr.and(ActiveEditorContext.isEqualTo(ChatEditorInput.EditorID), ChatContextKeys.newChatButtonExperimentIcon.isEqualTo('new-session')),
					order: 1
				}],
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), ACTIVE_GROUP, { pinned: true } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class NewChatEditorCommentIconAction extends Action2 {
		constructor() {
			super({
				id: ACTION_ID_OPEN_CHAT + '.commentIcon',
				title: localize2('interactiveSession.open', "New Chat Editor"),
				icon: Codicon.comment,
				f1: false,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
				menu: [{
					id: MenuId.EditorTitle,
					group: 'navigation',
					when: ContextKeyExpr.and(ActiveEditorContext.isEqualTo(ChatEditorInput.EditorID), ChatContextKeys.newChatButtonExperimentIcon.isEqualTo('comment')),
					order: 1
				}],
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), ACTIVE_GROUP, { pinned: true } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class NewChatEditorToSideAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.openChatToSide',
				title: localize2('interactiveSession.openToSide', "New Chat Editor to the Side"),
				f1: true,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), SIDE_GROUP, { pinned: true } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class NewChatWindowAction extends Action2 {
		constructor() {
			super({
				id: `workbench.action.newChatWindow`,
				title: localize2('interactiveSession.newChatWindow', "New Chat Window"),
				f1: true,
				category: CHAT_CATEGORY,
				precondition: ChatContextKeys.enabled,
				menu: [{
					id: MenuId.ChatTitleBarMenu,
					group: 'b_new',
					order: 1
				}, {
					id: MenuId.ChatNewMenu,
					group: '2_new',
					order: 3
				}]
			});
		}

		async run(accessor: ServicesAccessor) {
			const widgetService = accessor.get(IChatWidgetService);
			await widgetService.openSession(LocalChatSessionUri.getNewSessionUri(), AUX_WINDOW_GROUP, { pinned: true, auxiliary: { compact: true, bounds: { width: 640, height: 640 } } } satisfies IChatEditorOptions);
		}
	});

	registerAction2(class ClearChatInputHistoryAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.clearInputHistory',
				title: localize2('interactiveSession.clearHistory.label', "Clear Input History"),
				precondition: ChatContextKeys.enabled,
				category: CHAT_CATEGORY,
				f1: true,
			});
		}
		async run(accessor: ServicesAccessor, ...args: unknown[]) {
			const historyService = accessor.get(IChatWidgetHistoryService);
			historyService.clearHistory();
		}
	});

	registerAction2(class FocusChatAction extends EditorAction2 {
		constructor() {
			super({
				id: 'chat.action.focus',
				title: localize2('actions.interactiveSession.focus', 'Focus Chat List'),
				precondition: ContextKeyExpr.and(ChatContextKeys.inChatInput),
				category: CHAT_CATEGORY,
				keybinding: [
					// On mac, require that the cursor is at the top of the input, to avoid stealing cmd+up to move the cursor to the top
					{
						when: ContextKeyExpr.and(ChatContextKeys.inputCursorAtTop, ChatContextKeys.inQuickChat.negate()),
						primary: KeyMod.CtrlCmd | KeyCode.UpArrow,
						weight: KeybindingWeight.EditorContrib,
					},
					// On win/linux, ctrl+up can always focus the chat list
					{
						when: ContextKeyExpr.and(ContextKeyExpr.or(IsWindowsContext, IsLinuxContext), ChatContextKeys.inQuickChat.negate()),
						primary: KeyMod.CtrlCmd | KeyCode.UpArrow,
						weight: KeybindingWeight.EditorContrib,
					},
					{
						when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.inQuickChat),
						primary: KeyMod.CtrlCmd | KeyCode.DownArrow,
						weight: KeybindingWeight.WorkbenchContrib,
					}
				]
			});
		}

		runEditorCommand(accessor: ServicesAccessor, editor: ICodeEditor): void | Promise<void> {
			const editorUri = editor.getModel()?.uri;
			if (editorUri) {
				const widgetService = accessor.get(IChatWidgetService);
				widgetService.getWidgetByInputUri(editorUri)?.focusResponseItem();
			}
		}
	});

	registerAction2(class FocusMostRecentlyFocusedChatAction extends EditorAction2 {
		constructor() {
			super({
				id: 'workbench.chat.action.focusLastFocused',
				title: localize2('actions.interactiveSession.focusLastFocused', 'Focus Last Focused Chat List Item'),
				precondition: ContextKeyExpr.and(ChatContextKeys.inChatInput),
				category: CHAT_CATEGORY,
				keybinding: [
					// On mac, require that the cursor is at the top of the input, to avoid stealing cmd+up to move the cursor to the top
					{
						when: ContextKeyExpr.and(ChatContextKeys.inputCursorAtTop, ChatContextKeys.inQuickChat.negate()),
						primary: KeyMod.CtrlCmd | KeyCode.UpArrow | KeyMod.Shift,
						weight: KeybindingWeight.EditorContrib + 1,
					},
					// On win/linux, ctrl+up can always focus the chat list
					{
						when: ContextKeyExpr.and(ContextKeyExpr.or(IsWindowsContext, IsLinuxContext), ChatContextKeys.inQuickChat.negate()),
						primary: KeyMod.CtrlCmd | KeyCode.UpArrow | KeyMod.Shift,
						weight: KeybindingWeight.EditorContrib + 1,
					},
					{
						when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.inQuickChat),
						primary: KeyMod.CtrlCmd | KeyCode.DownArrow | KeyMod.Shift,
						weight: KeybindingWeight.WorkbenchContrib + 1,
					}
				]
			});
		}

		runEditorCommand(accessor: ServicesAccessor, editor: ICodeEditor): void | Promise<void> {
			const editorUri = editor.getModel()?.uri;
			if (editorUri) {
				const widgetService = accessor.get(IChatWidgetService);
				widgetService.getWidgetByInputUri(editorUri)?.focusResponseItem(true);
			}
		}
	});

	registerAction2(class FocusChatInputAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.focusInput',
				title: localize2('interactiveSession.focusInput.label', "Focus Chat Input"),
				f1: false,
				keybinding: [
					{
						primary: KeyMod.CtrlCmd | KeyCode.DownArrow,
						weight: KeybindingWeight.WorkbenchContrib,
						when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.inChatInput.negate(), ChatContextKeys.inQuickChat.negate()),
					},
					{
						when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.inChatInput.negate(), ChatContextKeys.inQuickChat),
						primary: KeyMod.CtrlCmd | KeyCode.UpArrow,
						weight: KeybindingWeight.WorkbenchContrib,
					}
				]
			});
		}
		run(accessor: ServicesAccessor, ...args: unknown[]) {
			const widgetService = accessor.get(IChatWidgetService);
			widgetService.lastFocusedWidget?.focusInput();
		}
	});

	registerAction2(class FocusTodosViewAction extends Action2 {
		static readonly ID = 'workbench.action.chat.focusTodosView';

		constructor() {
			super({
				id: FocusTodosViewAction.ID,
				title: localize2('interactiveSession.focusTodosView.label', "Toggle Focus Between TODOs and Input"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.chatModeKind.isEqualTo(ChatModeKind.Agent),
				keybinding: [{
					weight: KeybindingWeight.WorkbenchContrib + 1,
					primary: KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.KeyT,
					when: ContextKeyExpr.or(
						ContextKeyExpr.and(ChatContextKeys.inChatInput, ChatContextKeys.chatModeKind.isEqualTo(ChatModeKind.Agent)),
						ContextKeyExpr.and(ChatContextKeys.inChatTodoList, ChatContextKeys.chatModeKind.isEqualTo(ChatModeKind.Agent))
					),
				}]
			});
		}

		run(accessor: ServicesAccessor): void {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;

			if (!widget || !widget.toggleTodosViewFocus()) {
				alert(localize('chat.todoList.focusUnavailable', "No agent todos to focus right now."));
			}
		}
	});

	registerAction2(class FocusQuestionCarouselAction extends Action2 {
		static readonly ID = 'workbench.action.chat.focusQuestionCarousel';

		constructor() {
			super({
				id: FocusQuestionCarouselAction.ID,
				title: localize2('interactiveSession.focusQuestionCarousel.label', "Chat: Toggle Focus Between Question and Input"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.inChatSession,
				keybinding: [{
					weight: KeybindingWeight.WorkbenchContrib,
					primary: KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.KeyA,
					when: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.Editing.hasQuestionCarousel),
				}]
			});
		}

		run(accessor: ServicesAccessor): void {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;

			if (!widget || !widget.toggleQuestionCarouselFocus()) {
				alert(localize('chat.questionCarousel.focusUnavailable', "No chat question to focus right now."));
			}
		}
	});

	registerAction2(class PreviousQuestionCarouselQuestionAction extends Action2 {
		static readonly ID = 'workbench.action.chat.previousQuestion';

		constructor() {
			super({
				id: PreviousQuestionCarouselQuestionAction.ID,
				title: localize2('interactiveSession.previousQuestion.label', "Chat: Previous Question"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.Editing.hasQuestionCarousel),
				keybinding: [{
					weight: KeybindingWeight.WorkbenchContrib,
					primary: KeyMod.Alt | KeyCode.KeyP,
					when: ContextKeyExpr.and(ChatContextKeys.inChatQuestionCarousel, ChatContextKeys.Editing.hasQuestionCarousel),
				}]
			});
		}

		run(accessor: ServicesAccessor): void {
			const widgetService = accessor.get(IChatWidgetService);
			widgetService.lastFocusedWidget?.navigateToPreviousQuestion();
		}
	});

	registerAction2(class NextQuestionCarouselQuestionAction extends Action2 {
		static readonly ID = 'workbench.action.chat.nextQuestion';

		constructor() {
			super({
				id: NextQuestionCarouselQuestionAction.ID,
				title: localize2('interactiveSession.nextQuestion.label', "Chat: Next Question"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(ChatContextKeys.inChatSession, ChatContextKeys.Editing.hasQuestionCarousel),
				keybinding: [{
					weight: KeybindingWeight.WorkbenchContrib,
					primary: KeyMod.Alt | KeyCode.KeyN,
					when: ContextKeyExpr.and(ChatContextKeys.inChatQuestionCarousel, ChatContextKeys.Editing.hasQuestionCarousel),
				}]
			});
		}

		run(accessor: ServicesAccessor): void {
			const widgetService = accessor.get(IChatWidgetService);
			widgetService.lastFocusedWidget?.navigateToNextQuestion();
		}
	});

	registerAction2(class RefinePlanAction extends Action2 {
		static readonly ID = 'workbench.action.chat.refinePlan';

		constructor() {
			super({
				id: RefinePlanAction.ID,
				title: localize2('interactiveSession.refinePlan.label', "Chat: Refine Plan"),
				tooltip: localize('refinePlan', "Refine Plan"),
				icon: Codicon.refresh,
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(
					ChatContextKeys.inChatSession,
					ContextKeyExpr.or(
						ChatContextKeys.chatModeName.isEqualTo('Plan'),
						ChatContextKeys.chatModeName.isEqualTo('Planner'),
						ChatContextKeys.chatModeName.isEqualTo('planner')
					)
				),
				menu: {
					id: MenuId.ChatInputSecondary,
					order: 11,
					group: 'navigation',
					when: ContextKeyExpr.and(
						ChatContextKeys.enabled,
						ChatContextKeys.inChatSession,
						ContextKeyExpr.or(
							ChatContextKeys.chatModeName.isEqualTo('Plan'),
							ChatContextKeys.chatModeName.isEqualTo('Planner'),
							ChatContextKeys.chatModeName.isEqualTo('planner')
						),
						ChatContextKeys.location.isEqualTo(ChatAgentLocation.Chat),
						ChatContextKeys.inQuickChat.negate()
					)
				}
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;
			if (!widget || !(await widget.refinePlan())) {
				alert(localize('chat.refinePlan.unavailable', "Enter or revise a planning request first to refine the plan."));
			}
		}
	});

	registerAction2(class PreviousPlanningPhaseAction extends Action2 {
		static readonly ID = 'workbench.action.chat.previousPlanningPhase';

		constructor() {
			super({
				id: PreviousPlanningPhaseAction.ID,
				title: localize2('interactiveSession.previousPlanningPhase.label', "Chat: Previous Planning Phase"),
				tooltip: localize('previousPlanningPhase', "Previous Phase"),
				icon: Codicon.arrowLeft,
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(
					ChatContextKeys.inChatSession,
					ContextKeyExpr.or(
						ChatContextKeys.chatModeName.isEqualTo('Plan'),
						ChatContextKeys.chatModeName.isEqualTo('Planner'),
						ChatContextKeys.chatModeName.isEqualTo('planner')
					)
				),
				menu: {
					id: MenuId.ChatInputSecondary,
					order: 11.1,
					group: 'navigation',
					when: ContextKeyExpr.and(
						ChatContextKeys.enabled,
						ChatContextKeys.inChatSession,
						ContextKeyExpr.or(
							ChatContextKeys.chatModeName.isEqualTo('Plan'),
							ChatContextKeys.chatModeName.isEqualTo('Planner'),
							ChatContextKeys.chatModeName.isEqualTo('planner')
						),
						ChatContextKeys.location.isEqualTo(ChatAgentLocation.Chat),
						ChatContextKeys.inQuickChat.negate()
					)
				}
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;
			if (!widget || !(await widget.retreatPlanPhase())) {
				alert(localize('chat.previousPlanningPhase.unavailable', "There is no earlier planning phase available."));
			}
		}
	});

	registerAction2(class NextPlanningPhaseAction extends Action2 {
		static readonly ID = 'workbench.action.chat.nextPlanningPhase';

		constructor() {
			super({
				id: NextPlanningPhaseAction.ID,
				title: localize2('interactiveSession.nextPlanningPhase.label', "Chat: Next Planning Phase"),
				tooltip: localize('nextPlanningPhase', "Next Phase"),
				icon: Codicon.arrowRight,
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(
					ChatContextKeys.inChatSession,
					ContextKeyExpr.or(
						ChatContextKeys.chatModeName.isEqualTo('Plan'),
						ChatContextKeys.chatModeName.isEqualTo('Planner'),
						ChatContextKeys.chatModeName.isEqualTo('planner')
					)
				),
				menu: {
					id: MenuId.ChatInputSecondary,
					order: 11.2,
					group: 'navigation',
					when: ContextKeyExpr.and(
						ChatContextKeys.enabled,
						ChatContextKeys.inChatSession,
						ContextKeyExpr.or(
							ChatContextKeys.chatModeName.isEqualTo('Plan'),
							ChatContextKeys.chatModeName.isEqualTo('Planner'),
							ChatContextKeys.chatModeName.isEqualTo('planner')
						),
						ChatContextKeys.location.isEqualTo(ChatAgentLocation.Chat),
						ChatContextKeys.inQuickChat.negate()
					)
				}
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;
			if (!widget || !(await widget.advancePlanPhase())) {
				alert(localize('chat.nextPlanningPhase.unavailable', "There is no later planning phase available."));
			}
		}
	});

	registerAction2(class FocusTipAction extends Action2 {
		static readonly ID = 'workbench.action.chat.focusTip';

		constructor() {
			super({
				id: FocusTipAction.ID,
				title: localize2('interactiveSession.focusTip.label', "Chat: Toggle Focus Between Tip and Input"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.inChatSession,
				keybinding: [{
					weight: KeybindingWeight.WorkbenchContrib,
					primary: KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.Slash,
					when: ContextKeyExpr.or(
						ChatContextKeys.inChatSession,
						ChatContextKeys.inChatTip
					),
				}]
			});
		}

		run(accessor: ServicesAccessor): void {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget;

			if (!widget || !widget.toggleTipFocus()) {
				alert(localize('chat.tip.focusUnavailable', "No chat tip."));
			}
		}
	});

	registerAction2(class ShowContextUsageAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.showContextUsage',
				title: localize2('interactiveSession.showContextUsage.label', "Show Context Window Usage"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.enabled,
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const widgetService = accessor.get(IChatWidgetService);
			const widget = widgetService.lastFocusedWidget ?? (await widgetService.revealWidget());
			widget?.input.showContextUsageDetails();
		}
	});

	registerAction2(class ToggleShowContextUsageAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.toggleShowContextUsage',
				title: localize2('chat.showContextUsage', "Show Context Usage"),
				category: CHAT_CATEGORY,
				toggled: ContextKeyExpr.equals(`config.${ChatConfiguration.ChatContextUsageEnabled}`, true),
				menu: {
					id: MenuId.ChatWelcomeContext,
					group: '1_display',
					order: 1,
					when: ChatContextKeys.inChatEditor.negate()
				}
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const configurationService = accessor.get(IConfigurationService);
			const currentValue = configurationService.getValue<boolean>(ChatConfiguration.ChatContextUsageEnabled);
			await configurationService.updateValue(ChatConfiguration.ChatContextUsageEnabled, !currentValue);
		}
	});

	const nonEnterpriseCopilotUsers = ContextKeyExpr.and(ChatContextKeys.enabled, ContextKeyExpr.notEquals(`config.${defaultChat.completionsAdvancedSetting}.authProvider`, defaultChat.provider.enterprise.id));
	registerAction2(class extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.manageSettings',
				title: localize2('manageChat', "Manage Chat"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ContextKeyExpr.and(
					ContextKeyExpr.or(
						ChatContextKeys.Entitlement.planFree,
						ChatContextKeys.Entitlement.planPro,
						ChatContextKeys.Entitlement.planProPlus
					),
					nonEnterpriseCopilotUsers
				),
				menu: {
					id: MenuId.ChatTitleBarMenu,
					group: 'y_manage',
					order: 1,
					when: nonEnterpriseCopilotUsers
				}
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const openerService = accessor.get(IOpenerService);
			openerService.open(URI.parse(defaultChat.manageSettingsUrl));
		}
	});

	registerAction2(class ShowExtensionsUsingCopilot extends Action2 {

		constructor() {
			super({
				id: 'workbench.action.chat.showExtensionsUsingCopilot',
				title: localize2('showCopilotUsageExtensions', "Show Extensions using Copilot"),
				f1: true,
				category: EXTENSIONS_CATEGORY,
				precondition: ChatContextKeys.enabled
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const extensionsWorkbenchService = accessor.get(IExtensionsWorkbenchService);
			extensionsWorkbenchService.openSearch(`@contribute:${CopilotUsageExtensionFeatureId}`);
		}
	});

	registerAction2(class ConfigureCopilotCompletions extends Action2 {

		constructor() {
			super({
				id: 'workbench.action.chat.configureCodeCompletions',
				title: localize2('configureCompletions', "Configure Inline Suggestions..."),
				precondition: ContextKeyExpr.and(
					ChatContextKeys.Setup.installed,
					ChatContextKeys.Setup.disabled.negate(),
					ChatContextKeys.Setup.untrusted.negate()
				),
				menu: {
					id: MenuId.ChatTitleBarMenu,
					group: 'f_completions',
					order: 10,
				}
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			commandService.executeCommand(defaultChat.completionsMenuCommand);
		}
	});

	registerAction2(class ShowQuotaExceededDialogAction extends Action2 {

		constructor() {
			super({
				id: OPEN_CHAT_QUOTA_EXCEEDED_DIALOG,
				title: localize('upgradeChat', "Upgrade GitHub Copilot Plan")
			});
		}

		override async run(accessor: ServicesAccessor) {
			const chatEntitlementService = accessor.get(IChatEntitlementService);
			const commandService = accessor.get(ICommandService);
			const dialogService = accessor.get(IDialogService);
			const telemetryService = accessor.get(ITelemetryService);

			let message: string;
			const chatQuotaExceeded = chatEntitlementService.quotas.chat?.percentRemaining === 0;
			const completionsQuotaExceeded = chatEntitlementService.quotas.completions?.percentRemaining === 0;
			if (chatQuotaExceeded && !completionsQuotaExceeded) {
				message = localize('chatQuotaExceeded', "You've reached your monthly chat messages quota. You still have free inline suggestions available.");
			} else if (completionsQuotaExceeded && !chatQuotaExceeded) {
				message = localize('completionsQuotaExceeded', "You've reached your monthly inline suggestions quota. You still have free chat messages available.");
			} else {
				message = localize('chatAndCompletionsQuotaExceeded', "You've reached your monthly chat messages and inline suggestions quota.");
			}

			if (chatEntitlementService.quotas.resetDate) {
				const dateFormatter = chatEntitlementService.quotas.resetDateHasTime ? safeIntl.DateTimeFormat(language, { year: 'numeric', month: 'long', day: 'numeric', hour: 'numeric', minute: 'numeric' }) : safeIntl.DateTimeFormat(language, { year: 'numeric', month: 'long', day: 'numeric' });
				const quotaResetDate = new Date(chatEntitlementService.quotas.resetDate);
				message = [message, localize('quotaResetDate', "The allowance will reset on {0}.", dateFormatter.value.format(quotaResetDate))].join(' ');
			}

			const free = chatEntitlementService.entitlement === ChatEntitlement.Free;
			const upgradeToPro = free ? localize('upgradeToPro', "Upgrade to GitHub Copilot Pro (your first 30 days are free) for:\n- Unlimited inline suggestions\n- Unlimited chat messages\n- Access to premium models") : undefined;

			await dialogService.prompt({
				type: 'none',
				message: localize('copilotQuotaReached', "GitHub Copilot Quota Reached"),
				cancelButton: {
					label: localize('dismiss', "Dismiss"),
					run: () => { /* noop */ }
				},
				buttons: [
					{
						label: free ? localize('upgradePro', "Upgrade to GitHub Copilot Pro") : localize('upgradePlan', "Upgrade GitHub Copilot Plan"),
						run: () => {
							const commandId = 'workbench.action.chat.upgradePlan';
							telemetryService.publicLog2<WorkbenchActionExecutedEvent, WorkbenchActionExecutedClassification>('workbenchActionExecuted', { id: commandId, from: 'chat-dialog' });
							commandService.executeCommand(commandId);
						}
					},
				],
				custom: {
					icon: Codicon.copilotWarningLarge,
					markdownDetails: coalesce([
						{ markdown: new MarkdownString(message, true) },
						upgradeToPro ? { markdown: new MarkdownString(upgradeToPro, true) } : undefined
					])
				}
			});
		}
	});

	registerAction2(class ResetTrustedToolsAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.resetTrustedTools',
				title: localize2('resetTrustedTools', "Reset Tool Confirmations"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}
		override run(accessor: ServicesAccessor): void {
			accessor.get(ILanguageModelToolsConfirmationService).resetToolAutoConfirmation();
			accessor.get(INotificationService).info(localize('resetTrustedToolsSuccess', "Tool confirmation preferences have been reset."));
		}
	});

	registerAction2(class GenerateInstructionsAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_AGENT_INSTRUCTIONS_COMMAND_ID,
				title: localize2('generateInstructions', "Generate Agent Instructions"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/init',
				isPartialQuery: false,
			});
		}
	});

	registerAction2(class GenerateInstructionAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_ON_DEMAND_INSTRUCTIONS_COMMAND_ID,
				title: localize2('generateOnDemandInstructions', "Generate On-Demand Instructions"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/create-instructions ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class GeneratePromptAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_PROMPT_COMMAND_ID,
				title: localize2('generatePrompt', "Generate Prompt File"),
				shortTitle: localize2('generatePrompt.short', "Generate Prompt"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/create-prompt ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class GenerateSkillAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_SKILL_COMMAND_ID,
				title: localize2('generateSkill', "Generate Skill"),
				shortTitle: localize2('generateSkill.short', "Generate Skill"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/create-skill ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class GenerateAgentAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_AGENT_COMMAND_ID,
				title: localize2('generateAgent', "Generate Custom Agent"),
				shortTitle: localize2('generateAgent.short', "Generate Agent"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/create-agent ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class GenerateHookAction extends Action2 {
		constructor() {
			super({
				id: GENERATE_HOOK_COMMAND_ID,
				title: localize2('generateHook', "Generate Hook"),
				shortTitle: localize2('generateHook.short', "Generate Hook"),
				category: CHAT_CATEGORY,
				icon: Codicon.sparkle,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				mode: 'agent',
				query: '/create-hook ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class InsertForkConversationSlashCommandAction extends Action2 {
		constructor() {
			super({
				id: INSERT_FORK_CONVERSATION_COMMAND_ID,
				title: localize2('insertForkConversationSlashCommand', "Insert Fork Command"),
				shortTitle: localize2('insertForkConversationSlashCommand.short', "Insert /fork"),
				category: CHAT_CATEGORY,
				icon: Codicon.repoForked,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				query: '/fork ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class InsertTroubleshootSlashCommandAction extends Action2 {
		constructor() {
			super({
				id: INSERT_TROUBLESHOOT_COMMAND_ID,
				title: localize2('insertTroubleshootSlashCommand', "Insert Troubleshoot Command"),
				shortTitle: localize2('insertTroubleshootSlashCommand.short', "Insert /troubleshoot"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.enabled
			});
		}

		async run(accessor: ServicesAccessor): Promise<void> {
			const commandService = accessor.get(ICommandService);
			await commandService.executeCommand('workbench.action.chat.open', {
				query: '/troubleshoot ',
				isPartialQuery: true,
			});
		}
	});

	registerAction2(class OpenChatFeatureSettingsAction extends Action2 {
		constructor() {
			super({
				id: 'workbench.action.chat.openFeatureSettings',
				title: localize2('openChatFeatureSettings', "Chat Settings"),
				shortTitle: localize('openChatFeatureSettings.short', "Chat Settings"),
				category: CHAT_CATEGORY,
				f1: true,
				precondition: ChatContextKeys.enabled,
				menu: [{
					id: CHAT_CONFIG_MENU_ID,
					when: ContextKeyExpr.and(ChatContextKeys.enabled, ContextKeyExpr.equals('view', ChatViewId)),
					order: 15,
					group: '3_configure'
				},
				{
					id: MenuId.ChatWelcomeContext,
					group: '2_settings',
					order: 1
				},
				{
					id: MenuId.ViewTitle,
					when: ContextKeyExpr.and(ChatContextKeys.enabled, ContextKeyExpr.equals('view', ChatViewId), ContextKeyExpr.has(`config.${ChatConfiguration.ChatCustomizationMenuEnabled}`)),
					order: 15,
					group: '3_configure'
				}]
			});
		}

		override async run(accessor: ServicesAccessor): Promise<void> {
			const preferencesService = accessor.get(IPreferencesService);
			preferencesService.openSettings({ query: '@feature:chat ' });
		}
	});

	function revivePlanningSessionResource(sessionResource: URI | string): URI {
		return typeof sessionResource === 'string'
			? URI.parse(sessionResource)
			: URI.revive(sessionResource);
	}

	function getPlanningPlanText(chatService: IChatService, sessionResource: URI, requestId: string): string | undefined {
		const request = chatService.getSession(sessionResource)?.getRequests().find(candidate => candidate.id === requestId);
		return extractPlanningPlanText(request?.response?.entireResponse)
			?? extractPlanningPlanText(request?.response?.response);
	}

	interface IPlanningPlanWebviewQuestion {
		readonly id: string;
		readonly type: IChatQuestion['type'];
		readonly title: string;
		readonly message?: string;
		readonly description?: string;
		readonly options?: readonly { readonly id: string; readonly label: string; readonly value: string }[];
		readonly defaultValue?: string | string[];
		readonly allowFreeformInput?: boolean;
		readonly required?: boolean;
	}

	interface IPlanningPlanWebviewStep {
		readonly id: string;
		readonly index: number;
		readonly label: string;
		readonly text: string;
		readonly sectionTitle?: string;
		readonly kind: IChatPlanningPlanEditorStep['kind'];
		readonly questions?: readonly IPlanningPlanWebviewQuestion[];
	}

	function getPlanningPlanEditorSteps(currentPlanText: string, planSteps?: readonly IChatPlanningPlanEditorStep[]): readonly IChatPlanningPlanEditorStep[] {
		const extractedSteps = extractPlanningPlanSteps(currentPlanText, 24);
		return planSteps?.length
			? planSteps
			: extractedSteps.length > 0
				? extractedSteps.map(step => ({
					id: `step-${step.index}`,
					index: step.index,
					label: step.label,
					text: step.text,
					sectionTitle: step.sectionTitle,
					kind: step.kind,
				}))
				: [{
					index: 1,
					id: 'step-1',
					label: localize('openPlanningPlan.wholePlanFallbackLabel', 'Review the whole plan'),
					text: currentPlanText,
					kind: 'step' as const,
				}];
	}

	function toPlanningPlanWebviewQuestion(question: IChatQuestion): IPlanningPlanWebviewQuestion {
		const message = typeof question.message === 'string' ? question.message : question.message?.value;
		return {
			id: question.id,
			type: question.type,
			title: question.title,
			...(message ? { message } : {}),
			...(question.description ? { description: question.description } : {}),
			...(question.options?.length ? { options: question.options.map(option => ({ id: option.id, label: option.label, value: option.value })) } : {}),
			...(question.defaultValue !== undefined ? { defaultValue: question.defaultValue } : {}),
			...(question.allowFreeformInput !== undefined ? { allowFreeformInput: question.allowFreeformInput } : {}),
			...(question.required !== undefined ? { required: question.required } : {}),
		};
	}

	function toPlanningPlanWebviewStep(step: IChatPlanningPlanEditorStep): IPlanningPlanWebviewStep {
		const questions = step.questions ?? [];
		return {
			id: step.id,
			index: step.index,
			label: step.label,
			text: step.text,
			...(step.sectionTitle ? { sectionTitle: step.sectionTitle } : {}),
			kind: step.kind,
			questions: questions.map(toPlanningPlanWebviewQuestion),
		};
	}

	function toWebviewScriptLiteral(value: unknown): string {
		return JSON.stringify(value).replace(/[<>&\u2028\u2029]/g, character => {
			switch (character) {
				case '<': return '\\u003c';
				case '>': return '\\u003e';
				case '&': return '\\u0026';
				case '\u2028': return '\\u2028';
				case '\u2029': return '\\u2029';
				default: return character;
			}
		});
	}

	function getPlanningPlanWebviewStateKey(planText: string): string {
		let hash = 0;
		for (let i = 0; i < planText.length; i++) {
			hash = ((hash << 5) - hash + planText.charCodeAt(i)) | 0;
		}
		return `${planText.length}:${hash.toString(36)}`;
	}

	function getPlanningPlanControlsStateKey(visiblePlanSteps: readonly IChatPlanningPlanEditorStep[]): string {
		return visiblePlanSteps.map(step => [
			step.index,
			...(step.questions ?? []).map(question => `${question.id}:${question.title}:${question.type}:${question.options?.map(option => option.value).join(',') ?? ''}`)
		].join('|')).join(';');
	}

	function isPlanningPlanWebviewSubmitMessage(message: unknown): message is { readonly type: 'apply' | 'continue' | 'regenerateControls'; readonly answers?: IChatQuestionAnswers } {
		return typeof message === 'object'
			&& message !== null
			&& ('type' in message)
			&& (message.type === 'apply' || message.type === 'continue' || message.type === 'regenerateControls');
	}

	function createPlanningPlanEditorInput(title: string, requestId: string, contents: string): IUntitledTextResourceEditorInput {
		return {
			resource: URI.from({
				scheme: Schemas.untitled,
				path: `/${title.replace(/[\\/:*?"<>|]+/g, '-')}-${requestId.slice(-6)}.md`,
			}),
			contents,
			languageId: 'markdown',
		};
	}

	function buildPlanningPlanWebviewHtml(currentPlanText: string, previousPlanText: string | undefined, planSteps: readonly IChatPlanningPlanEditorStep[], canSubmit: boolean): string {
		const nonce = generateUuid();
		const visiblePlanText = isUsablePlanningPlanText(currentPlanText) ? currentPlanText : '';
		const visiblePreviousPlanText = previousPlanText && isUsablePlanningPlanText(previousPlanText) ? previousPlanText : undefined;
		const visiblePlanSteps = visiblePlanText ? planSteps : [];
		const changeSummary = summarizePlanningPlanChanges(visiblePreviousPlanText, visiblePlanText);
		const controlsStateKey = getPlanningPlanControlsStateKey(visiblePlanSteps);
		const data = {
			currentPlanText: visiblePlanText,
			previousPlanText: visiblePreviousPlanText,
			stateKey: getPlanningPlanWebviewStateKey(visiblePlanText),
			controlsStateKey,
			changeSummary,
			steps: visiblePlanSteps.map(step => toPlanningPlanWebviewStep(step)),
			canSubmit,
			regenerateControlsAnswerKey: planningPlanRegenerateControlsAnswerKey,
			regenerateControlsAnswerValue: planningPlanRegenerateControlsAnswerValue,
			labels: {
				title: localize('openPlanningPlan.webviewTitle', 'Plan Canvas'),
				changes: localize('openPlanningPlan.webviewChanges', 'Changes in This Revision'),
				currentPlan: localize('openPlanningPlan.webviewCurrentPlan', 'Current Plan'),
				planCanvas: localize('openPlanningPlan.webviewPlanCanvas', 'Plan Canvas'),
				inlineQuestions: localize('openPlanningPlan.webviewInlineQuestions', 'Plan Controls'),
				stepControls: localize('openPlanningPlan.webviewStepControls', 'Step {0} Controls'),
				chooseStep: localize('openPlanningPlan.webviewChooseStep', 'Select a step to shape it.'),
				changeSomethingElse: localize('openPlanningPlan.webviewChangeSomethingElse', 'Change something else'),
				changeSomethingElsePlaceholder: localize('openPlanningPlan.webviewChangeSomethingElsePlaceholder', 'Describe another change for this step.'),
				changeThisPart: localize('openPlanningPlan.webviewOpenControls', 'Change'),
				closeControls: localize('openPlanningPlan.webviewCloseControls', 'Close'),
				question: localize('openPlanningPlan.webviewQuestion', 'Control'),
				options: localize('openPlanningPlan.webviewOptions', 'Options'),
				freeform: localize('openPlanningPlan.webviewFreeform', 'Additional answer or note'),
				noGeneratedControls: localize('openPlanningPlan.webviewNoGeneratedControls', 'Generating suggestions...'),
				continue: localize('openPlanningPlan.webviewContinue', 'Continue'),
				regenerateControls: localize('openPlanningPlan.webviewRegenerateControls', 'Refresh controls'),
				applyEdits: localize('openPlanningPlan.webviewApplyEdits', 'Apply edits'),
				submitted: localize('openPlanningPlan.webviewSubmitted', 'Submitted.'),
				applyingEdits: localize('openPlanningPlan.webviewApplyingEdits', 'Updating plan...'),
				regeneratingControls: localize('openPlanningPlan.webviewRegeneratingControls', 'Refreshing controls...'),
				readOnly: localize('openPlanningPlan.webviewReadOnly', 'Controls are inactive.'),
			}
		};

		return `<!DOCTYPE html>
			<html lang="${language}">
			<head>
				<meta charset="UTF-8">
				<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
				<title>${data.labels.title}</title>
				<style nonce="${nonce}">
					:root {
						color-scheme: light dark;
					}

					body {
						margin: 0;
						padding: 0;
						color: var(--vscode-editor-foreground);
						background: var(--vscode-editor-background);
						font-family: var(--vscode-font-family);
						font-size: var(--vscode-font-size);
						line-height: 1.45;
					}

					button, textarea, input {
						font-family: inherit;
						font-size: inherit;
					}

					.plan-shell {
						box-sizing: border-box;
						width: 100%;
						margin: 0;
						padding: 14px 16px 24px;
					}

					.plan-header {
						display: flex;
						justify-content: space-between;
						gap: 12px;
						align-items: flex-start;
						border-bottom: 1px solid var(--vscode-editorWidget-border);
						padding-bottom: 10px;
						margin-bottom: 14px;
					}

					h1, h2, h3 {
						margin: 0;
						font-weight: 600;
						line-height: 1.3;
						letter-spacing: 0;
					}

					h1 {
						font-size: 18px;
					}

					h2 {
						font-size: 14px;
						margin: 16px 0 8px;
					}

					h3 {
						font-size: 13px;
					}

					.actions {
						display: flex;
						gap: 8px;
						flex-wrap: wrap;
						justify-content: flex-end;
					}

					.button {
						min-height: 28px;
						border: 1px solid transparent;
						border-radius: 4px;
						padding: 4px 12px;
						color: var(--vscode-button-foreground);
						background: var(--vscode-button-background);
						cursor: pointer;
					}

					.button.secondary {
						color: var(--vscode-button-secondaryForeground);
						background: var(--vscode-button-secondaryBackground);
					}

					.button:hover:not(:disabled) {
						background: var(--vscode-button-hoverBackground);
					}

					.button.secondary:hover:not(:disabled) {
						background: var(--vscode-button-secondaryHoverBackground);
					}

					.button:disabled {
						opacity: 0.5;
						cursor: default;
					}

					.status {
						color: var(--vscode-descriptionForeground);
						margin-top: 8px;
						min-height: 20px;
					}

					.readonly {
						border: 1px solid var(--vscode-inputValidation-warningBorder);
						background: var(--vscode-inputValidation-warningBackground);
						color: var(--vscode-inputValidation-warningForeground);
						border-radius: 4px;
						padding: 8px 10px;
						margin-bottom: 16px;
					}

					.markdown {
						border-left: 3px solid var(--vscode-textBlockQuote-border);
						padding-left: 12px;
						color: var(--vscode-foreground);
					}

					.markdown p {
						margin: 6px 0;
						overflow-wrap: anywhere;
					}

					.markdown ul,
					.markdown ol {
						margin: 6px 0 6px 20px;
						padding: 0;
					}

					.markdown li {
						margin: 4px 0;
						overflow-wrap: anywhere;
					}

					.markdown pre {
						white-space: pre-wrap;
						overflow-wrap: anywhere;
						background: var(--vscode-textCodeBlock-background);
						border-radius: 4px;
						padding: 8px;
					}

					.diff {
						display: grid;
						gap: 4px;
						border: 1px solid var(--vscode-editorWidget-border);
						border-radius: 6px;
						padding: 8px;
						background: var(--vscode-sideBar-background);
					}

					.diff-line {
						white-space: pre-wrap;
						overflow-wrap: anywhere;
					}

					.diff-line.added {
						color: var(--vscode-gitDecoration-addedResourceForeground);
					}

					.diff-line.removed {
						color: var(--vscode-gitDecoration-deletedResourceForeground);
					}

					.control-stack {
						display: flex;
						flex-direction: column;
						gap: 8px;
					}

					.plan-workspace {
						display: grid;
						grid-template-columns: minmax(0, 1fr);
						gap: 12px;
						align-items: start;
					}

					.plan-workspace.controls-open {
						grid-template-columns: minmax(320px, 1fr) minmax(260px, 340px);
					}

					.plan-canvas {
						border: 1px solid var(--vscode-editorWidget-border);
						border-radius: 6px;
						background: var(--vscode-editor-background);
						padding: 12px;
						min-height: 60vh;
						min-width: 0;
					}

					.plan-canvas h2 {
						margin-top: 0;
					}

					.plan-canvas .markdown {
						border-left: 0;
						padding-left: 0;
					}

					.plan-markdown {
						box-sizing: border-box;
						min-height: 56vh;
						padding: 10px 12px;
						border: 1px solid var(--vscode-input-border);
						border-radius: 4px;
						background: var(--vscode-editor-background);
						color: var(--vscode-editor-foreground);
						overflow-wrap: anywhere;
					}

					.plan-markdown h2,
					.plan-markdown h3,
					.plan-markdown p,
					.plan-markdown li {
						border-left: 3px solid transparent;
						padding-left: 6px;
					}

					.plan-markdown h2 {
						margin: 10px 0 6px;
						font-size: 16px;
					}

					.plan-markdown h3 {
						margin: 8px 0 4px;
						font-size: 14px;
					}

					.plan-markdown p {
						margin: 5px 0;
					}

					.plan-step-line {
						cursor: pointer;
						border-radius: 4px;
						position: relative;
						padding-right: 72px;
					}

					.plan-step-line:hover,
					.plan-step-line:focus-within {
						background: var(--vscode-list-hoverBackground);
					}

					.plan-step-line.plan-step-kind-step:hover,
					.plan-step-line.plan-step-kind-step:focus-within {
						border-left-color: var(--vscode-focusBorder);
					}

					.plan-step-line.plan-step-kind-verification:hover,
					.plan-step-line.plan-step-kind-verification:focus-within {
						border-left-color: var(--vscode-testing-iconPassed);
					}

					.plan-step-line.plan-step-kind-decision:hover,
					.plan-step-line.plan-step-kind-decision:focus-within {
						border-left-color: var(--vscode-charts-yellow);
					}

					.plan-step-line.plan-step-kind-guardrail:hover,
					.plan-step-line.plan-step-kind-guardrail:focus-within {
						border-left-color: var(--vscode-editorWarning-foreground);
					}

					.plan-step-action {
						position: absolute;
						right: 4px;
						top: 50%;
						transform: translateY(-50%);
						min-height: 22px;
						border: 1px solid var(--vscode-button-border, transparent);
						border-radius: 4px;
						padding: 2px 7px;
						color: var(--vscode-button-secondaryForeground);
						background: var(--vscode-button-secondaryBackground);
						opacity: 0;
						pointer-events: none;
						cursor: pointer;
					}

					.plan-step-line:hover .plan-step-action,
					.plan-step-line:focus-within .plan-step-action {
						opacity: 1;
						pointer-events: auto;
					}

					.plan-change-highlight {
						background: var(--vscode-editorInfo-background);
						border-left-color: var(--vscode-editorInfo-foreground) !important;
					}

					.plan-editor-panel {
						border: 1px solid var(--vscode-editorWidget-border);
						border-radius: 6px;
						background: var(--vscode-sideBar-background);
						padding: 10px;
						position: sticky;
						top: 16px;
						max-height: calc(100vh - 32px);
						overflow: auto;
					}

					.plan-editor-panel h2 {
						margin-top: 0;
					}

					.plan-editor-header {
						display: flex;
						gap: 8px;
						align-items: center;
						justify-content: space-between;
						margin-bottom: 10px;
					}

					.plan-editor-header h2 {
						margin: 0;
					}

					.control-card {
						border: 1px solid var(--vscode-editorWidget-border);
						border-radius: 6px;
						background: var(--vscode-editor-background);
						padding: 8px;
					}

					.control-card h3 {
						margin: 0 0 8px;
						font-size: 13px;
						font-weight: 600;
					}

					.selected-step {
						display: flex;
						align-items: center;
						gap: 8px;
						margin-bottom: 10px;
					}

					.selected-step-index {
						display: inline-flex;
						align-items: center;
						justify-content: center;
						width: 22px;
						height: 22px;
						border-radius: 50%;
						background: var(--vscode-badge-background);
						color: var(--vscode-badge-foreground);
						font-size: 12px;
						font-weight: 600;
						flex: 0 0 auto;
					}

					.selected-step-title {
						min-width: 0;
						font-weight: 600;
						overflow: hidden;
						text-overflow: ellipsis;
						white-space: nowrap;
					}

					.selected-plan-part {
						margin: -4px 0 10px 30px;
						color: var(--vscode-descriptionForeground);
						font-size: 12px;
						overflow-wrap: anywhere;
					}

					.control-note {
						color: var(--vscode-descriptionForeground);
						margin: 0;
					}

					.question-stack {
						display: flex;
						flex-direction: column;
						gap: 8px;
					}

					.question-block {
						margin: 0;
						border-top: 1px solid var(--vscode-editorWidget-border);
						padding-top: 8px;
					}

					.question-title {
						font-weight: 600;
						margin-bottom: 4px;
					}

					.question-message,
					.question-description {
						color: var(--vscode-descriptionForeground);
						margin: 4px 0;
						overflow-wrap: anywhere;
					}

					.option-list {
						display: flex;
						flex-wrap: wrap;
						gap: 6px 10px;
						margin-top: 6px;
					}

					label.option {
						display: inline-flex;
						gap: 6px;
						align-items: center;
						max-width: 100%;
						border: 1px solid var(--vscode-editorWidget-border);
						border-radius: 4px;
						padding: 4px 7px;
						background: var(--vscode-sideBar-background);
						cursor: pointer;
					}

					label.option:hover {
						border-color: var(--vscode-focusBorder);
					}

					label.option:has(input:checked) {
						border-color: var(--vscode-focusBorder);
						background: var(--vscode-list-activeSelectionBackground);
						color: var(--vscode-list-activeSelectionForeground);
					}

					textarea {
						box-sizing: border-box;
						width: 100%;
						min-height: 34px;
						resize: vertical;
						border: 1px solid var(--vscode-input-border);
						border-radius: 4px;
						background: var(--vscode-input-background);
						color: var(--vscode-input-foreground);
						padding: 6px 8px;
						line-height: 1.4;
					}

					.text-answer,
					.freeform-answer {
						margin-top: 8px;
					}

					@media (max-width: 560px) {
						.plan-shell {
							padding: 14px;
						}

						.plan-header {
							display: block;
						}

						.actions {
							justify-content: flex-start;
							margin-top: 12px;
						}
						.question-block {
							margin-left: 0;
						}
					}

					@media (max-width: 640px) {
						.plan-workspace,
						.plan-workspace.controls-open {
							grid-template-columns: minmax(0, 1fr);
						}

						.plan-editor-panel {
							position: static;
						}
					}
				</style>
			</head>
			<body>
				<div id="app"></div>
				<script nonce="${nonce}">
					const vscode = acquireVsCodeApi();
					const data = ${toWebviewScriptLiteral(data)};
					const previousState = vscode.getState?.() || {};
					const savedState = previousState.planStateKey === data.stateKey ? previousState : { planStateKey: data.stateKey };
					if (!data.currentPlanText && typeof savedState.currentPlanText === 'string') {
						data.currentPlanText = savedState.currentPlanText;
					}
					if (!data.previousPlanText && typeof savedState.previousPlanText === 'string') {
						data.previousPlanText = savedState.previousPlanText;
					}
					savedState.currentPlanText = data.currentPlanText;
					savedState.previousPlanText = data.previousPlanText;
					savedState.controlsStateKey = data.controlsStateKey;
					const root = document.getElementById('app');
					const controlMap = new Map();
					let submitted = savedState.submitted === true;
					let selectedStepId = savedState.selectedStepId;
					let controlsOpen = savedState.controlsOpen === true && data.canSubmit;
					document.addEventListener('input', markDirty, true);
					document.addEventListener('change', markDirty, true);

					function markDirty(event) {
						savedState.dirty = true;
						const key = event.target?.dataset?.answerKey;
						if (key) {
							savedState.dirtyKeys = savedState.dirtyKeys || {};
							savedState.dirtyKeys[key] = true;
						}
					}

					function isDirtyKey(key) {
						return savedState.dirtyKeys?.[key] === true;
					}

					function el(tag, className, text) {
						const node = document.createElement(tag);
						if (className) {
							node.className = className;
						}
						if (text !== undefined) {
							node.textContent = text;
						}
						return node;
					}

					function render() {
						root.textContent = '';
						controlMap.clear();
						const shell = el('main', 'plan-shell');
						root.appendChild(shell);

						const header = el('header', 'plan-header');
						const title = el('h1', undefined, data.labels.title);
						header.appendChild(title);
						shell.appendChild(header);

						const workspace = el('section', controlsOpen ? 'plan-workspace controls-open' : 'plan-workspace');

						const canvas = el('section', 'plan-canvas');
						canvas.appendChild(el('h2', undefined, data.labels.planCanvas));
						const diff = renderChangeSummary();
						if (diff) {
							canvas.appendChild(diff);
						}
						const planCanvas = el('div', 'plan-markdown');
						planCanvas.id = 'plan-markdown';
						renderPlanCanvas(planCanvas, data.currentPlanText, getChangedLineSet(data.currentPlanText, data.previousPlanText));
						canvas.appendChild(planCanvas);
						workspace.appendChild(canvas);

						if (controlsOpen) {
							const editorPanel = el('aside', 'plan-editor-panel');
							const editorHeader = el('div', 'plan-editor-header');
							editorHeader.appendChild(el('h2', undefined, data.labels.inlineQuestions));
							const closeButton = el('button', 'button secondary', data.labels.closeControls);
							closeButton.type = 'button';
							closeButton.addEventListener('click', closeControls);
							editorHeader.appendChild(closeButton);
							editorPanel.appendChild(editorHeader);
							const controlsHost = el('div');
							controlsHost.id = 'plan-editor-controls';
							renderSelectedStepControls(controlsHost);
							editorPanel.appendChild(controlsHost);

							const actions = el('div', 'actions');
							const regenerateControlsButton = el('button', 'button secondary', data.labels.regenerateControls);
							const continueButton = el('button', 'button secondary', data.labels.continue);
							const applyButton = el('button', 'button apply-button', data.labels.applyEdits);
							regenerateControlsButton.type = 'button';
							continueButton.type = 'button';
							applyButton.type = 'button';
							regenerateControlsButton.disabled = submitted;
							continueButton.disabled = submitted;
							applyButton.disabled = submitted || !hasEdits();
							regenerateControlsButton.addEventListener('click', () => submit('regenerateControls'));
							continueButton.addEventListener('click', () => submit('continue'));
							applyButton.addEventListener('click', () => submit('apply'));
							actions.appendChild(regenerateControlsButton);
							actions.appendChild(continueButton);
							actions.appendChild(applyButton);
							editorPanel.appendChild(actions);
							workspace.appendChild(editorPanel);
						}
						shell.appendChild(workspace);

						const status = el('div', 'status', submitted ? data.labels.submitted : '');
						status.id = 'plan-status';
						shell.appendChild(status);
						updateApplyButton();
						updateDisabledState();
					}

					function renderChangeSummary() {
						if (!data.changeSummary || !(data.changeSummary.added?.length || data.changeSummary.removed?.length)) {
							return undefined;
						}

						const section = el('section');
						section.appendChild(el('h3', undefined, data.labels.changes));
						const diff = el('div', 'diff');
						for (const line of data.changeSummary.added || []) {
							diff.appendChild(el('div', 'diff-line added', '+ ' + line));
						}
						for (const line of data.changeSummary.removed || []) {
							diff.appendChild(el('div', 'diff-line removed', '- ' + line));
						}
						section.appendChild(diff);
						return section;
					}

					function normalizeLine(text) {
						return String(text || '').trim().replace(/\\s+/g, ' ');
					}

					function getChangedLineSet(text, previousText) {
						if (!previousText) {
							return new Set();
						}
						const previousLines = new Set(String(previousText || '').split(/\\r?\\n/g).map(normalizeLine).filter(Boolean));
						return new Set(String(text || '').split(/\\r?\\n/g).map(normalizeLine).filter(line => line && !previousLines.has(line)));
					}

					function markIfChanged(node, raw, changedLines) {
						if (changedLines?.has(normalizeLine(raw))) {
							node.classList.add('plan-change-highlight');
						}
						return node;
					}

					function withMarkdownPrefix(node, prefix) {
						node.dataset.markdownPrefix = prefix;
						return node;
					}

					function appendEditableBlock(container, node, raw, changedLines) {
						container.appendChild(markIfChanged(node, raw, changedLines));
					}

					function selectStep(stepId, selectedPart) {
						const shouldRender = !controlsOpen;
						controlsOpen = true;
						savedState.controlsOpen = true;
						selectedStepId = stepId;
						savedState.selectedStepId = stepId;
						if (selectedPart) {
							savedState.selectedPartByStep = savedState.selectedPartByStep || {};
							savedState.selectedPartByStep[stepId] = selectedPart;
						}
						persistState();
						if (shouldRender) {
							render();
						} else {
							updateSelectedStepRendering();
						}
					}

					function closeControls() {
						controlsOpen = false;
						savedState.controlsOpen = false;
						persistState();
						render();
					}

					function clearRenderedStepControls() {
						for (const key of Array.from(controlMap.keys())) {
							if (String(key).startsWith('plan-editor-step-custom-') || String(key).startsWith('plan-editor-question-')) {
								controlMap.delete(key);
							}
						}
					}

					function updateSelectedStepRendering() {
						for (const node of document.querySelectorAll('.plan-step-line.active')) {
							node.classList.remove('active');
						}
						for (const node of document.querySelectorAll('.plan-step-line')) {
							if (node.dataset.stepId === selectedStepId) {
								node.classList.add('active');
							}
						}
						const controlsHost = document.getElementById('plan-editor-controls');
						if (controlsHost) {
							clearRenderedStepControls();
							controlsHost.textContent = '';
							renderSelectedStepControls(controlsHost);
						}
						updateApplyButton();
						updateDisabledState();
					}

					function normalizeComparableText(text) {
						return normalizeLine(String(text || '')
							.replace(/^\\s{0,3}(?:(?:[-*+])\\s+|\\d+[.)]\\s+)/, '')
							.replace(/^#{1,6}\\s+/, '')
							.replace(/\\*\\*/g, ''));
					}

					function tokenOverlapScore(left, right) {
						const leftTokens = new Set(normalizeComparableText(left).toLowerCase().match(/[a-z0-9]{4,}/g) || []);
						const rightTokens = new Set(normalizeComparableText(right).toLowerCase().match(/[a-z0-9]{4,}/g) || []);
						let score = 0;
						for (const token of leftTokens) {
							if (rightTokens.has(token)) {
								score++;
							}
						}
						return score;
					}

					function getPlanCanvasSectionKind(title) {
						if (/\\b(step|steps|implementation|approach|tasks?|work plan|execution)\\b/i.test(title || '')) {
							return 'step';
						}
						if (/\\b(verification|validate|validation|tests?|checks?|qa)\\b/i.test(title || '')) {
							return 'verification';
						}
						if (/\\b(decisions?|questions?|open items?|clarifications?)\\b/i.test(title || '')) {
							return 'decision';
						}
						if (/\\b(risks?|guardrails?|constraints?|assumptions?)\\b/i.test(title || '')) {
							return 'guardrail';
						}
						if (/\\b(relevant files?|files?|context|references?|dependencies?)\\b/i.test(title || '')) {
							return 'other';
						}
						return 'step';
					}

					function isActionablePlanSection(kind) {
						return !!kind && kind !== 'other';
					}

					function getDisplayHeadingTitle(title) {
						return String(title || '').replace(/^plan\\s*:\\s*/i, '').trim() || title;
					}

					function getStepForLine(raw, currentSectionTitle, currentSectionKind) {
						if (!isActionablePlanSection(currentSectionKind)) {
							return undefined;
						}

						const normalizedLine = normalizeLine(String(raw || '').replace(/^\\s{0,3}(?:(?:[-*+])\\s+|\\d+[.)]\\s+)/, '').replace(/\\*\\*/g, ''));
						if (!normalizedLine) {
							return undefined;
						}

						let bestStep;
						let bestScore = 0;
						const normalizedCurrentSection = normalizeComparableText(currentSectionTitle);
						for (const step of data.steps) {
							if (step.kind === 'other') {
								continue;
							}
							const label = normalizeComparableText(step.label);
							const text = normalizeComparableText(step.text);
							const section = normalizeComparableText(step.sectionTitle);
							let score = tokenOverlapScore(normalizedLine, [section, label, text].filter(Boolean).join(' '));
							if (label && (normalizedLine === label || normalizedLine.includes(label) || label.includes(normalizedLine))) {
								score += 20;
							}
							if (text && (text.includes(normalizedLine) || normalizedLine.includes(text))) {
								score += 12;
							}
							if (section && normalizedCurrentSection && (section === normalizedCurrentSection || section.includes(normalizedCurrentSection) || normalizedCurrentSection.includes(section))) {
								score += 4;
							}
							if (step.kind === currentSectionKind) {
								score += 6;
							} else {
								score -= 4;
							}
							if (score > bestScore) {
								bestScore = score;
								bestStep = step;
							}
						}
						return bestScore >= 8 ? bestStep : undefined;
					}

					function getActionablePlanSteps() {
						return data.steps.filter(step => step.kind !== 'other');
					}

					function markStepNode(node, raw, step) {
						if (!data.canSubmit || !step) {
							return node;
						}

						const selectedPart = (normalizeComparableText(raw) || step.label).slice(0, 240);
						node.classList.add('plan-step-line');
						node.classList.add('plan-step-kind-' + (step.kind || 'step'));
						if (selectedStepId === step.id) {
							node.classList.add('active');
						}
						node.tabIndex = 0;
						node.setAttribute('role', 'button');
						node.dataset.stepId = step.id;
						node.dataset.selectedPart = selectedPart;
						node.addEventListener('click', () => selectStep(step.id, selectedPart));
						node.addEventListener('keydown', event => {
							if (event.key === 'Enter' || event.key === ' ') {
								event.preventDefault();
								selectStep(step.id, selectedPart);
							}
						});
						const action = el('button', 'plan-step-action', data.labels.changeThisPart);
						action.type = 'button';
						action.addEventListener('click', event => {
							event.stopPropagation();
							selectStep(step.id, selectedPart);
						});
						node.appendChild(action);
						return node;
					}

					function appendInlineMarkdown(container, text) {
						const value = String(text || '');
						const tokenPattern = /(\\x60[^\\x60]+\\x60|\\*\\*[^*]+\\*\\*)/g;
						let lastIndex = 0;
						let match = tokenPattern.exec(value);
						while (match) {
							if (match.index > lastIndex) {
								container.appendChild(document.createTextNode(value.slice(lastIndex, match.index)));
							}
							const token = match[0];
							if (token.startsWith('\\x60')) {
								container.appendChild(el('code', undefined, token.slice(1, -1)));
							} else {
								container.appendChild(el('strong', undefined, token.slice(2, -2)));
							}
							lastIndex = match.index + token.length;
							match = tokenPattern.exec(value);
						}
						if (lastIndex < value.length) {
							container.appendChild(document.createTextNode(value.slice(lastIndex)));
						}
						return container;
					}

					function inlineMarkdownElement(tag, className, text) {
						return appendInlineMarkdown(el(tag, className), text);
					}

					function renderPlanCanvas(container, text, changedLines = new Set()) {
						container.textContent = '';
						const lines = String(text || '').split(/\\r?\\n/g);
						const actionableSteps = getActionablePlanSteps();
						let nextActionableStepIndex = 0;
						let list;
						let pre;
						let orderedIndex = 1;
						let currentSectionTitle = '';
						let currentSectionKind = '';
						function getStepForActionableRow(rowText) {
							const matchedStep = getStepForLine(rowText, currentSectionTitle, currentSectionKind);
							if (matchedStep) {
								const matchedIndex = actionableSteps.findIndex(step => step.id === matchedStep.id);
								if (matchedIndex >= nextActionableStepIndex) {
									nextActionableStepIndex = matchedIndex + 1;
								}
								return matchedStep;
							}

							return actionableSteps[nextActionableStepIndex++];
						}
						for (const raw of lines) {
							const line = raw.trimEnd();
							if (line.startsWith('\\x60\\x60\\x60')) {
								if (pre) {
									appendEditableBlock(container, pre, pre.textContent || currentSectionTitle, changedLines);
									pre = undefined;
								} else {
									pre = withMarkdownPrefix(el('pre'), '\\x60\\x60\\x60');
									pre.textContent = '';
								}
								continue;
							}
							if (pre) {
								pre.textContent += (pre.textContent ? '\\n' : '') + raw;
								continue;
							}
							if (!line.trim()) {
								if (!isActionablePlanSection(currentSectionKind)) {
									list = undefined;
									orderedIndex = 1;
								}
								continue;
							}
							const heading = /^(#{1,6})\\s+(.+)$/.exec(line);
							if (heading) {
								list = undefined;
								orderedIndex = 1;
								currentSectionTitle = heading[2].replace(/#+\\s*$/, '').trim();
								currentSectionKind = /^plan\\b/i.test(currentSectionTitle) ? '' : getPlanCanvasSectionKind(currentSectionTitle);
								appendEditableBlock(container, withMarkdownPrefix(inlineMarkdownElement(heading[1].length <= 2 ? 'h2' : 'h3', undefined, getDisplayHeadingTitle(currentSectionTitle)), heading[1] + ' '), raw, changedLines);
								continue;
							}
							const boldSection = /^\\*\\*(.+?)\\*\\*\\s*:?$/.exec(line.trim());
							if (boldSection) {
								list = undefined;
								orderedIndex = 1;
								currentSectionTitle = boldSection[1].trim();
								currentSectionKind = /^plan\\b/i.test(currentSectionTitle) ? '' : getPlanCanvasSectionKind(currentSectionTitle);
								appendEditableBlock(container, inlineMarkdownElement('h2', undefined, getDisplayHeadingTitle(currentSectionTitle)), raw, changedLines);
								continue;
							}
							const planTitle = /^Plan\\s*:\\s*(.+)$/i.exec(line.trim());
							if (planTitle) {
								list = undefined;
								orderedIndex = 1;
								currentSectionTitle = '';
								currentSectionKind = '';
								appendEditableBlock(container, inlineMarkdownElement('h2', undefined, planTitle[1].trim()), raw, changedLines);
								continue;
							}
							const ordered = /^\\d+[.)]\\s+(.+)$/.exec(line.trim());
							const unordered = /^[-*+]\\s+(.+)$/.exec(line.trim());
							if (ordered || unordered) {
								const listType = isActionablePlanSection(currentSectionKind) || ordered ? 'ol' : 'ul';
								if (!list || list.tagName.toLowerCase() !== listType) {
									list = el(listType);
									container.appendChild(list);
									orderedIndex = 1;
								}
								const itemPrefix = ordered ? String(orderedIndex++) + '. ' : '- ';
								const itemText = (ordered || unordered)[1];
								const currentStep = isActionablePlanSection(currentSectionKind) ? getStepForActionableRow(itemText) : undefined;
								list.appendChild(markIfChanged(markStepNode(withMarkdownPrefix(inlineMarkdownElement('li', undefined, itemText), itemPrefix), itemText, currentStep), raw, changedLines));
								continue;
							}
							if (isActionablePlanSection(currentSectionKind)) {
								if (!list || list.tagName.toLowerCase() !== 'ol') {
									list = el('ol');
									container.appendChild(list);
									orderedIndex = 1;
								}
								const currentStep = getStepForActionableRow(line);
								list.appendChild(markIfChanged(markStepNode(withMarkdownPrefix(inlineMarkdownElement('li', undefined, line), String(orderedIndex++) + '. '), line, currentStep), raw, changedLines));
								continue;
							}
							list = undefined;
							orderedIndex = 1;
							appendEditableBlock(container, withMarkdownPrefix(inlineMarkdownElement('p', undefined, line), ''), raw, changedLines);
						}
						if (pre) {
							appendEditableBlock(container, pre, pre.textContent || currentSectionTitle, changedLines);
						}
					}

					function renderSelectedStepControls(container) {
						const selectedStep = data.steps.find(step => step.id === selectedStepId);
						if (!selectedStep) {
							container.appendChild(el('p', 'control-note', data.labels.chooseStep));
							return;
						}

						const header = el('div', 'selected-step');
						header.appendChild(el('span', 'selected-step-index', String(selectedStep.index)));
						header.appendChild(el('div', 'selected-step-title', selectedStep.label));
						container.appendChild(header);
						const selectedPart = savedState.selectedPartByStep?.[selectedStep.id];
						if (selectedPart && selectedPart !== selectedStep.label) {
							container.appendChild(el('div', 'selected-plan-part', selectedPart));
						}

						const card = el('section', 'control-card');
						const stack = el('div', 'question-stack');
						if (!(selectedStep.questions || []).length) {
							stack.appendChild(el('p', 'control-note', data.labels.noGeneratedControls));
						}
						for (const question of selectedStep.questions || []) {
							stack.appendChild(renderQuestion(selectedStep, question));
						}
						stack.appendChild(renderStepOpenResponse(selectedStep));
						card.appendChild(stack);
						container.appendChild(card);
					}

					function renderStepOpenResponse(step) {
						const key = 'plan-editor-step-custom-' + step.id;
						const block = el('div', 'question-block');
						const label = el('label', 'question-title', data.labels.changeSomethingElse);
						label.htmlFor = key;
						block.appendChild(label);
						const textarea = el('textarea', 'freeform-answer');
						textarea.id = key;
						textarea.dataset.answerKey = key;
						textarea.rows = 3;
						textarea.placeholder = data.labels.changeSomethingElsePlaceholder;
						textarea.value = savedState[key] || '';
						textarea.addEventListener('input', () => {
							savedState[key] = textarea.value;
							persistState();
							updateApplyButton();
						});
						controlMap.set(key, { text: textarea });
						block.appendChild(textarea);
						return block;
					}

					function renderQuestion(step, question) {
						const block = el('div', 'question-block');
						block.appendChild(el('div', 'question-title', question.title));
						if (question.message && question.message !== question.title) {
							block.appendChild(el('div', 'question-message', question.message));
						}
						if (question.description) {
							block.appendChild(el('div', 'question-description', question.description));
						}
						const key = 'plan-editor-question-' + step.id + '-' + question.id;
						if (question.type === 'text') {
							const textarea = el('textarea', 'text-answer');
							textarea.dataset.answerKey = key;
							textarea.rows = 2;
							textarea.value = savedState[key] || '';
							textarea.addEventListener('input', () => {
								savedState[key] = textarea.value;
								persistState();
								updateApplyButton();
							});
							controlMap.set(key, { question, text: textarea });
							block.appendChild(textarea);
							return block;
						}
						const optionList = el('div', 'option-list');
						const defaults = new Set(Array.isArray(question.defaultValue) ? question.defaultValue : question.defaultValue ? [question.defaultValue] : []);
						const saved = savedState[key];
						const selectedValues = new Set(Array.isArray(saved?.selectedValues) ? saved.selectedValues : saved?.selectedValue ? [saved.selectedValue] : []);
						for (const option of question.options || []) {
							const label = el('label', 'option');
							const input = document.createElement('input');
							input.type = question.type === 'singleSelect' ? 'radio' : 'checkbox';
							input.name = key;
							input.dataset.answerKey = key;
							input.value = option.value;
							input.checked = selectedValues.has(option.value) || selectedValues.has(option.label) || (!saved && (defaults.has(option.label) || defaults.has(option.value)));
							input.addEventListener('change', () => {
								saveQuestionChoiceState(key, question);
								updateApplyButton();
							});
							label.appendChild(input);
							label.appendChild(document.createTextNode(option.label));
							optionList.appendChild(label);
						}
						block.appendChild(optionList);
						const freeform = el('textarea', 'freeform-answer');
						freeform.dataset.answerKey = key;
						freeform.rows = 2;
						freeform.value = saved?.freeformValue || '';
						freeform.disabled = question.allowFreeformInput === false;
						freeform.placeholder = data.labels.changeSomethingElsePlaceholder;
						freeform.addEventListener('input', () => {
							saveQuestionChoiceState(key, question);
							updateApplyButton();
						});
						if (question.allowFreeformInput !== false) {
							block.appendChild(freeform);
						}
						controlMap.set(key, { question, optionList, freeform });
						saveQuestionChoiceState(key, question, false);
						return block;
					}

					function saveQuestionChoiceState(key, question, persist = true) {
						const control = controlMap.get(key);
						if (!control) {
							return;
						}
						if (question.type === 'text') {
							savedState[key] = control.text.value;
						} else {
							const selected = Array.from(control.optionList.querySelectorAll('input:checked')).map(input => input.value);
							if (question.type === 'singleSelect') {
								savedState[key] = {
									selectedValue: selected[0],
									...(control.freeform?.value.trim() ? { freeformValue: control.freeform.value.trim() } : {}),
								};
							} else {
								savedState[key] = {
									selectedValues: selected,
									...(control.freeform?.value.trim() ? { freeformValue: control.freeform.value.trim() } : {}),
								};
							}
						}
						if (persist) {
							persistState();
						}
					}

					function collectAnswers() {
						const answers = {};
						for (const step of data.steps) {
							let stepHasAnswer = false;
							const stepCustomKey = 'plan-editor-step-custom-' + step.id;
							if (isDirtyKey(stepCustomKey)) {
								const value = controlMap.get(stepCustomKey)?.text?.value.trim() ?? (typeof savedState[stepCustomKey] === 'string' ? savedState[stepCustomKey].trim() : '');
								if (value) {
									answers[stepCustomKey] = value;
									stepHasAnswer = true;
								}
							}
							for (const question of step.questions || []) {
								const questionKey = 'plan-editor-question-' + step.id + '-' + question.id;
								if (!isDirtyKey(questionKey)) {
									continue;
								}
								const control = controlMap.get(questionKey);
								if (question.type === 'text') {
									const value = control?.text?.value.trim() ?? (typeof savedState[questionKey] === 'string' ? savedState[questionKey].trim() : '');
									if (value) {
										answers[questionKey] = value;
										stepHasAnswer = true;
									}
								} else {
									saveQuestionChoiceState(questionKey, question, false);
									const answer = savedState[questionKey];
									if (question.type === 'singleSelect') {
										if (answer?.selectedValue || answer?.freeformValue) {
											answers[questionKey] = answer;
											stepHasAnswer = true;
										}
									} else if (answer?.selectedValues?.length || answer?.freeformValue) {
										answers[questionKey] = answer;
										stepHasAnswer = true;
									}
								}
							}
							const selectedPart = savedState.selectedPartByStep?.[step.id];
							if (stepHasAnswer && selectedPart) {
								answers['plan-editor-selected-part-' + step.id] = selectedPart;
							}
						}
						return answers;
					}

					function hasEdits() {
						if (savedState.dirty !== true) {
							return false;
						}
						for (const step of data.steps) {
							const stepCustomKey = 'plan-editor-step-custom-' + step.id;
							if (isDirtyKey(stepCustomKey)) {
								const value = controlMap.get(stepCustomKey)?.text?.value.trim() ?? (typeof savedState[stepCustomKey] === 'string' ? savedState[stepCustomKey].trim() : '');
								if (value) {
									return true;
								}
							}
							for (const question of step.questions || []) {
								const questionKey = 'plan-editor-question-' + step.id + '-' + question.id;
								if (!isDirtyKey(questionKey)) {
									continue;
								}
								const control = controlMap.get(questionKey);
								const saved = savedState[questionKey];
								if (question.type === 'text') {
									const value = control?.text?.value.trim() ?? (typeof saved === 'string' ? saved.trim() : '');
									if (value) {
										return true;
									}
									continue;
								}
								const selected = control
									? Array.from(control.optionList.querySelectorAll('input:checked')).map(input => input.value)
									: Array.isArray(saved?.selectedValues)
										? saved.selectedValues
										: saved?.selectedValue ? [saved.selectedValue] : [];
								const freeform = control?.freeform?.value.trim() ?? (typeof saved?.freeformValue === 'string' ? saved.freeformValue.trim() : '');
								if (selected.length || freeform) {
									return true;
								}
							}
						}
						return false;
					}

					function updateApplyButton() {
						const button = document.querySelector('.apply-button');
						if (button) {
							button.disabled = !data.canSubmit || submitted || !hasEdits();
						}
					}

					function updateDisabledState() {
						if (!submitted) {
							return;
						}
						for (const control of document.querySelectorAll('button, textarea, input')) {
							control.disabled = true;
						}
					}

					function submit(type) {
						if (!data.canSubmit || submitted) {
							return;
						}
						const answers = collectAnswers();
						if (type === 'regenerateControls') {
							answers[data.regenerateControlsAnswerKey] = data.regenerateControlsAnswerValue;
						}
						submitted = true;
						savedState.submitted = true;
						persistState();
						updateDisabledState();
						const status = document.getElementById('plan-status');
						if (status) {
							status.textContent = type === 'regenerateControls'
								? data.labels.regeneratingControls
								: type === 'apply'
									? data.labels.applyingEdits
									: data.labels.submitted;
						}
						vscode.postMessage({ type, answers });
					}

					function persistState() {
						vscode.setState?.(savedState);
					}

					window.addEventListener('message', event => {
						if (event.data?.type === 'planUpdate' && typeof event.data.planText === 'string') {
							if (!event.data.planText.trim()) {
								return;
							}
							const previousPlanText = typeof event.data.previousPlanText === 'string' ? event.data.previousPlanText : data.currentPlanText;
							data.previousPlanText = previousPlanText;
							data.currentPlanText = event.data.planText;
							if (Array.isArray(event.data.planSteps)) {
								data.steps = event.data.planSteps;
								if (selectedStepId && !data.steps.some(step => step.id === selectedStepId)) {
									selectedStepId = undefined;
									savedState.selectedStepId = undefined;
									savedState.selectedPartByStep = {};
								}
							}
							if (typeof event.data.controlsStateKey === 'string') {
								data.controlsStateKey = event.data.controlsStateKey;
								savedState.controlsStateKey = data.controlsStateKey;
							}
							savedState.previousPlanText = data.previousPlanText;
							savedState.currentPlanText = data.currentPlanText;
							persistState();
							const planMarkdown = document.getElementById('plan-markdown');
							if (planMarkdown) {
								renderPlanCanvas(planMarkdown, data.currentPlanText, getChangedLineSet(data.currentPlanText, data.previousPlanText));
							}
							if (controlsOpen) {
								updateSelectedStepRendering();
							}
							const status = document.getElementById('plan-status');
							if (status) {
								status.textContent = !data.canSubmit && event.data.isComplete
									? ''
									: event.data.isComplete ? data.labels.submitted : data.labels.applyingEdits;
							}
							return;
						}
						if (event.data?.type === 'submitted') {
							submitted = true;
							savedState.submitted = true;
							persistState();
							updateDisabledState();
							const status = document.getElementById('plan-status');
							if (status) {
								status.textContent = data.labels.submitted;
							}
						}
					});

					render();
				</script>
			</body>
			</html>`;
	}

	async function openPlanningPlanWebview(accessor: ServicesAccessor, args: IPlanningPlanCommandArgs, group: typeof ACTIVE_GROUP | typeof SIDE_GROUP): Promise<void> {
		const chatService = accessor.get(IChatService);
		const webviewWorkbenchService = accessor.get(IWebviewWorkbenchService);
		const notificationService = accessor.get(INotificationService);
		const sessionResource = revivePlanningSessionResource(args.sessionResource);
		const webviewKey = sessionResource.toString();
		const planText = args.planText ?? getPlanningPlanText(chatService, sessionResource, args.requestId);
		const previousPlanText = args.previousPlanText ?? (args.previousRequestId ? getPlanningPlanText(chatService, sessionResource, args.previousRequestId) : undefined);
		if (!planText) {
			notificationService.warn(localize('openPlanningPlan.missing', 'The selected plan is no longer available in this chat session.'));
			return;
		}

		const planSteps = getPlanningPlanEditorSteps(planText, args.planSteps);
		const title = localize('openPlanningPlan.editorTitle', 'Plan Canvas');
		const existing = planningPlanWebviews.get(webviewKey);
		let webviewInput = existing?.input;
		if (webviewInput && !webviewInput.isDisposed()) {
			existing?.listener.dispose();
			webviewWorkbenchService.revealWebview(webviewInput, group, false);
		} else {
			webviewInput = webviewWorkbenchService.openWebview({
				providedViewType: PLANNING_PLAN_WEBVIEW_VIEW_TYPE,
				title,
				options: {
					purpose: WebviewContentPurpose.CustomEditor,
					enableFindWidget: true,
					retainContextWhenHidden: true,
				},
				contentOptions: {
					allowScripts: true,
					allowForms: true,
				},
				extension: undefined,
			}, PLANNING_PLAN_WEBVIEW_VIEW_TYPE, title, undefined, { group });
		}

		const listener = webviewInput.webview.onMessage(event => {
			if (!isPlanningPlanWebviewSubmitMessage(event.message)) {
				return;
			}
			if (!args.planEditorResolveId) {
				notificationService.warn(localize('openPlanningPlan.notInteractive', 'This plan is no longer connected to an active chat review.'));
				return;
			}

			chatService.notifyQuestionCarouselAnswer(args.requestId, args.planEditorResolveId, event.message.answers);
			if (event.message.type === 'continue') {
				void webviewInput.webview.postMessage({ type: 'submitted' });
			}
		});
		planningPlanWebviews.set(webviewKey, { input: webviewInput, listener });
		webviewInput.webview.onDidDispose(() => {
			listener.dispose();
			if (planningPlanWebviews.get(webviewKey)?.input === webviewInput) {
				planningPlanWebviews.delete(webviewKey);
			}
		});
		webviewInput.webview.setHtml(buildPlanningPlanWebviewHtml(planText, previousPlanText, planSteps, !!args.planEditorResolveId));
	}

	async function updatePlanningPlanWebview(accessor: ServicesAccessor, args: IPlanningPlanUpdateCommandArgs | undefined): Promise<void> {
		if (!args?.sessionResource || typeof args.planText !== 'string' || !isUsablePlanningPlanText(args.planText)) {
			return;
		}

		const sessionResource = revivePlanningSessionResource(args.sessionResource);
		const entry = planningPlanWebviews.get(sessionResource.toString());
		if (!entry || entry.input.isDisposed()) {
			if (entry?.input.isDisposed()) {
				planningPlanWebviews.delete(sessionResource.toString());
			}
			if (args.requestId) {
				await openPlanningPlanWebview(accessor, {
					sessionResource,
					requestId: args.requestId,
					planText: args.planText,
					previousPlanText: args.previousPlanText,
					planSteps: args.planSteps,
				}, ACTIVE_GROUP);
			}
			return;
		}

		const planSteps = getPlanningPlanEditorSteps(args.planText, args.planSteps);
		await entry.input.webview.postMessage({
			type: 'planUpdate',
			planText: args.planText,
			previousPlanText: args.previousPlanText && isUsablePlanningPlanText(args.previousPlanText) ? args.previousPlanText : undefined,
			planSteps: planSteps.map(step => toPlanningPlanWebviewStep(step)),
			controlsStateKey: getPlanningPlanControlsStateKey(planSteps),
			isComplete: args.isComplete === true,
		});
	}

	registerAction2(class OpenPlanningPlanAction extends Action2 {
		constructor() {
			super({
				id: OPEN_PLANNING_PLAN_ACTION_ID,
				title: localize2('openPlanningPlan', 'Open Current Plan'),
				category: CHAT_CATEGORY,
				f1: false,
			});
		}

		override async run(accessor: ServicesAccessor, args?: IPlanningPlanCommandArgs): Promise<void> {
			if (!args?.requestId || !args.sessionResource) {
				return;
			}
			await openPlanningPlanWebview(accessor, args, ACTIVE_GROUP);
		}
	});

	registerAction2(class UpdatePlanningPlanAction extends Action2 {
		constructor() {
			super({
				id: UPDATE_PLANNING_PLAN_ACTION_ID,
				title: localize2('updatePlanningPlan', 'Update Current Plan Canvas'),
				category: CHAT_CATEGORY,
				f1: false,
			});
		}

		override async run(accessor: ServicesAccessor, args?: IPlanningPlanUpdateCommandArgs): Promise<void> {
			await updatePlanningPlanWebview(accessor, args);
		}
	});

	registerAction2(class OpenPlanningPlanToSideAction extends Action2 {
		constructor() {
			super({
				id: OPEN_PLANNING_PLAN_TO_SIDE_ACTION_ID,
				title: localize2('openPlanningPlanToSide', 'Open Plan Beside'),
				category: CHAT_CATEGORY,
				f1: false,
			});
		}

		override async run(accessor: ServicesAccessor, args?: IPlanningPlanCommandArgs): Promise<void> {
			if (!args?.requestId || !args.sessionResource) {
				return;
			}
			await openPlanningPlanWebview(accessor, args, SIDE_GROUP);
		}
	});

	registerAction2(class OpenPlanningPlanDiffAction extends Action2 {
		constructor() {
			super({
				id: OPEN_PLANNING_PLAN_DIFF_ACTION_ID,
				title: localize2('openPlanningPlanDiff', 'Compare Plan Revisions'),
				category: CHAT_CATEGORY,
				f1: false,
			});
		}

		override async run(accessor: ServicesAccessor, args?: IPlanningPlanCommandArgs): Promise<void> {
			if (!args?.requestId || !args?.previousRequestId || !args.sessionResource) {
				return;
			}

			const chatService = accessor.get(IChatService);
			const editorService = accessor.get(IEditorService);
			const notificationService = accessor.get(INotificationService);
			const sessionResource = revivePlanningSessionResource(args.sessionResource);
			const previousPlanText = args.previousPlanText ?? getPlanningPlanText(chatService, sessionResource, args.previousRequestId);
			const currentPlanText = args.planText ?? getPlanningPlanText(chatService, sessionResource, args.requestId);
			if (!previousPlanText || !currentPlanText) {
				notificationService.warn(localize('openPlanningPlanDiff.missing', 'The previous or current plan is no longer available in this chat session.'));
				return;
			}

			const diffInput: IResourceDiffEditorInput = {
				original: createPlanningPlanEditorInput(localize('openPlanningPlanDiff.previousTitle', 'Previous Plan'), args.previousRequestId, previousPlanText),
				modified: createPlanningPlanEditorInput(localize('openPlanningPlanDiff.currentTitle', 'Current Plan'), args.requestId, currentPlanText),
				label: localize('openPlanningPlanDiff.diffLabel', 'Plan Changes'),
				options: { pinned: true },
			};
			await editorService.openEditor(diffInput);
		}
	});

	// When customizations menu is enabled, show a direct gear action to open the Customizations editor
	MenuRegistry.appendMenuItem(MenuId.ViewTitle, {
		command: {
			id: AICustomizationManagementCommands.OpenEditor,
			title: localize2('openChatCustomizations', "Open Customizations"),
			category: CHAT_CATEGORY,
			icon: Codicon.gear
		},
		group: 'navigation',
		when: ContextKeyExpr.and(
			ChatContextKeys.enabled,
			ContextKeyExpr.equals('view', ChatViewId),
			ContextKeyExpr.has(`config.${ChatConfiguration.ChatCustomizationMenuEnabled}`)
		),
		order: 6
	});

	// When customizations menu is disabled, show the legacy gear submenu
	MenuRegistry.appendMenuItem(MenuId.ViewTitle, {
		submenu: CHAT_CONFIG_MENU_ID,
		title: localize2('config.label', "Configure Chat"),
		group: 'navigation',
		when: ContextKeyExpr.and(ContextKeyExpr.equals('view', ChatViewId), ContextKeyExpr.has(`config.${ChatConfiguration.ChatCustomizationMenuEnabled}`).negate()),
		icon: Codicon.gear,
		order: 6
	});
}

export function stringifyItem(item: IChatRequestViewModel | IChatResponseViewModel, includeName = true): string {
	if (isRequestVM(item)) {
		return (includeName ? `${item.username}: ` : '') + item.messageText;
	} else {
		return (includeName ? `${item.username}: ` : '') + item.response.toString();
	}
}

export interface IToolFilteringOptions {
	allTools: IToolData[];
	allToolSets: IToolSet[];
	toolsInclude?: string[];
	toolsExclude?: string[];
}

export interface IToolFilteringResult {
	enablementMap: Map<IToolData | IToolSet, boolean>;
	unknownIdentifiers: string[];
}

/**
 * Computes the tool enablement map based on include/exclude filters.
 *
 * Resolution algorithm:
 * 1. If `toolsInclude` is specified, start with only those tools/toolsets enabled
 * 2. If `toolsExclude` is specified, remove those tools/toolsets
 * 3. Explicit tool references in `toolsInclude` override toolset exclusions
 * 4. Explicit tool exclusions always win
 * 5. Toolset enablement is calculated based on whether all member tools are enabled
 *
 * @throws Error if filtering results in zero enabled tools
 */
export function computeToolEnablementMap(options: IToolFilteringOptions): IToolFilteringResult {
	const { allTools, allToolSets, toolsInclude, toolsExclude } = options;

	const enablementMap = new Map<IToolData | IToolSet, boolean>();
	const matchedIdentifiers = new Set<string>();

	// Helper to check if a tool matches any identifier (by id or toolReferenceName)
	const toolMatches = (tool: IToolData, identifiers: Set<string>): boolean => {
		if (identifiers.has(tool.id)) {
			matchedIdentifiers.add(tool.id);
			return true;
		}
		if (tool.toolReferenceName && identifiers.has(tool.toolReferenceName)) {
			matchedIdentifiers.add(tool.toolReferenceName);
			return true;
		}
		return false;
	};

	// Helper to check if a toolset matches any identifier (by id or referenceName)
	const toolSetMatches = (toolSet: IToolSet, identifiers: Set<string>): boolean => {
		if (identifiers.has(toolSet.id)) {
			matchedIdentifiers.add(toolSet.id);
			return true;
		}
		if (identifiers.has(toolSet.referenceName)) {
			matchedIdentifiers.add(toolSet.referenceName);
			return true;
		}
		return false;
	};

	// Track which tools are explicitly referenced in toolsInclude
	const explicitlyIncludedTools = new Set<IToolData>();

	// Step 1: Build initial set based on toolsInclude
	if (toolsInclude) {
		const includeSet = new Set(toolsInclude);

		// First, process toolsets - if a toolset matches, enable all its tools
		for (const toolSet of allToolSets) {
			if (toolSetMatches(toolSet, includeSet)) {
				for (const tool of toolSet.getTools()) {
					enablementMap.set(tool, true);
				}
			}
		}

		// Then process individual tools
		for (const tool of allTools) {
			if (toolMatches(tool, includeSet)) {
				enablementMap.set(tool, true);
				explicitlyIncludedTools.add(tool);
			} else if (!enablementMap.has(tool)) {
				enablementMap.set(tool, false);
			}
		}
		// Also process tools from toolsets that may not be in allTools
		for (const toolSet of allToolSets) {
			for (const tool of toolSet.getTools()) {
				if (toolMatches(tool, includeSet)) {
					enablementMap.set(tool, true);
					explicitlyIncludedTools.add(tool);
				} else if (!enablementMap.has(tool)) {
					enablementMap.set(tool, false);
				}
			}
		}
	} else {
		// No toolsInclude specified - start with all tools enabled
		for (const tool of allTools) {
			enablementMap.set(tool, true);
		}
		for (const toolSet of allToolSets) {
			for (const tool of toolSet.getTools()) {
				enablementMap.set(tool, true);
			}
		}
	}

	// Step 2: Remove tools matching toolsExclude
	if (toolsExclude) {
		const excludeSet = new Set(toolsExclude);

		// First, process toolsets - if a toolset matches, disable all its tools
		// (unless explicitly included as individual tools)
		for (const toolSet of allToolSets) {
			if (toolSetMatches(toolSet, excludeSet)) {
				for (const tool of toolSet.getTools()) {
					// Explicit tool reference overrides toolset exclusion
					if (!explicitlyIncludedTools.has(tool)) {
						enablementMap.set(tool, false);
					}
				}
			}
		}

		// Then process individual tools - explicit exclusion always wins
		for (const tool of allTools) {
			if (toolMatches(tool, excludeSet)) {
				enablementMap.set(tool, false);
			}
		}
		for (const toolSet of allToolSets) {
			for (const tool of toolSet.getTools()) {
				if (toolMatches(tool, excludeSet)) {
					enablementMap.set(tool, false);
				}
			}
		}
	}

	// Collect unknown identifiers
	const allIdentifiers = new Set([...(toolsInclude ?? []), ...(toolsExclude ?? [])]);
	const unknownIdentifiers: string[] = [];
	for (const identifier of allIdentifiers) {
		if (!matchedIdentifiers.has(identifier)) {
			unknownIdentifiers.push(identifier);
		}
	}

	// Validate at least one tool is enabled
	const enabledToolCount = Array.from(enablementMap.entries()).filter(([item, enabled]) => enabled && !isToolSet(item)).length;
	if (enabledToolCount === 0) {
		throw new Error('Tool filtering resulted in zero enabled tools. At least one tool must be enabled.');
	}

	// Calculate toolset enablement based on whether all member tools are enabled
	for (const toolSet of allToolSets) {
		const toolSetTools = Array.from(toolSet.getTools());
		const allToolsEnabled = toolSetTools.length > 0 && toolSetTools.every(t => enablementMap.get(t) === true);
		enablementMap.set(toolSet, allToolsEnabled);
	}

	return { enablementMap, unknownIdentifiers };
}


/**
 * Returns whether we can continue clearing/switching chat sessions, false to cancel.
 */
export async function handleCurrentEditingSession(model: IChatModel, phrase: string | undefined, dialogService: IDialogService): Promise<boolean> {
	return showClearEditingSessionConfirmation(model, dialogService, { messageOverride: phrase });
}

/**
 * Returns whether we can switch the agent, based on whether the user had to agree to clear the session, false to cancel.
 */
export async function handleModeSwitch(
	accessor: ServicesAccessor,
	fromMode: ChatModeKind,
	toMode: ChatModeKind,
	requestCount: number,
	model: IChatModel | undefined,
): Promise<false | { needToClearSession: boolean }> {
	if (!model?.editingSession || fromMode === toMode) {
		return { needToClearSession: false };
	}

	const dialogService = accessor.get(IDialogService);
	const needToClearEdits = (fromMode === ChatModeKind.Edit || toMode === ChatModeKind.Edit) && requestCount > 0;
	if (needToClearEdits) {
		// Switching into or out of edit mode, ask to discard the session
		const phrase = localize('switchMode.confirmPhrase', "Switching agents will end your current edit session.");

		const currentEdits = model.editingSession.entries.get();
		const undecidedEdits = currentEdits.filter((edit) => edit.state.get() === ModifiedFileEntryState.Modified);
		if (undecidedEdits.length > 0) {
			if (!await handleCurrentEditingSession(model, phrase, dialogService)) {
				return false;
			}

			return { needToClearSession: true };
		} else {
			const confirmation = await dialogService.confirm({
				title: localize('agent.newSession', "Start new session?"),
				message: localize('agent.newSessionMessage', "Changing the agent will end your current edit session. Would you like to change the agent?"),
				primaryButton: localize('agent.newSession.confirm', "Yes"),
				type: 'info'
			});
			if (!confirmation.confirmed) {
				return false;
			}

			return { needToClearSession: true };
		}
	}

	return { needToClearSession: false };
}

export interface IClearEditingSessionConfirmationOptions {
	titleOverride?: string;
	messageOverride?: string;
	isArchiveAction?: boolean;
}

/**
 * Clears the current chat session and starts a new one, preserving
 * the session type (e.g. Claude, Cloud, Background) for non-local sessions
 * in the sidebar.
 */
export async function clearChatSessionPreservingType(widget: IChatWidget, viewsService: IViewsService, sessionType?: string): Promise<void> {
	const currentResource = widget.viewModel?.model.sessionResource;
	const newSessionType = sessionType ?? (currentResource ? getChatSessionType(currentResource) : localChatSessionType);
	if (isIChatViewViewContext(widget.viewContext) && newSessionType !== localChatSessionType) {
		// For the sidebar, we need to explicitly load a session with the same type
		const newResource = URI.from({ scheme: newSessionType, path: `/untitled-${generateUuid()}` });
		const view = await viewsService.openView(ChatViewId) as ChatViewPane;
		await view.loadSession(newResource);
	} else {
		// For the editor, widget.clear() already preserves the session type via clearChatEditor
		await widget.clear();
	}
}


// --- Chat Submenus in various Components

MenuRegistry.appendMenuItem(MenuId.EditorContext, {
	submenu: MenuId.ChatTextEditorMenu,
	group: '1_chat',
	order: 5,
	title: localize('generateCode', "Generate Code"),
	when: ContextKeyExpr.and(
		ChatContextKeys.Setup.hidden.negate(),
		ChatContextKeys.Setup.disabled.negate()
	)
});

// --- Chat Default Visibility

registerAction2(class ToggleDefaultVisibilityAction extends Action2 {
	constructor() {
		super({
			id: 'workbench.action.chat.toggleDefaultVisibility',
			title: localize2('chat.toggleDefaultVisibility.label', "Show View by Default"),
			toggled: ContextKeyExpr.equals('config.workbench.secondarySideBar.defaultVisibility', 'hidden').negate(),
			f1: false,
			menu: {
				id: MenuId.ViewTitle,
				when: ContextKeyExpr.and(
					ContextKeyExpr.equals('view', ChatViewId),
					ChatContextKeys.panelLocation.isEqualTo(ViewContainerLocation.AuxiliaryBar),
				),
				order: 0,
				group: '5_configure'
			},
		});
	}

	async run(accessor: ServicesAccessor) {
		const configurationService = accessor.get(IConfigurationService);

		const currentValue = configurationService.getValue<'hidden' | unknown>('workbench.secondarySideBar.defaultVisibility');
		configurationService.updateValue('workbench.secondarySideBar.defaultVisibility', currentValue !== 'hidden' ? 'hidden' : 'visible');
	}
});

registerAction2(class EditToolApproval extends Action2 {
	constructor() {
		super({
			id: 'workbench.action.chat.editToolApproval',
			title: localize2('chat.editToolApproval.label', "Manage Tool Approval"),
			metadata: {
				description: localize2('chat.editToolApproval.description', "Edit/manage the tool approval and confirmation preferences for AI chat agents."),
			},
			precondition: ChatContextKeys.enabled,
			f1: true,
			category: CHAT_CATEGORY,
		});
	}

	async run(accessor: ServicesAccessor, scope?: 'workspace' | 'profile' | 'session'): Promise<void> {
		const confirmationService = accessor.get(ILanguageModelToolsConfirmationService);
		const toolsService = accessor.get(ILanguageModelToolsService);
		confirmationService.manageConfirmationPreferences([...toolsService.getAllToolsIncludingDisabled()], scope ? { defaultScope: scope } : undefined);
	}
});
