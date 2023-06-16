import {
	Activity,
	CardFactory,
	MessageFactory,
	TurnContext,
	TeamsActivityHandler,
	MemoryStorage,
	ActionTypes,
} from "botbuilder";
import {
	CommandMessage,
	TeamsFxBotCommandHandler,
	TriggerPatterns,
} from "@microsoft/teamsfx";
import * as Util from "util";
import * as path from "path";

// import * as debug from "debug";
// const log = debug("msteams");
// IMPORTATION STARTS
import {
	Application,
	ConversationHistory,
	DefaultPromptManager,
	DefaultTurnState,
	OpenAIModerator,
	OpenAIPlanner,
	AI,
} from "@microsoft/botbuilder-m365";
// eslint-disable-next-line @typescript-eslint/no-empty-interface
interface ConversationState {}
type ApplicationTurnState = DefaultTurnState<ConversationState>;

// Create AI components
const planner = new OpenAIPlanner({
	apiKey: "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",
	defaultModel: "text-davinci-003",
	logRequests: true,
});
const moderator = new OpenAIModerator({
	apiKey: "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",
	moderate: "both",
});
const promptManager = new DefaultPromptManager(
	path.join(__dirname, "../prompts")
);

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
	storage,
	ai: {
		planner,
		moderator,
		promptManager,
		prompt: "chat",
		history: {
			assistantHistoryType: "text",
		},
	},
});

export class convoBot extends TeamsActivityHandler {
	triggerPatterns: TriggerPatterns = "mentionme";
	async handleCommandReceived(
		context: TurnContext,
		message: CommandMessage
	): Promise<string | Partial<Activity> | void> {
		console.log(`App received message: ${message.text}`);

		if (message.text === "mentionme") {
			await this.handleMessageMentionMeOneOnOne(context);
		} else if (message.text.endsWith("</at> mentionme")) {
			await this.handleMessageMentionMeChannelConversation(context);
		} else {
			await this.handleAiBot(context);
		}
	}

	private async handleAiBot(context: TurnContext): Promise<void> {
		app.ai.action(AI.FlaggedInputActionName, async (context, state, data) => {
			await context.sendActivity(
				`I'm sorry your message was flagged: ${JSON.stringify(data)}`
			);
			return false;
		});

		app.ai.action(AI.FlaggedOutputActionName, async (context, state, data) => {
			await context.sendActivity(`I'm not allowed to talk about such things.`);
			return false;
		});

		app.message("/history", async (context, state) => {
			const history = ConversationHistory.toString(state, 2000, "\n\n");
			await context.sendActivity(history);
		});
	}

	private async handleMessageMentionMeOneOnOne(
		context: TurnContext
	): Promise<void> {
		const mention = {
			mentioned: context.activity.from,
			text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
			type: "mention",
		};

		const replyActivity = MessageFactory.text(
			`Hi ${mention.text} from 1:1 chat.`
		);
		replyActivity.entities = [mention];
		await context.sendActivity(replyActivity);
	}

	private async handleMessageMentionMeChannelConversation(
		context: TurnContext
	): Promise<void> {
		const mention = {
			mentioned: context.activity.from,
			text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
			type: "text",
		};

		const replyActivity = MessageFactory.text(`Hi ${mention.text}!`);
		replyActivity.entities = [mention];
		const followUpActivity = MessageFactory.text(
			`*We are in a channel conversion group chat!*`
		);
		await context.sendActivities([replyActivity, followUpActivity]);
	}
}
