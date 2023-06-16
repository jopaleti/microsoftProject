import * as restify from "restify";
import { commandApp } from "./internal/initialize";
import { TeamsBot } from "./teamsBot";

// Import required packages IMPORTATION START
import { config } from "dotenv";
import * as path from "path";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
	CloudAdapter,
	ConfigurationBotFrameworkAuthentication,
	ConfigurationServiceClientCredentialFactory,
	MemoryStorage,
} from "botbuilder";

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
	{},
	new ConfigurationServiceClientCredentialFactory({
		MicrosoftAppId: process.env.BOT_ID,
		MicrosoftAppPassword: process.env.BOT_PASSWORD,
		MicrosoftAppType: "MultiTenant",
	})
);
// IMPORTATION ENDS
// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// This template uses `restify` to serve HTTP responses.
// Create a restify server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
	console.log(`\nApp Started, ${server.name} listening to ${server.url}`);
});

// Register an API endpoint with `restify`. Teams sends messages to your application
// through this endpoint.
//
// The Teams Toolkit bot registration configures the bot with `/api/messages` as the
// Bot Framework endpoint. If you customize this route, update the Bot registration
// in `templates/azure/provision/botservice.bicep`.
const teamsBot = new TeamsBot();
server.post("/api/messages", async (req, res) => {
	await commandApp.requestHandler(req, res, async (context) => {
		await teamsBot.run(context);
	});
});
