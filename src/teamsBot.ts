import { TeamsActivityHandler, MessageFactory } from "botbuilder";
export class TeamsBot extends TeamsActivityHandler {
	constructor() {
		super();

		this.onMessage(async (context, next) => {
			const { text } = context.activity;

			// Check if the message contains a date
			const dateRegex = /(\d{4})-(\d{2})-(\d{2})/;
			const dateMatch = text.match(dateRegex);

			if (dateMatch) {
				const [_, year, month, day] = dateMatch;

				// Add the date to the calendar
				const event = {
					subject: "New Event",
					start: {
						dateTime: `${year}-${month}-${day}T08:00:00`,
						timeZone: "Pacific Standard Time",
					},
					end: {
						dateTime: `${year}-${month}-${day}T09:00:00`,
						timeZone: "Pacific Standard Time",
					},
					body: {
						contentType: "text",
						content: "Event description",
					},
				};

				// await this.createEventInCalendar(event);
				await context.sendActivity("Event added to the calendar.");
			}

			// Continue processing other activities
			await next();
		});
		
	}
}
