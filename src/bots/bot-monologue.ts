import {
    AI,
    Plan,
    PredictedSayCommand
} from "@microsoft/teams-ai";
import { TurnContext } from "botbuilder";
import { adapter } from "../index";
import config from "../config";
import { GraphClientService } from "../services/graphClientService";
import IGraphMeetingTime from "../interfaces/IGraphMeetingTime";
import { MeetingTimeSuggestion } from "@microsoft/microsoft-graph-types";
import { appBuilderService } from "../services/appBuilderService";
import { ApplicationTurnState, moderator, planner, scopes, storage } from "appContext";

export const run = (context: TurnContext) => app.run(context);

const app = new appBuilderService<ApplicationTurnState>(storage,
    planner,
    moderator,
    adapter,
    config,
    scopes)
    .app;

export default app;

let graphClient: GraphClientService;

app.authentication.get('graph').onUserSignInSuccess(async (context: TurnContext, state: ApplicationTurnState) => {
    if (!graphClient) {
        graphClient = new GraphClientService(state.temp.authTokens['graph']!);

        // This will indicate the user that the plan is being processed during init
        await context.sendActivities([
            { type: 'typing' },
            { type: 'delay', value: 2000 }
        ])

        // Here we'll init the LLM at first user's interaction
        // Otherwise the user will have to send another message
        // This is because of the silent SSO event performed here with onUserSignInSuccess
        await app.ai.doAction(context, state, "greetings");

        // Clear the conversation state after greetings
        // So that the model won't be "stuck" for further requests
        // await state.deleteConversationState();
    }
    else {
        // If token has expired, update the Graph client
        graphClient = new GraphClientService(state.temp.authTokens['graph']!);
    }
});

app.authentication.get('graph').onUserSignInFailure(async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity('Failed to log in');
});

app.ai.action("greetings", async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity('Hello! How can I help you today?');
    return AI.StopCommandName;
  });
  
  app.ai.action("getUserInfo", async (context: TurnContext, state: ApplicationTurnState) => {
    if (state.conversation.UserInfo === undefined) {
      if (!await app.authentication.get('graph').isUserSignedIn) {
        await app.authentication.signUserIn(context, state, 'graph');
      }
      const me = await graphClient.getMe();
  
      state.conversation.UserInfo = {
        mail: me.mail,
        name: me.displayName
      };
    }
  
    const userInfo = state.conversation.UserInfo;
  
    await context.sendActivity('You are ' + userInfo.name + ' and your email is ' + userInfo.mail + '. How can I help you today?');
  
    return AI.StopCommandName;
  });
  
  app.ai.action("getUserColleagues", async (context: TurnContext, state: ApplicationTurnState) => {
    if (state.conversation.colleagues === undefined) {
      const people = await graphClient.getMyPeople();
  
      state.conversation.colleagues = people.map(person => person.displayName);
    }
  
    const colleagues = state.conversation.colleagues;
  
    await context.sendActivity(`You are collaborating at least with ${colleagues.length} people: <b>${colleagues.join('</b>, <b>')}</b>`);
  
    return AI.StopCommandName;
  });
  
  app.ai.action("getUserUnreadEmails", async (context: TurnContext, state: ApplicationTurnState) => {
    if (state.conversation.nbUnreadEmails === undefined) {
      const emails = await graphClient.getMyUnreadEmails();
  
      state.conversation.nbUnreadEmails = emails['@odata.count'];
    }
  
    const unreadEmails = state.conversation.nbUnreadEmails;
  
    await context.sendActivity(`You have ${unreadEmails} unread emails`);
  
    return AI.StopCommandName;
  });
  
  planner.prompts.addFunction('getDate', async (context, state, data) => {
    return new Date();
  });
  
  app.ai.action("findMeetingTimes", async (context: TurnContext, state: ApplicationTurnState, parameters: IGraphMeetingTime) => {
    if (parameters.colleague === '') {
      await context.sendActivity('You need to specify a colleague');
      // return 'You need to specify a colleague';
      return AI.StopCommandName;
    }
  
    const meetingSuggestions: MeetingTimeSuggestion[] = await graphClient.findMeetingTimes(parameters);
  
    if (meetingSuggestions.length === 0) {
      await context.sendActivity(`No meeting times found with ${parameters.colleague}`);
      return AI.StopCommandName;
    }
  
    let message = `I found ${meetingSuggestions.length} meeting times for you with ${parameters.colleague}:`;
    for (const meetingSuggestion of meetingSuggestions.filter(meeting => meeting.confidence === 100)) {
      const startDate = new Date(meetingSuggestion.meetingTimeSlot.start.dateTime);
      const endDate = new Date(meetingSuggestion.meetingTimeSlot.end.dateTime);
      message += `\n * <b>${startDate.toDateString()}</b>: from ${startDate.toTimeString().split(' ')[0]} to ${endDate.toTimeString().split(' ')[0]}`;
    }
  
    await context.sendActivity(message);
  
    return AI.StopCommandName;
  });
  
  // List for /reset command and then delete the conversation state
  app.message('/reset', async (context: TurnContext, state: ApplicationTurnState) => {
    state.deleteConversationState();
    await app.ai.doAction(context, state, "greetings");
  });
  
  app.messageReactions("reactionsAdded", async (context, state) => {
    await context.sendActivity(`I see ${context.activity.from.name} reacted to a message with ${context.activity.reactionsAdded?.map(reaction => reaction.type).join(', ')}`);
  });