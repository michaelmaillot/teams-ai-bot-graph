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
import { ApplicationTurnState, moderator, planner, scopes, storage } from "../appContext";

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
    const plan = await app.ai.planner.beginTask(context, state, app.ai);
    // The command [1] matches with the first SAY command in the generated DO / SAY plan
    // It means that the actions.json must be registered
    // Otherwise, it will use the auto-generated SAY command ([0])
    const welcomeCommand: PredictedSayCommand = (plan.commands.length > 1 ? plan.commands[1] : plan.commands[0]) as PredictedSayCommand;

    await context.sendActivity(welcomeCommand.response);

    return AI.StopCommandName;
});

app.ai.action("getUserInfo", async (context: TurnContext, state: ApplicationTurnState) => {

    if (state.conversation.UserInfo === undefined) {
        const me = await graphClient.getMe();

        state.conversation.UserInfo = {
            mail: me.mail,
            name: me.displayName
        };

        await sendActivityFromPlanner(context, state);

        return AI.StopCommandName;
    }

    // If the action is triggered once again
    // We willingly return nothing from the action
    // So that the AI system will handle the answer based on the updated prompt
    return '';
});

app.ai.action("getUserColleagues", async (context: TurnContext, state: ApplicationTurnState) => {

    if (state.conversation.colleagues === undefined) {
        const people = await graphClient.getMyPeople();

        state.conversation.colleagues = people.map(person => person.displayName);

        await sendActivityFromPlanner(context, state);

        return AI.StopCommandName;
    }

    return '';
});

app.ai.action("getUserUnreadEmails", async (context: TurnContext, state: ApplicationTurnState) => {

    if (state.conversation.nbUnreadEmails === undefined) {
        const emails = await graphClient.getMyUnreadEmails();

        state.conversation.nbUnreadEmails = emails['@odata.count'];

        await sendActivityFromPlanner(context, state);

        return AI.StopCommandName;
    }

    return '';
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

app.messageReactions("reactionsAdded", async (context, state) => {
    await context.sendActivity(`I see ${context.activity.from.name} reacted to a message with ${context.activity.reactionsAdded?.map(reaction => reaction.type).join(', ')}`);
});

async function sendActivityFromPlanner(context: TurnContext, state: ApplicationTurnState): Promise<void> {
    let plan: Plan;
    let nbRetries = 0;
    const maxRetries = 3;

    do {
        // Start a new completion task from the planner
        // (beginTask method triggers in fact continueTask under the hood)
        plan = await app.ai.planner.continueTask(context, state, app.ai);

        // if the answer is generated from a DO / SAY plan match, process it
        if (plan.commands.length > 1 && plan.commands[1]?.type === "SAY") {
            await context.sendActivity((plan.commands[1] as PredictedSayCommand).response);
            break;
        }
        else {
            nbRetries++;
        }
    }
    while (nbRetries < maxRetries);

    if (nbRetries === maxRetries) {
        await context.sendActivity((plan.commands[0] as PredictedSayCommand).response);
    }
} 