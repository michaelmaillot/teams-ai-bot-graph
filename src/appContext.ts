import { MemoryStorage } from "botbuilder";
import * as path from "path";

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
import { ActionPlanner, OpenAIModel, PromptManager, DefaultConversationState, AzureContentSafetyModerator, ModerationSeverity, TurnState } from "@microsoft/teams-ai";

import config from "./config";

// Create AI components
export const model = new OpenAIModel({
  // Use OpenAI
  // apiKey: config.openAIKey,
  // defaultModel: "gpt-3.5-turbo",

  // Uncomment the following lines to use Azure OpenAI
  azureApiKey: config.openAIKey,
  azureDefaultDeployment: config.openAIDeployment,
  azureEndpoint: config.openAIEndpoint,
  azureApiVersion: '2024-02-01',

  useSystemMessages: true,
  logRequests: true,
});

export const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../src/prompts"),
});

export const planner = new ActionPlanner({
  model,
  prompts,
  defaultPrompt: "chat",
});

planner.prompts.addFunction('getDate', async (_context, _state, _data) => {
  return new Date();
});

// Define storage and application
export const storage = new MemoryStorage();

export const moderator = new AzureContentSafetyModerator({
  apiKey: config.azContentSafetyKey,
  endpoint: config.azContentSafetyEndpoint,
  model: "gpt-3.5-turbo",
  moderate: 'both',
  categories: [
    {
      category: 'Hate',
      severity: ModerationSeverity.High
    },
    {
      category: 'SelfHarm',
      severity: ModerationSeverity.High
    },
    {
      category: 'Sexual',
      severity: ModerationSeverity.High
    },
    {
      category: 'Violence',
      severity: ModerationSeverity.High
    }
  ]
});

export interface ConversationState extends DefaultConversationState {
  graphToken: string;
  UserInfo: {
    mail: string;
    name: string;
  };
  colleagues: string[];
  nbUnreadEmails: number;
}

export type ApplicationTurnState = TurnState<ConversationState>;

export const scopes = ['User.Read', 'User.ReadBasic.All', 'People.Read', 'Calendars.Read.Shared', 'Mail.Read'];

