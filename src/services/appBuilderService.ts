import { AI, Application, ApplicationBuilder, DefaultConversationState, DefaultTempState, DefaultUserState, Moderator, Planner, TeamsAdapter, TurnState } from "@microsoft/teams-ai";
import { Storage } from "botbuilder";

export class appBuilderService<T extends TurnState<DefaultConversationState, DefaultUserState, DefaultTempState>> {
  private _app: Application<T>;

  constructor(storage: Storage, planner: Planner, moderator: Moderator, adapter: TeamsAdapter, config: any, permissions: string[]) {
    this._app = new ApplicationBuilder<T>()
      .withStorage(storage)
      .withAIOptions({
        planner,
        moderator
      })
      .withAuthentication(adapter, {
        settings: {
          graph: {
            scopes: permissions,
            msalConfig: {
              auth: {
                clientId: config.appId,
                clientSecret: config.appPassword,
                authority: config.authority,
              }
            },
            signInLink: `https://${config.botDomain}/auth-start.html`,
            endOnInvalidMessage: true
          }
        }
      })
      .build();

    this._app.ai.action(AI.FlaggedInputActionName, async (context, _state, data) => {
      await context.sendActivity(`I'm sorry your message was flagged: ${JSON.stringify(data)}`);
      return AI.StopCommandName;
    });

    this._app.ai.action(AI.FlaggedOutputActionName, async (context, _state, data) => {
      await context.sendActivity(`I'm not allowed to talk about such things.`);
      await context.sendActivity(`I'm sorry the output message was flagged: ${JSON.stringify(data)}`);
      return AI.StopCommandName;
    });
  }

  public get app(): Application<T> {
    return this._app;
  }
}