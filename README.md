# Overview of the AI Chat Bot Graph sample

This sample is related to [this article](https://michaelmaillot.github.io/articles/20240327-getting-started-with-teams-ai/) and is ready-to-use once configured.

This sample showcases a bot app that responds to user questions like an AI assistant. This enables your users to talk with the AI assistant in Teams to find information.

The app is built using the Teams AI library, which provides the capabilities to build AI-based Teams applications.

- [Overview of the AI Chat Bot Graph sample](#overview-of-the-ai-chat-bot-graph-sample)
  - [Get started with the AI Chat Bot Graph sample](#get-started-with-the-ai-chat-bot-graph-sample)
    - [Local running](#local-running)
    - [(optional) Deployment](#optional-deployment)
  - [What's included in the sample](#whats-included-in-the-sample)
    - [index.ts](#indexts)
  - [Additional information and references](#additional-information-and-references)

## Get started with the AI Chat Bot Graph sample

> **Prerequisites**
>
> To run the AI Chat Bot Graph sample in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 16, 18
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-cli)
> - An Azure subscription to host resources such as [Azure OpenAI](https://learn.microsoft.com/azure/ai-services/openai/how-to/create-resource?pivots=web-portal) and [Azure AI Content Safety](https://learn.microsoft.com/azure/ai-services/content-safety/)

### Local running

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) to both Microsoft 365 and Azure if you haven't already.
3. Add *env/.env.local*, fill in following info:
   ```ini
    # This file includes environment variables that can be committed to git. It's gitignored by default because it represents your local development environment.

    # Built-in environment variables
    TEAMSFX_ENV=local
    APP_NAME_SUFFIX=local
    
    # Generated during provision, you can also add your own variables.
    BOT_ID=
    TEAMS_APP_ID=
    BOT_DOMAIN=
    BOT_ENDPOINT=
    AAD_APP_TENANT_ID=
    AAD_APP_OBJECT_ID=
    AAD_APP_CLIENT_ID=
    AAD_APP_ACCESS_AS_USER_PERMISSION_ID=
    AAD_APP_OAUTH_AUTHORITY=
    AAD_APP_OAUTH_AUTHORITY_HOST=
   ```
4. Add *env/.env.local.user*, fill in following secret info:
   1. Azure OpenAI key `SECRET_AZURE_OPENAI_API_KEY=<your-key>`
   2. Azure OpenAI endpoint `SECRET_AZURE_OPENAI_ENDPOINT=<your-endpoint>`
   3. Azure OpenAI deployment `SECRET_AZURE_OPENAI_DEPLOYMENT=<your-deployment-model>`
   4. Azure AI Content Safety key `SECRET_AZURE_CONTENT_SAFETY_KEY=<your-key>`
   5. Azure AI Content Safety endpoint `SECRET_AZURE_CONTENT_SAFETY_ENDPOINT=<your-endpoint>`
   6. (provisioned automatically) Bot password `SECRET_BOT_PASSWORD=`
   7. (provisioned automatically) Teams App update time `TEAMS_APP_UPDATE_TIME=`
   8. (provisioned automatically) Entra ID OBO Client ID `SECRET_AAD_APP_CLIENT_SECRET=`
5. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
6. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
7. You will receive a welcome message from the bot, or send any message to get a response.

**Congratulations**! You are running an application that can now interact with users in Teams:

![ai chat bot](https://user-images.githubusercontent.com/7642967/258726187-8306610b-579e-4301-872b-1b5e85141eff.png)

### (optional) Deployment

The sample is ready for being deployed. But first, a *.env.dev.user* file has to be created in the `/env` folder, with following secret info:

```ini
# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# Secrets. Keys prefixed with `SECRET_` will be masked in Teams Toolkit logs.
SECRET_OPENAI_API_KEY=<your-key>
SECRET_AZURE_OPENAI_ENDPOINT=<your-endpoint>
SECRET_AZURE_CONTENT_SAFETY_KEY=<your-key>
SECRET_AZURE_CONTENT_SAFETY_ENDPOINT=<your-endpoint>
SECRET_BOT_PASSWORD=
SECRET_AAD_APP_CLIENT_SECRET=
TEAMS_APP_UPDATE_TIME=
```

Then from Teams Toolkit, select `dev` environment and click in the order the following actions in `LIFECYCLE` tab:

- Provision
- Deploy
- Publish

The app should be available in you organization!

## What's included in the sample

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the Teams application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `public`     | Sign-in and redirection pages for SSO               |
| `src`        | The source code for the application                 |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/index.ts`| Sets up and configures the AI Chat Bot Graph.|
|`src/appContext.ts`| Handles AI resources configuration for the AI Chat Bot Graph.|
|`src/config.ts`| Defines the environment variables.|
|`src/bots/bot-sequence.ts`| Defines the actions available and the Bot behavior using the `sequence` augmentation.|
|`src/bots/bot-monologue.ts`| Defines the actions available and the Bot behavior using the `monologue` augmentation.|
|`src/prompts/chat/skprompt.txt`| Defines the prompt.|
|`src/prompts/chat/config.json`| Configures the prompt.|
|`src/prompts/chat/actions.json`| Defines the available augmentation actions.|
|`src/services/appBuilderService.ts`| Handles the SSO authentication and the Azure AI Content Safety flagged input / output events.|
|`src/services/graphClientService.ts`| Contains everything related to Microsoft Graph (token init, methods,...).|

| File                                 | Contents                                           |
| - | - |
|`teamsapp.yml`|This is the main Teams Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging.|
|`teamsapp.testtool.yml`| This overrides `teamsapp.yml` with actions that enable local execution and debugging in Teams App Test Tool. WARNING: current version of Teams App Test Tool doesn't work with SSO|

### index.ts

By default, the sample is referring to the `bot-sequence.ts` augmentation file. To use the `bot-monologue.ts` augmentation file, update the following files:

- Update the `index.ts` file:
   
```typescript
// Import required packages
//...

// To use bot-monologue, comment the bot-sequence and uncomment bot-monologue
import app from "../src/bots/bot-sequence";
// import app from "../src/bots/bot-monologue";
```

- Update the `src/prompts/chat/config.json` file:

```json
{
  "schema": 1.1,
  "description": "AI Bot",
  "type": "completion",
  "completion": {
    "completion_type": "chat",
    "max_input_tokens": 4000,
    "max_tokens": 1500,
    "temperature": 0.3,
    "top_p": 0.0,
    "presence_penalty": 0.6,
    "frequency_penalty": 0.0,
    "stop_sequences": []
  },
  "augmentation": {
      "augmentation_type": "sequence" // Replace "sequence" value by "monologue"
  }
}
```

## Additional information and references

- [Teams AI library](https://aka.ms/teams-ai-library)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-cli)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)