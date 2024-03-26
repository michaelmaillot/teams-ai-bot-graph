const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  botDomain: process.env.BOT_DOMAIN,
  botEndpoint: process.env.BOT_ENDPOINT,
  openAIKey: process.env.AZURE_OPENAI_API_KEY,
  openAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  openAIDeployment: process.env.AZURE_OPENAI_DEPLOYMENT,
  appId: process.env.AAD_APP_CLIENT_ID,
  appPassword: process.env.AAD_APP_CLIENT_SECRET,
  tenantId: process.env.AAD_APP_TENANT_ID,
  authority: process.env.AAD_APP_OAUTH_AUTHORITY,
  azContentSafetyEndpoint: process.env.AZURE_CONTENT_SAFETY_ENDPOINT,
  azContentSafetyKey: process.env.AZURE_CONTENT_SAFETY_KEY
};

export default config;
