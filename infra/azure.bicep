@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@description('Required when create Azure Bot service')
param botAadAppClientId string

@secure()
@description('Required by Bot Framework package in your bot project')
param botAadAppClientSecret string

param webAppSKU string

@maxLength(42)
param botDisplayName string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location

@secure()
param azOpenAiKey string
param azOpenAiEndpoint string
param azOpenAiDeployment string
param aadTenantId string
param aadAppClientId string
param aadAppOauthAuthority string
@secure()
param aadAppClientSecret string
param azContentSafetyEndpoint string
@secure()
param azContentSafetyKey string

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your bot
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      ftpsState: 'FtpsOnly'
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    botAadAppClientId: botAadAppClientId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// Create and update app settings for web app
module appSettings 'appsettings.bicep' = {
  name: '${webAppName}-appsettings'
  params: {
    webAppName: webApp.name
    currentAppSettings: list(resourceId('Microsoft.Web/sites/config', webApp.name, 'appsettings'), '2021-02-01').properties
    appSettings: {
      WEBSITE_RUN_FROM_PACKAGE: '1'
      WEBSITE_NODE_DEFAULT_VERSION: '~18'
      RUNNING_ON_AZURE: '1'
      BOT_ID: botAadAppClientId
      BOT_PASSWORD: botAadAppClientSecret
      BOT_DOMAIN: webApp.properties.defaultHostName
      AZURE_OPENAI_API_KEY: azOpenAiKey
      AZURE_OPENAI_ENDPOINT: azOpenAiEndpoint
      AZURE_OPENAI_DEPLOYMENT: azOpenAiDeployment
      AAD_APP_TENANT_ID: aadTenantId
      AAD_APP_CLIENT_ID: aadAppClientId
      AAD_APP_OAUTH_AUTHORITY: aadAppOauthAuthority
      AAD_APP_CLIENT_SECRET: aadAppClientSecret
      AZURE_CONTENT_SAFETY_ENDPOINT: azContentSafetyEndpoint
      AZURE_CONTENT_SAFETY_KEY: azContentSafetyKey
    }
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
