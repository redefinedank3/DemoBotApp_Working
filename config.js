const config = {
  MicrosoftAppId: process.env.CLIENT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.TENANT_ID,
  MicrosoftAppPassword: process.env.CLIENT_PASSWORD,
  externalApiEndpoint: process.env.EXTERNAL_API_ENDPOINT || 'https://jsonplaceholder.typicode.com/todos/1'
};

module.exports = config;
