const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TableStorage } = require('botbuilder-azure');
const restify = require('restify');

const development = (process.env.NODE_ENV === 'development');
if (development) {
  require('dotenv').config();
}

// Create server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`${server.name} listening to ${server.url}`);
});

// Create adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword
});

// Add conversation state middleware
let conversationState = null;
if (development) {
  conversationState = new ConversationState(new MemoryStorage());
  adapter.use(conversationState);
} else {
  const tableName = 'botdata';
  const azureStorage = new TableStorage({
    tableName,
    storageAccountOrConnectionString: process.env.AzureWebJobsStorage
  });
  conversationState = new ConversationState(azureStorage);
  adapter.use(conversationState);
}

// Listen for incoming requests
server.post('/api/messages', (req, res) => {
  // Route received request to adapter for processing
  adapter.processActivity(req, res, (context) => {
    if (context.activity.type === 'message') {
      const state = conversationState.get(context);
      const count = state.count === undefined ? state.count = 0 : ++state.count;
      return context.sendActivity(`${count}: You said "${context.activity.text}"`);
    } else {
      return context.sendActivity(`[${context.activity.type} event detected]`);
    }
  });
});
