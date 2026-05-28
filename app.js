const { stripMentionsText } = require("@microsoft/teams.api");
const { App } = require("@microsoft/teams.apps");
const { LocalStorage } = require("@microsoft/teams.common");
const config = require("./config");
const { ManagedIdentityCredential } = require("@azure/identity");
const { MessageFactory } = require("botbuilder"); // Required to build Suggested Actions

// Create storage for conversation history
const storage = new LocalStorage();

const createTokenFactory = () => {
  return async (scope, tenantId) => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// Create the app with storage
const app = new App({
  ...credentialOptions,
  storage,
});

const getConversationState = (conversationId) => {
  let state = storage.get(conversationId);
  if (!state) {
    state = { count: 0 };
    storage.set(conversationId, state);
  }
  return state;
};

// --- CONFIGURABLE HELPER PROMPTS & KEYWORDS ---
// List of all keywords that are allowed to trigger the external API call
const HELPER_KEYWORDS = ["fetch", "call api", "get data", "trigger api", "run check"];

// Helper function to generate clickable prompt options for the user
const sendHelperPrompts = async (context, introductoryText) => {
  const actions = [
    { title: "Get Data From API", type: "imBack", value: "Get Data" },
    { title: "Trigger API Check", type: "imBack", value: "Trigger API" },
    { title: "Get some Unit info ", type: "imBack", value: "/runtime" }
  ];
  
  const message = MessageFactory.suggestedActions(actions, introductoryText);
  await context.send(message);
};

// Welcome user and send helper prompts when they join or start a new conversation
app.on("conversationUpdate", async (context) => {
  if (context.activity.membersAdded && context.activity.membersAdded.length > 0) {
    for (let member of context.activity.membersAdded) {
      if (member.id !== context.activity.recipient.id) {
        await sendHelperPrompts(context, "Welcome! I am your integration assistant. Choose an option below or type a command to begin:");
      }
    }
  }
});

app.on("message", async (context) => {
  const activity = context.activity;
  const text = stripMentionsText(activity).trim();
  const lowerText = text.toLowerCase();

  if (text === "/reset") {
    storage.delete(activity.conversation.id);
    await context.send("Ok I've deleted the current conversation state.");
    return;
  }

  if (text === "/count") {
    const state = getConversationState(activity.conversation.id);
    await context.send(`The count is ${state.count}`);
    return;
  }

  if (text === "/diag") {
    await context.send(JSON.stringify(activity));
    return;
  }

  if (text === "/state") {
    const state = getConversationState(activity.conversation.id);
    await context.send(JSON.stringify(state));
    return;
  }

  if (text === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0",
    };
    await context.send(JSON.stringify(runtime));
    return;
  }

  // --- DYNAMIC KEYWORD MATCHING FOR API TRIGGER ---
  // Checks if the user's typed input or clicked button value matches any keyword in our array
  const isApiTrigger = HELPER_KEYWORDS.some(keyword => lowerText.includes(keyword));

  if (isApiTrigger) {
    await context.send("Fetching data from the external endpoint...");

    try {
      const targetEndpoint = process.env.EXTERNAL_API_ENDPOINT || "https://jsonplaceholder.typicode.com/todos/1";
      const response = await fetch(targetEndpoint);

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();

      const replyText = `**API Response Received Successfully!**\n\n` +
                        `• **ID:** ${data.id}\n` +
                        `• **Title:** ${data.title}\n` +
                        `• **Completed Status:** ${data.completed}`;

      await context.send(replyText);
      
      // Re-prompt the user with the action buttons for their next interaction
      await sendHelperPrompts(context, "What would you like to do next?");
      return; 
    } catch (error) {
      console.error("Outbound API integration error:", error);
      await context.send(`Failed to fetch data from endpoint. Error details: ${error.message}`);
      return;
    }
  }
  // --- END OF DYNAMIC KEYWORD MATCHING ---

  // Default echo behavior
  const state = getConversationState(activity.conversation.id);
  state.count++;
  await context.send(`[${state.count}] you said: ${text}`);
  
  // Provide helper choices if they type a random message so they don't get stuck
  await sendHelperPrompts(context, "Need help? Select an option below:");
});

module.exports = app;