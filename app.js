/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);

bot.dialog('/', [
    function (session) {
        session.beginDialog('askName');
        var text = teams.TeamsMessage.getTextWithoutMentions(session.message);
        var command = text.split(" ")[0];
        var extras = text.split(command+" ")[1];

        switch(command){
            case "help":
                help(extras,session);
                break;                
        }
    },
    function (session, results) {
        session.endDialog('Hello %s!', results.response);
    }
]);
bot.dialog('askName', [
    function (session) {
        builder.Prompts.text(session, 'Hi! What is your name?');
    },
    function (session, results) {
        session.endDialogWithResult(results);
    }
]);

function help(targets,session){
        var helpText = "**You can do the following commands:**\n\n";
        helpText += ". \n\n";
        helpText += "**help:** Displays this help\n\n";
        helpText += "**oncall [group]:** Displays who's on call\n\n";
        helpText += "**engage [group]:** Invite people to the chat\n\n";
        helpText += "**confCall:** Creates a conference bridge\n\n";

        postToChannel(session,helpText,"markdown");
    }

function postToChannel(session, text,type){
        var msg = new builder.Message(session);
        msg.text(text);
        if(!!type){
            console.log(type);
            msg.textFormat(type);
        }
        msg.textLocale('en-US');
        console.log(msg);
        bot.send(msg);
    }

