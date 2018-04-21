

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
//var bot = new builder.UniversalBot(connector);

var bot = new builder.UniversalBot(connector, function (session) {
   var text = session.message.text;
   var command = text.split(" ")[0];
   var extras = text.split(command+" ")[1];
   switch(command){
      case "help":
         //session.send("You said" + text);
         help(extras,session);
         break;
   }
});
bot.set('storage', tableStorage);

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

bot.dialog('engageButtonClick', [
        function (session, args, next) {

            var utterance = args.intent.matched[0];
            var engageMethod = /(SMS|E-Mail|Any Method)/i.exec(utterance);
            var engageType = /\b(Critical Incident|Invite to chat)\b/i.exec(utterance);
            var recipientType = /\b(Directly)\b/i.exec(utterance);

             var contactType = session.dialogData.contactType = {
                utterance: utterance,
                endpoint: "engage",
                engageMethod: engageMethod ? engageMethod[0].toLowerCase() : null,
                engageType: engageType ? engageType[0].toLowerCase() : null,
                target: utterance.split(" ")[1] ? utterance.split(" ")[1] : null,
                recipientType: recipientType ? recipientType[0].toLowerCase()+" " : "",
            };

            //TODO: ensure group exists

            if(contactType.engageType){
                next();
            }else{
                var msg = new builder.Message(session);
                msg.attachments([
                    new builder.HeroCard(session)
                        .title("Engagement Type")
                        .subtitle("Choose the type of engagement")
                        .buttons([
                            builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident", "Critical Incident"),
                            builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Invite to chat", "Invite to chat")
                        ])
                ]);
                session.send(msg).endDialog();
            } 
        },
        function (session, args, next) {
            var contactType = session.dialogData.contactType;
            var utterance = contactType.utterance;

            if(contactType.engageType == "critical incident"){

                var engagePriority = /(High|Medium|Low)/i.exec(utterance);
                contactType.engagePriority = engagePriority ? engagePriority[0].toLowerCase() : null
                session.dialogData.contactType = contactType;

                if(contactType.engagePriority){
                    next();
                }else{
                    var msg = new builder.Message(session);
                    msg.attachments([
                        new builder.HeroCard(session)
                            .title("Incident Priority")
                            .subtitle("Choose the priority of incident")
                            .buttons([
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with High Priority", "High"),
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with Medium Priority", "Medium"),
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with Low Priority", "Low")
                            ])
                    ]);
                    session.send(msg).endDialog();
                } 
            }else{
                next();
            }
        },
        function (session, results) {
            var contactType = session.dialogData.contactType;
            contactType.recipientType = contactType.recipientType.trim();

            if(contactType.engageType == "critical incident"){

                var args = {
                    data: contactType,
                    headers: { "Content-Type": "application/json" }
                };

                xmatters.xmattersInstance.post(xmattersConfig.url + "/api/integration/1/functions/"+xmattersConfig.integrationCodes.engage_incident+"/triggers", args, function (data, response) {
                    session.send(contactType.target+" engaged").endDialog();
                });

            }else{
                savedAddress = session.message.address;
                savedSession = session;
                var direct = true;
                if(contactType.recipientType == ""){
                    direct = false;
                }
                
                engage(contactType.target,session,direct);
            }
        }
    ]).triggerAction({ matches: /(Engage)\s(.*).*/i });

function engage(targets,session,direct){
        console.log("engage");

        if(!direct){
            xmatters.groupsExists(targets, savedSession, bot, builder, function(newTargets,invalidGroups){

                var args = {
                    headers: { "Content-Type": "application/json" },
                    parameters: { text: newTargets}, // this is serialized as URL parameters
                    data: { text: newTargets }
                };
                client.post(xmattersConfig.url+"/api/integration/1/functions/8a347908-ceb4-4a79-a12b-5a34a476823d/triggers", args, function (data, response) {
                    if(!!data.requestId){
                        postToChannel(session,newTargets + " has been invited to the channel");
                    }
                });
            });
        }else{
            var args = {
                headers: { "Content-Type": "application/json" },
                parameters: { text: targets}, // this is serialized as URL parameters
                data: { text: targets }
            };
            client.post(xmattersConfig.url+"/api/integration/1/functions/8a347908-ceb4-4a79-a12b-5a34a476823d/triggers", args, function (data, response) {
                if(!!data.requestId){
                    postToChannel(session,targets + " has been invited to the channel");
                }
            });
        }
    }

    // Inbound functions
    function respond(req, res, next) {
        var msg = new builder.Message().address(savedAddress);
        msg.text(req.body.text);
        msg.textLocale('en-US');
        bot.send(msg);

        next();
    }

    // Export the connector for any downstream integration - e.g. registering a messaging extension
    module.exports.connector = connector;
};





        
       
