/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var adal = require('adal-node');
var teams = require('botbuilder-teams');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
var builder_cognitiveservices = require("botbuilder-cognitiveservices");
const MicrosoftGraph = require('@microsoft/microsoft-graph-client');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    // openIdMetadata: process.env.BotOpenIdMetadata
});


var clientID = process.env.clientID;
var appSecret = process.env.AppSecret;
var bot = new builder.UniversalBot(connector, function (session) {
    connector.fetchChannelList(session.message.address.serviceUrl, teamId, function (err, result) {
        if (err) {
            session.endDialog('Error fetching team channel information');
        }
        else {
            var AuthenticationContext = adal.AuthenticationContext;
            var authorityHostUrl = 'https://login.windows.net';
            var tenant = 'http://sep007.sharepoint.com/';

            var authoriotyUrl = authorityHostUrl + '/' + tenant;
            var applicationId = clientID;
            var clientSecret = appSecret;
            const resource = "https://graph.microsoft.com";
            var messageText = session.message.text;
            messageText = messageText.substring(mentionString.length);
            var context = new AuthenticationContext(authoriotyUrl);
            context.acquireTokenWithClientCredentials(
                resource,
                applicationId,
                clientSecret,
                function (err, tokenResponse) {
                    if (err) {
                        console.log('well that did not work: ' + err.stack);
                    } else {
                        let client = MicrosoftGraph.Client.init({
                            authProvider: (done) => {
                                done(null, tokenResponse.accessToken);
                            }
                        });
                        client.api('https://graph.microsoft.com/beta/sites/sep007.sharepoint.com,052e0c8e-9e92-4165-8d9a-c4f405aaf2d8/lists/test/items/')
                            .version("beta")
                            .header("Content-type", "application/json")
                            .post({
                                "fields": {
                                    "ContentType": "Item",
                                    "Title": messageText
                                }
                            }).then((res) => {
                                session.send("your message has been posted to SharePoint");
                            }).catch((err) => {
                                console.log(err);
                                session.send("Oops ! error ocured");
                            });
                    }
                });
        }
    });
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());


