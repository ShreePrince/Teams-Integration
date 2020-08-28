// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, MessageFactory } = require('botbuilder');

var querystring = require('querystring');
var https = require('https');


class TeamsMessagingExtensionsActionBot extends TeamsActivityHandler {
    handleTeamsMessagingExtensionSubmitAction(context, action) {
        console.log("This is the Command ID:" + action.commandId);
        switch (action.commandId) {
            case 'createInc':
                return createCardCommand(context, action);
            case 'createCard':
                return EXECUTERESTTOSNOW(context, action);
            default:
                throw new Error('NotImplemented');
        }
    }
}

function createCardCommand(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const heroCard = CardFactory.heroCard(data.title, data.text);
    heroCard.content.subtitle = data.subTitle;
    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };
    //Execute REST

    var host = 'https://dev63884.service-now.com';
    var username = 'admin';
    var password = 'Prince@1995';
    // var apiKey = '*****';
     var sessionId = null;
    // var deckId = '68DC5A20-EE4F-11E2-A00C-0858C0D5C2ED';

    performRequest('/api/now/table/incident', 'POST', {
        username: username,
        password: password,
    }, function (data) {
        console.log('Result:', data.result);
    });



    function performRequest(endpoint, method, data, success) {
        var dataString = JSON.stringify(data);
        var headers = {};

        if (method == 'GET') {
            endpoint += '?' + querystring.stringify(data);
        }
        else {
            headers = {
                'Content-Type': 'application/json',
                'Content-Length': dataString.length
            };
        }
        var options = {
            host: host,
            path: endpoint,
            method: method,
            headers: headers
        };

        var req = https.request(options, function (res) {
            res.setEncoding('utf-8');

            var responseString = '';

            res.on('data', function (data) {
                responseString += data;
            });

            res.on('end', function () {
                console.log(responseString);
                var responseObject = JSON.parse(responseString);
                success(responseObject);
            });
        });

        req.write(dataString);
        req.end();
    }
    //return data
    return {
        composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [
                attachment
            ]
        }
    };
}

function EXECUTERESTTOSNOW(context, action) {

    //Fetch data
    const data = action.data;
    console.log(data);

    // The user has chosen to share a message by choosing the 'Share Message' context menu command.
    let userName = 'unknown';
    if (action.messagePayload.from &&
        action.messagePayload.from.user &&
        action.messagePayload.from.user.displayName) {
        userName = action.messagePayload.from.user.displayName;
    }



    // This Messaging Extension example allows the user to check a box to include an image with the
    // shared message.  This demonstrates sending custom parameters along with the message payload.


    let images = [];
    const includeImage = action.data.includeImage;
    if (includeImage === 'true') {
        images = ['https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU'];
    }
    const heroCard = CardFactory.heroCard(`${userName} originally sent this message:`,
        action.messagePayload.body.content,
        images);

    if (action.messagePayload.attachments && action.messagePayload.attachments.length > 0) {
        // This sample does not add the MessagePayload Attachments.  This is left as an
        // exercise for the user.
        heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
    }

    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

    return {
        composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [
                attachment
            ]
        }
    };
}

module.exports.TeamsMessagingExtensionsActionBot = TeamsMessagingExtensionsActionBot;
