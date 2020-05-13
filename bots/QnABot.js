// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');

class QnABot extends ActivityHandler {
    constructor() {
        super();

        try {
            this.qnaMaker = new QnAMaker({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
                endpointKey: process.env.QnAEndpointKey,
                host: process.env.QnAEndpointHostName
            });
        } catch (err) {
            console.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
        }

        // If a new user is added to the conversation, send them a greeting message
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity('¡Bienvenido a la muestra de QnA Maker! Hazme una pregunta y trataré de responderla.');
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        // When a user sends a message, perform a call to the QnA Maker service to retrieve matching Question and Answer pairs.
        this.onMessage(async (context, next) => {
            if (!process.env.QnAKnowledgebaseId || !process.env.QnAEndpointKey || !process.env.QnAEndpointHostName) {
                let unconfiguredQnaMessage = 'NOTE: \r\n' + 
                    'QnA Maker no está configurado. Para habilitar todas las capacidades, agregue `QnAKnowledgebaseId`, `QnAEndpointKey` and `QnAEndpointHostName` to the .env file. \r\n' +
                    'You may visit www.qnamaker.ai to create a QnA Maker knowledge base.'

                 await context.sendActivity(unconfiguredQnaMessage)
            }
            else {
                console.log('Calling QnA Maker');
    
                const qnaResults = await this.qnaMaker.getAnswers(context);
    
                // Si se recibió una respuesta de QnA Maker, envíe la respuesta al usuario.
                if (qnaResults[0]) {
                    await context.sendActivity(qnaResults[0].answer);
    
                // If no answers were returned from QnA Maker, reply with help.
                } else {
                    await context.sendActivity('No se encontraron respuestas de QnA Maker.');
                }
    
            }
            
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.QnABot = QnABot;
