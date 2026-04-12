// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

const { QnAMaker } = require('botbuilder-ai');
const DentistScheduler = require('./dentistscheduler');
const IntentRecognizer = require("./intentrecognizer")

class DentistBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        // call the parent constructor
        super();
        if (!configuration) throw new Error('[QnaMakerBot]: Missing parameter. configuration is required');

        // create a QnAMaker connector
        this.QnAMaker = new QnAMaker(configuration.QnAConfiguration, qnaOptions);
       
        // create a DentistScheduler connector
        this.DentistScheduler = new DentistScheduler(configuration.SchedulerConfiguration);

        // create a IntentRecognizer connector
        this.IntentRecognizer = new IntentRecognizer(configuration.LuisConfiguration);


        this.onMessage(async (context, next) => {
            // send user input to QnA Maker and collect the response in a variable
            // don't forget to use the 'await' keyword
            const qnaResults = await this.QnAMaker.getAnswers(context);

            // send user input to IntentRecognizer and collect the response in a variable
            // don't forget 'await'
            const luisResult = this.IntentRecognizer.isConfigured
                ? await this.IntentRecognizer.executeLuisQuery(context)
                : undefined;

            // determine which service to respond with based on the results from LUIS //

            // if(top intent is intentA and confidence greater than 50){
            //  doSomething();
            //  await context.sendActivity();
            //  await next();
            //  return;
            // }
            // else {...}
            const topIntent = luisResult?.prediction?.topIntent;
            const intentScore = topIntent && luisResult?.prediction?.intents
                ? luisResult.prediction.intents[topIntent]?.score || 0
                : 0;

            if (topIntent && intentScore >= 0.5) {
                if (topIntent === 'GetAvailability') {
                    const availabilityText = await this.DentistScheduler.getAvailability();
                    await context.sendActivity(MessageFactory.text(availabilityText, availabilityText));
                    await next();
                    return;
                }

                if (topIntent === 'ScheduleAppointment') {
                    const requestedTime = this.IntentRecognizer.getTimeEntity(luisResult);
                    if (!requestedTime) {
                        const promptText = 'What time would you like to schedule your appointment?';
                        await context.sendActivity(MessageFactory.text(promptText, promptText));
                        await next();
                        return;
                    }

                    const scheduleText = await this.DentistScheduler.scheduleAppointment(requestedTime);
                    await context.sendActivity(MessageFactory.text(scheduleText, scheduleText));
                    await next();
                    return;
                }
            }

            if (qnaResults && qnaResults.length > 0) {
                await context.sendActivity(MessageFactory.text(qnaResults[0].answer));
                await next();
                return;
            }

            const fallbackText = 'Sorry, I did not understand that. Can you rephrase?';
            await context.sendActivity(MessageFactory.text(fallbackText, fallbackText));
            await next();
    });

        this.onMembersAdded(async (context, next) => {
        const membersAdded = context.activity.membersAdded;
        //write a custom greeting
        const welcomeText = 'Welcome to Contoso Dentistry. How can I help you today?';
        for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
            if (membersAdded[cnt].id !== context.activity.recipient.id) {
                await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
            }
        }
        // by calling next() you ensure that the next BotHandler is run.
        await next();
    });
    }
}

module.exports.DentistBot = DentistBot;
