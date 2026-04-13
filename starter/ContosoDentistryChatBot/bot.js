// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// @ts-check

const { ActivityHandler, MessageFactory } = require('botbuilder');

// CLU and CQA are called via direct REST API to ensure
// we always use the latest supported API version.

class DentistBot extends ActivityHandler {
    constructor() {
        super();

        // --- CLU Setup (replaces LUIS) ---
        this.cluEndpoint = process.env.LANGUAGE_ENDPOINT;
        this.cluKey = process.env.LANGUAGE_KEY;
        this.cluProjectName = process.env.CLU_PROJECT_NAME;
        this.cluDeploymentName = process.env.CLU_DEPLOYMENT_NAME;

        // --- CQA Setup (replaces QnA Maker) ---
        this.cqaEndpoint = process.env.LANGUAGE_ENDPOINT;
        this.cqaKey = process.env.LANGUAGE_KEY;
        this.cqaProjectName = process.env.CQA_PROJECT_NAME;
        this.cqaDeploymentName = process.env.CQA_DEPLOYMENT_NAME || 'production';

        this.onMessage(async (context, next) => {
            const userMessage = context.activity.text?.trim();

            if (!userMessage) {
                await next();
                return;
            }

            try {
                // Step 1: Try CLU first to detect scheduling intents
                const cluResult = await this.callCLU(userMessage);
                const prediction = cluResult?.result?.prediction;
                const topIntent = prediction?.topIntent;
                const confidence = prediction?.intents?.find(
                    i => i.category === topIntent
                )?.confidenceScore ?? 0;

                if (topIntent === 'ScheduleAppointment' && confidence > 0.85) {
                    const entities = prediction?.entities ?? [];
                    const datetimeEntity = entities.find(e => e.category === 'datetime');
                    const when = datetimeEntity?.text ?? 'a time of your choosing';

                    await context.sendActivity(
                        MessageFactory.text(
                            `I can help you schedule an appointment for ${when}. ` +
                            `Please call our office or use our online portal to confirm your booking.`
                        )
                    );

                } else if (topIntent === 'GetAvailability' && confidence > 0.85) {
                    const entities = prediction?.entities ?? [];
                    const datetimeEntity = entities.find(e => e.category === 'datetime');
                    const when = datetimeEntity?.text ?? 'the requested time';

                    await context.sendActivity(
                        MessageFactory.text(
                            `Let me check availability for ${when}. ` +
                            `Please contact our office for real-time availability.`
                        )
                    );

                } else {
                    // Step 2: Fall back to CQA for FAQ-style questions
                    const cqaAnswer = await this.callCQA(userMessage);

                    if (cqaAnswer && cqaAnswer.confidence > 0.5) {
                        await context.sendActivity(
                            MessageFactory.text(cqaAnswer.answer)
                        );
                    } else {
                        await context.sendActivity(
                            MessageFactory.text(
                                "I'm sorry, I didn't understand that. " +
                                "I can help you schedule appointments, check availability, " +
                                "or answer questions about our dental office. How can I assist you?"
                            )
                        );
                    }
                }

            } catch (error) {
                console.error('Error processing message:', error);
                await context.sendActivity(
                    MessageFactory.text(
                        "I'm having trouble processing your request right now. Please try again."
                    )
                );
            }

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded ?? [];
            const welcomeText =
                'Hello and welcome to Contoso Dentistry! ' +
                'I can help you schedule appointments, check availability, ' +
                'or answer questions about our office. How can I help you today?';

            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            await next();
        });
    }

    /**
     * Call CLU via direct REST API (api-version=2023-04-01)
     * @param {string} utterance
     */
    async callCLU(utterance) {
        const endpoint = this.cluEndpoint.replace(/\/$/, '');
        const url = `${endpoint}/language/:analyze-conversations?api-version=2023-04-01`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Ocp-Apim-Subscription-Key': this.cluKey,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                kind: 'Conversation',
                analysisInput: {
                    conversationItem: {
                        id: '1',
                        participantId: 'user',
                        text: utterance
                    }
                },
                parameters: {
                    projectName: this.cluProjectName,
                    deploymentName: this.cluDeploymentName,
                    stringIndexType: 'TextElement_V8'
                }
            })
        });

        if (!response.ok) {
            const err = await response.text();
            throw new Error(`CLU request failed: ${response.status} ${err}`);
        }

        return response.json();
    }

    /**
     * Call CQA via direct REST API (api-version=2021-10-01)
     * @param {string} question
     */
    async callCQA(question) {
        const endpoint = this.cqaEndpoint.replace(/\/$/, '');
        const url = `${endpoint}/language/:query-knowledgebases` +
            `?projectName=${this.cqaProjectName}` +
            `&deploymentName=${this.cqaDeploymentName}` +
            `&api-version=2021-10-01`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Ocp-Apim-Subscription-Key': this.cqaKey,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                question,
                top: 1,
                confidenceScoreThreshold: 0.3
            })
        });

        if (!response.ok) {
            const err = await response.text();
            throw new Error(`CQA request failed: ${response.status} ${err}`);
        }

        const data = await response.json();
        const topAnswer = data?.answers?.[0];

        if (topAnswer && topAnswer.answer !== 'No good match found in KB.') {
            return {
                answer: topAnswer.answer,
                confidence: topAnswer.confidenceScore
            };
        }

        return null;
    }
}

module.exports.DentistBot = DentistBot;