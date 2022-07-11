// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { MessageFactory, InputHints, MessagingExtensionActionResponse } from 'botbuilder';

export class CardResponseHelpers {
    static toTaskModuleResponse(cardAttachment): MessagingExtensionActionResponse {
        return {
            task: {
                type: 'continue',
                value: {
                    card: cardAttachment,
                    height: 450,
                    title: 'Task Module Fetch Example',
                    url: null,
                    width: 500
                }
            }
        };
    }

    static toComposeExtensionResultResponse(cardAttachment) {
        return {
          composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: [cardAttachment],
          },
        };
    }

    static toMessagingExtensionBotMessagePreviewResponse(cardAttachment) {
        return {
            composeExtension: {
                activityPreview: MessageFactory.attachment(cardAttachment, null, null, InputHints.ExpectingInput),
                type: 'botMessagePreview'
            }
        };
    }
}