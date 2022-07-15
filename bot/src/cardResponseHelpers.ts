// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { MessageFactory, InputHints, MessagingExtensionActionResponse, AdaptiveCardInvokeResponse } from 'botbuilder';

export class CardResponseHelpers {
  // https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/task-modules/task-modules-bots?tabs=nodejs#invoke-a-task-module-using-taskfetch
  // https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/task-modules/invoking-task-modules#the-taskinfo-object
  static toTaskModuleResponse(
    cardAttachment
  ): MessagingExtensionActionResponse {
    return {
      task: {
        type: "continue",
        value: {
          card: cardAttachment,
          height: 450,
          title: "Task Module Fetch Example",
          url: null,
          width: 500,
        },
      },
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
        activityPreview: MessageFactory.attachment(
          cardAttachment,
          null,
          null,
          InputHints.ExpectingInput
        ),
        type: "botMessagePreview",
      },
    };
  }

  // https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/universal-action-model#response-format
  static toBotInvokeRefreshResponse(card: Record<string, unknown>): AdaptiveCardInvokeResponse {
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: card,
    };
  }

  static toBotInvokeMessageResponse(message: string): AdaptiveCardInvokeResponse {
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.activity.message",
      value: message as any,
    };
  }
}