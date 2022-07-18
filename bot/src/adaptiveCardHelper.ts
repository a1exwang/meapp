// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { CardFactory } from "botbuilder";
import taskModuleComposeCardBasicInfo from "../adaptiveCards/taskModuleComposeCardBasicInfo.json";
import searchCard from "../adaptiveCards/search/search.json";
import searchResultsCard from "../adaptiveCards/search/searchResults.json";
import searchResultCard from "../adaptiveCards/search/searchResult.json";
import approvalBaseCard from "../adaptiveCards/approval/approvalBase.json";
import approvalForSender from "../adaptiveCards/approval/approvalForSender.json";
import approvalForApprover from "../adaptiveCards/approval/approvalForApprover.json";
import approvalApproved from "../adaptiveCards/approval/approvalApproved.json";
import approvalRejected from "../adaptiveCards/approval/approvalRejected.json";

export enum CardID {
  ApprovalBase = "approvalBase",
  ApprovalForApprover = "approvalForApprover",
  ApprovalForSender = "approvalForSender",
  BotSearchCard = "BotSearchCard",
  BotSearchResultsCard = "BotSearchResultsCard",
  BotSearchResultCard = "BotSearchResultCard",
}

export class AdaptiveCardHelper {
  static toSubmitExampleData(action) {
    const activityPreview = action.botActivityPreview[0];
    const attachmentContent = activityPreview.attachments[0].content;
    const userText = attachmentContent.body[1].text;
    const choiceSet = attachmentContent.body[3];
    const attributionFlag = attachmentContent.body[4].text.split(":")[1];
    return {
      MultiSelect: choiceSet.isMultiSelect ? "true" : "false",
      UserAttributionSelect: attributionFlag,
      Option1: choiceSet.choices[0].title,
      Option2: choiceSet.choices[1].title,
      Option3: choiceSet.choices[2].title,
      Question: userText,
    };
  }

  static createTaskModuleComposeCardBasicInfo() {
    return taskModuleComposeCardBasicInfo;  
  }

  static createTaskModuleComposeCardApprovers(
    title: string,
    description: string,
    approvers?: string[]
  ) {
    const approverPart = approvers
      ? {
          id: "approvers",
          type: "Input.ChoiceSet",
          isRequired: true,
          style: "expanded",
          isMultiSelect: true,
          choices: approvers.map((email) => {
            return { title: email, value: email };
          }),
        }
      : {
          id: "approvers",
          type: "Input.Text",
          spacing: "None",
          placeholder: "Input your approver email",
        };
    return CardFactory.adaptiveCard({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
      type: "AdaptiveCard",
      actions: [
        {
          data: {
            id: "submit",
            cardId: "taskModuleComposeCardApprovers",
            // filled from last step
            title: title,
            description: description,
          },
          title: "Confirm",
          type: "Action.Submit",
        },
      ],
      body: [
        {
          type: "TextBlock",
          text: "Step 2/2: Please enter the approvers:",
        },
        {
          type: "TextBlock",
          text: "Approver List (*)",
        },
        approverPart,
      ],
    });
  }

  static createTaskModuleComposeCardApproval(
    from: string,
    title: string,
    description: string,
    approvers: string[],
    userIds: string[]
  ) {
    return CardFactory.adaptiveCard({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
      type: "AdaptiveCard",
      refresh: {
        action: {
          type: "Action.Execute",
          title: "Refresh",
          verb: "refresh",
          data: {
            from: from,
            title: title,
            description: description,
            approvers: approvers,
          },
        },
        userIds: userIds,
      },
      actions: [
        {
          type: "Action.Submit",
          title: "Confirm Approval Request",
          data: {
            id: "submit",
            cardId: "taskModuleComposeCardApproval",
            title: title,
            description: description,
            approvers: approvers,
          },
        },
      ],
      body: [
        {
          type: "TextBlock",
          text: "Please take a final review of your approval request",
        },
        {
          type: "TextBlock",
          text: "From " + from,
        },
        {
          type: "TextBlock",
          text: "Title",
        },
        {
          id: "title",
          type: "Input.Text",
          isRequired: true,
          errorMessage: "Title is required",
          placeholder: "Input your approval request title",
          value: title,
        },
        {
          type: "TextBlock",
          text: "Description",
        },
        {
          id: "description",
          type: "Input.Text",
          isMultiline: true,
          placeholder: "Input your approval request description",
          value: description,
        },
        {
          type: "TextBlock",
          text: "Approvers",
        },
        {
          type: "FactSet",
          facts: approvers.map((email, index) => {
            return { title: index + 1, value: email };
          }),
        },
      ],
    });
  }

  static createAdaptiveCardEditor(
    userText = null,
    isMultiSelect = true,
    option1 = null,
    option2 = null,
    option3 = null
  ) {
    return CardFactory.adaptiveCard({
      actions: [
        {
          data: {
            submitLocation: "messagingExtensionFetchTask",
            id: "editor",
          },
          title: "Submit",
          type: "Action.Submit",
        },
      ],
      body: [
        {
          text: "This is an Adaptive Card within a Task Module",
          type: "TextBlock",
          weight: "bolder",
        },
        { type: "TextBlock", text: "Enter text for Question:" },
        {
          id: "Question",
          placeholder: "Question text here",
          type: "Input.Text",
          value: userText,
        },
        { type: "TextBlock", text: "Options for Question:" },
        { type: "TextBlock", text: "Is Multi-Select:" },
        {
          choices: [
            { title: "True", value: "true" },
            { title: "False", value: "false" },
          ],
          id: "MultiSelect",
          isMultiSelect: false,
          style: "expanded",
          type: "Input.ChoiceSet",
          value: isMultiSelect ? "true" : "false",
        },
        {
          id: "Option1",
          placeholder: "Option 1 here",
          type: "Input.Text",
          value: option1,
        },
        {
          id: "Option2",
          placeholder: "Option 2 here",
          type: "Input.Text",
          value: option2,
        },
        {
          id: "Option3",
          placeholder: "Option 3 here",
          type: "Input.Text",
          value: option3,
        },
        {
          type: "TextBlock",
          text: "Do you want to send this card on behalf of the User?",
        },
        {
          choices: [
            { title: "Yes", value: "true" },
            { title: "No", value: "false" },
          ],
          id: "UserAttributionSelect",
          isMultiSelect: false,
          style: "expanded",
          type: "Input.ChoiceSet",
          value: isMultiSelect ? "true" : "false",
        },
      ],
      type: "AdaptiveCard",
      version: "1.0",
    });
  }

  static createAdaptiveCardAttachment(data) {
    return CardFactory.adaptiveCard({
      actions: [
        {
          type: "Action.Submit",
          title: "Submit",
          data: { submitLocation: "messagingExtensionSubmit" },
        },
      ],
      body: [
        {
          text: "Adaptive Card from Task Module",
          type: "TextBlock",
          weight: "bolder",
        },
        { text: `${data.Question}`, type: "TextBlock", id: "Question" },
        { id: "Answer", placeholder: "Answer here...", type: "Input.Text" },
        {
          choices: [
            { title: data.Option1, value: data.Option1 },
            { title: data.Option2, value: data.Option2 },
            { title: data.Option3, value: data.Option3 },
          ],
          id: "Choices",
          isMultiSelect: data.MultiSelect,
          style: "expanded",
          type: "Input.ChoiceSet",
        },
        {
          text: `Sending card on behalf of user is set to: ${data.UserAttributionSelect}`,
          type: "TextBlock",
          id: "AttributionChoice",
        },
      ],
      type: "AdaptiveCard",
      version: "1.0",
    });
  }

  static createBotUserSpecificViewCardApprovalBase(data: {
    from: string;
    title: string;
    description: string;
    approvers: string[];
    approverComments: string[];
    userIds: string[];
  }) {
    return AdaptiveCards.declare(approvalBaseCard).render(data);
  }

  static createBotUserSpecificViewCardApprovalForSender(data: {
    from: string;
    title: string;
    description: string;
    approvers: string[];
    approverComments: { email: string; comment: string }[];
    refresh: boolean;
    userIds: string[];
  }) {
    return AdaptiveCards.declare(approvalForSender).render(data);
  }

  static createBotUserSpecificViewCardApprovalForApprover(data: {
    from: string;
    title: string;
    description: string;
    approvers: string[];
    approverComments: { email: string; comment: string }[];
    userIds: string[];
  }) {
    return AdaptiveCards.declare(approvalForApprover).render(data);
  }

  static createBotUserSpecificViewCardApprovalCanceled(data: {
    from: string;
    title: string;
    description: string;
    approvers: string[];
  }) {
    // transform data
    const cardData = JSON.parse(JSON.stringify(data));
    cardData["approvers"] = data.approvers.map((item, index) => {
      return { title: `${index + 1}`, value: item };
    });

    // render card
    const template = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: "Cancelled: Approval request from ${from}",
        },
        {
          type: "TextBlock",
          text: "Title",
        },
        {
          type: "TextBlock",
          text: "${title}",
        },
        {
          type: "TextBlock",
          text: "Description",
        },
        {
          type: "TextBlock",
          text: "${description}",
        },
        {
          type: "TextBlock",
          text: "Approvers",
        },
        {
          type: "FactSet",
          facts: [
            {
              $data: "${approvers}",
              title: "${title}",
              value: "${value}",
            },
          ],
        },
      ],
    };
    return AdaptiveCards.declare(template).render(cardData);
  }

  static createBotUserSpecificViewCardApprovalApproved(data: {
    from: string;
    title: string;
    description: string;
    approverComments: { email: string; comment: string }[];
  }) {
    return AdaptiveCards.declare(approvalApproved).render(data);
  }

  static createBotUserSpecificViewCardApprovalRejected(data: {
    from: string;
    title: string;
    description: string;
    rejectedBy: string;
    comment: string;
  }) {
    return AdaptiveCards.declare(approvalRejected).render(data);
  }

  static createBotSearchCard(data: {manufacturers: {name: string}[]}) {
    return AdaptiveCards.declare(searchCard).render({
      ...data,
      cardId: CardID.BotSearchCard,
    });
  }

  static createBotSearchResultsCard(data: {results: {id: string, name: string, description: string}[]}) {
    return AdaptiveCards.declare(searchResultsCard).render({
      ...data,
      cardId: CardID.BotSearchResultsCard,
    });
  }

  static createBotSearchResultCard(data: {id: string, name: string, description: string}) {
    return AdaptiveCards.declare(searchResultCard).render({
      ...data,
      cardId: CardID.BotSearchResultCard,
    });
  }

}
