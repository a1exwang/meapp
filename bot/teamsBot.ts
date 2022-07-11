import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  BotHandler,
  Activity,
  TeamsInfo,
  ChannelAccount,
  TeamsChannelAccount,
  Attachment,
  MessageFactory,
  ConversationReference,
  ConversationParameters,
} from "botbuilder";
import { ConnectorClient } from "botframework-connector";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import rawLearnCard from "./adaptiveCards/learn.json";
import rawTaskCard from "./adaptiveCards/task.json";
import rawTaskResponseCard from "./adaptiveCards/taskResponse.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { AdaptiveCardHelper } from "./adaptiveCardHelper";
import { CardResponseHelpers } from "./cardResponseHelpers";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      // TODO: support command bot trigger
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }
      await next();
    });

    this.onTurn(async (context: TurnContext, next) => {
      // Debug log activity info
      console.log(JSON.stringify(context.activity, null, 2));
      await next();
    });
  }

  // Message Extension handlers

  // Message Extension task/fetch
  // Called when user click message extension icon and select a command with fetchTask == true.
  // The returned adaptive card will be rendered in the task module.
  async handleTeamsMessagingExtensionFetchTask(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    if (action.commandId === "taskModuleCompose") {
      const adaptiveCard = AdaptiveCardHelper.createTaskModuleComposeCardBasicInfo();
      return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
    } /* if (action.commandId === "taskModuleBot") */ else {
      // In order to use the bot to send message, the bot needs to be in the team
      if (context.activity.conversation.conversationType === "channel") {
        let members: TeamsChannelAccount[] = [];
        try {
          members = await TeamsBot.getTeamMembers(context);
        } catch (e) {
          // if failed, assuming the bot is not added to the team
          const adaptiveCard = CardFactory.adaptiveCard({
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            version: "1.4",
            type: "AdaptiveCard",
            body: [
              {
                type: "TextBlock",
                text: "Looks like you haven't used Disco in this team/chat",
              },
            ],
            actions: [
              {
                type: "Action.Submit",
                title: "Continue",
                data: {
                  msteams: {
                    justInTimeInstall: true,
                  },
                },
              },
            ],
          });
          return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
        }
        const adaptiveCard = AdaptiveCardHelper.createTaskModuleComposeCardBasicInfo();
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
      } else {
        // assume personal chat
        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor();
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
      }
    }
  }


  // Called when user click submit in task module adaptive card.
  // The return value can be:
  //    1. TaskInfo object with type == 'continue': multi-step task module.
  //    2. ComposeExtension object with type == 'botMessagePreview': Task module adaptive card preview/editor
  //    3. ComposeExtension object with type == 'result': Insert the returned adaptive card into compose area.
  //    4. ComposeExtension object with type == 'auth': Auth
  //    5. Empty object. Close the task module. This can be used to let the bot send adaptive card.
  // See https://docs.microsoft.com/en-us/microsoftteams/platform/resources/messaging-extension-v3/create-extensions?tabs=typescript#responding-to-submit
  public async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: any
  ): Promise<any> {
    switch (action.commandId) {
      case "taskModuleCompose":
      case "taskModuleBot":
        const data: {cardId: string, title: string, description: string} = action.data;
        if (data.cardId === "taskModuleComposeCardBasicInfo") {
          let adaptiveCard: Attachment;
          try {
            const members = await TeamsBot.getTeamMembers(context);
            adaptiveCard = AdaptiveCardHelper.createTaskModuleComposeCardApprovers(data.title, data.description, members.map((member) => member.email));
          } catch (e) {
            adaptiveCard = AdaptiveCardHelper.createTaskModuleComposeCardApprovers(data.title, data.description);
          }
          return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
        } else {
          // Send Approver list:
          //  For task module compose card, just return the card to insert to compose area
          //  For bot, let bot send the card

          // adaptive card ChoiceSet result is separated by ','
          const approvers = action.data.approvers.split(",");

          // Only approver and sender should refresh
          const refreshUserIds: string[] = (await TeamsBot.getTeamMembers(context))
            .filter(
              (item) =>
                approvers.indexOf(item.email) !== -1 ||
                item.aadObjectId === context.activity.from.aadObjectId
            )
            .map((item) => item.id);

          const adaptiveCard =
            AdaptiveCardHelper.createTaskModuleComposeCardApproval(
              context.activity.from.name,
              action.data.title,
              action.data.description,
              approvers,
              action.commandId === "taskModuleBot",
              refreshUserIds,
            );
          if (action.commandId === "taskModuleCompose") {
            return CardResponseHelpers.toComposeExtensionResultResponse(adaptiveCard);
          } else {
            await context.sendActivity({
              attachments: [adaptiveCard],
            });
            return {};
          }
        }
      case "staticParameters":
        {
          const data = action.data;
          const heroCard = CardFactory.heroCard(data.title, data.text);
          heroCard.content.subtitle = data.subTitle;
          const attachment = {
            contentType: heroCard.contentType,
            content: heroCard.content,
            preview: heroCard,
          };
          return CardResponseHelpers.toComposeExtensionResultResponse(attachment)
        }
      case "taskModule":
        if (action.data.id === "editor") {
          const submittedData = action.data;
          const adaptiveCard =
            AdaptiveCardHelper.createAdaptiveCardAttachment(submittedData);
          return CardResponseHelpers.toMessagingExtensionBotMessagePreviewResponse(
            adaptiveCard
          );
        } else {
        }
      default:
        throw new Error("NotImplemented");
    }
  }

  // TODO: support preview/edit
  async handleTeamsMessagingExtensionBotMessagePreviewEdit(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    // The data has been returned to the bot in the action structure.
    const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

    // This is a preview edit call and so this time we want to re-create the adaptive card editor.
    const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor(
      submitData.Question,
      submitData.MultiSelect.toLowerCase() === "true",
      submitData.Option1,
      submitData.Option2,
      submitData.Option3
    );

    return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
  }

  // TODO: support preview/edit
  async handleTeamsMessagingExtensionBotMessagePreviewSend(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    // The data has been returned to the bot in the action structure.
    const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

    // This is a send so we are done and we will create the adaptive card editor.
    const adaptiveCard =
      AdaptiveCardHelper.createAdaptiveCardAttachment(submitData);
    var responseActivity: Partial<Activity> = { type: "message", attachments: [adaptiveCard] };
    if (submitData.UserAttributionSelect === "true") {
      responseActivity = {
        type: "message",
        attachments: [adaptiveCard],
        channelData: {
          onBehalfOf: [
            {
              itemId: 0,
              mentionType: "person",
              mri: context.activity.from.id,
              displayName: context.activity.from.name,
            },
          ],
        },
      };
    }
    await context.sendActivity(responseActivity);

    return undefined;
  }

  async handleTeamsMessagingExtensionCardButtonClicked(
    context: TurnContext,
    cardData: any
  ): Promise<void> {
    const members = await TeamsBot.getTeamMembers(context);
    const data: { title: string; description: string; approvers: string[] } =
      cardData;
    const deeplink = TeamsBot.generateDeeplink(context);
    for (const member of members) {
      if (data.approvers.indexOf(member.email) !== -1) {
        const conversationReference = await this.newConversation(context, member);
        // do not await to prevent adaptive card timeout
        context.adapter.continueConversation(
          conversationReference,
          async (personalContext: TurnContext) => {
            console.log(`Sending to ${member.email}`);
            const activity = MessageFactory.text(
              `Please approve "${data.title}". Click here for details: ${deeplink}`
            );
            await personalContext.sendActivity(activity);
          }
        );
      }
    }

    await context.sendActivity("Approvers are notified.")
  }

  // Bot handlers
  // Bot adaptive card invoke
  // TODO: support adaptive card refresh
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    return { statusCode: 200, type: undefined, value: undefined };
  }

  // Utilities
  private static async getTeamMembers(context: TurnContext): Promise<TeamsChannelAccount[]> {
    const members: TeamsChannelAccount[] = [];
    let continuationToken = undefined;
    do {
      var pagedMembers = await TeamsInfo.getPagedMembers(
        context,
        100,
        continuationToken
      );
      continuationToken = pagedMembers.continuationToken;
      members.push(...pagedMembers.members);
    } while (continuationToken !== undefined);
    return members;
  }

  // See https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links?tabs=teamsjs-v2#generate-deep-links-to-channel-conversation
  private static generateDeeplink(context: TurnContext): string {
    const parentMessageId = context.activity.replyToId;
    const deeplink = `https://teams.microsoft.com/l/message/${context.activity.channelData.channel.id}/${parentMessageId}?tenantId=${context.activity.conversation.tenantId}&parentMessageId=${parentMessageId}`;
    return deeplink;
  }

  // Helper method to send notification
  private async newConversation(context: TurnContext, user: ChannelAccount): Promise<ConversationReference> {
    const reference = TurnContext.getConversationReference(context.activity);
    const personalConversation = JSON.parse(JSON.stringify(reference));

    const connectorClient: ConnectorClient = context.turnState.get(
      context.adapter.ConnectorClientKey
    );
    const conversation = await connectorClient.conversations.createConversation(
      {
        isGroup: false,
        tenantId: context.activity.conversation.tenantId,
        bot: context.activity.recipient,
        members: [user],
        channelData: {},
      } as ConversationParameters
    );
    personalConversation.conversation.id = conversation.id;

    return personalConversation;
  }

}
