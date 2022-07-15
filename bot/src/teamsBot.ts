import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  Activity,
  TeamsChannelAccount,
  Attachment,
  MessageFactory,
  ConversationState,
  UserState,
  StatePropertyAccessor,
} from "botbuilder";
import rawWelcomeCard from "../adaptiveCards/welcome.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { AdaptiveCardHelper, CardID } from "./adaptiveCardHelper";
import { CardResponseHelpers } from "./cardResponseHelpers";
import {
  getTeamMembers,
  getUserInfoFromAadObjectId,
  newConversation,
  generateDeeplink,
  deepClone,
} from "./utils";

interface WorkflowInfo {
  workflowState: string;
  userStates: {};
}

interface UserProfile {}

export class TeamsBot extends TeamsActivityHandler {
  private static readonly ConverstationStateWorkflowInfo = "workflowInfo";
  private static readonly UserStateProfile = "userProfile";

  private conversationState: ConversationState;
  private userState: UserState;
  private conversationStateAccessor: StatePropertyAccessor<WorkflowInfo>;
  private userStateAccessor: StatePropertyAccessor<UserProfile>;

  constructor(conversationState: ConversationState, userState: UserState) {
    super();
    this.conversationState = conversationState;
    this.userState = userState;
    this.conversationStateAccessor =
      conversationState.createProperty<WorkflowInfo>(
        TeamsBot.ConverstationStateWorkflowInfo
      );
    this.userStateAccessor = userState.createProperty<UserState>(
      TeamsBot.UserStateProfile
    );

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

      if (txt === "welcome") {
        const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
        await context.sendActivity({
          attachments: [CardFactory.adaptiveCard(card)],
        });
      } else if (txt.match(/find\s+part/i)) {
        // Developers can query custom data to get search options.
        const manufacturers = [
          { name: "Bosch" },
          { name: "Microsoft" },
          { name: "Others" },
        ];

        await context.sendActivity({
          attachments: [
            CardFactory.adaptiveCard(
              AdaptiveCardHelper.createBotSearchCard({
                manufacturers,
              })
            ),
          ],
        });
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
      const adaptiveCard =
        AdaptiveCardHelper.createTaskModuleComposeCardBasicInfo();
      return CardResponseHelpers.toTaskModuleResponse(CardFactory.adaptiveCard(adaptiveCard));
    } /* if (action.commandId === "taskModuleBot") */ else {
      // In order to use the bot to send message, the bot needs to be in the team
      if (context.activity.conversation.conversationType === "channel") {
        let members: TeamsChannelAccount[] = [];
        try {
          members = await getTeamMembers(context);
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
        const adaptiveCard =
          AdaptiveCardHelper.createTaskModuleComposeCardBasicInfo();
        return CardResponseHelpers.toTaskModuleResponse(CardFactory.adaptiveCard(adaptiveCard));
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
        const data: { cardId: string; title: string; description: string } =
          action.data;
        if (data.cardId === "taskModuleComposeCardBasicInfo") {
          let adaptiveCard: Attachment;
          try {
            const members = await getTeamMembers(context);
            adaptiveCard =
              AdaptiveCardHelper.createTaskModuleComposeCardApprovers(
                data.title,
                data.description,
                members
                  .filter(
                    (item) =>
                      // cannot self approve
                      item.aadObjectId !== context.activity.from.aadObjectId
                  )
                  .map((member) => member.email)
                  .sort()
              );
          } catch (e) {
            adaptiveCard =
              AdaptiveCardHelper.createTaskModuleComposeCardApprovers(
                data.title,
                data.description
              );
          }
          return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
        } else {
          // Send Approver list:
          //  For task module compose card, just return the card to insert to compose area
          //  For bot, let bot send the card

          // adaptive card ChoiceSet result is separated by ','
          const approvers = action.data.approvers.split(",");

          // Only approver and sender should refresh
          const teamMembers = await getTeamMembers(context);
          const refreshUserIds: string[] = teamMembers
            .filter(
              (item) =>
                approvers.indexOf(item.email) !== -1 ||
                item.aadObjectId === context.activity.from.aadObjectId
            )
            .map((item) => item.id);
          // sender must be in the team
          const sender = teamMembers.filter(
            (item) => item.aadObjectId === context.activity.from.aadObjectId
          )[0].email;

          const card =
            AdaptiveCardHelper.createBotUserSpecificViewCardApprovalBase({
              ...action.data,
              approvers: approvers,
              approverComments: [],
              from: sender,
              userIds: refreshUserIds,
            });
          if (action.commandId === "taskModuleCompose") {
            return CardResponseHelpers.toComposeExtensionResultResponse(
              CardFactory.adaptiveCard(card)
            );
          } else {
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)],
            });
            return {};
          }
        }
      case "staticParameters": {
        const data = action.data;
        const heroCard = CardFactory.heroCard(data.title, data.text);
        heroCard.content.subtitle = data.subTitle;
        const attachment = {
          contentType: heroCard.contentType,
          content: heroCard.content,
          preview: heroCard,
        };
        return CardResponseHelpers.toComposeExtensionResultResponse(attachment);
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
  async handleTeamsMessagingExtensionBotMessagePreviewEdit(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
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
  async handleTeamsMessagingExtensionBotMessagePreviewSend(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    // The data has been returned to the bot in the action structure.
    const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

    // This is a send so we are done and we will create the adaptive card editor.
    const adaptiveCard =
      AdaptiveCardHelper.createAdaptiveCardAttachment(submitData);
    var responseActivity: Partial<Activity> = {
      type: "message",
      attachments: [adaptiveCard],
    };
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
    const members = await getTeamMembers(context);
    const data: { title: string; description: string; approvers: string[] } =
      cardData;
    const deeplink = generateDeeplink(context);
    for (const member of members) {
      if (data.approvers.indexOf(member.email) !== -1) {
        const conversationReference = await newConversation(context, member);
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

    await context.sendActivity("Approvers are notified.");
  }

  // Bot handlers
  // Bot adaptive card invoke
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    const cardId = invokeValue.action.data.cardId;
    const verb = invokeValue.action.verb;
    if (cardId === "approvalBase") {
      if (verb === "refresh") {
        // if (invokeValue.trigger === "automatic")
        const account = await getUserInfoFromAadObjectId(
          context,
          context.activity.from.aadObjectId
        );

        // TODO: support later steps
        let card;
        const cardData = deepClone(invokeValue.action.data);
        // user specific view
        cardData.userIds = [context.activity.from.id];

        if (account.email === invokeValue.action.data.from) {
          // refresh for sender
          // TODO: fix type
          card =
            AdaptiveCardHelper.createBotUserSpecificViewCardApprovalForSender(
              invokeValue.action.data as any
            );
        } else {
          // for approver
          card =
            AdaptiveCardHelper.createBotUserSpecificViewCardApprovalForApprover(
              invokeValue.action.data as any
            );
        }
        return CardResponseHelpers.toBotInvokeRefreshResponse(card);
      } else {
        throw new Error("Base card cannot have action other than refresh");
      }
    } else if (cardId === "approvalForSender") {
      if (verb === "update") {
        const adaptiveCard =
          AdaptiveCardHelper.createBotUserSpecificViewCardApprovalBase(
            invokeValue.action.data as any
          );
        return await this.refreshAllCards(context, adaptiveCard);
      } else if (verb === "cancel") {
        const teamMembers = await getTeamMembers(context);
        const sender = teamMembers.filter(
          (item) => item.aadObjectId === context.activity.from.aadObjectId
        )[0].email;
        const adaptiveCard =
          AdaptiveCardHelper.createBotUserSpecificViewCardApprovalCanceled({
            ...invokeValue.action.data,
            from: sender,
          } as any);
          return await this.refreshAllCards(context, adaptiveCard);
      } else if (verb === "refresh") {
        const adaptiveCard =
          AdaptiveCardHelper.createBotUserSpecificViewCardApprovalForSender(
            invokeValue.action.data as any
          );
        return CardResponseHelpers.toBotInvokeRefreshResponse(adaptiveCard);
      } else {
        throw new Error("Sender card: Unknown verb " + verb);
      }
    } else if (cardId === "approvalForApprover") {
      if (verb === "approve") {
        const teamMembers = await getTeamMembers(context);
        const sender = teamMembers.filter(
          (item) => item.aadObjectId === context.activity.from.aadObjectId
        )[0].email;
        const data: {
          from: string;
          title: string;
          description: string;
          // remaining approvers
          approvers: string[];
          // already approved approvers
          approverComments: { email: string; comment: string }[];
          // current approver comment
          comment: string;
        } = invokeValue.action.data as any;
        // the last approver will complete the request
        let adaptiveCard;
        if (data.approvers.length === 1) {
          const cardData = {
            from: data.from,
            title: data.title,
            description: data.description,
            approverComments: [
              ...data.approverComments,
              { email: sender, comment: data.comment },
            ],
          };
          adaptiveCard =
            AdaptiveCardHelper.createBotUserSpecificViewCardApprovalApproved(
              cardData
            );
        } else {
          const newApprovers = data.approvers.filter(
            (approver) => approver != sender
          );
          // Only approver and sender should refresh
          const teamMembers = await getTeamMembers(context);
          const refreshUserIds: string[] = teamMembers
            .filter(
              (item) =>
                newApprovers.indexOf(item.email) !== -1 ||
                item.aadObjectId === context.activity.from.aadObjectId
            )
            .map((item) => item.id);
          const newData = {
            ...deepClone(data),
            // remove sender from approver list and add to approverComments list
            approverComments: [
              ...data.approverComments,
              { email: sender, comment: data.comment },
            ],
            approvers: newApprovers,
            userIds: refreshUserIds,
          };
          adaptiveCard =
            AdaptiveCardHelper.createBotUserSpecificViewCardApprovalForApprover(
              newData
            );
        }
        return this.refreshAllCards(context, adaptiveCard);
      } else if (verb === "reject") {
        const teamMembers = await getTeamMembers(context);
        const sender = teamMembers.filter(
          (item) => item.aadObjectId === context.activity.from.aadObjectId
        )[0].email;
        const adaptiveCard =
          AdaptiveCardHelper.createBotUserSpecificViewCardApprovalRejected({
            ...invokeValue.action.data,
            rejectedBy: sender,
          } as any);
          return await this.refreshAllCards(context, adaptiveCard);
      } else if (verb === "refresh") {
        const adaptiveCard =
          AdaptiveCardHelper.createBotUserSpecificViewCardApprovalForApprover(
            invokeValue.action.data as any
          );
        return CardResponseHelpers.toBotInvokeRefreshResponse(adaptiveCard);
      } else {
        throw new Error("Approver card: Unknown verb " + verb);
      }
    } else if (cardId === CardID.BotSearchCard) {
      // Developers can query data here and get results
      const results = [
        { id: "1", name: "Part 1", description: "Description of part 1" },
        { id: "2", name: "Part 2", description: "Description of part 2" },
        { id: "3", name: "Part 3", description: "Description of part 3" },
      ];

      const card = AdaptiveCardHelper.createBotSearchResultsCard({ results });
      if (invokeValue.action.data.searchResult === "refresh") {
        return CardResponseHelpers.toBotInvokeRefreshResponse(card);
      } else {
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
        return CardResponseHelpers.toBotInvokeMessageResponse("Your query request was sent.");
      }
    } else if (cardId === CardID.BotSearchResultsCard) {
      // Developer can query single data with ID
      const itemId = invokeValue.action.data.id as string;
      const cardData = {
        id: itemId,
        name: "Part " + itemId,
        description: "Description of part " + itemId,
      };

      const card = AdaptiveCardHelper.createBotSearchResultCard(cardData);
      return CardResponseHelpers.toBotInvokeRefreshResponse(card);
      // return CardResponseHelpers.toEmptyBotInvokeResponse();
    } else if (cardId === CardID.BotSearchResultCard) {
      if (verb === "procurementRequest") {
        const teamMembers = await getTeamMembers(context);

        // Developers can pre-fill some workflow fields based on the search result:
        const approvers = teamMembers.map((member) => member.email).filter((email) => email?.startsWith("alex2"));
        const title = "Procurement request approval";
        const description = `Asset ID: ${invokeValue.action.data.id}`;

        const card = await this.buildApprovalBaseCard(context, title, description, approvers);

        return this.refreshAllCards(context, card);
      } else {
        throw new Error("unknown verb '" + verb + "'");
      }
    } else {
      throw new Error("Unknown card " + cardId);
    }
  }

  // Helper function to replace current card with a new adaptive card, for all users.
  private async refreshAllCards(context: TurnContext, adaptiveCard: Record<string, unknown>) {
    const cardAttachment = MessageFactory.attachment(CardFactory.adaptiveCard(adaptiveCard));
    cardAttachment.id = context.activity.replyToId;
    // See https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/update-and-delete-bot-messages?tabs=typescript#update-cards
    await context.updateActivity(cardAttachment);
    // NOTE: also need to return the exact card otherwise refresh will fail
    return CardResponseHelpers.toBotInvokeRefreshResponse(cardAttachment);
  }

  // helper method to build the base card from scratch
  private async buildApprovalBaseCard(context: TurnContext, title: string, description: string, approvers: string[]) {
    const teamMembers = await getTeamMembers(context);

    // construct workflow
    const refreshUserIds: string[] = teamMembers
      .filter(
        (item) =>
          approvers.indexOf(item.email) !== -1 ||
          item.aadObjectId === context.activity.from.aadObjectId
      )
      .map((item) => item.id);

    // sender must be in the team
    const sender = teamMembers.filter(
      (item) => item.aadObjectId === context.activity.from.aadObjectId
    )[0].email;
    const card = AdaptiveCardHelper.createBotUserSpecificViewCardApprovalBase({
      from: sender,
      title: title,
      description: description,
      approvers: approvers,
      approverComments: [],
      userIds: refreshUserIds,
    });
    return card;
  }

}
