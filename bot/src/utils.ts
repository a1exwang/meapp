import {
  ChannelAccount,
  ConversationParameters,
  ConversationReference,
  TeamsChannelAccount,
  TeamsInfo,
  TurnContext,
} from "botbuilder";
import { ConnectorClient } from "botframework-connector";

export function deepClone<T>(object: T): T {
  return JSON.parse(JSON.stringify(object));
}

export async function getTeamMembers(
  context: TurnContext
): Promise<TeamsChannelAccount[]> {
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
export function generateDeeplink(context: TurnContext): string {
  const parentMessageId = context.activity.replyToId;
  const deeplink = `https://teams.microsoft.com/l/message/${context.activity.channelData.channel.id}/${parentMessageId}?tenantId=${context.activity.conversation.tenantId}&parentMessageId=${parentMessageId}`;
  return deeplink;
}

export async function getUserInfoFromAadObjectId(
  context: TurnContext,
  aadObjectId: string
): Promise<TeamsChannelAccount> {
  const members = await this.getTeamMembers(context);
  const result = members.filter((item) => item.aadObjectId === aadObjectId);
  if (result.length === 0) {
    throw new Error("User not found " + aadObjectId);
  }
  return result[0];
}

// Helper method to send notification
export async function newConversation(
  context: TurnContext,
  user: ChannelAccount
): Promise<ConversationReference> {
  const reference = TurnContext.getConversationReference(context.activity);
  const personalConversation = JSON.parse(JSON.stringify(reference));

  const connectorClient: ConnectorClient = context.turnState.get(
    context.adapter.ConnectorClientKey
  );
  const conversation = await connectorClient.conversations.createConversation({
    isGroup: false,
    tenantId: context.activity.conversation.tenantId,
    bot: context.activity.recipient,
    members: [user],
    channelData: {},
  } as ConversationParameters);
  personalConversation.conversation.id = conversation.id;

  return personalConversation;
}
