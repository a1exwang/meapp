import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import {
  AdaptiveCardInvokeResponse,
  AdaptiveCardInvokeValue,
  Middleware,
  TeamsChannelAccount,
  TurnContext,
} from "botbuilder";
import { AdaptiveCardHelper, CardID } from "../adaptiveCardHelper";
import { CardResponseHelpers } from "../cardResponseHelpers";
import { deepClone, getTeamMembers } from "../utils";

type ActionHandler = () => Promise<WorkflowStep>;

export interface ApprovalInput {
  from: string;
  title: string;
  description: string;
  // remaining approvers
  approvers: string[];
  // already approved approvers
  approverComments: { email: string; comment: string }[];
  // current approver comment
  comment: string;
}

export interface ApprovalOutput {
  from: string;
  title: string;
  description: string;
  // remaining approvers
  approvers: string[];
  // already approved approvers
  approverComments: { email: string; comment: string }[];
}
export enum ApprovalVerb {
  Approve = "approve",
  Reject = "reject",
}

export interface WorkflowStepData<InputType = any, VerbType = string> {
  cardId: string;
  role: string;
  verb: VerbType;
  customData: InputType;
  sender: TeamsChannelAccount;
}

export enum WorkflowStepResponseType {
  Refresh = "Refresh",
  UpdateActivity = "UpdateActivity",
  Reply = "Reply",
}

export type WorkflowStepResponsePayload =
  | { type: WorkflowStepResponseType.Refresh | WorkflowStepResponseType.Reply }
  | { type: WorkflowStepResponseType.UpdateActivity; message: string };

export interface WorkflowStepOutput<T> {
  cardId: string;
  responseType: WorkflowStepResponsePayload;
  data: T;
  refreshUserIds?: string[];
}

export type WorkflowActionHandler = (
  context: TurnContext,
  input: WorkflowStepData<ApprovalInput, ApprovalVerb>
) => Promise<WorkflowStepOutput<ApprovalOutput>>;

export abstract class WorkflowStep {
  // actionHandlers: {[cardId: string]}
  // actions: { [cardId: string]: { [verb: string]: WorkflowActionHandler } };
  // abstract readonly cards: { [cardId: string]: Record<string, unknown> };

  constructor() {
    // this.actions = {
    //   [CardID.ApprovalForApprover]: {
    //     [ApprovalVerb.Approve]: this.handleWorkflow,
    //   },
    // };
  }

  actions: {
    [cardId: string]: {
      [verb: string]: (
        context: TurnContext,
        data: WorkflowStepData<ApprovalInput, ApprovalVerb>
      ) => Promise<WorkflowStepOutput<ApprovalOutput>>;
    };
  };

  cards: {
    [cardId: string]: Record<string, unknown>,
  }

  async dispatchAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    const teamMembers = await getTeamMembers(context);
    const sender = teamMembers.filter(
      (item) => item.aadObjectId === context.activity.from.aadObjectId
    )[0];

    const verb = invokeValue.action.verb as ApprovalVerb;
    const inputData: WorkflowStepData<
      ApprovalInput,
      ApprovalVerb
    > = {
      cardId: invokeValue.action.data.cardId as string,
      verb: verb,
      role: await this.getSenderRole(context),
      customData: invokeValue.action
        .data as any as ApprovalInput,
      sender: sender,
    };

    const outputData = await this.handleWorkflow(context, inputData);

    const adaptiveCard = await this.renderAdaptiveCard(context, outputData);

    switch (outputData.responseType.type) {
      case WorkflowStepResponseType.Refresh:
        // TODO:
        return CardResponseHelpers.toBotInvokeRefreshResponse(adaptiveCard);
      case WorkflowStepResponseType.UpdateActivity:
      case WorkflowStepResponseType.Reply:
        throw new Error("not implemented");
      default:
        throw new Error("unknown response type");
    }
  }

  async getSenderRole(context: TurnContext): Promise<string> {
    return "unknown";
  }

  // default implem
  async buildAdaptiveCardTemplate(
    cardId: string,
    result: WorkflowStepOutput<ApprovalOutput>
  ): Promise<unknown> {
    const cards = {
      [CardID.ApprovalForApprover]:
        AdaptiveCardHelper.createBotUserSpecificViewCardApprovalApproved(
          result.data
        ),
      [CardID.ApprovalBase]:
        AdaptiveCardHelper.createBotUserSpecificViewCardApprovalApproved(
          result.data
        ),
    };

    return cards[cardId];
  }

  async buildAdaptiveCardData(
    cardId: string,
    result: WorkflowStepOutput<ApprovalOutput>
  ) {
    return result.data;
  }

  async isMatch(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<boolean> {
    return true;
  }

  async handleWorkflow(
    context: TurnContext,
    data: WorkflowStepData<ApprovalInput, ApprovalVerb>
  ): Promise<WorkflowStepOutput<ApprovalOutput>> {
    const action = this.actions[data.cardId]?.[data.verb];
    if (!action) {
      throw new Error(`Action and verb not found ${data.cardId}, ${data.verb}`);
    }

    return await action(context, data);
  }

  async renderAdaptiveCard(
    context: TurnContext,
    outputData: WorkflowStepOutput<ApprovalOutput>
  ) {
    if (!(outputData.cardId in this.cards)) {
      throw new Error("Card not found " + outputData.cardId);
    }
    return AdaptiveCards.declare(this.cards[outputData.cardId]).render(
      outputData.data
    );
  }
}

export class Workflow {
  constructor(name: string, steps: WorkflowStep[]) {
  }

  async startWorkflow(context: TurnContext, initalData: WorkflowStepData): Promise<void> {

  }

  async isWorkflow(context: TurnContext, data: AdaptiveCardInvokeValue): Promise<boolean> {
    return true;
  }

  async handleWorkflow(context: TurnContext, data: AdaptiveCardInvokeValue): Promise<boolean> {
    return true;
  }
}

export class WorkflowMiddleware implements Middleware {
  workflows: Workflow[] = [];
  constructor(workflows: Workflow[]) {
    this.workflows = workflows;
  }

  async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    if (context.activity.type === "invoke") {
      const invokeData: AdaptiveCardInvokeValue = context.activity.value;
      // TODO: option workflow match check (maybe workflow id/name + map?)
      for (const workflow of this.workflows) {
        if (await workflow.isWorkflow(context, invokeData)) {
          const handled = await workflow.handleWorkflow(context, invokeData);
          if (handled) {
            return;
          }
        }
      }
    }
    await next();
  }
}