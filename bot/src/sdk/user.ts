import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { TurnContext } from "botbuilder";
import { AdaptiveCardHelper, CardID } from "../adaptiveCardHelper";
import { deepClone, getTeamMembers } from "../utils";
import {
  ApprovalInput as ApprovalInput,
  ApprovalOutput as ApprovalOutput,
  ApprovalVerb,
  Workflow,
  WorkflowMiddleware,
  WorkflowStep,
  WorkflowStepData,
  WorkflowStepOutput,
  WorkflowStepResponseType as WorkflowResponseType,
} from "./sdk";
import approvalForApproverTemplate from "../../adaptiveCards/approval/approvalForApprover.json";
import approvalForSenderTemplate from "../../adaptiveCards/approval/approvalForSender.json";

// class WorkflowStepSender extends WorkflowStep {
//   readonly cards = {
//     [CardID.ApprovalForSender]: senderCard,
//   }

//   async handleWorkflow(
//     context: TurnContext,
//     data: WorkflowStepData<ApprovalApproverApproveInput, ApprovalVerb>
//   ): Promise<WorkflowStepOutput<ApprovalApproverApproveOutput>> {
//     return {
//       cardId: CardID.ApprovalBase,
//       responseType: { type: WorkflowResponseType.Reply },
//       result: {...some result},
//     };
//   }
// }

// class WorkflowStepApprover extends WorkflowStep {

//   readonly cards = {
//     [CardID.ApprovalForApprover]: approverCard,
//   }

//   async handleWorkflow(
//     context: TurnContext,
//     data: WorkflowStepData<ApprovalApproverApproveInput, ApprovalVerb>
//   ): Promise<WorkflowStepOutput<ApprovalApproverApproveOutput>> {

//     const result: ApprovalApproverApproveOutput = deepClone(data.customData);
//     result.approverComments = [
//       ...data.customData.approverComments,
//       { email: data.sender.email, comment: data.customData.comment },
//     ];

//     if (data.customData.approvers.length === 1) {
//       // If the last approver approves, the request is passed.
//       return {
//         cardId: CardID.ApprovalForApprover,
//         responseType: { type: WorkflowResponseType.Refresh },
//         result,
//       };
//     } else {
//       // Otherwise, remove this approver from approver list.
//       const newApprovers = data.customData.approvers.filter(
//         (approver) => approver != data.sender.email
//       );
//       const teamMembers = await getTeamMembers(context);
//       const refreshUserIds: string[] = teamMembers
//         .filter(
//           (item) =>
//             newApprovers.indexOf(item.email) !== -1 ||
//             item.aadObjectId === context.activity.from.aadObjectId
//         )
//         .map((item) => item.id);
//       result.approvers = newApprovers;
//       return {
//         cardId: CardID.ApprovalBase,
//         responseType: { type: WorkflowResponseType.Refresh },
//         result: result,
//         refreshUserIds: refreshUserIds,
//       };
//     }
//   }
// }

// const workflowBot = new WorkflowBot({
//   adapterConfig: {
//     appId: process.env.BOT_ID,
//     appPassword: process.env.BOT_PASSWORD,
//   },
//   workflowBot: {
//     enabled: true,
//     workflows: [new Workflow([new WorkflowStepApprover(), new WorkflowStepSender()])],
//   },
// });

export const workflowMiddleware = new WorkflowMiddleware([
  new Workflow("approval", []),
]);

class WorkflowStepApprover extends WorkflowStep {
  constructor() {
    super();

    this.cards = {
      [CardID.ApprovalForApprover]: approvalForApproverTemplate,
      [CardID.ApprovalForSender]: approvalForSenderTemplate,
    };
    this.actions = {
      [CardID.ApprovalBase]: {
        [ApprovalVerb.Approve]: this.handleApprove.bind(this),
      },
    };
  }

  async handleApprove(
    context: TurnContext,
    data: WorkflowStepData<ApprovalInput, ApprovalVerb>
  ): Promise<WorkflowStepOutput<ApprovalOutput>> {
    const output: ApprovalOutput = deepClone(data.customData);

    /// Bussiness logic
    output.approverComments = [
      ...data.customData.approverComments,
      { email: data.sender.email, comment: data.customData.comment },
    ];

    if (data.customData.approvers.length === 1) {
      // If the last approver approves, the request is passed.
      return {
        cardId: CardID.ApprovalForApprover,
        responseType: { type: WorkflowResponseType.Refresh },
        data: output,
      };
    } else {
      // Otherwise, remove this approver from approver list.
      const newApprovers = data.customData.approvers.filter(
        (approver) => approver != data.sender.email
      );
      const teamMembers = await getTeamMembers(context);
      const refreshUserIds: string[] = teamMembers
        .filter(
          (item) =>
            newApprovers.indexOf(item.email) !== -1 ||
            item.aadObjectId === context.activity.from.aadObjectId
        )
        .map((item) => item.id);
      output.approvers = newApprovers;

      // build a model and send to views (adaptive card renderer)
      return {
        cardId: CardID.ApprovalBase,
        responseType: { type: WorkflowResponseType.Refresh },
        data: output,
        refreshUserIds: refreshUserIds,
      };
    }
  }
}