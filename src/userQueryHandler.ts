import { Activity, TurnContext, ActivityTypes } from "botbuilder";
import {
    CommandMessage,
    TriggerPatterns,
    TeamsFxBotSsoCommandHandler,
    TeamsBotSsoPromptTokenResponse,
    OnBehalfOfUserCredential,
    createMicrosoftGraphClientWithCredential,
    MessageBuilder
} from "@microsoft/teamsfx";
import oboAuthConfig from "./internal/authConfig";
import profileCard from "./adaptiveCards/profile.json";
import { OpenAIHelper } from './helpers/openai';
import { GraphHelper } from "./helpers/graph";
import { ProfileCardData } from "./models/profileCardData";

export class UserQueryHandler implements TeamsFxBotSsoCommandHandler {

    triggerPatterns: TriggerPatterns = ".*";

    openAIMessasges: any[] = [];
    pictureUrl: string = "";
    showMyDetailsAsCard: boolean = false;


    async callFunction(functionName: string, functionArguments: any) {

        console.log('funciontName: ' + functionName);

        let functionResult = null;

        if (functionName === "getMyDetails") {
            let getNameOnly = functionArguments.getNameOnly;
            functionResult = await GraphHelper.getMyDetails(getNameOnly);
            if (getNameOnly) {
                this.showMyDetailsAsCard = true;
                this.pictureUrl = await GraphHelper.getMyPhoto();
            }
        }

        if (functionName === "getMyEvents") {
            functionResult = await GraphHelper.getMyEvents(functionArguments.getFutureEventsOnly);
        }

        if (functionName === "createCalendarEvent") {
            functionResult = await GraphHelper.createCalendarEvent(
                functionArguments.subject,
                functionArguments.attendees,
                functionArguments.startDateTime,
                functionArguments.endDateTime,
                functionArguments.location
            );
        }

        if (functionName === "getMyTodoTaskList") {
            functionResult = await GraphHelper.getMyTodoTaskList();
        }

        if (functionName === "createTodoTaskList") {
            functionResult = await GraphHelper.createTodoTaskList(functionArguments.displayName);
        }

        if (functionName === "getListTasks") {
            functionResult = await GraphHelper.getListTasks(functionArguments.getTasksByStatus, functionArguments.idTodoList);
        }

        if (functionName === "createTodoTask") {
            functionResult = await GraphHelper.createTodoTask(
                functionArguments.title,
                functionArguments.idTodoList
            );
        }

        if (functionName === "getMyDriveDocuments") {
            functionResult = await GraphHelper.getMyDriveDocuments();
        }

        if (functionName === "sendEmail") {
            console.log(functionArguments);
            functionResult = await GraphHelper.sendEmail(functionArguments.to, functionArguments.subject, functionArguments.body);
        }

        if (functionName === "getContactByName") {
            console.log(functionArguments);
            functionResult = await GraphHelper.getContactByName(functionArguments.name);
        }

        return functionResult;
    }

    async processOpenAIResponse(messages: any[], response: any) {

        // if respose in null or undefined, return TRY_LATER_MESSAGE
        if (!response) {
            return OpenAIHelper.TRY_LATER_MESSAGE;
        }

        try {

            const response_choice = response["choices"][0];
            const response_message = response_choice["message"];
            const response_finish_reason = response_choice["finish_reason"];

            switch (response_finish_reason) {
                case "stop": {
                    const responseText = response_message["content"];
                    return responseText;
                }
                case "function_call": {
                    const function_name = response_message["function_call"]["name"];
                    const function_arguments = response_message["function_call"]["arguments"];
                    const function_arguments_json = JSON.parse(function_arguments);

                    switch (function_name) {
                        case "getMyDetails":
                        case "getMyEvents":
                        case "createCalendarEvent":
                        case "getMyTodoTaskList":
                        case "createTodoTaskList":
                        case "getListTasks":
                        case "createTodoTask":
                        case "getMyDriveDocuments":
                        case "sendEmail":
                        case "getContactByName": {
                            const functionResult = await this.callFunction(function_name, function_arguments_json);
                            console.log("functionResult: " + functionResult);
                            const assistantMessage = OpenAIHelper.getAssistantMessage(function_name, function_arguments_json);
                            const functionMessage = OpenAIHelper.getFunctionMessage(function_name, functionResult);
                            messages.push(assistantMessage);
                            messages.push(functionMessage);

                            const secondResponse = await OpenAIHelper.callOpenAI(messages);
                            return await this.processOpenAIResponse(messages, secondResponse);
                        }

                        case "showFunnyMessage": {
                            const funnyMessage = function_arguments_json.funnyMessage;
                            return funnyMessage;
                        }
                        default:
                            return OpenAIHelper.TRY_LATER_MESSAGE;
                    }
                }
                default:
                    return OpenAIHelper.TRY_LATER_MESSAGE;
            }

        } catch (error) {
            console.error('Error userQueryHalder: ' + error);
            return OpenAIHelper.TRY_LATER_MESSAGE;
        }
    }

    async handleCommandReceived(
        context: TurnContext,
        message: CommandMessage,
        tokenResponse: TeamsBotSsoPromptTokenResponse,
    ): Promise<string | Partial<Activity> | void> {
        await context.sendActivity({ type: ActivityTypes.Typing });

        // Init OnBehalfOfUserCredential instance with SSO token
        const oboCredential = new OnBehalfOfUserCredential(tokenResponse.ssoToken, oboAuthConfig);
        // Add scope for your Azure AD app. For example: Mail.Read, etc.
        const graphClient = createMicrosoftGraphClientWithCredential(oboCredential, [
            "User.Read",
            "Calendars.Read",
            "Calendars.ReadWrite",
            "Tasks.Read",
            "Tasks.ReadWrite",
            "Mail.Read",
            "Mail.Send",
            "Files.Read",
            "People.Read",
            "People.Read.All",
        ]);

        GraphHelper.setGraphClient(graphClient);

        // set showMyDetailsAsCard to false
        this.showMyDetailsAsCard = false;

        // initialize openAI
        await OpenAIHelper.initialize();

        // clear openAI messages
        this.openAIMessasges = [];

        // get the system message and add it to openAI messages
        const systemMessage = OpenAIHelper.getSystemMessage();
        this.openAIMessasges.push(systemMessage);

        // get the user message and add it to openAI messages
        const userMessage = OpenAIHelper.getUserMessage(message.text);
        this.openAIMessasges.push(userMessage);

        // call openAI
        const response = await OpenAIHelper.callOpenAI(this.openAIMessasges);
        const finalResponse = await this.processOpenAIResponse(this.openAIMessasges, response);

        if (this.showMyDetailsAsCard) {
            const profileCardData: ProfileCardData = {
                pictureUrl: this.pictureUrl,
                details: finalResponse
            }
            return MessageBuilder.attachAdaptiveCard<ProfileCardData>(profileCard, profileCardData);
        }

        await context.sendActivity({
            type: "message",
            text: finalResponse,
            textFormat: "markdown"
        });
    }
}
