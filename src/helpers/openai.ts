import { Configuration, OpenAIApi } from "openai"
import config from "../internal/config";

export class OpenAIHelper {

    private static configuration: Configuration;
    private static openai: OpenAIApi;
    public static TRY_LATER_MESSAGE = "Sorry, I am unable to process your query at the moment. Please try again later.";
    public static SYSTEM_MESSAGE = `You are a personal assistant. Your final reply must be in markdown format. Use ** for bold and * for italics and emojis where needed. For events and tasks note that today is ${new Date()}.`;
    public static FUNCTIONS = [
        {
            "name": "getMyDetails",
            "description": "Get the details of the current user",
            "parameters": {
                "type": "object",
                "properties": {
                    "getNameOnly": {
                        "type": "boolean",
                        "description": "Get user's name only"
                    }
                },
                "required": [
                    "getNameOnly"
                ]
            }
        },
        {
            "name": "getMyEvents",
            "description": "Get the events in a calendar of the current user. Show information in a ordered list.",
            "parameters": {
                "type": "object",
                "properties": {
                    "getFutureEventsOnly": {
                        "type": "boolean",
                        "description": "Get future events only"
                    }
                },
                "required": [
                    "getFutureEventsOnly"
                ]
            }
        },
        {
            "name": "createCalendarEvent",
            "description": "Create an event in a calendar of the current user. Required contact information. For that use function getContactByName. Build object attendees with the result of getContactByName function and use the next format: attendees: [{ emailAddress: { address: 'emailAddress', name: 'name' }, type: 'required' }]",
            "parameters": {
                "type": "object",
                "required": [
                    "subject",
                    "attendees",
                    "startDateTime",
                    "endDateTime",
                    "location"
                ],
                "properties": {
                    "subject": {
                        "type": "string",
                        "description": "Subject of the event"
                    },
                    "attendees": {
                        "type": "object",
                        "description": "Attendees of the event. Each attendee is an object with emailAddress and type properties. emailAddress is an object with address and name properties. type is a string. Use function getContactByName to get the emailAddress of a contact.",
                    },
                    "startDateTime": {
                        "type": "string",
                        "description": "Start date and time of the event. Take information from user's query and convert it to a date and time. Format: YYYY-MM-DDTHH:MM:SS"
                    },
                    "endDateTime": {
                        "type": "string",
                        "description": "Start date and time of the event. Take information from user's query and convert it to a date and time. Format: YYYY-MM-DDTHH:MM:SS"
                    },
                    "location": {
                        "type": "string",
                        "description": "Location of the event"
                    }
                }
            }
        },
        {
            "name": "getMyTodoTaskList",
            "description": "Get the todo task lists from the Microsoft Todo of the current user. If the user query needs to show the todo list tasks then use the function getListTasks. If the user query needs to create a todo task then use the function createTodoTask.",
            "parameters": {
                "type": "object",
                "properties": {
                }
            }
        },
        {
            "name": "createTodoTaskList",
            "description": "Create a todo task list in the Microsoft Todo of the current user. This function is only use for a list of tasks.",
            "parameters": {
                "type": "object",
                "required": [
                    "displayName"
                ],
                "properties": {
                    "displayName": {
                        "type": "string",
                        "description": "Name of the todo task list. Send the name with the first word capitalized."
                    }
                }
            }
        },
        {
            "name": "getListTasks",
            "description": "Get list of tasks completed or not completed from a todo task list of the Microsoft Todo of the current user. For that use function getMyTodoTaskList to get the id of the todo task list.",
            "parameters": {
                "type": "object",
                "required": [
                    "getTasksByStatus",
                    "idTodoList"
                ],
                "properties": {
                    "getTasksByStatus": {
                        "type": "string",
                        "description": "Identify if the query is to get incomplete tasks or complete tasks"
                    },
                    "idTodoList": {
                        "type": "string",
                        "description": "Id of the todo task list. Get this id from the result of the function getMyTodoTaskList in the array of the key called 'value'."
                    }
                }
            }
        },
        {
            "name": "createTodoTask",
            "description": "Create a task for a task list in the Microsoft todo of the current user. For that, always use function getMyTodoTaskList to get the id of the todo task list.",
            "parameters": {
                "type": "object",
                "required": [
                    "title",
                    "idTodoList"
                ],
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "Title of the task. Send the name with the first word capitalized."
                    },
                    "idTodoList": {
                        "type": "string",
                        "description": "Id of the todo task list. Get this id from the result of the function getMyTodoTaskList in the array of the key called 'value'."
                    }
                }
            }
        },
        {
            "name": "getMyDriveDocuments",
            "description": "Get the documents from the OneDrive of the current user",
            "parameters": {
                "type": "object",
                "properties": {
                }
            }
        },
        {
            "name": "sendEmail",
            "description": "Send an email to a recipient",
            "parameters": {
                "type": "object",
                "required": [
                    "to",
                    "subject",
                    "body"
                ],
                "properties": {
                    "to": {
                        "type": "string",
                        "description": "Email address of the recipient"
                    },
                    "subject": {
                        "type": "string",
                        "description": "Subject of the email"
                    },
                    "body": {
                        "type": "string",
                        "description": "Body of the email"
                    }
                }
            }
        },
        {
            "name": "getContactByName",
            "description": "Get the details of a contact by name",
            "parameters": {
                "type": "object",
                "required": [
                    "name"
                ],
                "properties": {
                    "name": {
                        "type": "string",
                        "description": "Name of the contact"
                    }
                }
            }
        },
        {
            "name": "showFunnyMessage",
            "description": "If user's query is not related to work based personal assistance then show a funny message",
            "parameters": {
                "type": "object",
                "required": [
                    "funnyMessage"
                ],
                "properties": {
                    "funnyMessage": {
                        "type": "string",
                        "description": "A funny/sarcastic message to say why user's query is not related to work based personal assistance. Max 20 words."
                    }
                }
            }
        }
    ];

    public static async initialize() {
        this.configuration = new Configuration({
            apiKey: config.openaiAPIKey
        });
        this.openai = new OpenAIApi(this.configuration);
    }

    public static async callOpenAI(messages: any[]) {

        try {
            const response = await this.openai.createChatCompletion({
                model: config.gptModel,
                messages,
                functions: this.FUNCTIONS,
                max_tokens: 512,
                temperature: 0.5,
                top_p: 1
            });

            return response.data;

        } catch (error) {
            console.error("Error initialize: " + error);
            return null;
        }
    }

    // getAssistantMessage
    public static getAssistantMessage(functionName: string, functionArguments: any) {
        return {
            role: 'assistant',
            content: "",
            function_call: {
                name: functionName,
                arguments: JSON.stringify(functionArguments)
            }
        };
    }

    // getFunctionMessage
    public static getFunctionMessage(functionName: string, functionResult: any) {
        return {
            role: 'function',
            name: functionName,
            content: JSON.stringify(functionResult)
        };
    }

    // getUserMessage
    public static getUserMessage(userMessage: string) {
        return {
            role: 'user',
            content: userMessage
        };
    }

    // getSystemMessage
    public static getSystemMessage() {
        return {
            role: 'system',
            content: this.SYSTEM_MESSAGE
        };
    }
}