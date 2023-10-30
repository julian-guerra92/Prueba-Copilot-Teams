import { ResponseType } from "@microsoft/microsoft-graph-client";
import { ContactEventCalendar } from "../models/contactEventCalendar";

export class GraphHelper {
    static graphClient: any;

    static async setGraphClient(graphClient: any) {
        this.graphClient = graphClient;
    }

    static async getMyDetails(getNameOnly: boolean) {
        const me = await this.graphClient.api("/me").get();

        if (me) {
            if (getNameOnly) {
                return me.displayName;
            } else {
                return me;
            }
        } else {
            return null;
        }
    }

    static async getMyPhoto() {
        let photoBinary: ArrayBuffer;
        try {
            photoBinary = await this.graphClient
                .api("/me/photo/$value")
                .responseType(ResponseType.ARRAYBUFFER)
                .get();
        } catch {
            return;
        }

        const buffer = Buffer.from(photoBinary);

        return "data:image/png;base64," + buffer.toString("base64");
    }

    static async getMyEvents(futureEventsOnly: boolean) {
        const userEvents = await this.graphClient.api("/me/events").select(["subject", "start", "end", "attendees", "location"]).get();

        if (userEvents) {
            if (futureEventsOnly) {
                return userEvents.value.filter((event: any) => {
                    return new Date(event.end.dateTime) > new Date();
                });
            }
            return userEvents;
        } else {
            return null;
        }
    }

    static async createCalendarEvent(subject: string, attendees: ContactEventCalendar[], startDateTime: string, endDateTime: string, location: string) {
        const event = {
            subject: subject,
            attendees: attendees,
            start: {
                dateTime: startDateTime,
                timeZone: "UTC"
            },
            end: {
                dateTime: endDateTime,
                timeZone: "UTC"
            },
            location: {
                displayName: location
            }
        };

        const result = await this.graphClient.api("/me/events").post({ event: event });
        if (result) {
            return "Event created successfully";
        } else {
            return null;
        }
    }

    static async getMyTodoTaskList() {
        const userTodoTaskLists = await this.graphClient.api("/me/todo/lists").get();
        if (userTodoTaskLists) {
            return userTodoTaskLists;
        } else {
            return null;
        }
    }

    static async createTodoTaskList(name: string) {
        const todoTaskList = {
            displayName: name
        };
        const result = await this.graphClient.api("/me/todo/lists").post({ todoTaskList: todoTaskList });
        if (result) {
            return "Todo task list created successfully";
        } else {
            return null;
        }
    }

    static async getMyTodoTasks(getIncompleteTasksOnly: boolean, idTodoList: string) {
        const userTodoTasks = await this.graphClient.api(`/me/todo/${idTodoList}/tasks`).get();
        if (userTodoTasks) {
            if (getIncompleteTasksOnly) {
                return userTodoTasks.value.filter((task: any) => {
                    return task.status !== "completed";
                });
            }
            return userTodoTasks;
        } else {
            return null;
        }
    }

    static async createTodoTask(title: string, startDateTime: string, dueDateTime: string, idTodoList: string) {
        const task = {
            title: title,
            startDateTime: startDateTime,
            dueDateTime: dueDateTime
        };
        const result = await this.graphClient.api(`/me/todo/${idTodoList}/tasks`).post({ task: task });
        if (result) {
            return "Task created successfully";
        } else {
            return null;
        }
    }

    static async getMyDriveDocuments() {
        const userDriveItems = await this.graphClient.api("/me/drive/root/children")
            .select([
                "name",
                "webUrl",
                "@microsoft.graph.downloadUrl",
                "createdBy",
                "lastModifiedBy"
            ]).get();

        if (userDriveItems) {
            return userDriveItems;
        } else {
            return null;
        }
    }

    static async sendEmail(to: string, subject: string, body: string) {
        const email = {
            subject: subject,
            toRecipients: [
                {
                    emailAddress: {
                        address: to
                    }
                }
            ],
            body: {
                content: body,
                contentType: "text"
            }
        };

        const result = await this.graphClient.api("/me/sendMail").post({ message: email });
        if (result) {
            return "Email sent successfully";
        } else {
            return null;
        }
    }

    static async getContactByName(name: string) {
        const contacts = await this.graphClient.api("/me/people").search(name).get();
        console.log(contacts);

        if (contacts) {
            return contacts;
        } else {
            return null;
        }
    }

}