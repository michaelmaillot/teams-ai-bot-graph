import IGraphMeetingTime from "@interfaces/IGraphMeetingTime";
import { Client } from "@microsoft/microsoft-graph-client";
import { Person, User, MeetingTimeSuggestionsResult, MeetingTimeSuggestion } from "@microsoft/microsoft-graph-types";

export class GraphClientService {
  private _token: string;
  graphClient: Client;
  constructor(token: string) {
    if (!token || !token.trim()) {
      throw new Error('GraphClient: Invalid token received.');
    }

    this._token = token;

    this.graphClient = Client.init({
      authProvider: (done: (arg0: null, arg1: string) => void) => {
        done(null, this._token);
      }
    });
  }

  getToken(): string {
    return this._token;
  }

  /**
   * Collects information about the user in the bot.
   * @returns {Promise<User>} - The user's information.
   */
  async getMe(): Promise<User> {
    return await this.graphClient
      .api('/me')
      .get().then((res: User) => {
        return res;
      });
  }

  /**
   * Gets the people in the user's organization.
   * @returns {Promise<Person[]>} - A list of people in the user's organization.
   */
  async getMyPeople(): Promise<Person[]> {
    return await this.graphClient
      .api('/me/people')
      .get().then((res: any) => {
        return res.value as Person[];
      });
  }

  /**
   * Gets the user's unread emails.
   * @returns {Promise<any>} - The user's unread emails.
   */
  async getMyUnreadEmails(): Promise<any> {
    const res = await this.graphClient
    .api('/me/messages')
    .filter("isRead ne true")
    .count(true)
    .top(5)
    .get();

    return res;
  }

  /**
   * Finds meeting times for the user and a colleague.
   * @param {IGraphMeetingTime} meetingTimeOptions - The meeting time options.
   * @returns {Promise<MeetingTimeSuggestion[]>} - The meeting time suggestions.
   */
  async findMeetingTimes(meetingTimeOptions: IGraphMeetingTime): Promise<MeetingTimeSuggestion[]> {
    const attendee = await this.getPerson(meetingTimeOptions.colleague);
    
    if (!attendee) return [];
    
    let date = new Date();
    const startTime = meetingTimeOptions.startTime ?? date.toISOString();
    const endTime = meetingTimeOptions.endTime ?? new Date(date.setDate(date.getDate() + 2)).toISOString();
    const duration = meetingTimeOptions.duration ?? 30;

    const requestBody = {
      attendees: [
        {
          emailAddress: {
            address: attendee.mail,
            name: attendee.displayName
          },
          type: "required"
        }
      ],
      timeConstraint: {
        timeslots: [
          {
            start: {
              dateTime: startTime,//"2022-01-01T09:00:00",
              timeZone: "Central European Standard Time"
            },
            end: {
              dateTime: endTime,//"2022-01-01T18:00:00",
              timeZone: "Central European Standard Time"
            }
          }
        ]
      },
      meetingDuration: "PT" + duration + "M",
      returnSuggestionReasons: true,
      minimumAttendeePercentage: 100
    };

    try {
      const response: MeetingTimeSuggestionsResult = await this.graphClient
        .api('/me/findMeetingTimes')
        .header('Prefer', 'outlook.timezone="Central European Standard Time"')
        .post(requestBody);

      if (response.emptySuggestionsReason) {
        console.log(`No meeting times available. Empty suggestions reason: ${response.emptySuggestionsReason}`);
        return [];
      }

      response.meetingTimeSuggestions.forEach((suggestion) => {
        console.log(`Suggested meeting time: ${suggestion.meetingTimeSlot.start.dateTime}`);
      });

      return response.meetingTimeSuggestions;
    } catch (error) {
      console.error(error);
      throw new Error('Failed to find meeting times.');
    }
  }

  /**
   * Gets a person by their display name.
   * @param {string} displayName - The display name of the person.
   * @returns {Promise<User>} - The person's information.
   */
  async getPerson(displayName: string): Promise<User> {
    return await this.graphClient
      .api(`/users`)
      .filter(`startswith(displayName, '${displayName}')`)
      .select('displayName,mail')
      .get().then((res) => {
        return res.value[0];
      });
  }
}