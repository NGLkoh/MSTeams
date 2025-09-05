
  // <GetUserSnippet>
  import { Client, GraphRequestOptions, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';
  import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
  import { endOfWeek, startOfWeek } from 'date-fns';
  import { zonedTimeToUtc } from 'date-fns-tz';
  import { User, Event } from '@microsoft/microsoft-graph-types';

  
  let graphClient: Client | undefined = undefined;

  function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    if (!graphClient) {
      graphClient = Client.initWithMiddleware({
        authProvider: authProvider
      });
    }

    return graphClient;
  }

  export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
    ensureClient(authProvider);

    // Return the /me API endpoint result as a User object
    const user: User = await graphClient!.api('/me')
      // Only retrieve the specific fields needed
      .select('displayName,mail,mailboxSettings,userPrincipalName')
      .get();

    return user;
  }
  // </GetUserSnippet>

  // <GetUserWeekCalendarSnippet>
  export async function getUserWeekCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string, startDate: Date, endDate:Date): Promise<Event[]> {
    ensureClient(authProvider);

    // Generate startDateTime and endDateTime query params
    // to display a 7-day window
    const startDateTime = zonedTimeToUtc(startDate, timeZone).toISOString();
    const endDateTime = zonedTimeToUtc(endDate, timeZone).toISOString();


    // GET /me/calendarview?startDateTime=''&endDateTime=''
    // &$select=subject,organizer,start,end
    // &$orderby=start/dateTime
    // &$top=50
    var response: PageCollection = await graphClient!
      .api('/me/calendarview')
      .header('Prefer', `outlook.timezone="${timeZone}"`)
      .query({ startDateTime: startDateTime, endDateTime: endDateTime })
      .select('subject,organizer,start,end,recurrence')
      .orderby('start/dateTime')
      .top(25)
      .get();

    if (response["@odata.nextLink"]) {
      // Presence of the nextLink property indicates more results are available
      // Use a page iterator to get all results
      var events: Event[] = [];

      // Must include the time zone header in page
      // requests too
      var options: GraphRequestOptions = {
        headers: { 'Prefer': `outlook.timezone="${timeZone}"` }
      };

      var pageIterator = new PageIterator(graphClient!, response, (event) => {
        events.push(event);
        return true;
      }, options);

      await pageIterator.iterate();

      const schedulerEvents: any = events.map((event) => ({
        ...event,
        subject: event.subject,
        start: event.start?.dateTime,
        end: event.end?.dateTime,
        startTimezone: event.start?.timeZone,
        endTimezone: event.end?.timeZone,
        recurrence: event.recurrence
      }));
      return schedulerEvents;
    } else {
      const schedulerEvents = response.value.map((event) => ({
        ...event,
        subject: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        startTimezone: event.start.timeZone,
        endTimezone: event.end.timeZone,
        recurrence: event.recurrence
      }));
      return schedulerEvents;
    }
  }
  // </GetUserWeekCalendarSnippet>

  export async function createCalendarSubscription(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    ensureClient(authProvider);

    const subscription = {
      changeType: "created,updated,deleted",
      notificationUrl: "http://localhost:3000/api/callback", // ngrok public URL
      resource: "me/events",
      expirationDateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(), // 1 hour max for personal accounts
      clientState: "secretClientValue"
    };

    return await graphClient!.api("/subscriptions").post(subscription);
  }

// Enhanced Teams Meeting Creation - GraphService.ts additions

export async function createValidTeamsMeeting(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[] = [],
  body: string = "",
  timezone: string = "UTC"
) {
  ensureClient(authProvider);

  console.log("Creating valid Teams meeting with proper join URL...", { subject, start, end, attendeeEmails });

  // Method 1: Use the /me/onlineMeetings endpoint directly (most reliable for getting join URLs)
  try {
    console.log("Attempting direct online meeting creation via /me/onlineMeetings...");
    
    const onlineMeetingRequest = {
      subject: subject,
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      participants: {
        organizer: {
          identity: {
            user: {
              id: "current" // Will be resolved to current user
            }
          }
        },
        attendees: attendeeEmails.map(email => ({
          identity: {
            user: {
              id: email
            }
          }
        }))
      }
    };

    const onlineMeeting = await graphClient!.api("/me/onlineMeetings").post(onlineMeetingRequest);
    console.log("Direct online meeting created successfully:", onlineMeeting);

    // Now create the calendar event that references this online meeting
    const calendarEventRequest = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: `
          <div>
            ${body ? `<p>${body}</p>` : ''}
            <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
              <p><strong>Microsoft Teams Meeting</strong></p>
              <p><a href="${onlineMeeting.joinWebUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
              ${onlineMeeting.videoTeleconferenceId ? `<p><strong>Conference ID:</strong> ${onlineMeeting.videoTeleconferenceId}</p>` : ''}
              ${onlineMeeting.audioConferencing?.tollNumber ? `<p><strong>Phone:</strong> ${onlineMeeting.audioConferencing.tollNumber}</p>` : ''}
            </div>
          </div>
        `
      },
      start: {
        dateTime: start.toISOString(),
        timeZone: "UTC"
      },
      end: {
        dateTime: end.toISOString(),
        timeZone: "UTC"
      },
      location: {
        displayName: "Microsoft Teams Meeting",
        locationUri: onlineMeeting.joinWebUrl
      },
      attendees: attendeeEmails.map(email => ({
        emailAddress: {
          address: email,
          name: email.split('@')[0]
        },
        type: "required"
      })),
      // Link the calendar event to the online meeting
      onlineMeetingUrl: onlineMeeting.joinWebUrl
    };

    const calendarEvent = await graphClient!.api("/me/events").post(calendarEventRequest);
    console.log("Calendar event created with online meeting link:", calendarEvent);

    return {
      success: true,
      method: "direct_online_meeting",
      id: onlineMeeting.id,
      eventId: calendarEvent.id,
      subject: subject,
      joinWebUrl: onlineMeeting.joinWebUrl,
      joinUrl: onlineMeeting.joinWebUrl,
      meetingId: onlineMeeting.videoTeleconferenceId || onlineMeeting.id,
      conferenceId: onlineMeeting.audioConferencing?.conferenceId,
      tollNumber: onlineMeeting.audioConferencing?.tollNumber,
      tollFreeNumber: onlineMeeting.audioConferencing?.tollFreeNumber,
      dialInUrl: onlineMeeting.audioConferencing?.dialInUrl,
      startTime: { dateTime: start.toISOString() },
      endTime: { dateTime: end.toISOString() },
      attendees: calendarEvent.attendees,
      webLink: calendarEvent.webLink,
      onlineMeeting: onlineMeeting,
      calendarEvent: calendarEvent
    };

  } catch (directError: any) {
    console.log("Direct online meeting creation failed:", directError.message);
    
    // Method 2: Calendar-first approach with retry logic
    try {
      console.log("Attempting calendar-first approach with Teams integration...");
      
      const calendarEventWithTeams = {
        subject: subject,
        body: {
          contentType: "HTML",
          content: body || "Microsoft Teams meeting"
        },
        start: {
          dateTime: start.toISOString(),
          timeZone: "UTC"
        },
        end: {
          dateTime: end.toISOString(),
          timeZone: "UTC"
        },
        attendees: attendeeEmails.map(email => ({
          emailAddress: {
            address: email,
            name: email.split('@')[0]
          },
          type: "required"
        })),
        isOnlineMeeting: true,
        onlineMeetingProvider: "teamsForBusiness"
      };

      const createdEvent = await graphClient!.api("/me/events").post(calendarEventWithTeams);
      console.log("Calendar event created, waiting for Teams integration...");

      // Retry logic to wait for Teams meeting to be provisioned
      let retryCount = 0;
      const maxRetries = 5;
      let eventWithMeeting = null;

      while (retryCount < maxRetries) {
        await new Promise(resolve => setTimeout(resolve, 2000 + (retryCount * 1000))); // Increasing delay
        
        try {
          eventWithMeeting = await graphClient!.api(`/me/events/${createdEvent.id}`)
            .select('id,subject,start,end,onlineMeeting,attendees,webLink,isOnlineMeeting')
            .get();

          if (eventWithMeeting.onlineMeeting && eventWithMeeting.onlineMeeting.joinUrl) {
            console.log(`Teams meeting provisioned after ${retryCount + 1} retries`);
            break;
          }
        } catch (fetchError) {
          console.log(`Retry ${retryCount + 1} failed:`, fetchError);
        }
        
        retryCount++;
      }

      if (eventWithMeeting && eventWithMeeting.onlineMeeting && eventWithMeeting.onlineMeeting.joinUrl) {
        return {
          success: true,
          method: "calendar_with_retry",
          id: eventWithMeeting.id,
          eventId: eventWithMeeting.id,
          subject: eventWithMeeting.subject,
          joinWebUrl: eventWithMeeting.onlineMeeting.joinUrl,
          joinUrl: eventWithMeeting.onlineMeeting.joinUrl,
          meetingId: eventWithMeeting.onlineMeeting.conferenceId || eventWithMeeting.id,
          conferenceId: eventWithMeeting.onlineMeeting.conferenceId,
          tollNumber: eventWithMeeting.onlineMeeting.tollNumber,
          tollFreeNumber: eventWithMeeting.onlineMeeting.tollFreeNumber,
          startTime: eventWithMeeting.start,
          endTime: eventWithMeeting.end,
          attendees: eventWithMeeting.attendees,
          webLink: eventWithMeeting.webLink,
          onlineMeeting: eventWithMeeting.onlineMeeting,
          calendarEvent: eventWithMeeting,
          retriesNeeded: retryCount + 1
        };
      } else {
        throw new Error("Teams meeting was not provisioned after maximum retries");
      }

    } catch (calendarError: any) {
      console.log("Calendar-first approach failed:", calendarError.message);
      
      // Method 3: Try the application-level communications API
      try {
        console.log("Attempting application communications API...");
        
        const communicationsMeetingRequest = {
          subject: subject,
          startDateTime: start.toISOString(),
          endDateTime: end.toISOString()
        };

        const communicationsMeeting = await graphClient!.api("/app/onlineMeetings").post(communicationsMeetingRequest);
        
        // Create calendar event with the meeting link
        const calendarEventForComms = {
          subject: subject,
          body: {
            contentType: "HTML",
            content: `
              <div>
                ${body ? `<p>${body}</p>` : ''}
                <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
                  <p><strong>Microsoft Teams Meeting</strong></p>
                  <p><a href="${communicationsMeeting.joinWebUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
                  <p><strong>Meeting ID:</strong> ${communicationsMeeting.id}</p>
                </div>
              </div>
            `
          },
          start: {
            dateTime: start.toISOString(),
            timeZone: "UTC"
          },
          end: {
            dateTime: end.toISOString(),
            timeZone: "UTC"
          },
          location: {
            displayName: "Microsoft Teams Meeting",
            locationUri: communicationsMeeting.joinWebUrl
          },
          attendees: attendeeEmails.map(email => ({
            emailAddress: {
              address: email,
              name: email.split('@')[0]
            },
            type: "required"
          }))
        };

        const commsCalendarEvent = await graphClient!.api("/me/events").post(calendarEventForComms);

        return {
          success: true,
          method: "communications_api",
          id: communicationsMeeting.id,
          eventId: commsCalendarEvent.id,
          subject: subject,
          joinWebUrl: communicationsMeeting.joinWebUrl,
          joinUrl: communicationsMeeting.joinWebUrl,
          meetingId: communicationsMeeting.id,
          startTime: { dateTime: start.toISOString() },
          endTime: { dateTime: end.toISOString() },
          attendees: commsCalendarEvent.attendees,
          webLink: commsCalendarEvent.webLink,
          onlineMeeting: communicationsMeeting,
          calendarEvent: commsCalendarEvent
        };

      } catch (commsError: any) {
        console.error("All methods failed:", commsError);
        throw new Error(`Failed to create valid Teams meeting using all available methods. Last error: ${commsError.message}. Required permissions: OnlineMeetings.ReadWrite, Calendars.ReadWrite, User.Read`);
      }
    }
  }
}

// Helper function to verify Teams meeting is valid
export async function verifyTeamsMeetingUrl(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  meetingId: string
): Promise<boolean> {
  ensureClient(authProvider);
  
  try {
    // Try to fetch the meeting details to verify it exists and has a join URL
    const meeting = await graphClient!.api(`/me/onlineMeetings/${meetingId}`).get();
    return !!(meeting && (meeting.joinWebUrl || meeting.joinUrl));
  } catch (error) {
    console.log("Meeting verification failed:", error);
    return false;
  }
}


export async function detectAccountType(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  const graphClient = Client.initWithMiddleware({ authProvider });
  
  try {
    const user = await graphClient.api('/me').select('userPrincipalName,mail,accountEnabled').get();
    
    // Check if it's a personal account or business account
    const isPersonalAccount = user.userPrincipalName?.includes('@outlook.com') || 
                             user.userPrincipalName?.includes('@hotmail.com') ||
                             user.userPrincipalName?.includes('@live.com') ||
                             user.mail?.includes('@outlook.com') ||
                             user.mail?.includes('@hotmail.com') ||
                             user.mail?.includes('@live.com');
    
    return {
      isPersonalAccount,
      userPrincipalName: user.userPrincipalName,
      mail: user.mail,
      expectedUrlFormat: isPersonalAccount ? 'teams.live.com' : 'teams.microsoft.com'
    };
  } catch (error) {
    console.error('Error detecting account type:', error);
    return {
      isPersonalAccount: false,
      expectedUrlFormat: 'teams.microsoft.com'
    };
  }
}

// Enhanced meeting creation that tries to generate the URL format you want
export async function createTeamsMeetingWithSpecificFormat(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[] = [],
  body: string = "",
  timezone: string
) {
  const graphClient = Client.initWithMiddleware({ authProvider });
  
  console.log("Creating Teams meeting with specific URL format...");
  
  // First, detect the account type
  const accountInfo = await detectAccountType(authProvider);
  console.log("Account type detected:", accountInfo);
  
  // Method 1: Try creating online meeting directly
  try {
    const onlineMeetingRequest = {
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      subject: subject
    };

    console.log("Creating online meeting...");
    const onlineMeeting = await graphClient.api("/me/onlineMeetings").post(onlineMeetingRequest);
    
    const joinUrl = onlineMeeting.joinWebUrl || onlineMeeting.joinUrl;
    
    if (joinUrl) {
      console.log("Generated URL:", joinUrl);
      console.log("URL format matches expected:", joinUrl.includes(accountInfo.expectedUrlFormat));
      
      // Create calendar event
      const calendarEventRequest = {
        subject: subject,
        body: {
          contentType: "HTML",
          content: `
            <div>
              ${body ? `<p>${body}</p>` : ''}
              <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
                <p><strong>Microsoft Teams Meeting</strong></p>
                <p><a href="${joinUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
                ${onlineMeeting.videoTeleconferenceId ? `<p><strong>Conference ID:</strong> ${onlineMeeting.videoTeleconferenceId}</p>` : ''}
              </div>
            </div>
          `
        },
        start: {
          dateTime: start.toISOString(),
          timeZone: "UTC"
        },
        end: {
          dateTime: end.toISOString(),
          timeZone: "UTC"
        },
        location: {
          displayName: "Microsoft Teams Meeting",
          locationUri: joinUrl
        },
        attendees: attendeeEmails.map(email => ({
          emailAddress: {
            address: email,
            name: email.split('@')[0]
          },
          type: "required"
        })),
        isOnlineMeeting: true,
        onlineMeetingProvider: "teamsForBusiness"
      };

      const calendarEvent = await graphClient.api("/me/events").post(calendarEventRequest);
      
      return {
        success: true,
        method: "direct_online_meeting",
        accountType: accountInfo,
        eventId: calendarEvent.id,
        meetingId: onlineMeeting.id,
        subject: subject,
        joinWebUrl: joinUrl,
        joinUrl: joinUrl,
        urlFormat: joinUrl.includes('teams.live.com') ? 'teams.live.com' : 
                   joinUrl.includes('teams.microsoft.com') ? 'teams.microsoft.com' : 'unknown',
        conferenceId: onlineMeeting.videoTeleconferenceId,
        startTime: { dateTime: start.toISOString() },
        endTime: { dateTime: end.toISOString() },
        attendees: calendarEvent.attendees,
        webLink: calendarEvent.webLink,
        onlineMeeting: onlineMeeting,
        calendarEvent: calendarEvent
      };
    }
  } catch (error) {
    console.log("Direct online meeting creation failed:", error);
  }

  // Method 2: Try using consumer APIs if personal account
  if (accountInfo.isPersonalAccount) {
    try {
      console.log("Trying consumer-specific endpoints for personal account...");
      
      // For personal accounts, try different endpoint patterns
      const consumerMeetingRequest = {
        startTime: start.toISOString(),
        endTime: end.toISOString(),
        subject: subject
      };

      // Try alternative endpoints that might generate teams.live.com URLs
      let consumerMeeting;
      
      try {
        // Try the communications endpoint
        consumerMeeting = await graphClient.api("/communications/onlineMeetings").post(consumerMeetingRequest);
      } catch (commError) {
        console.log("Communications endpoint failed, trying app endpoint...");
        
        try {
          // Try app-level endpoint
          consumerMeeting = await graphClient.api("/app/onlineMeetings").post(consumerMeetingRequest);
        } catch (appError) {
          console.log("App endpoint also failed");
          throw appError;
        }
      }
      
      if (consumerMeeting && (consumerMeeting.joinWebUrl || consumerMeeting.joinUrl)) {
        const joinUrl = consumerMeeting.joinWebUrl || consumerMeeting.joinUrl;
        console.log("Consumer meeting URL generated:", joinUrl);
        
        // Create calendar event with the consumer meeting
        const calendarEventRequest = {
          subject: subject,
          body: {
            contentType: "HTML",
            content: `
              <div>
                ${body ? `<p>${body}</p>` : ''}
                <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
                  <p><strong>Microsoft Teams Meeting</strong></p>
                  <p><a href="${joinUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
                  <p><strong>Meeting ID:</strong> ${consumerMeeting.id}</p>
                </div>
              </div>
            `
          },
          start: {
            dateTime: start.toISOString(),
            timeZone: "UTC"
          },
          end: {
            dateTime: end.toISOString(),
            timeZone: "UTC"
          },
          location: {
            displayName: "Microsoft Teams Meeting",
            locationUri: joinUrl
          },
          attendees: attendeeEmails.map(email => ({
            emailAddress: {
              address: email,
              name: email.split('@')[0]
            },
            type: "required"
          }))
        };

        const calendarEvent = await graphClient.api("/me/events").post(calendarEventRequest);

        return {
          success: true,
          method: "consumer_api",
          accountType: accountInfo,
          eventId: calendarEvent.id,
          meetingId: consumerMeeting.id,
          subject: subject,
          joinWebUrl: joinUrl,
          joinUrl: joinUrl,
          urlFormat: joinUrl.includes('teams.live.com') ? 'teams.live.com' : 
                     joinUrl.includes('teams.microsoft.com') ? 'teams.microsoft.com' : 'unknown',
          startTime: { dateTime: start.toISOString() },
          endTime: { dateTime: end.toISOString() },
          attendees: calendarEvent.attendees,
          webLink: calendarEvent.webLink,
          onlineMeeting: consumerMeeting,
          calendarEvent: calendarEvent
        };
      }
    } catch (consumerError) {
      console.log("Consumer API approach failed:", consumerError);
    }
  }

  // Method 3: Fallback to calendar-first approach
  try {
    console.log("Fallback: Calendar-first approach...");
    
    const calendarEvent = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body || "Microsoft Teams meeting"
      },
      start: {
        dateTime: start.toISOString(),
        timeZone: "UTC"
      },
      end: {
        dateTime: end.toISOString(),
        timeZone: "UTC"
      },
      attendees: attendeeEmails.map(email => ({
        emailAddress: {
          address: email,
          name: email.split('@')[0]
        },
        type: "required"
      })),
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    };

    const createdEvent = await graphClient.api("/me/events").post(calendarEvent);
    
    // Wait for Teams meeting to be provisioned
    for (let i = 0; i < 10; i++) {
      await new Promise(resolve => setTimeout(resolve, 2000 + (i * 1000)));
      
      try {
        const eventWithMeeting = await graphClient.api(`/me/events/${createdEvent.id}`)
          .select('id,subject,start,end,onlineMeeting,attendees,webLink,isOnlineMeeting')
          .expand('onlineMeeting')
          .get();

        if (eventWithMeeting.onlineMeeting && 
            (eventWithMeeting.onlineMeeting.joinUrl || eventWithMeeting.onlineMeeting.joinWebUrl)) {
          
          const joinUrl = eventWithMeeting.onlineMeeting.joinUrl || eventWithMeeting.onlineMeeting.joinWebUrl;
          console.log(`Teams meeting provisioned after ${i + 1} retries:`, joinUrl);
          
          return {
            success: true,
            method: "calendar_with_retry",
            accountType: accountInfo,
            eventId: eventWithMeeting.id,
            meetingId: eventWithMeeting.onlineMeeting.conferenceId || eventWithMeeting.id,
            subject: eventWithMeeting.subject,
            joinWebUrl: joinUrl,
            joinUrl: joinUrl,
            urlFormat: joinUrl.includes('teams.live.com') ? 'teams.live.com' : 
                       joinUrl.includes('teams.microsoft.com') ? 'teams.microsoft.com' : 'unknown',
            conferenceId: eventWithMeeting.onlineMeeting.conferenceId,
            startTime: eventWithMeeting.start,
            endTime: eventWithMeeting.end,
            attendees: eventWithMeeting.attendees,
            webLink: eventWithMeeting.webLink,
            onlineMeeting: eventWithMeeting.onlineMeeting,
            calendarEvent: eventWithMeeting,
            retriesNeeded: i + 1
          };
        }
      } catch (fetchError) {
        console.log(`Retry ${i + 1} failed:`, fetchError);
      }
    }
    
    throw new Error("Teams meeting was not provisioned after maximum retries");
    
  } catch (calendarError) {
    console.error("Calendar-first approach failed:", calendarError);
    throw new Error(`All meeting creation methods failed: ${calendarError}`);
  }
}

// Function to specifically target personal Teams meetings
export async function createPersonalTeamsMeeting(
authProvider: AuthCodeMSALBrowserAuthenticationProvider, subject: string, start: Date, end: Date, attendeeEmails: string[] = [], body: string = "", timezone: string) {
  const graphClient = Client.initWithMiddleware({ authProvider });
  
  console.log("Attempting to create personal Teams meeting (teams.live.com format)...");
  
  // Verify this is a personal account
  const accountInfo = await detectAccountType(authProvider);
  
  if (!accountInfo.isPersonalAccount) {
    console.warn("This appears to be a business account. Personal Teams meeting URLs (teams.live.com) are typically only available for personal Microsoft accounts.");
  }
  
  // Try multiple approaches specific to personal accounts
  const approaches = [
    {
      name: "Consumer OnlineMeetings API",
      endpoint: "/me/onlineMeetings",
      requestBody: {
        startDateTime: start.toISOString(),
        endDateTime: end.toISOString(),
        subject: subject,
        // Add parameters that might influence URL format for personal accounts
        meetingType: "meetNow" // This might help generate personal meeting URLs
      }
    },
    {
      name: "Communications API",
      endpoint: "/communications/onlineMeetings",
      requestBody: {
        startTime: start.toISOString(),
        endTime: end.toISOString(),
        subject: subject
      }
    }
  ];
  
  for (const approach of approaches) {
    try {
      console.log(`Trying ${approach.name}...`);
      
      const meeting = await graphClient.api(approach.endpoint).post(approach.requestBody);
      
      const joinUrl = meeting.joinWebUrl || meeting.joinUrl;
      
      if (joinUrl) {
        console.log(`${approach.name} generated URL:`, joinUrl);
        
        if (joinUrl.includes('teams.live.com')) {
          console.log("✅ Successfully generated teams.live.com URL!");
        } else {
          console.log("⚠️ Generated URL uses different format:", joinUrl);
        }
        
        // Create calendar event
        const calendarEventRequest = {
          subject: subject,
          body: {
            contentType: "HTML",
            content: `
              <div>
                ${body ? `<p>${body}</p>` : ''}
                <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
                  <p><strong>Microsoft Teams Meeting</strong></p>
                  <p><a href="${joinUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
                  <p><strong>Meeting ID:</strong> ${meeting.id}</p>
                </div>
              </div>
            `
          },
          start: {
            dateTime: start.toISOString(),
            timeZone: "UTC"
          },
          end: {
            dateTime: end.toISOString(),
            timeZone: "UTC"
          },
          location: {
            displayName: "Microsoft Teams Meeting",
            locationUri: joinUrl
          },
          attendees: attendeeEmails.map(email => ({
            emailAddress: {
              address: email,
              name: email.split('@')[0]
            },
            type: "required"
          }))
        };

        const calendarEvent = await graphClient.api("/me/events").post(calendarEventRequest);

        return {
          success: true,
          method: approach.name.toLowerCase().replace(/\s+/g, '_'),
          accountType: accountInfo,
          eventId: calendarEvent.id,
          meetingId: meeting.id,
          subject: subject,
          joinWebUrl: joinUrl,
          joinUrl: joinUrl,
          urlFormat: joinUrl.includes('teams.live.com') ? 'teams.live.com' : 
                     joinUrl.includes('teams.microsoft.com') ? 'teams.microsoft.com' : 'unknown',
          isDesiredFormat: joinUrl.includes('teams.live.com'),
          startTime: { dateTime: start.toISOString() },
          endTime: { dateTime: end.toISOString() },
          attendees: calendarEvent.attendees,
          webLink: calendarEvent.webLink,
          onlineMeeting: meeting,
          calendarEvent: calendarEvent
        };
      }
    } catch (error) {
      console.log(`${approach.name} failed:`, error);
    }
  }
  
  throw new Error("Failed to create personal Teams meeting with teams.live.com URL format. This may require a personal Microsoft account or specific tenant configuration.");
}
// SOLUTION: Reliable Teams Meeting Creation with Join URL
// Replace your existing Teams meeting functions with these

export async function createTeamsMeetingReliable(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[] = [],
  body: string = ""
) {
  ensureClient(authProvider);
  
  console.log("Creating Teams meeting with reliable URL generation...");

  // STEP 1: Create the online meeting FIRST (this ensures we get a join URL)
  try {
    const onlineMeetingRequest = {
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      subject: subject
    };

    console.log("Creating online meeting...", onlineMeetingRequest);
    const onlineMeeting = await graphClient!.api("/me/onlineMeetings").post(onlineMeetingRequest);
    
    if (!onlineMeeting.joinWebUrl && !onlineMeeting.joinUrl) {
      throw new Error("Online meeting created but no join URL received");
    }
    
    const joinUrl = onlineMeeting.joinWebUrl || onlineMeeting.joinUrl;
    console.log("Online meeting created with join URL:", joinUrl);

    // STEP 2: Create calendar event that references the online meeting
    const calendarEventRequest = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: `
          <div>
            ${body ? `<p>${body}</p>` : ''}
            <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
              <p><strong>Microsoft Teams Meeting</strong></p>
              <p><a href="${joinUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
              ${onlineMeeting.videoTeleconferenceId ? `<p><strong>Conference ID:</strong> ${onlineMeeting.videoTeleconferenceId}</p>` : ''}
              ${onlineMeeting.tollNumber ? `<p><strong>Phone:</strong> ${onlineMeeting.tollNumber}</p>` : ''}
            </div>
          </div>
        `
      },
      start: {
        dateTime: start.toISOString(),
        timeZone: "UTC"
      },
      end: {
        dateTime: end.toISOString(),
        timeZone: "UTC"
      },
      location: {
        displayName: "Microsoft Teams Meeting",
        locationUri: joinUrl
      },
      attendees: attendeeEmails.map(email => ({
        emailAddress: {
          address: email,
          name: email.split('@')[0]
        },
        type: "required"
      })),
      // Mark as online meeting
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    };

    const calendarEvent = await graphClient!.api("/me/events").post(calendarEventRequest);
    console.log("Calendar event created successfully");

    return {
      success: true,
      eventId: calendarEvent.id,
      meetingId: onlineMeeting.id,
      subject: subject,
      joinWebUrl: joinUrl,
      joinUrl: joinUrl,
      conferenceId: onlineMeeting.videoTeleconferenceId,
      startTime: { dateTime: start.toISOString() },
      endTime: { dateTime: end.toISOString() },
      attendees: calendarEvent.attendees,
      webLink: calendarEvent.webLink,
      onlineMeeting: onlineMeeting,
      calendarEvent: calendarEvent
    };

  } catch (error: any) {
    console.error("Failed to create Teams meeting:", error);
    throw new Error(`Teams meeting creation failed: ${error.message}`);
  }
}

// Alternative approach using the beta endpoint (often more reliable)
export async function createTeamsMeetingBetaEndpoint(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[] = [],
  body: string = ""
) {
  ensureClient(authProvider);
  
  try {
    // Use BETA endpoint for more reliable Teams integration
    const meetingRequest = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body || "Microsoft Teams meeting"
      },
      start: {
        dateTime: start.toISOString(),
        timeZone: "UTC"
      },
      end: {
        dateTime: end.toISOString(),
        timeZone: "UTC"
      },
      attendees: attendeeEmails.map(email => ({
        emailAddress: {
          address: email,
          name: email.split('@')[0]
        },
        type: "required"
      })),
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    };

    console.log("Creating event via beta endpoint...");
    const response = await graphClient!.api("/beta/me/events").post(meetingRequest);
    
    // Beta endpoint often returns meeting details immediately
    if (response.onlineMeeting && (response.onlineMeeting.joinUrl || response.onlineMeeting.joinWebUrl)) {
      const joinUrl = response.onlineMeeting.joinUrl || response.onlineMeeting.joinWebUrl;
      
      return {
        success: true,
        method: "beta_endpoint",
        eventId: response.id,
        meetingId: response.onlineMeeting.conferenceId || response.id,
        subject: response.subject,
        joinWebUrl: joinUrl,
        joinUrl: joinUrl,
        conferenceId: response.onlineMeeting.conferenceId,
        startTime: response.start,
        endTime: response.end,
        attendees: response.attendees,
        webLink: response.webLink,
        onlineMeeting: response.onlineMeeting,
        calendarEvent: response
      };
    } else {
      throw new Error("Beta endpoint did not return meeting details immediately");
    }

  } catch (error: any) {
    console.error("Beta endpoint failed:", error);
    throw error;
  }
}

// Function to check if Teams meetings are properly configured for your tenant
export async function validateTeamsMeetingSetup(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  ensureClient(authProvider);
  
  const validationResults = {
    canCreateOnlineMeetings: false,
    canCreateCalendarEvents: false,
    teamsLicenseValid: false,
    organizationPolicies: 'unknown',
    error: null as string | null
  };

  try {
    // Test 1: Can we create online meetings?
    console.log("Testing online meetings creation...");
    const testMeeting = {
      startDateTime: new Date(Date.now() + 3600000).toISOString(),
      endDateTime: new Date(Date.now() + 7200000).toISOString(),
      subject: "Validation Test - Will Delete"
    };

    const meeting = await graphClient!.api("/me/onlineMeetings").post(testMeeting);
    
    if (meeting.joinWebUrl || meeting.joinUrl) {
      validationResults.canCreateOnlineMeetings = true;
      console.log("✅ Online meetings working");
      
      // Clean up
      try {
        await graphClient!.api(`/me/onlineMeetings/${meeting.id}`).delete();
      } catch (deleteError) {
        console.log("Could not delete test meeting");
      }
    }

  } catch (error: any) {
    validationResults.error = error.message;
    console.error("❌ Online meetings failed:", error.message);
    
    // Check for specific error patterns
    if (error.message.includes('Forbidden')) {
      validationResults.organizationPolicies = 'blocked';
    } else if (error.message.includes('license')) {
      validationResults.teamsLicenseValid = false;
    }
  }

  // Test 2: Can we create calendar events?
  try {
    console.log("Testing calendar events creation...");
    const testEvent = {
      subject: "Calendar Test - Will Delete",
      start: {
        dateTime: new Date(Date.now() + 3600000).toISOString(),
        timeZone: "UTC"
      },
      end: {
        dateTime: new Date(Date.now() + 7200000).toISOString(),
        timeZone: "UTC"
      }
    };

    const event = await graphClient!.api("/me/events").post(testEvent);
    validationResults.canCreateCalendarEvents = true;
    console.log("✅ Calendar events working");
    
    // Clean up
    try {
      await graphClient!.api(`/me/events/${event.id}`).delete();
    } catch (deleteError) {
      console.log("Could not delete test event");
    }

  } catch (error: any) {
    console.error("❌ Calendar events failed:", error.message);
  }

  return validationResults;
}

// Simplified function that tries multiple approaches
export async function createTeamsMeetingWithFallback(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[] = [],
  body: string = ""
) {
  console.log("Attempting Teams meeting creation with multiple fallback methods...");

  // Method 1: Try the reliable approach (online meeting first)
  try {
    console.log("Method 1: Online meeting first approach...");
    return await createTeamsMeetingReliable(authProvider, subject, start, end, attendeeEmails, body);
  } catch (error1: any) {
    console.log("Method 1 failed:", error1.message);
  }

  // Method 2: Try beta endpoint
  try {
    console.log("Method 2: Beta endpoint approach...");
    return await createTeamsMeetingBetaEndpoint(authProvider, subject, start, end, attendeeEmails, body);
  } catch (error2: any) {
    console.log("Method 2 failed:", error2.message);
  }

  // Method 3: Try calendar-first with retry
  try {
    console.log("Method 3: Calendar-first with retry...");
    return await createCalendarFirstWithRetry(authProvider, subject, start, end, attendeeEmails, body);
  } catch (error3: any) {
    console.log("Method 3 failed:", error3.message);
  }

  throw new Error("All Teams meeting creation methods failed. Your organization may have restrictions on creating Teams meetings via API, or you may need additional licenses/permissions.");
}

// Helper function for calendar-first approach with retry logic
async function createCalendarFirstWithRetry(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[],
  body: string
) {
  ensureClient(authProvider);

  const calendarEvent = {
    subject: subject,
    body: {
      contentType: "HTML",
      content: body || "Microsoft Teams meeting"
    },
    start: {
      dateTime: start.toISOString(),
      timeZone: "UTC"
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: "UTC"
    },
    attendees: attendeeEmails.map(email => ({
      emailAddress: {
        address: email,
        name: email.split('@')[0]
      },
      type: "required"
    })),
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness"
  };

  const createdEvent = await graphClient!.api("/me/events").post(calendarEvent);
  console.log("Calendar event created, waiting for Teams meeting provisioning...");

  // Retry logic to wait for Teams meeting URL
  for (let i = 0; i < 10; i++) {
    await new Promise(resolve => setTimeout(resolve, 2000 + (i * 1000))); // Increasing delay
    
    try {
      const eventWithMeeting = await graphClient!.api(`/me/events/${createdEvent.id}`)
        .select('id,subject,start,end,onlineMeeting,attendees,webLink,isOnlineMeeting')
        .expand('onlineMeeting')
        .get();

      if (eventWithMeeting.onlineMeeting && (eventWithMeeting.onlineMeeting.joinUrl || eventWithMeeting.onlineMeeting.joinWebUrl)) {
        console.log(`Teams meeting provisioned after ${i + 1} retries`);
        
        const joinUrl = eventWithMeeting.onlineMeeting.joinUrl || eventWithMeeting.onlineMeeting.joinWebUrl;
        
        return {
          success: true,
          method: "calendar_with_retry",
          eventId: eventWithMeeting.id,
          meetingId: eventWithMeeting.onlineMeeting.conferenceId || eventWithMeeting.id,
          subject: eventWithMeeting.subject,
          joinWebUrl: joinUrl,
          joinUrl: joinUrl,
          conferenceId: eventWithMeeting.onlineMeeting.conferenceId,
          startTime: eventWithMeeting.start,
          endTime: eventWithMeeting.end,
          attendees: eventWithMeeting.attendees,
          webLink: eventWithMeeting.webLink,
          onlineMeeting: eventWithMeeting.onlineMeeting,
          calendarEvent: eventWithMeeting,
          retriesNeeded: i + 1
        };
      }
    } catch (fetchError) {
      console.log(`Retry ${i + 1} failed:`, fetchError);
    }
  }

  throw new Error("Teams meeting URL was not generated after maximum retries");
}

  const buildRecurrence = (rule: string, event: any) => {
    const [freqPart, intervalPart, countPart] = rule.split(';');
    const frequency = freqPart.split('=')[1].toLowerCase();
    const interval = parseInt(intervalPart.split('=')[1]);
    const count = parseInt(countPart.split('=')[1]);

    return {
      pattern: {
        type: frequency,
        interval: interval,
      },
      range: {
        type: "numbered",
        startDate: event.start.toISOString().split('T')[0],
        numberOfOccurrences: count,
      },
    };
  };

  const timeZone = (tz: string | undefined) => (tz === null || tz === undefined) ? 'UTC' : tz;

// helpers
const withTeams = (makeTeams?: boolean) =>
  makeTeams
    ? { isOnlineMeeting: true as const, onlineMeetingProvider: "teamsForBusiness" as const }
    : {};

export async function CreateEvent(event: any, makeTeams = true) {
  const event1 = {
    subject: `${event.subject}`,
    start: {
      dateTime: event.start.toISOString(),
      timeZone: timeZone(event.StartTimezone), // your helper
    },
    end: {
      dateTime: event.end.toISOString(),
      timeZone: timeZone(event.EndTimezone),
    },
    ...(event.RecurrenceRule && { recurrence: buildRecurrence(event.RecurrenceRule, event) }),
    ...withTeams(makeTeams),
  };

  const created = await graphClient!.api("/me/events").post(event1);

  // IMPORTANT: expand the onlineMeeting nav property to actually get the joinUrl
  const createdWithMeeting = await graphClient!
    .api(`/me/events/${created.id}`)
    .expand("onlineMeeting")
    .select("id,subject,start,end,webLink,isOnlineMeeting,onlineMeetingProvider,onlineMeeting")
    .get();

  // If still empty (can happen while Teams provisions), try the direct relationship
  let joinUrl =
    createdWithMeeting?.onlineMeeting?.joinUrl ??
    createdWithMeeting?.onlineMeeting?.joinWebUrl;

  if (!joinUrl) {
    try {
      const om = await graphClient!.api(`/me/events/${created.id}/onlineMeeting`).get();
      joinUrl = om?.joinUrl ?? om?.joinWebUrl;
    } catch {
      // swallow – it may not exist yet
    }
  }

  return { ...createdWithMeeting, joinUrl: joinUrl ?? null };
}

export async function updateEvent(event: any, makeTeams = true) {
  if (!event || !event.id) throw new Error("Event ID is required.");

  const event1 = {
    subject: `${event.subject}`,
    start: {
      dateTime: event.start.toISOString(),
      timeZone: timeZone(event.StartTimezone),
    },
    end: {
      dateTime: event.end.toISOString(),
      timeZone: timeZone(event.EndTimezone),
    },
    ...(event.RecurrenceRule && { recurrence: buildRecurrence(event.RecurrenceRule, event) }),
    ...withTeams(makeTeams),
  };

  await graphClient!.api(`/me/events/${event.id}`).patch(event1);

  const updated = await graphClient!
    .api(`/me/events/${event.id}`)
    .expand("onlineMeeting")
    .select("id,subject,start,end,webLink,isOnlineMeeting,onlineMeetingProvider,onlineMeeting")
    .get();

  let joinUrl = updated?.onlineMeeting?.joinUrl ?? updated?.onlineMeeting?.joinWebUrl;
  if (!joinUrl) {
    try {
      const om = await graphClient!.api(`/me/events/${event.id}/onlineMeeting`).get();
      joinUrl = om?.joinUrl ?? om?.joinWebUrl;
    } catch {}
  }

  return { ...updated, joinUrl: joinUrl ?? null };
}

const fetchEventWithJoin = async (eventId: string) => {
  // First try expand
  const ev = await graphClient!
    .api(`/me/events/${eventId}`)
    .expand("onlineMeeting")
    .select("id,subject,start,end,attendees,webLink,isOnlineMeeting,onlineMeetingProvider,onlineMeeting")
    .get();

  let joinUrl = ev?.onlineMeeting?.joinUrl ?? ev?.onlineMeeting?.joinWebUrl;

  if (!joinUrl) {
    // Fallback: call the relationship
    try {
      const om = await graphClient!.api(`/me/events/${eventId}/onlineMeeting`).get();
      joinUrl = om?.joinUrl ?? om?.joinWebUrl;
    } catch {}
  }

  return { ...ev, joinUrl: joinUrl ?? null };
};



  export async function deleteEvent(id: any) {
    return await graphClient!.api(`/me/events/${id}`).delete();
  }

  // Get all active subscriptions for the current user
  export async function listSubscriptions(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    ensureClient(authProvider);

    const response = await graphClient!.api("/subscriptions").get();
    return response.value; // returns array of subscriptions
  }

  // Function to check what permissions the current token has
  export async function checkPermissions(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    ensureClient(authProvider);
    
    const permissions = {
      calendars: false,
      onlineMeetings: false,
      mail: false,
      user: false,
      communications: false
    };
    
    try {
      // Test user permission
      await graphClient!.api("/me").get();
      permissions.user = true;
    } catch (e) {
      console.log("User.Read permission not available");
    }
    
    try {
      // Test calendar permission
      await graphClient!.api("/me/calendar").get();
      permissions.calendars = true;
    } catch (e) {
      console.log("Calendars.ReadWrite permission not available");
    }
    
    try {
      // Test online meetings permission
      await graphClient!.api("/me/onlineMeetings").get();
      permissions.onlineMeetings = true;
    } catch (e) {
      console.log("OnlineMeetings.ReadWrite permission not available");
    }
    
    try {
      // Test communications permission
      await graphClient!.api("/communications/onlineMeetings").get();
      permissions.communications = true;
    } catch (e) {
      console.log("OnlineMeetings.ReadWrite (Communications) permission not available");
    }
    
    try {
      // Test mail permission
      await graphClient!.api("/me/mailFolders").get();
      permissions.mail = true;
    } catch (e) {
      console.log("Mail.Read permission not available");
    }
    
    return permissions;
  }

  // Utility function to get meeting details by event ID
  export async function getMeetingDetails(
    authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    eventId: string
  ) {
    ensureClient(authProvider);
    
    try {
      const event = await graphClient!.api(`/me/events/${eventId}`)
        .select('id,subject,start,end,onlineMeeting,attendees,webLink,location')
        .get();
      
      return {
        id: event.id,
        subject: event.subject,
        startTime: event.start,
        endTime: event.end,
        joinWebUrl: event.onlineMeeting?.joinUrl,
        meetingId: event.onlineMeeting?.conferenceId || event.id,
        attendees: event.attendees,
        webLink: event.webLink,
        location: event.location,
        onlineMeeting: event.onlineMeeting
      };
    } catch (error: any) {
      console.error("Error fetching meeting details:", error);
      throw new Error(`Failed to fetch meeting details: ${error.message}`);
    }
  }



  // Teams Meeting Utility Functions - Additional GraphService.ts functions

// Function to check if a user has the necessary permissions
export async function checkTeamsMeetingPermissions(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider
): Promise<{
  canCreateMeetings: boolean;
  canCreateCalendar: boolean;
  canReadUser: boolean;
  missingPermissions: string[];
  recommendations: string[];
}> {
  ensureClient(authProvider);
  
  const results = {
    canCreateMeetings: false,
    canCreateCalendar: false,
    canReadUser: false,
    missingPermissions: [] as string[],
    recommendations: [] as string[]
  };
  
  // Test user read permission
  try {
    await graphClient!.api("/me").select("displayName,mail").get();
    results.canReadUser = true;
  } catch (error: any) {
    results.missingPermissions.push("User.Read");
    console.log("User.Read permission missing:", error.message);
  }
  
  // Test calendar permissions
  try {
    await graphClient!.api("/me/calendar").get();
    results.canCreateCalendar = true;
  } catch (error: any) {
    results.missingPermissions.push("Calendars.ReadWrite");
    console.log("Calendars.ReadWrite permission missing:", error.message);
  }
  
  // Test online meetings permission
  try {
    const testMeeting = {
      startDateTime: new Date(Date.now() + 3600000).toISOString(), // 1 hour from now
      endDateTime: new Date(Date.now() + 7200000).toISOString(),   // 2 hours from now
      subject: "Permission Test Meeting - Will be deleted"
    };
    
    const meeting = await graphClient!.api("/me/onlineMeetings").post(testMeeting);
    
    // Clean up test meeting
    try {
      await graphClient!.api(`/me/onlineMeetings/${meeting.id}`).delete();
    } catch (deleteError) {
      console.log("Could not delete test meeting:", deleteError);
    }
    
    results.canCreateMeetings = true;
  } catch (error: any) {
    results.missingPermissions.push("OnlineMeetings.ReadWrite");
    console.log("OnlineMeetings.ReadWrite permission missing:", error.message);
  }
  
  // Generate recommendations
  if (!results.canCreateMeetings) {
    results.recommendations.push("Request OnlineMeetings.ReadWrite permission from your administrator");
    results.recommendations.push("This permission is essential for creating Teams meetings with join URLs");
  }
  
  if (!results.canCreateCalendar) {
    results.recommendations.push("Request Calendars.ReadWrite permission to create calendar events");
  }
  
  if (!results.canReadUser) {
    results.recommendations.push("Request User.Read permission for basic user information");
  }
  
  if (results.missingPermissions.length === 0) {
    results.recommendations.push("All required permissions are available!");
  }
  
  return results;
}

// Function to validate a meeting before creation
export function validateMeetingData(
  subject: string,
  start: Date,
  end: Date,
  attendeeEmails: string[]
): { isValid: boolean; errors: string[] } {
  const errors: string[] = [];
  
  // Validate subject
  if (!subject || subject.trim().length === 0) {
    errors.push("Meeting subject is required");
  } else if (subject.length > 255) {
    errors.push("Meeting subject is too long (max 255 characters)");
  }
  
  // Validate dates
  if (isNaN(start.getTime())) {
    errors.push("Invalid start date");
  }
  
  if (isNaN(end.getTime())) {
    errors.push("Invalid end date");
  }
  
  if (start.getTime() && end.getTime()) {
    if (start >= end) {
      errors.push("Start time must be before end time");
    }
    
    if (start.getTime() < Date.now() - 300000) { // 5 minutes ago
      errors.push("Start time cannot be more than 5 minutes in the past");
    }
    
    const duration = end.getTime() - start.getTime();
    const maxDuration = 24 * 60 * 60 * 1000; // 24 hours
    
    if (duration > maxDuration) {
      errors.push("Meeting duration cannot exceed 24 hours");
    }
    
    if (duration < 60000) { // 1 minute
      errors.push("Meeting must be at least 1 minute long");
    }
  }
  
  // Validate attendee emails
  attendeeEmails.forEach((email, index) => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (email && !emailRegex.test(email)) {
      errors.push(`Invalid email format: ${email}`);
    }
  });
  
  if (attendeeEmails.length > 250) {
    errors.push("Too many attendees (maximum 250)");
  }
  
  return {
    isValid: errors.length === 0,
    errors
  };
}

// Function to get detailed meeting information
export async function getDetailedMeetingInfo(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  eventId: string
): Promise<any> {
  ensureClient(authProvider);
  
  try {
    // Get the calendar event
    const event = await graphClient!.api(`/me/events/${eventId}`)
      .select('id,subject,start,end,attendees,organizer,location,body,webLink,isOnlineMeeting,onlineMeetingProvider')
      .expand('onlineMeeting')
      .get();
    
    let onlineMeetingDetails = null;
    
    // Try to get online meeting details if available
    if (event.onlineMeeting) {
      onlineMeetingDetails = event.onlineMeeting;
    } else if (event.isOnlineMeeting) {
      try {
        onlineMeetingDetails = await graphClient!.api(`/me/events/${eventId}/onlineMeeting`).get();
      } catch (error) {
        console.log("Could not fetch online meeting details:", error);
      }
    }
    
    return {
      event,
      onlineMeetingDetails,
      hasValidJoinUrl: !!(onlineMeetingDetails?.joinUrl || onlineMeetingDetails?.joinWebUrl),
      joinUrl: onlineMeetingDetails?.joinUrl || onlineMeetingDetails?.joinWebUrl,
      meetingId: onlineMeetingDetails?.conferenceId || onlineMeetingDetails?.videoTeleconferenceId || event.id,
      conferenceId: onlineMeetingDetails?.conferenceId,
      audioConferencing: onlineMeetingDetails?.audioConferencing
    };
  } catch (error: any) {
    console.error("Error fetching detailed meeting info:", error);
    throw new Error(`Failed to get meeting details: ${error.message}`);
  }
}

// Function to repair a meeting that doesn't have a proper Teams join URL
export async function repairTeamsMeeting(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  eventId: string
): Promise<any> {
  ensureClient(authProvider);
  
  try {
    // Get the existing event
    const existingEvent = await graphClient!.api(`/me/events/${eventId}`)
      .select('id,subject,start,end,attendees,body,location')
      .get();
    
    console.log("Repairing Teams meeting for event:", existingEvent.id);
    
    // Create a new online meeting
    const onlineMeetingRequest = {
      subject: existingEvent.subject,
      startDateTime: existingEvent.start.dateTime,
      endDateTime: existingEvent.end.dateTime
    };
    
    const newOnlineMeeting = await graphClient!.api("/me/onlineMeetings").post(onlineMeetingRequest);
    
    // Update the calendar event with the new meeting details
    const updateRequest = {
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
      location: {
        displayName: "Microsoft Teams Meeting",
        locationUri: newOnlineMeeting.joinWebUrl
      },
      body: {
        contentType: "HTML",
        content: `
          ${existingEvent.body?.content || ''}
          <div style="margin: 20px 0; padding: 15px; background-color: #f3f2f1; border-left: 4px solid #6264a7;">
            <p><strong>Microsoft Teams Meeting</strong></p>
            <p><a href="${newOnlineMeeting.joinWebUrl}" style="color: #6264a7; font-weight: bold;">Join Microsoft Teams Meeting</a></p>
            <p><strong>Meeting ID:</strong> ${newOnlineMeeting.videoTeleconferenceId || newOnlineMeeting.id}</p>
          </div>
        `
      }
    };
    
    const updatedEvent = await graphClient!.api(`/me/events/${eventId}`).patch(updateRequest);
    
    return {
      repaired: true,
      eventId: updatedEvent.id,
      joinWebUrl: newOnlineMeeting.joinWebUrl,
      meetingId: newOnlineMeeting.videoTeleconferenceId || newOnlineMeeting.id,
      onlineMeeting: newOnlineMeeting,
      updatedEvent: updatedEvent
    };
    
  } catch (error: any) {
    console.error("Error repairing Teams meeting:", error);
    throw new Error(`Failed to repair Teams meeting: ${error.message}`);
  }
}

// Function to bulk create multiple Teams meetings
export async function bulkCreateTeamsMeetings(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  meetings: Array<{
    subject: string;
    start: Date;
    end: Date;
    attendeeEmails?: string[];
    body?: string;
  }>
): Promise<Array<{ success: boolean; meeting?: any; error?: string; originalIndex: number }>> {
  ensureClient(authProvider);
  
  const results = [];
  
  for (let i = 0; i < meetings.length; i++) {
    const meeting = meetings[i];
    
    try {
      console.log(`Creating meeting ${i + 1} of ${meetings.length}: ${meeting.subject}`);
      
      const createdMeeting = await createValidTeamsMeeting(
        authProvider,
        meeting.subject,
        meeting.start,
        meeting.end,
        meeting.attendeeEmails || [],
        meeting.body || ""
      );
      
      results.push({
        success: true,
        meeting: createdMeeting,
        originalIndex: i
      });
      
      // Add delay between requests to avoid rate limiting
      if (i < meetings.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
      
    } catch (error: any) {
      console.error(`Failed to create meeting ${i + 1}:`, error);
      results.push({
        success: false,
        error: error.message,
        originalIndex: i
      });
    }
  }
  
  return results;
}

// Fix 1: Correct the testAllPermissions function return types
export async function testAllPermissions(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  ensureClient(authProvider);
  
  const results = {
    userRead: { success: false, error: null as string | null },
    calendarRead: { success: false, error: null as string | null },
    calendarWrite: { success: false, error: null as string | null },
    onlineMeetingsRead: { success: false, error: null as string | null },
    onlineMeetingsWrite: { success: false, error: null as string | null }
  };

  // Test User.Read
  try {
    const user = await graphClient!.api('/me').select('displayName,mail').get();
    results.userRead.success = true;
    console.log('✅ User.Read working:', user.displayName);
  } catch (error: any) {
    results.userRead.error = error.message;
    console.error('❌ User.Read failed:', error.message);
  }

  // Test Calendars.Read
  try {
    const calendar = await graphClient!.api('/me/calendar').get();
    results.calendarRead.success = true;
    console.log('✅ Calendars.Read working');
  } catch (error: any) {
    results.calendarRead.error = error.message;
    console.error('❌ Calendars.Read failed:', error.message);
  }

  // Test Calendars.ReadWrite (try to get events)
  try {
    const events = await graphClient!.api('/me/events').top(1).get();
    results.calendarWrite.success = true;
    console.log('✅ Calendar access working');
  } catch (error: any) {
    results.calendarWrite.error = error.message;
    console.error('❌ Calendar access failed:', error.message);
  }

  // Test OnlineMeetings.Read
  try {
    const meetings = await graphClient!.api('/me/onlineMeetings').top(1).get();
    results.onlineMeetingsRead.success = true;
    console.log('✅ OnlineMeetings.Read working');
  } catch (error: any) {
    results.onlineMeetingsRead.error = error.message;
    console.error('❌ OnlineMeetings.Read failed:', error.message);
  }

  // Test OnlineMeetings.ReadWrite (try to create a test meeting)
  try {
    const testMeeting = {
      startDateTime: new Date(Date.now() + 3600000).toISOString(),
      endDateTime: new Date(Date.now() + 7200000).toISOString(),
      subject: 'Permission Test - Will Delete'
    };

    const meeting = await graphClient!.api('/me/onlineMeetings').post(testMeeting);
    results.onlineMeetingsWrite.success = true;
    console.log('✅ OnlineMeetings.ReadWrite working:', meeting.id);

    // Clean up
    try {
      await graphClient!.api(`/me/onlineMeetings/${meeting.id}`).delete();
      console.log('✅ Test meeting cleaned up');
    } catch (deleteError) {
      console.log('⚠️ Could not delete test meeting:', deleteError);
    }

  } catch (error: any) {
    results.onlineMeetingsWrite.error = error.message;
    console.error('❌ OnlineMeetings.ReadWrite failed:', error.message);
  }

  return results;
}

// Fix 2: Update the testTeamsMeetingCapabilities function with proper type assertions
export async function testTeamsMeetingCapabilities(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider
): Promise<{
  overallStatus: 'working' | 'partial' | 'not_working';
  tests: Array<{
    name: string;
    status: 'passed' | 'failed' | 'warning';
    message: string;
    details?: any;
  }>;
}> {
  ensureClient(authProvider);
  
  const tests: Array<{
    name: string;
    status: 'passed' | 'failed' | 'warning';
    message: string;
    details?: any;
  }> = [];
  
  // Test 1: Check permissions
  try {
    const permissions = await checkTeamsMeetingPermissions(authProvider);
    
    if (permissions.canCreateMeetings && permissions.canCreateCalendar && permissions.canReadUser) {
      tests.push({
        name: "Permissions Check",
        status: 'passed' as const,
        message: "All required permissions are available",
        details: permissions
      });
    } else {
      tests.push({
        name: "Permissions Check",
        status: 'failed' as const,
        message: `Missing permissions: ${permissions.missingPermissions.join(', ')}`,
        details: permissions
      });
    }
  } catch (error: any) {
    tests.push({
      name: "Permissions Check",
      status: 'failed' as const,
      message: `Permission check failed: ${error.message}`
    });
  }
  
  // Test 2: Try creating a test meeting
  try {
    const testStart = new Date(Date.now() + 3600000); // 1 hour from now
    const testEnd = new Date(Date.now() + 7200000);   // 2 hours from now
    
    const testMeeting = await createValidTeamsMeeting(
      authProvider,
      "Test Meeting - Will be deleted",
      testStart,
      testEnd,
      [],
      "This is a test meeting to verify Teams integration"
    );
    
    if (testMeeting.joinWebUrl) {
      tests.push({
        name: "Meeting Creation Test",
        status: 'passed' as const,
        message: "Successfully created Teams meeting with join URL",
        details: {
          method: testMeeting.method,
          hasJoinUrl: !!testMeeting.joinWebUrl,
          meetingId: testMeeting.meetingId
        }
      });
      
      // Clean up test meeting
      try {
        await graphClient!.api(`/me/events/${testMeeting.eventId}`).delete();
      } catch (deleteError) {
        console.log("Could not delete test meeting:", deleteError);
      }
    } else {
      tests.push({
        name: "Meeting Creation Test",
        status: 'warning' as const,
        message: "Meeting created but no join URL generated",
        details: testMeeting
      });
    }
    
  } catch (error: any) {
    tests.push({
      name: "Meeting Creation Test",
      status: 'failed' as const,
      message: `Meeting creation failed: ${error.message}`
    });
  }
  
  // Determine overall status
  const passedTests = tests.filter(t => t.status === 'passed').length;
  const failedTests = tests.filter(t => t.status === 'failed').length;
  
  let overallStatus: 'working' | 'partial' | 'not_working';
  
  if (failedTests === 0) {
    overallStatus = 'working';
  } else if (passedTests > 0) {
    overallStatus = 'partial';
  } else {
    overallStatus = 'not_working';
  }
  
  return {
    overallStatus,
    tests
  };
}

// Fix 3: Improve the getTokenInfo function to handle MSAL properly
export async function getTokenInfo(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  try {
    // Get the MSAL instance from the auth provider
    const msalInstance = (authProvider as any).msalInstance;
    const accounts = msalInstance.getAllAccounts();
    
    if (accounts.length === 0) {
      console.error('No accounts found. User needs to sign in.');
      return false;
    }
    
    console.log('Current account:', accounts[0].username);
    
    // Try to get a fresh token with specific scopes
    const tokenRequest = {
      scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Calendars.ReadWrite',
        'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
      ],
      account: accounts[0],
      forceRefresh: true
    };
    
    const response = await msalInstance.acquireTokenSilent(tokenRequest);
    console.log('Token acquired successfully');
    
    // Decode the token to see what permissions it contains (just for debugging)
    if (response.accessToken) {
      const tokenParts = response.accessToken.split('.');
      if (tokenParts.length === 3) {
        const payload = JSON.parse(atob(tokenParts[1]));
        console.log('Token scopes:', payload.scp || payload.roles);
        console.log('Token expires:', new Date(payload.exp * 1000));
        console.log('Token audience:', payload.aud);
      }
    }
    
    return true;
  } catch (error: any) {
    console.error('Token acquisition failed:', error);
    
    // If silent token acquisition fails, might need interactive login
    if (error.name === 'InteractionRequiredAuthError') {
      console.log('Interactive login required');
      // You might want to trigger interactive login here
    }
    
    return false;
  }
}

// Fix 4: Add a function to force token refresh and re-authentication
export async function forceTokenRefresh(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  try {
    const msalInstance = (authProvider as any).msalInstance;
    const accounts = msalInstance.getAllAccounts();
    
    if (accounts.length === 0) {
      // No accounts, need to login
      const loginRequest = {
        scopes: [
          'https://graph.microsoft.com/User.Read',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
        ]
      };
      
      const response = await msalInstance.loginPopup(loginRequest);
      console.log('Interactive login successful:', response.account?.username);
      return true;
    } else {
      // Force refresh existing token
      const tokenRequest = {
        scopes: [
          'https://graph.microsoft.com/User.Read',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
        ],
        account: accounts[0],
        forceRefresh: true
      };
      
      try {
        const response = await msalInstance.acquireTokenSilent(tokenRequest);
        console.log('Token refresh successful');
        return true;
      } catch (silentError: any) {
        if (silentError.name === 'InteractionRequiredAuthError') {
          // Need interactive token acquisition
          const response = await msalInstance.acquireTokenPopup(tokenRequest);
          console.log('Interactive token acquisition successful');
          return true;
        }
        throw silentError;
      }
    }
  } catch (error: any) {
    console.error('Token refresh failed:', error);
    return false;
  }
}

// Fix 5: Add a comprehensive permission diagnostic function
export async function diagnosePermissionIssues(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  console.log('🔍 Starting permission diagnosis...');
  
  // Step 1: Check token
  const tokenValid = await getTokenInfo(authProvider);
  console.log('Token status:', tokenValid ? '✅ Valid' : '❌ Invalid');
  
  // Step 2: Test all permissions
  const permissionResults = await testAllPermissions(authProvider);
  
  // Step 3: Generate report
  const report = {
    tokenValid,
    permissions: permissionResults,
    recommendations: [] as string[]
  };
  
  // Generate recommendations based on failures
  if (!permissionResults.userRead.success) {
    report.recommendations.push('❌ User.Read permission missing - check Azure app registration');
  }
  
  if (!permissionResults.calendarRead.success || !permissionResults.calendarWrite.success) {
    report.recommendations.push('❌ Calendar permissions missing - ensure Calendars.ReadWrite is granted');
  }
  
  if (!permissionResults.onlineMeetingsRead.success || !permissionResults.onlineMeetingsWrite.success) {
    report.recommendations.push('❌ OnlineMeetings permissions missing - this is critical for Teams meeting creation');
    report.recommendations.push('💡 Try waiting 10-15 minutes after granting permissions for them to propagate');
    report.recommendations.push('💡 Clear browser cache and sign out/in again');
  }
  
  if (report.recommendations.length === 0) {
    report.recommendations.push('✅ All permissions appear to be working correctly!');
  }
  
  return report;
}