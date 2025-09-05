import { useEffect, useRef, useState } from "react";
import { findIana } from "windows-iana";
import { Event } from "@microsoft/microsoft-graph-types";
import { AuthenticatedTemplate } from "@azure/msal-react";

import {
  CreateEvent,
  deleteEvent,
  getUserWeekCalendar,
  updateEvent,
} from "./GraphService";
import { useAppContext } from "./AppContext";
import "./Scheduler.css";
import "./App.css";

import {
  ScheduleComponent,
  Day,
  Week,
  WorkWeek,
  Month,
  Agenda,
  Inject,
} from "@syncfusion/ej2-react-schedule";

export default function Scheduler() {
  const app = useAppContext();

  const [events, setEvents] = useState<Event[]>();
  const [isLoading, setIsLoading] = useState(false);
  let scheduleObj = useRef<ScheduleComponent>(null);

  useEffect(() => {
    const loadEvents = async () => {
      if (app.user && !events && !isLoading) {
        try {
          setIsLoading(true);
          
          let startDate: any = scheduleObj.current?.getCurrentViewDates()[0];
          let endDate: any = scheduleObj.current
            ?.getCurrentViewDates()
            .slice(-1)[0];
          
          // Add null checks
          if (!startDate || !endDate) {
            console.warn('Schedule dates not available yet');
            setIsLoading(false);
            return;
          }
          
          const ianaTimeZones = findIana(app.user?.timeZone!);
          const fetchedEvents = await getUserWeekCalendar(
            app.authProvider!,
            ianaTimeZones[0].valueOf(),
            startDate,
            endDate
          );
          
          // Validate events structure
          const validEvents = fetchedEvents?.filter(event => 
            event && typeof event === 'object'
          ) || [];
          
          console.log('Events data:', validEvents);
          setEvents(validEvents);
        } catch (err) {
          const error = err as Error;
          app.displayError!(error.message);
        } finally {
          setIsLoading(false);
        }
      }
    };

    loadEvents();
  }, [app.user, events, app.authProvider, app.displayError, isLoading]); // Add dependencies

  const fieldsData = {
    id: "Id",
    subject: { name: "subject" },
    startTime: { name: "start" },
    endTime: { name: "end" },
  };

  const onActionComplete = async (args: any): Promise<void> => {
    let startDate: any;
    let endDate: any;
    
    try {
      if (args.requestType === "eventCreated") {
        const event = args.data[0];
        await CreateEvent(event);
      } else if (args.requestType === "eventChanged") {
        const event = args.data[0];
        await updateEvent(event);
      } else if (args.requestType === "eventRemoved") {
        var eventId = args.data[0].id;
        await deleteEvent(eventId);
      }
      
      // Re-fetch events after any modification
      if (app.user) {
        const ianaTimeZones = findIana(app.user?.timeZone!);
        startDate = scheduleObj.current?.getCurrentViewDates()[0];
        endDate = scheduleObj.current?.getCurrentViewDates().slice(-1)[0];

        // Add null checks
        if (startDate && endDate) {
          const updatedEvents = await getUserWeekCalendar(
            app.authProvider!,
            ianaTimeZones[0].valueOf(),
            startDate,
            endDate
          );
          
          // Validate events structure
          const validUpdatedEvents = updatedEvents?.filter(event => 
            event && typeof event === 'object'
          ) || [];
          
          setEvents(validUpdatedEvents);
        }
      }
    } catch (err) {
      const error = err as Error;
      app.displayError!(error.message);
    }
  };

  const eventSettings = { 
    dataSource: events || [], // Ensure it's never undefined
    fields: fieldsData 
  };

  return (
    <AuthenticatedTemplate>
      {app.user && (
        <ScheduleComponent
          ref={scheduleObj}
          eventSettings={eventSettings}
          height="100vh"
          actionComplete={onActionComplete}
        >
          <Inject services={[Day, Week, WorkWeek, Month, Agenda]} />
        </ScheduleComponent>
      )}
    </AuthenticatedTemplate>
  );
}