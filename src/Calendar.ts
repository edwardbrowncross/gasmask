const calendars: Record<string, CalendarCalendar & { _events: Record<string, CalendarEvent> }> = {};

type ResponseStatus = 'needsAction' | 'declined' | 'tentative' | 'accepted';

type CalendarEvent = {
  id?: string,
  summary?: string,
  description?: string,
  location?: string,
  start: {
    dateTime: string,
  },
  end: {
    dateTime: string,
  },
  attendees?: {
    email: string,
    optional?: boolean,
    responseStatus?: ResponseStatus,
    comment?: string,
  }[],
}

class Events {
  static get(calId: string, eventId: string): CalendarEvent {
    if (!calendars[calId]) {
      throw new Error(`Calendar ${calId} not found`);
    }
    const evt = calendars[calId]._events[eventId];
    if (!evt) {
      throw new Error(`Event ${eventId} not found`);
    }
    return evt;
  }

  static remove(calId: string, eventId: string): void {
    // Calling get here makes it throw when parameters are invalid
    Events.get(calId, eventId);
    delete calendars[calId]._events[eventId];
  }

  static insert(event: CalendarEvent, calId: string): CalendarEvent {
    if (!calendars[calId]) {
      throw new Error(`Calendar ${calId} not found`);
    }
    const newEvent = {
      id: Math.random().toString(36).slice(2, 9),
      summary: '',
      description: '',
      location: '',
      attendees: [] as CalendarEvent['attendees'],
      ...event,
    }
    newEvent.attendees = newEvent.attendees!.map((suppliedAttendee) => ({
      optional: false,
      responseStatus: 'needsAction',
      comment: '',
      ...suppliedAttendee
    }));
    calendars[calId]._events[newEvent.id] = newEvent;
    return newEvent;
  }

  static patch(event: Partial<CalendarEvent>, calId: string, eventId: string): CalendarEvent {
    const existingEvent = Events.get(calId, eventId);
    const newEvent = {
      ...existingEvent,
      ...event,
    }
    calendars[calId]._events[eventId] = newEvent;
    return newEvent;
  }
}

type CalendarCalendar = {
  id: string,
  summary: string,
}

class Calendars {
  static insert(calendar: CalendarCalendar): CalendarCalendar {
    calendars[calendar.id] = {
      ...calendar,
      _events: {},
    };
    return calendars[calendar.id];
  }
}

/**
 * Advanced Calendar API
 * @link https://developers.google.com/apps-script/advanced/calendar
 */
export default class Calendar {
  static Events = Events;
  static Calendars = Calendars;
}