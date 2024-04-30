import defineAction from '../../../../helpers/define-action.js';
import { DateTime } from 'luxon';

const MAX_EVENTS = 50;
const FORMAT_TIMESTAMP = 'yyyy-MM-dd\'T\'HH:mm:ssZZ';

export default defineAction({
  name: 'List Events',
  key: 'listEvents',
  description: 'Lists events from a calendar.',
  arguments: [
    {
      label: 'Calendar',
      key: 'calendarId',
      type: 'dropdown',
      required: true,
      description: '',
      variables: false,
      source: {
        type: 'query',
        name: 'getDynamicData',
        arguments: [
          {
            name: 'key',
            value: 'listCalendars'
          }
        ]
      }
    },
    {
      label: 'Look Forward Interval',
      key: 'interval',
      type: 'dropdown',
      required: true,
      value: null,
      description: 'The interval to look forward for events.',
      variables: true,
      options: [
        {
          label: 'This week',
          value: 'week'
        },
        {
          label: 'This Month',
          value: 'month'
        }
      ]
    }
  ],

  async run($) {
    const calendarId = $.step.parameters.calendarId;

    const start = this.getStartOf($.step.parameters.interval);
    const end = this.getEndOf(start, $.step.parameters.interval);

    const params = {
      pageToken: undefined,
      orderBy: 'startTime',
      timeMin: start.toFormat(FORMAT_TIMESTAMP),
      timeMax: end.toFormat(FORMAT_TIMESTAMP),
      singleEvents: true
    };

    const results = [];

    do {
      const {data} = await $.http.get(`/v3/calendars/${calendarId}/events`, {
        params
      });
      params.pageToken = data.nextPageToken;

      if (!data.items?.length) {
        continue;
      }

      for (const event of data.items) {
        results.push(this.asResult(event));
        if (results.length >= MAX_EVENTS) {
          break;
        }
      }
    } while (params.pageToken);

    $.setActionItem({raw: {events: results}});
  },

  getStartOf(intervalType) {
    return DateTime.now().startOf(intervalType);
  },

  getEndOf(start, intervalType) {
    return start.endOf(intervalType);
  },

  asResult(event) {
    return {
      title: event.summary,
      description: event.description,
      location: event.location,
      start: event.start,
      end: event.end,
      recurrence: event.recurrence,
      attendees: event.attendees?.map((a) => ({
        name: a.displayName, email: a.email, status: a.responseStatus
      }))
    };
  }
});
