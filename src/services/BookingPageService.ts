/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { google } from 'googleapis';
import { logToFile } from '../utils/logger';
import { gaxiosOptions } from '../utils/GaxiosConfig';
import codeSource from '../booking-page/Code.gs';
import indexSource from '../booking-page/Index.html';

// Google Calendar has no API for its native appointment schedules, so a
// booking page is deployed as an Apps Script web app that runs as the user.
export const BOOKING_PAGE_SCOPES = [
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.deployments',
];

const WEEKDAYS = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'] as const;
export type Weekday = (typeof WEEKDAYS)[number];

const APPS_SCRIPT_SETTINGS_URL = 'https://script.google.com/home/usersettings';

export interface CreateBookingPageInput {
  title: string;
  description?: string;
  durationMinutes?: number;
  days?: Weekday[];
  startTime?: string;
  endTime?: string;
  timeZone?: string;
  windowDays?: number;
  minNoticeHours?: number;
  bufferMinutes?: number;
  calendarId?: string;
  checkCalendarIds?: string[];
  addMeetLink?: boolean;
}

export interface DeleteBookingPageInput {
  scriptId: string;
}

function parseTimeOfDay(value: string): number {
  const match = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(value);
  if (!match) {
    throw new Error(`Invalid time "${value}". Use 24-hour HH:MM, e.g. 09:00.`);
  }
  return Number(match[1]) * 60 + Number(match[2]);
}

export class BookingPageService {
  constructor(private authManager: any) {}

  private async getClientOptions() {
    const auth = await this.authManager.getAuthenticatedClient();
    return { ...gaxiosOptions, auth };
  }

  private errorResponse(tool: string, error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logToFile(`Error during ${tool}: ${errorMessage}`);
    const help = /Apps Script API/i.test(errorMessage)
      ? `Turn on the Apps Script API at ${APPS_SCRIPT_SETTINGS_URL} and try again.`
      : undefined;
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify({ error: errorMessage, help }),
        },
      ],
    };
  }

  createBookingPage = async (input: CreateBookingPageInput) => {
    const {
      title,
      description = '',
      durationMinutes = 30,
      days = ['MON', 'TUE', 'WED', 'THU', 'FRI'],
      startTime = '09:00',
      endTime = '17:00',
      windowDays = 14,
      minNoticeHours = 12,
      bufferMinutes = 0,
      calendarId = 'primary',
      addMeetLink = true,
    } = input;
    logToFile(`Creating booking page: ${title}`);

    try {
      const startMinutes = parseTimeOfDay(startTime);
      const endMinutes = parseTimeOfDay(endTime);
      if (startMinutes + durationMinutes > endMinutes) {
        throw new Error(
          'endTime must be at least one slot later than startTime.',
        );
      }

      const options = await this.getClientOptions();
      let timeZone = input.timeZone;
      if (!timeZone) {
        const calendar = google.calendar({ version: 'v3', ...options });
        const res = await calendar.calendars.get({ calendarId });
        timeZone = res.data.timeZone || 'UTC';
      }

      const config = {
        title,
        description,
        durationMinutes,
        days,
        startMinutes,
        endMinutes,
        timeZone,
        windowDays,
        minNoticeHours,
        bufferMinutes,
        calendarId,
        checkCalendarIds: input.checkCalendarIds?.length
          ? input.checkCalendarIds
          : [calendarId],
        addMeetLink,
      };
      const manifest = {
        timeZone,
        runtimeVersion: 'V8',
        exceptionLogging: 'STACKDRIVER',
        dependencies: {
          enabledAdvancedServices: [
            { userSymbol: 'Calendar', serviceId: 'calendar', version: 'v3' },
          ],
        },
        oauthScopes: [
          'https://www.googleapis.com/auth/calendar.events',
          'https://www.googleapis.com/auth/calendar.freebusy',
        ],
        webapp: { executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' },
      };

      const script = google.script({ version: 'v1', ...options });
      const project = await script.projects.create({
        requestBody: { title: `Booking page: ${title}` },
      });
      const scriptId = project.data.scriptId;
      if (!scriptId) {
        throw new Error('Apps Script did not return a script ID.');
      }
      logToFile(`Created Apps Script project: ${scriptId}`);

      await script.projects.updateContent({
        scriptId,
        requestBody: {
          files: [
            {
              name: 'appsscript',
              type: 'JSON',
              source: JSON.stringify(manifest, null, 2),
            },
            {
              name: 'Code',
              type: 'SERVER_JS',
              // A function replacement keeps `$` sequences in the config literal.
              source: codeSource.replace('__BOOKING_CONFIG__', () =>
                JSON.stringify(config, null, 2),
              ),
            },
            { name: 'Index', type: 'HTML', source: indexSource },
          ],
        },
      });
      const version = await script.projects.versions.create({
        scriptId,
        requestBody: { description: 'Booking page' },
      });
      const deployment = await script.projects.deployments.create({
        scriptId,
        requestBody: {
          versionNumber: version.data.versionNumber,
          manifestFileName: 'appsscript',
          description: 'Booking page',
        },
      });
      const deploymentId = deployment.data.deploymentId;
      const bookingUrl =
        deployment.data.entryPoints?.find((entry) => entry.webApp?.url)?.webApp
          ?.url || `https://script.google.com/macros/s/${deploymentId}/exec`;
      logToFile(`Deployed booking page: ${bookingUrl}`);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              bookingUrl,
              scriptId,
              deploymentId,
              editorUrl: `https://script.google.com/d/${scriptId}/edit`,
              config,
              nextStep:
                'The calendar owner must open bookingUrl once while signed in and approve the calendar permissions. Until then the page does not work for visitors.',
            }),
          },
        ],
      };
    } catch (error) {
      return this.errorResponse('calendar.createBookingPage', error);
    }
  };

  deleteBookingPage = async ({ scriptId }: DeleteBookingPageInput) => {
    logToFile(`Deleting booking page: ${scriptId}`);
    try {
      const options = await this.getClientOptions();
      const drive = google.drive({ version: 'v3', ...options });
      // Trashing the script project takes its web app offline.
      await drive.files.update({
        fileId: scriptId,
        requestBody: { trashed: true },
      });
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              scriptId,
              status:
                'Booking page moved to Drive trash and taken offline. Events already booked are unchanged.',
            }),
          },
        ],
      };
    } catch (error) {
      return this.errorResponse('calendar.deleteBookingPage', error);
    }
  };
}
