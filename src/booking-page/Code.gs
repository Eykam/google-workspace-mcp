/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * Booking page web app, deployed by the google-workspace MCP server
 * (calendar.createBookingPage). Runs as the calendar owner; visitors are
 * anonymous. The placeholder below is replaced with a JSON object at deploy time.
 */
var CONFIG = __BOOKING_CONFIG__;

var MINUTE_MS = 60 * 1000;
var DAY_MS = 24 * 60 * MINUTE_MS;
var WEEKDAYS = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the page on load. */
function getPageData() {
  return {
    title: CONFIG.title,
    description: CONFIG.description,
    durationMinutes: CONFIG.durationMinutes,
    slots: availableSlots_().map(function (slot) {
      return slot.toISOString();
    }),
  };
}

/** Called by the page when a visitor confirms a slot. */
function book(request) {
  request = request || {};
  // Honeypot field: real visitors never see or fill it.
  if (request.website) {
    return { ok: false, reason: 'invalid' };
  }
  var name = String(request.name || '').trim();
  var email = String(request.email || '').trim();
  var notes = String(request.notes || '').trim();
  var start = new Date(request.start);
  if (
    !name ||
    name.length > 100 ||
    email.length > 254 ||
    !/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email) ||
    notes.length > 1000 ||
    isNaN(start.getTime())
  ) {
    return { ok: false, reason: 'invalid' };
  }

  // Serialize bookings so two visitors cannot take the same slot.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var stillFree = availableSlots_().some(function (slot) {
      return slot.getTime() === start.getTime();
    });
    if (!stillFree) {
      return { ok: false, reason: 'taken' };
    }

    var end = new Date(start.getTime() + CONFIG.durationMinutes * MINUTE_MS);
    var event = {
      summary: CONFIG.title + ' (' + name + ')',
      description:
        'Booked by ' + name + ' <' + email + '> via booking page.' +
        (notes ? '\n\nNotes:\n' + notes : ''),
      start: { dateTime: start.toISOString() },
      end: { dateTime: end.toISOString() },
      attendees: [{ email: email, displayName: name }],
      extendedProperties: { private: { bookingPage: 'true' } },
    };
    var options = { sendUpdates: 'all' };
    if (CONFIG.addMeetLink) {
      event.conferenceData = {
        createRequest: {
          requestId: Utilities.getUuid(),
          conferenceSolutionKey: { type: 'hangoutsMeet' },
        },
      };
      options.conferenceDataVersion = 1;
    }
    var created = Calendar.Events.insert(event, CONFIG.calendarId, options);
    return {
      ok: true,
      start: start.toISOString(),
      end: end.toISOString(),
      meetLink: created.hangoutLink || null,
    };
  } finally {
    lock.releaseLock();
  }
}

/** Slots inside the configured hours that do not collide with a busy block. */
function availableSlots_() {
  var slots = candidateSlots_();
  if (slots.length === 0) {
    return slots;
  }
  var durationMs = CONFIG.durationMinutes * MINUTE_MS;
  var bufferMs = CONFIG.bufferMinutes * MINUTE_MS;
  var busy = busyBlocks_(
    new Date(slots[0].getTime() - bufferMs),
    new Date(slots[slots.length - 1].getTime() + durationMs + bufferMs),
  );
  return slots.filter(function (slot) {
    var from = slot.getTime() - bufferMs;
    var to = slot.getTime() + durationMs + bufferMs;
    return !busy.some(function (block) {
      return from < block.end && to > block.start;
    });
  });
}

/** Every slot the weekly hours allow inside the booking window, in order. */
function candidateSlots_() {
  var tz = CONFIG.timeZone;
  var now = new Date();
  var earliest = now.getTime() + CONFIG.minNoticeHours * 60 * MINUTE_MS;
  var latest = now.getTime() + CONFIG.windowDays * DAY_MS;
  // Anchor on local noon so stepping by 24h never skips or repeats a date
  // across a daylight-saving change.
  var todayNoon = Utilities.parseDate(
    Utilities.formatDate(now, tz, 'yyyy-MM-dd') + ' 12:00',
    tz,
    'yyyy-MM-dd HH:mm',
  );
  var slots = [];
  for (var i = 0; i <= CONFIG.windowDays; i++) {
    var noon = new Date(todayNoon.getTime() + i * DAY_MS);
    var weekday = Utilities.formatDate(noon, tz, 'EEE').toUpperCase();
    if (CONFIG.days.indexOf(weekday) === -1 || WEEKDAYS.indexOf(weekday) === -1) {
      continue;
    }
    var day = Utilities.formatDate(noon, tz, 'yyyy-MM-dd');
    for (
      var minutes = CONFIG.startMinutes;
      minutes + CONFIG.durationMinutes <= CONFIG.endMinutes;
      minutes += CONFIG.durationMinutes
    ) {
      var start = Utilities.parseDate(
        day + ' ' + pad_(Math.floor(minutes / 60)) + ':' + pad_(minutes % 60),
        tz,
        'yyyy-MM-dd HH:mm',
      );
      if (start.getTime() >= earliest && start.getTime() <= latest) {
        slots.push(start);
      }
    }
  }
  return slots;
}

function busyBlocks_(timeMin, timeMax) {
  var response = Calendar.Freebusy.query({
    timeMin: timeMin.toISOString(),
    timeMax: timeMax.toISOString(),
    items: CONFIG.checkCalendarIds.map(function (id) {
      return { id: id };
    }),
  });
  var blocks = [];
  Object.keys(response.calendars || {}).forEach(function (id) {
    var calendar = response.calendars[id];
    // Fail closed: offering slots we could not check risks a double booking.
    if (calendar.errors && calendar.errors.length) {
      throw new Error('Could not read availability for calendar ' + id);
    }
    (calendar.busy || []).forEach(function (block) {
      blocks.push({
        start: new Date(block.start).getTime(),
        end: new Date(block.end).getTime(),
      });
    });
  });
  return blocks;
}

function pad_(n) {
  return (n < 10 ? '0' : '') + n;
}
