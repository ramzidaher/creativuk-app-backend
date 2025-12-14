#!/usr/bin/env node
// Prints every user's calendar (current week) in CLI.
//
// App-only Graph auth (client credentials). Requires:
//  - Application permissions: User.Read.All (or Directory.Read.All), Calendars.Read (or ReadWrite)
//  - Admin consent granted
//  - Optional: Exchange RBAC for Applications or App Access Policy to scope which mailboxes are readable.
//
// Node 18+ (uses global fetch)

import 'dotenv/config';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { DateTime } from 'luxon';

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TIMEZONE = process.env.TIMEZONE || 'Europe/London';
const MAX_USERS = Number(process.env.MAX_USERS || 50);
const EXCLUDE_NO_MAIL = String(process.env.EXCLUDE_NO_MAIL || 'true').toLowerCase() === 'true';

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error('Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in environment.');
  process.exit(1);
}

const cca = new ConfidentialClientApplication({
  auth: {
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET
  }
});

async function getToken() {
  const res = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default']
  });
  if (!res) {
    throw new Error('Failed to acquire token');
  }
  return res.accessToken;
}

async function gjson(url: string, token: string, options: any = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${res.status} ${res.statusText}: ${text}`);
  }
  return res.json();
}

// Compute current week [Mon 00:00, Sun 23:59:59.999] in TIMEZONE
function weekWindow(zone) {
  // Luxon: Monday=1..Sunday=7
  const now = DateTime.now().setZone(zone);
  const dow = now.weekday;                  // 1..7
  const start = now.minus({ days: dow - 1 }).startOf('day');  // Monday
  const end = start.plus({ days: 6 }).endOf('day');           // Sunday
  return {
    startISO: start.toISO({ suppressSeconds: true, suppressMilliseconds: true }),
    endISO: end.toISO({ suppressSeconds: true, suppressMilliseconds: true })
  };
}

interface User {
  id: string;
  name: string;
  addr: string;
}

function isValidUser(user: any): user is User {
  return user && user.id && typeof user.id === 'string' && user.id.length > 0;
}

async function listAllUsers(token: string, cap: number, excludeNoMail: boolean): Promise<User[]> {
  const users: User[] = [];
  let url = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,accountEnabled&$top=100';
  while (url && users.length < cap) {
    const data = await gjson(url, token);
    for (const u of data.value || []) {
      if (!u.accountEnabled) continue;
      const addr = u.mail || u.userPrincipalName;
      if (excludeNoMail && !addr) continue;
      
      const user = { 
        id: u.id, 
        name: u.displayName || u.userPrincipalName, 
        addr 
      };
      
      if (isValidUser(user)) {
        users.push(user);
        if (users.length >= cap) break;
      }
    }
    url = data['@odata.nextLink'] || null;
  }
  return users;
}

function maskIfPrivate(ev: any) {
  // Event resource exposes 'sensitivity' (normal/personal/private/confidential).
  // If private, hide subject/location.
  const sensitivity = (ev.sensitivity || '').toLowerCase();
  if (sensitivity === 'private') {
    return {
      ...ev,
      subject: '(busy)',
      location: { displayName: '' }
    };
  }
  return ev;
}

function fmt(dt: any) {
  // dt is { dateTime, timeZone }
  return `${dt.dateTime}`;
}

async function listCalendarForUser(token: string, userId: string, startISO: string, endISO: string) {
  const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/calendarView`;
  const qs = `?startDateTime=${encodeURIComponent(startISO)}&endDateTime=${encodeURIComponent(endISO)}&$select=subject,start,end,location,showAs,isAllDay,sensitivity,webLink,organizer`;
  const url = base + qs;

  const data = await gjson(url, token, {
    headers: { Prefer: `outlook.timezone="${TIMEZONE}"` }
  });

  const items = (data.value || []).map(maskIfPrivate).sort((a: any, b: any) => {
    return a.start.dateTime.localeCompare(b.start.dateTime);
  });

  return items;
}

(async () => {
  try {
    const { startISO, endISO } = weekWindow(TIMEZONE);
    console.log(`\n=== Tenant calendars for week ===`);
    console.log(`Range: ${startISO} -> ${endISO}  [${TIMEZONE}]`);
    console.log(`Limit: ${MAX_USERS} users\n`);

    const token = await getToken();

    // 1) Users
    const users = await listAllUsers(token, MAX_USERS, EXCLUDE_NO_MAIL);
    if (!users.length) {
      console.log('No users found (or all filtered).');
      process.exit(0);
    }

    // 2) Per-user calendar view
    for (const u of users) {
      console.log(`\n# ${u.name}  <${u.addr || 'no-mail'}>`);
      try {
        // @ts-ignore - u.id is guaranteed to be non-null due to type guard
        const events = await listCalendarForUser(token, u.id, startISO, endISO);
        if (!events.length) {
          console.log('  (no events)');
          continue;
        }
        for (const ev of events) {
          const allDay = ev.isAllDay ? '[All-day] ' : '';
          const line = `  - ${allDay}${fmt(ev.start)} -> ${fmt(ev.end)}  | ${ev.showAs || 'busy'}  | ${ev.subject || '(no subject)'}`;
          console.log(line);
        }
      } catch (e) {
        // Likely blocked by RBAC scope / policy for this mailbox; continue
        console.log(`  (skipped: ${e.message.split('\n')[0]})`);
      }
    }
    console.log('\nDone.\n');
  } catch (e) {
    console.error('\nError:', e.message || e);
    process.exit(1);
  }
})();

