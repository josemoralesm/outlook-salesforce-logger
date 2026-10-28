// Thin Salesforce REST client used by the task pane.
import { CONFIG } from "./config.js";
import { currentSession, refreshSession, AuthError } from "./auth.js";

async function request(path, { method = "GET", body } = {}, retried = false) {
  const s = currentSession();
  if (!s) throw new AuthError();
  const url = `${s.instanceUrl}/services/data/${CONFIG.apiVersion}${path}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${s.accessToken}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (res.status === 401 && !retried) {
    await refreshSession();
    return request(path, { method, body }, true);
  }
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const msg = Array.isArray(data) ? data.map((e) => e.message).join("; ") : data?.message;
    throw new Error(msg || `Salesforce returned an error (${res.status}).`);
  }
  return data;
}

const q = (v) => `'${String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'")}'`;

export async function query(soql) {
  const r = await request(`/query?q=${encodeURIComponent(soql)}`);
  return r.records;
}

export async function create(sobject, fields) {
  // Drop empty values so Salesforce applies its own defaults.
  const clean = Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== null && v !== undefined && v !== ""));
  const r = await request(`/sobjects/${sobject}`, { method: "POST", body: clean });
  return r.id;
}

export function recordUrl(id) {
  return `${currentSession()?.instanceUrl}/${id}`;
}

// ── People ────────────────────────────────────────────────────────────────────
export async function findPeopleByEmail(emails) {
  if (!emails.length) return { contacts: [], leads: [] };
  const list = emails.slice(0, 100).map(q).join(",");
  const [contacts, leads] = await Promise.all([
    query(`SELECT Id, Name, Email, Title, AccountId, Account.Name FROM Contact WHERE Email IN (${list}) ORDER BY LastModifiedDate DESC LIMIT 100`),
    query(`SELECT Id, Name, Email, Company, Status FROM Lead WHERE Email IN (${list}) AND IsConverted = false ORDER BY LastModifiedDate DESC LIMIT 100`),
  ]);
  return { contacts, leads };
}

export function createContact({ firstName, lastName, email, title, accountId }) {
  return create("Contact", { FirstName: firstName, LastName: lastName, Email: email, Title: title, AccountId: accountId });
}

export function createLead({ firstName, lastName, email, title, company }) {
  return create("Lead", { FirstName: firstName, LastName: lastName, Email: email, Title: title, Company: company });
}

// ── Related records (WhatId) ──────────────────────────────────────────────────
export async function openOpportunities(accountIds) {
  if (!accountIds.length) return [];
  return query(
    `SELECT Id, Name, StageName, AccountId FROM Opportunity WHERE AccountId IN (${accountIds.map(q).join(",")}) AND IsClosed = false ORDER BY CloseDate ASC LIMIT 25`
  );
}

export async function searchRelated(term) {
  // Strip SOSL reserved characters, then prefix-match.
  const clean = term.replace(/[?&|!{}[\]()^~*:\\"'+\-]/g, " ").replace(/\s+/g, " ").trim();
  if (clean.length < 2) return [];
  const sosl = `FIND {${clean}*} IN NAME FIELDS RETURNING Account(Id, Name), Opportunity(Id, Name, StageName), Case(Id, CaseNumber, Subject) LIMIT 20`;
  const r = await request(`/search?q=${encodeURIComponent(sosl)}`);
  return (r.searchRecords || []).map(toOption);
}

export function toOption(rec) {
  const type = rec.attributes?.type;
  let label = rec.Name;
  if (type === "Opportunity" && rec.StageName) label = `${rec.Name} · ${rec.StageName}`;
  if (type === "Case") label = `Case ${rec.CaseNumber}${rec.Subject ? ` · ${rec.Subject}` : ""}`;
  return { id: rec.Id, type, label };
}

// ── Activities ────────────────────────────────────────────────────────────────
export function logEmailTask({ subject, description, whoId, whatId, date }) {
  return create("Task", {
    Subject: `Email: ${subject}`.slice(0, 255),
    Description: description,
    WhoId: whoId,
    WhatId: whatId,
    ActivityDate: date,
    Status: CONFIG.taskStatus,
    Priority: CONFIG.taskPriority,
    TaskSubtype: "Email",
  });
}

export function logMeetingEvent({ subject, description, whoId, whatId, start, end, location }) {
  return create("Event", {
    Subject: subject.slice(0, 255),
    Description: description,
    WhoId: whoId,
    WhatId: whatId,
    StartDateTime: start.toISOString(),
    EndDateTime: end.toISOString(),
    Location: (location || "").slice(0, 255),
  });
}

// ── Files ─────────────────────────────────────────────────────────────────────
// Uploads a file to `recordId` (shows in its Files list) and also links it to the activity.
export async function uploadFile({ name, base64, recordId, activityId }) {
  const dot = name.lastIndexOf(".");
  const versionId = await create("ContentVersion", {
    Title: (dot > 0 ? name.slice(0, dot) : name).slice(0, 255),
    PathOnClient: name,
    VersionData: base64,
    FirstPublishLocationId: recordId,
  });
  if (activityId && activityId !== recordId) {
    try {
      const [v] = await query(`SELECT ContentDocumentId FROM ContentVersion WHERE Id = ${q(versionId)}`);
      await create("ContentDocumentLink", { ContentDocumentId: v.ContentDocumentId, LinkedEntityId: activityId, ShareType: "V" });
    } catch {
      /* file is still on the related record; the extra link is a convenience */
    }
  }
  return versionId;
}
