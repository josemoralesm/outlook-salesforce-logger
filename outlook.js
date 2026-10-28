// Reads the current Outlook item in a mode-independent way
// (read mode returns values directly; organizer/compose mode needs getAsync).

const MARKER = "sfLogged";

function asyncOf(start) {
  return new Promise((resolve, reject) =>
    start((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error?.message || "Outlook request failed."))
    )
  );
}

function prop(p) {
  return p && typeof p.getAsync === "function" ? asyncOf((cb) => p.getAsync(cb)) : Promise.resolve(p);
}

const fmt = (p) => (p.displayName && p.displayName !== p.emailAddress ? `${p.displayName} <${p.emailAddress}>` : p.emailAddress);

export function hasItem() {
  return !!Office.context.mailbox.item;
}

export async function readCurrentItem() {
  const it = Office.context.mailbox.item;
  const isMeeting = it.itemType === Office.MailboxEnums.ItemType.Appointment;
  const me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();

  const [subject, body] = await Promise.all([prop(it.subject), asyncOf((cb) => it.body.getAsync(Office.CoercionType.Text, cb))]);

  const result = { kind: isMeeting ? "meeting" : "email", subject: subject || "(no subject)", body: body || "" };
  let people;

  if (isMeeting) {
    const [start, end, location, required, optional, organizer] = await Promise.all([
      prop(it.start),
      prop(it.end),
      prop(it.location),
      prop(it.requiredAttendees),
      prop(it.optionalAttendees),
      prop(it.organizer),
    ]);
    Object.assign(result, { start: new Date(start), end: new Date(end), location: location || "" });
    people = [organizer, ...(required || []), ...(optional || [])];
    result.header = [
      organizer && `Organizer: ${fmt(organizer)}`,
      required?.length && `Required: ${required.map(fmt).join("; ")}`,
      optional?.length && `Optional: ${optional.map(fmt).join("; ")}`,
      `When: ${result.start.toLocaleString()} – ${result.end.toLocaleString()}`,
      result.location && `Where: ${result.location}`,
    ];
  } else {
    const to = it.to || [];
    const cc = it.cc || [];
    result.date = it.dateTimeCreated ? new Date(it.dateTimeCreated) : new Date();
    result.direction = it.from?.emailAddress?.toLowerCase() === me ? "sent" : "received";
    people = [it.from, ...to, ...cc];
    result.header = [
      it.from && `From: ${fmt(it.from)}`,
      to.length && `To: ${to.map(fmt).join("; ")}`,
      cc.length && `Cc: ${cc.map(fmt).join("; ")}`,
      `Date: ${result.date.toLocaleString()}`,
      it.internetMessageId && `Message-ID: ${it.internetMessageId}`,
    ];
  }
  result.header = result.header.filter(Boolean).join("\n");

  // Unique external participants (excluding the signed-in user).
  const seen = new Map();
  for (const p of people) {
    const email = p?.emailAddress?.toLowerCase();
    if (!email || email === me || seen.has(email)) continue;
    seen.set(email, { email, name: p.displayName && p.displayName !== p.emailAddress ? p.displayName : "" });
  }
  result.people = [...seen.values()];
  result.attachments = await listAttachments(it);
  return result;
}

async function listAttachments(it) {
  const list = Array.isArray(it.attachments)
    ? it.attachments
    : typeof it.getAttachmentsAsync === "function"
      ? await asyncOf((cb) => it.getAttachmentsAsync(cb))
      : [];
  return list
    .filter((a) => !a.isInline && a.attachmentType === Office.MailboxEnums.AttachmentType.File)
    .map((a) => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType }));
}

export async function getAttachmentBase64(id) {
  const c = await asyncOf((cb) => Office.context.mailbox.item.getAttachmentContentAsync(id, cb));
  if (c.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) throw new Error("This attachment type can't be uploaded.");
  return c.content;
}

// Remember on the Outlook item that it was logged, to warn about duplicates.
export async function getLoggedMarker() {
  try {
    const props = await asyncOf((cb) => Office.context.mailbox.item.loadCustomPropertiesAsync(cb));
    const v = props.get(MARKER);
    return v ? JSON.parse(v) : null;
  } catch {
    return null;
  }
}

export async function setLoggedMarker(value) {
  try {
    const props = await asyncOf((cb) => Office.context.mailbox.item.loadCustomPropertiesAsync(cb));
    props.set(MARKER, JSON.stringify(value));
    await asyncOf((cb) => props.saveAsync(cb));
  } catch {
    /* non-critical */
  }
}
