import { CONFIG } from "./config.js";
import { currentSession, signIn, signOut, AuthError } from "./auth.js";
import * as sf from "./salesforce.js";
import * as ol from "./outlook.js";

// ── State ─────────────────────────────────────────────────────────────────────
const fresh = () => ({
  item: null,
  contacts: [],
  leads: [],
  who: null, // { id, type }
  what: null, // { id, type, label }
  suggestions: [], // related-record options derived from matched contacts
  searchResults: [],
  adding: null, // email address with the "add to Salesforce" form open
  includeBody: true,
  files: new Set(),
  logged: null, // { id, at, kind } when this item was already logged
  busy: false,
});
let S = fresh();
let loadSeq = 0;

// ── DOM helpers ───────────────────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k.startsWith("on")) el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === "class") el.className = v;
    else if (k === "checked" || k === "disabled" || k === "value") el[k] = v;
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat(Infinity)) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : String(c));
  }
  return el;
}

function showView(name) {
  for (const v of ["loading", "signin", "empty", "main"]) $(`view-${v}`).hidden = v !== name;
  $("actions").hidden = name !== "main";
  const s = currentSession();
  $("account").hidden = !s || name === "signin";
  if (s) $("account-host").textContent = new URL(s.instanceUrl).hostname.split(".")[0];
}

function showError(message) {
  $("error").textContent = message || "";
  $("error").hidden = !message;
}

function setStatus(text) {
  $("status").textContent = text || "";
}

const sizeLabel = (b) => (b >= 1048576 ? `${(b / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(b / 1024))} KB`);
const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

function link(id, text) {
  return h("a", { href: sf.recordUrl(id), target: "_blank", rel: "noopener" }, text);
}

// ── Startup ───────────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showError("Open this add-in from Outlook.");
    return;
  }
  $("signin-btn").addEventListener("click", onSignIn);
  $("signout-btn").addEventListener("click", async () => {
    await signOut();
    load();
  });
  $("log-btn").addEventListener("click", onLog);
  $("search").addEventListener("input", onSearchInput);
  // Pinned task pane: reload when the user selects a different email.
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => load());
  } catch {
    /* not supported in this client */
  }
  load();
});

async function onSignIn() {
  showError("");
  $("signin-btn").disabled = true;
  try {
    await signIn();
    await load();
  } catch (e) {
    showError(e.message);
  } finally {
    $("signin-btn").disabled = false;
  }
}

async function load() {
  const seq = ++loadSeq;
  S = fresh();
  showError("");
  setStatus("");
  $("search").value = "";
  if (!ol.hasItem()) return showView("empty");
  if (!currentSession()) return showView("signin");
  showView("loading");

  try {
    const item = await ol.readCurrentItem();
    const logged = await ol.getLoggedMarker();
    const { contacts, leads } = await sf.findPeopleByEmail(item.people.map((p) => p.email));
    const accounts = new Map();
    for (const c of contacts) if (c.AccountId) accounts.set(c.AccountId, { id: c.AccountId, type: "Account", label: c.Account?.Name || "Account" });
    const opps = await sf.openOpportunities([...accounts.keys()]);
    if (seq !== loadSeq) return; // user moved to another item meanwhile

    Object.assign(S, { item, logged, contacts, leads });
    S.suggestions = [...accounts.values(), ...opps.map(sf.toOption)];
    for (const a of item.attachments) if (a.size <= CONFIG.maxAttachmentBytes) S.files.add(a.id);

    // Default "Name": first participant found in Salesforce (contacts before leads).
    const first = item.people.map((p) => matchFor(p.email)).find(Boolean);
    if (first) selectWho(first, false);

    showView("main");
    render();
  } catch (e) {
    if (seq === loadSeq) handleError(e);
  }
}

function handleError(e) {
  if (e instanceof AuthError) {
    showView("signin");
    showError(e.message);
  } else {
    if (!S.item) showView("empty");
    showError(e.message || String(e));
  }
}

// ── Selection logic ───────────────────────────────────────────────────────────
function matchFor(email) {
  const eq = (r) => r.Email && r.Email.toLowerCase() === email;
  return S.contacts.find(eq) || S.leads.find(eq) || null;
}

function selectWho(rec, rerender = true) {
  if (!rec) {
    S.who = null;
  } else {
    const type = rec.attributes.type;
    S.who = { id: rec.Id, type };
    if (type === "Lead") {
      S.what = null; // Salesforce doesn't allow a WhatId with a Lead
    } else if (rec.AccountId && (!S.what || S.what.type === "Account")) {
      S.what = S.suggestions.find((o) => o.id === rec.AccountId) || { id: rec.AccountId, type: "Account", label: rec.Account?.Name || "Account" };
    }
  }
  if (rerender) render();
}

function selectWhat(opt) {
  S.what = opt;
  if (opt && !S.suggestions.some((o) => o.id === opt.id)) S.suggestions.push(opt);
  S.searchResults = [];
  $("search").value = "";
  render();
}

// ── Rendering ─────────────────────────────────────────────────────────────────
function render() {
  renderSummary();
  renderLogged();
  renderPeople();
  renderRelated();
  renderSearch();
  renderOptions();
  renderActions();
}

function renderSummary() {
  const it = S.item;
  const when =
    it.kind === "meeting"
      ? `${it.start.toLocaleString([], { dateStyle: "medium", timeStyle: "short" })} – ${it.end.toLocaleTimeString([], { timeStyle: "short" })}`
      : `${it.direction === "sent" ? "Sent" : "Received"} ${it.date.toLocaleString([], { dateStyle: "medium", timeStyle: "short" })}`;
  $("summary").replaceChildren(
    h("div", { class: "kind" }, it.kind === "meeting" ? "Meeting → Event" : "Email → Task"),
    h("div", { class: "subject" }, it.subject),
    h("div", { class: "muted small" }, when, it.location ? ` · ${it.location}` : "")
  );
}

function renderLogged(result) {
  const el = $("logged");
  if (result) {
    el.replaceChildren(result);
  } else if (S.logged?.id) {
    el.replaceChildren(
      h("div", { class: "banner banner-warn" }, `Already logged on ${new Date(S.logged.at).toLocaleDateString()}. `, link(S.logged.id, "View in Salesforce"))
    );
  } else {
    el.replaceChildren();
  }
}

function personRow(p) {
  const m = matchFor(p.email);
  if (m) {
    const type = m.attributes.type;
    const selected = S.who?.id === m.Id;
    const detail = type === "Contact" ? m.Account?.Name : m.Company;
    return h(
      "label",
      { class: `row selectable${selected ? " selected" : ""}` },
      h("input", { type: "radio", name: "who", checked: selected, onchange: () => selectWho(m) }),
      h("div", { class: "main" }, h("div", { class: "name" }, m.Name), h("div", { class: "sub" }, [p.email, detail].filter(Boolean).join(" · "))),
      h("span", { class: "badge" }, type)
    );
  }
  const open = S.adding === p.email;
  return h(
    "div",
    { class: "row" },
    h(
      "div",
      { class: "main" },
      h("div", { class: "name" }, p.name || p.email),
      h("div", { class: "sub" }, p.name ? p.email : "Not in Salesforce"),
      open ? addForm(p) : null
    ),
    open
      ? null
      : h("button", { class: "secondary", type: "button", onclick: () => ((S.adding = p.email), renderPeople()) }, "Add")
  );
}

function renderPeople() {
  const people = S.item.people;
  const rows = people.map(personRow);
  if (!people.length) rows.push(h("div", { class: "muted small" }, "No external participants on this item."));
  rows.push(
    h(
      "label",
      { class: `row selectable${!S.who ? " selected" : ""}` },
      h("input", { type: "radio", name: "who", checked: !S.who, onchange: () => selectWho(null) }),
      h("div", { class: "main muted" }, "Don't link a person")
    )
  );
  $("people").replaceChildren(...rows);
}

function splitName(name, email) {
  const n = (name || "").trim();
  if (n.includes(",")) {
    const [last, first] = n.split(",").map((s) => s.trim());
    return { first, last };
  }
  const parts = n.split(/\s+/).filter(Boolean);
  if (parts.length >= 2) return { first: parts.slice(0, -1).join(" "), last: parts.at(-1) };
  if (parts.length === 1) return { first: "", last: parts[0] };
  const local = email.split("@")[0].split(/[._-]+/).filter(Boolean);
  const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
  return local.length >= 2 ? { first: cap(local[0]), last: cap(local.at(-1)) } : { first: "", last: cap(local[0] || email) };
}

function companyFromEmail(email) {
  const host = email.split("@")[1] || "";
  const generic = /^(gmail|outlook|hotmail|live|yahoo|icloud|aol|proton|protonmail|msn|me)\./i;
  if (generic.test(host)) return "";
  const label = host.split(".").slice(-2, -1)[0] || "";
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function addForm(p) {
  const { first, last } = splitName(p.name, p.email);
  const account = S.what?.type === "Account" ? S.what : null;
  const type = h("select", {}, h("option", { value: "Contact" }, "Contact"), h("option", { value: "Lead" }, "Lead"));
  type.value = account ? "Contact" : "Lead";
  const f = h("input", { type: "text", value: first });
  const l = h("input", { type: "text", value: last });
  const title = h("input", { type: "text" });
  const company = h("input", { type: "text", value: companyFromEmail(p.email) });
  const companyLabel = h("label", { class: "full" }, "Company *", company);
  const accountNote = h(
    "div",
    { class: "full small muted" },
    account ? `Will be added to account: ${account.label}` : "Tip: pick an account under “Related to” first to attach this contact to it."
  );
  const sync = () => {
    companyLabel.hidden = type.value !== "Lead";
    accountNote.hidden = type.value !== "Contact";
  };
  type.addEventListener("change", sync);
  sync();

  const save = h("button", { class: "primary", type: "button" }, "Save");
  save.addEventListener("click", async () => {
    showError("");
    if (!l.value.trim()) return showError("Last name is required.");
    if (type.value === "Lead" && !company.value.trim()) return showError("Company is required for a lead.");
    save.disabled = true;
    try {
      const base = { firstName: f.value.trim(), lastName: l.value.trim(), email: p.email, title: title.value.trim() };
      let rec;
      if (type.value === "Contact") {
        const id = await sf.createContact({ ...base, accountId: account?.id });
        rec = { attributes: { type: "Contact" }, Id: id, Name: `${base.firstName} ${base.lastName}`.trim(), Email: p.email, AccountId: account?.id, Account: account ? { Name: account.label } : null };
        S.contacts.push(rec);
      } else {
        const id = await sf.createLead({ ...base, company: company.value.trim() });
        rec = { attributes: { type: "Lead" }, Id: id, Name: `${base.firstName} ${base.lastName}`.trim(), Email: p.email, Company: company.value.trim() };
        S.leads.push(rec);
      }
      S.adding = null;
      selectWho(rec);
      setStatus(`${type.value} created.`);
    } catch (e) {
      handleError(e);
      save.disabled = false;
    }
  });

  return h(
    "div",
    { class: "form" },
    h("label", { class: "full" }, "Add as", type),
    h("label", {}, "First name", f),
    h("label", {}, "Last name *", l),
    h("label", { class: "full" }, "Title", title),
    companyLabel,
    accountNote,
    h(
      "div",
      { class: "full", style: "display:flex;gap:6px;justify-content:flex-end" },
      h("button", { class: "secondary", type: "button", onclick: () => ((S.adding = null), renderPeople()) }, "Cancel"),
      save
    )
  );
}

function renderRelated() {
  const leadSelected = S.who?.type === "Lead";
  $("search").disabled = leadSelected;
  if (leadSelected) {
    $("related").replaceChildren(h("div", { class: "muted small" }, "Activities on a lead can't also be related to an account, opportunity or case."));
    return;
  }
  const opts = [...S.suggestions];
  if (S.what && !opts.some((o) => o.id === S.what.id)) opts.unshift(S.what);
  const row = (opt) => {
    const selected = (S.what?.id || null) === (opt?.id || null);
    return h(
      "label",
      { class: `row selectable${selected ? " selected" : ""}` },
      h("input", { type: "radio", name: "what", checked: selected, onchange: () => selectWhat(opt) }),
      h("div", { class: "main" + (opt ? "" : " muted") }, h("div", { class: "name" }, opt ? opt.label : "Nothing")),
      opt ? h("span", { class: "badge" }, opt.type) : null
    );
  };
  $("related").replaceChildren(...opts.map(row), row(null));
}

let searchTimer;
function onSearchInput(e) {
  clearTimeout(searchTimer);
  const term = e.target.value;
  searchTimer = setTimeout(async () => {
    try {
      S.searchResults = term.trim().length >= 2 ? await sf.searchRelated(term) : [];
      if ($("search").value === term) renderSearch();
    } catch (err) {
      handleError(err);
    }
  }, 300);
}

function renderSearch() {
  const term = $("search").value.trim();
  if (!term) return $("search-results").replaceChildren();
  if (!S.searchResults.length) return $("search-results").replaceChildren(h("div", { class: "muted small" }, term.length < 2 ? "" : "No matches."));
  $("search-results").replaceChildren(
    ...S.searchResults.map((opt) =>
      h(
        "div",
        { class: "row selectable", role: "button", tabindex: "0", onclick: () => selectWhat(opt), onkeydown: (e) => e.key === "Enter" && selectWhat(opt) },
        h("div", { class: "main" }, h("div", { class: "name" }, opt.label)),
        h("span", { class: "badge" }, opt.type)
      )
    )
  );
}

function renderOptions() {
  const rows = [
    h(
      "label",
      { class: "row selectable" },
      h("input", { type: "checkbox", checked: S.includeBody, onchange: (e) => (S.includeBody = e.target.checked) }),
      h("div", { class: "main" }, h("div", { class: "name" }, S.item.kind === "meeting" ? "Meeting notes / invite text" : "Email body"))
    ),
  ];
  for (const a of S.item.attachments) {
    const tooBig = a.size > CONFIG.maxAttachmentBytes;
    rows.push(
      h(
        "label",
        { class: `row ${tooBig ? "disabled" : "selectable"}` },
        h("input", {
          type: "checkbox",
          checked: S.files.has(a.id),
          disabled: tooBig,
          onchange: (e) => (e.target.checked ? S.files.add(a.id) : S.files.delete(a.id)),
        }),
        h("div", { class: "main" }, h("div", { class: "name" }, a.name), h("div", { class: "sub" }, tooBig ? `${sizeLabel(a.size)} · too large to upload` : sizeLabel(a.size)))
      )
    );
  }
  $("options").replaceChildren(...rows);
}

function renderActions() {
  const btn = $("log-btn");
  btn.disabled = S.busy;
  btn.textContent = S.busy ? "Logging…" : S.logged ? "Log again" : S.item.kind === "meeting" ? "Log meeting" : "Log email";
}

// ── Logging ───────────────────────────────────────────────────────────────────
async function onLog() {
  if (S.busy) return;
  const it = S.item;
  showError("");
  S.busy = true;
  renderActions();
  setStatus("Creating activity…");

  try {
    const whoId = S.who?.id || null;
    const whatId = S.who?.type === "Lead" ? null : S.what?.id || null;
    let description = it.header;
    if (S.includeBody && it.body.trim()) description += `\n\n${it.body.trim()}`;
    description = description.slice(0, CONFIG.maxDescriptionChars);

    const activityId =
      it.kind === "meeting"
        ? await sf.logMeetingEvent({ subject: it.subject, description, whoId, whatId, start: it.start, end: it.end, location: it.location })
        : await sf.logEmailTask({ subject: it.subject, description, whoId, whatId, date: ymd(it.date) });

    const uploaded = [];
    const failed = [];
    const files = it.attachments.filter((a) => S.files.has(a.id));
    for (const [i, a] of files.entries()) {
      setStatus(`Uploading ${i + 1} of ${files.length}: ${a.name}`);
      try {
        const base64 = await ol.getAttachmentBase64(a.id);
        await sf.uploadFile({ name: a.name, base64, recordId: whatId || whoId || activityId, activityId });
        uploaded.push(a.name);
      } catch (e) {
        if (e instanceof AuthError) throw e;
        failed.push(`${a.name} (${e.message})`);
      }
    }

    S.logged = { id: activityId, at: new Date().toISOString(), kind: it.kind };
    await ol.setLoggedMarker(S.logged);
    setStatus("");
    renderLogged(
      h(
        "div",
        { class: "banner banner-ok" },
        `${it.kind === "meeting" ? "Meeting" : "Email"} logged`,
        uploaded.length ? ` with ${uploaded.length} file${uploaded.length > 1 ? "s" : ""}` : "",
        ". ",
        link(activityId, "View in Salesforce")
      )
    );
    if (failed.length) showError(`Some files weren't uploaded: ${failed.join("; ")}`);
  } catch (e) {
    setStatus("");
    handleError(e);
  } finally {
    S.busy = false;
    if (S.item) renderActions();
  }
}
