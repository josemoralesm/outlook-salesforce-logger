// Salesforce OAuth 2.0 web-server flow with PKCE, run inside an Office dialog.
// No client secret is used or stored; tokens live in this add-in's local storage.
import { CONFIG, redirectUri, baseUrl } from "./config.js";

const KEY = "sfOutlook.session";

export class AuthError extends Error {
  constructor(message = "Please sign in to Salesforce.") {
    super(message);
    this.name = "AuthError";
  }
}

export function getSession() {
  try {
    return JSON.parse(localStorage.getItem(KEY));
  } catch {
    return null;
  }
}

function saveSession(session) {
  try {
    localStorage.setItem(KEY, JSON.stringify(session));
  } catch {
    /* storage unavailable: session lasts until the pane closes */
  }
  memorySession = session;
}

let memorySession = null;
export function currentSession() {
  return getSession() || memorySession;
}

export async function signOut() {
  const s = currentSession();
  try {
    localStorage.removeItem(KEY);
  } catch {}
  memorySession = null;
  if (s?.refreshToken) {
    // Best effort: revoke the refresh token so it can't be reused.
    fetch(`${CONFIG.loginUrl}/services/oauth2/revoke`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ token: s.refreshToken }),
    }).catch(() => {});
  }
}

function b64url(buffer) {
  let s = "";
  for (const b of new Uint8Array(buffer)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function randomString(bytes) {
  const a = new Uint8Array(bytes);
  crypto.getRandomValues(a);
  return b64url(a);
}

async function tokenRequest(params) {
  const res = await fetch(`${CONFIG.loginUrl}/services/oauth2/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: new URLSearchParams(params),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(json.error_description || json.error || `Salesforce sign-in failed (${res.status}).`);
  return json;
}

function openDialog(url) {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url, { height: 70, width: 35, promptBeforeOpen: false }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(res.error?.message || "Could not open the sign-in window."));
        return;
      }
      const dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        try {
          resolve(JSON.parse(arg.message));
        } catch (e) {
          reject(e);
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        reject(new Error(arg.error === 12006 ? "The sign-in window was closed." : `Sign-in failed (code ${arg.error}).`));
      });
    });
  });
}

export async function signIn() {
  if (CONFIG.clientId.startsWith("PASTE_")) {
    throw new Error("The add-in isn't configured yet: set clientId in src/config.js.");
  }
  const verifier = randomString(48); // 64 chars
  const state = randomString(16);
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));

  const authorize = new URL(`${CONFIG.loginUrl}/services/oauth2/authorize`);
  authorize.search = new URLSearchParams({
    response_type: "code",
    client_id: CONFIG.clientId,
    redirect_uri: redirectUri(),
    scope: CONFIG.scopes,
    code_challenge: challenge,
    code_challenge_method: "S256",
    state,
  }).toString();

  // Office dialogs must open on the add-in's own domain first; start.html then redirects to Salesforce.
  const start = new URL("auth/start.html", baseUrl());
  start.searchParams.set("u", authorize.href);

  const result = await openDialog(start.href);
  if (result.error) throw new Error(result.error_description || result.error);
  if (result.state !== state || !result.code) throw new Error("Sign-in response didn't match. Please try again.");

  const t = await tokenRequest({
    grant_type: "authorization_code",
    code: result.code,
    client_id: CONFIG.clientId,
    redirect_uri: redirectUri(),
    code_verifier: verifier,
  });
  saveSession({ accessToken: t.access_token, refreshToken: t.refresh_token || null, instanceUrl: t.instance_url });
  return currentSession();
}

// Called when an API request returns 401.
export async function refreshSession() {
  const s = currentSession();
  if (!s?.refreshToken) {
    await signOut();
    throw new AuthError("Your Salesforce session expired. Please sign in again.");
  }
  try {
    const t = await tokenRequest({ grant_type: "refresh_token", client_id: CONFIG.clientId, refresh_token: s.refreshToken });
    // Salesforce may rotate the refresh token; keep the newest one.
    const next = { ...s, accessToken: t.access_token, refreshToken: t.refresh_token || s.refreshToken, instanceUrl: t.instance_url || s.instanceUrl };
    saveSession(next);
    return next;
  } catch {
    await signOut();
    throw new AuthError("Your Salesforce session expired. Please sign in again.");
  }
}
