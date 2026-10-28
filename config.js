// ─── Edit these values for your Salesforce org ────────────────────────────────
export const CONFIG = {
  // Consumer Key of your Salesforce External Client App (or Connected App).
  clientId: "3MVG9szVa2RxsqBbLdHxYuJESe5Tgi7n.9VOGfn.vXqZaQqUnsl6akIqiFSf_zSmCAQH.cvxU85mMlkwycjf.",

  // Where users sign in. Use your My Domain for production, e.g.
  // "https://surfaceoptics.my.salesforce.com", or "https://test.salesforce.com" for a sandbox.
  loginUrl: "https://surfaceoptics.my.salesforce.com",

  apiVersion: "v62.0",
  scopes: "api refresh_token",

  // Picklist values used when logging. Change if your org customised them.
  taskStatus: "Completed",
  taskPriority: "Normal",

  // Attachments larger than this are not offered for upload.
  maxAttachmentBytes: 25 * 1024 * 1024,

  // Salesforce Task/Event Description limit.
  maxDescriptionChars: 32000,
};

// OAuth callback page. Must match the Callback URL in the Salesforce app exactly.
export function redirectUri() {
  return new URL("auth/callback.html", baseUrl()).href;
}

// Folder the add-in is served from (works for both taskpane.html and auth/*.html).
export function baseUrl() {
  const here = new URL(".", window.location.href);
  return here.pathname.endsWith("/auth/") ? new URL("..", here).href : here.href;
}
