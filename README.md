# Log to Salesforce (Outlook add-in)

Hosted at https://josemoralesm.github.io/outlook-salesforce-logger/

## Before it works
1. **Salesforce app key:** open `config.js` and paste the Salesforce Consumer Key between the quotes on the `clientId` line, then commit and push.
2. **Salesforce admin settings:**
   - Callback URL: `https://josemoralesm.github.io/outlook-salesforce-logger/auth/callback.html`
   - CORS allowed origin: `https://josemoralesm.github.io`
   - Enable CORS for OAuth endpoints; require PKCE; no client secret required.
3. **Outlook:** add `manifest.xml` at https://aka.ms/olksideload → My add-ins → Add a custom add-in → Add from file.
