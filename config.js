// Public configuration. No secrets - a browser app cannot keep secrets,
// which is why this app uses MSAL with PKCE rather than a client secret.
//
// Replace these placeholders, or create a gitignored config.local.js that
// overrides them (see README). config.local.js is loaded after this file.

const CONFIG = {
  clientId: "YOUR_ENTRA_APP_CLIENT_ID",
  tenantId: "YOUR_TENANT_ID",

  // Microsoft Work IQ - these are Microsoft's own public identifiers
  workIqResourceAppId: "fdcc1f02-fc51-4226-8753-f668596af7f7",
  a2aEndpoint: "https://workiq.svc.cloud.microsoft/a2a/",
};

// Scope format is "<resource-app-id>/<scope-name>". Entra uses this to decide
// which API the token is minted for.
CONFIG.workIqScope = `${CONFIG.workIqResourceAppId}/WorkIQAgent.Ask`;
