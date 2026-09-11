// Handles Entra sign-in and token acquisition using MSAL.js.

let msalInstance = null;

// Build the MSAL client lazily, so it is only constructed when first needed
// (after the msal library and CONFIG have definitely loaded).
function getMsal() {
  if (!msalInstance) {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        // Tenant-specific authority restricts sign-in to your own tenant.
        // Use "https://login.microsoftonline.com/common" only if you
        // genuinely need users from other tenants.
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });
  }
  return msalInstance;
}

// MSAL v3 must finish initialising before any other call.
async function initAuth() {
  const client = getMsal();
  await client.initialize();
  await client.handleRedirectPromise();

  const accounts = client.getAllAccounts();
  if (accounts.length > 0) {
    client.setActiveAccount(accounts[0]);
  }
  return client.getActiveAccount();
}

async function signIn() {
  const client = getMsal();
  await client.initialize();
  const result = await client.loginPopup({ scopes: [CONFIG.workIqScope] });
  client.setActiveAccount(result.account);
  return result.account;
}

async function signOut() {
  await getMsal().logoutPopup();
}

// Returns a bearer token for the Work IQ API.
// Tries the silent path first (cache or refresh token); only prompts if that fails.
async function getWorkIqToken() {
  const client = getMsal();
  const account = client.getActiveAccount();
  if (!account) throw new Error("Not signed in");

  const request = { scopes: [CONFIG.workIqScope], account };

  try {
    const result = await client.acquireTokenSilent(request);
    return result.accessToken;
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      const result = await client.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw error;
  }
}
