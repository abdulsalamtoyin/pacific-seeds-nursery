// Microsoft sign-in, for OneDrive sync.
//
// MSAL is vendored like ExcelJS and loaded only when sync is configured.
// Sign-in uses a popup rather than a page redirect: inside a SharePoint page
// the app sits in a frame, and Microsoft's sign-in page refuses to load in one.
//
// Background syncs ask for tokens silently. If that needs the user (a password
// change, a new consent), it fails with NeedsSignInError instead of opening a
// popup nobody clicked for — browsers block those anyway.

import { SYNC_CONFIG } from "./sync-config.js";

// Named so sync.js can recognise it without importing MSAL-side code.
export class NeedsSignInError extends Error { name = "NeedsSignInError"; }

let msal = null;
let client = null;
let account = null;

export const syncConfigured = () =>
  Boolean(SYNC_CONFIG.clientId && SYNC_CONFIG.tenantId);

function loadMsal() {
  if (globalThis.msal) return Promise.resolve(globalThis.msal);
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = new URL("./vendor/msal-browser.min.js", import.meta.url).href;
    script.onload = () => resolve(globalThis.msal);
    script.onerror = () => reject(new Error(
      "The sign-in library (vendor/msal-browser.min.js) could not be loaded."));
    document.head.append(script);
  });
}

/** Set up MSAL and return the account already signed in here, if any. */
export async function initAuth() {
  msal = await loadMsal();
  client = new msal.PublicClientApplication({
    auth: {
      clientId: SYNC_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${SYNC_CONFIG.tenantId}`,
      redirectUri: new URL("./auth-redirect.html", import.meta.url).href,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await client.initialize();
  account = client.getActiveAccount() ?? client.getAllAccounts()[0] ?? null;
  if (account) client.setActiveAccount(account);
  return account;
}

/** Interactive — call from a click. */
export async function signIn() {
  const result = await client.loginPopup({ scopes: SYNC_CONFIG.scopes });
  account = result.account;
  client.setActiveAccount(account);
  return account;
}

export async function signOut() {
  const leaving = account;
  account = null;
  await client.logoutPopup({
    account: leaving,
    postLogoutRedirectUri: new URL("./auth-redirect.html", import.meta.url).href,
  });
}

export const currentAccount = () => account;

/** An access token for Graph, without user interaction. */
export async function getToken({ force = false } = {}) {
  if (!account) throw new NeedsSignInError("Not signed in to Microsoft.");
  try {
    const result = await client.acquireTokenSilent({
      scopes: SYNC_CONFIG.scopes, account, forceRefresh: force,
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      throw new NeedsSignInError("Microsoft needs you to sign in again.");
    }
    throw err;
  }
}
