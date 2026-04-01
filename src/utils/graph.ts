import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from "@azure/msal-browser";
import type { ParsedContact } from "./parser";

// -----------------------------------------------------------------------
// CONFIGURATION — fill these in after Azure App Registration
// -----------------------------------------------------------------------
const CLIENT_ID = "de056afa-b318-4d3b-bc1a-35e89d947f79";
const REDIRECT_URI = "https://outlook-contact.vercel.app/taskpane.html";
// Use "common" to support both personal Microsoft accounts and work/school accounts
const AUTHORITY = "https://login.microsoftonline.com/common";
// -----------------------------------------------------------------------

const SCOPES = ["Contacts.ReadWrite", "User.Read"];

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: AUTHORITY,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage" as const,
    storeAuthStateInCookie: false,
  },
};

let msalInstance: PublicClientApplication | null = null;

async function getMsalInstance(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(msalConfig);
    await msalInstance.initialize();
  }
  return msalInstance;
}

/**
 * Returns a valid access token. Tries silent acquisition first (cached token),
 * falls back to a popup login if needed.
 */
export async function getAccessToken(): Promise<string> {
  const msal = await getMsalInstance();

  const accounts: AccountInfo[] = msal.getAllAccounts();

  if (accounts.length > 0) {
    try {
      const result: AuthenticationResult = await msal.acquireTokenSilent({
        scopes: SCOPES,
        account: accounts[0],
      });
      return result.accessToken;
    } catch {
      // Silent acquisition failed (expired, consent needed) — fall through to popup
    }
  }

  // No cached account or silent failed — show login popup
  const result: AuthenticationResult = await msal.loginPopup({
    scopes: SCOPES,
  });
  return result.accessToken;
}

/**
 * Returns the signed-in account, or null if no one is signed in.
 */
export async function getSignedInAccount(): Promise<AccountInfo | null> {
  const msal = await getMsalInstance();
  const accounts = msal.getAllAccounts();
  return accounts[0] ?? null;
}

/**
 * Signs the user out and clears the token cache.
 */
export async function signOut(): Promise<void> {
  const msal = await getMsalInstance();
  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    await msal.logoutPopup({ account: accounts[0] });
  }
}

/**
 * Creates a contact in the signed-in user's Outlook Contacts via Microsoft Graph.
 * Throws an error (with message) if the API call fails.
 */
// Removes keys with null, undefined, or empty string values so the Graph API
// doesn't reject the request due to unexpected null fields.
function omitEmpty(obj: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(obj).filter(([, v]) => v !== null && v !== undefined && v !== "")
  );
}

export async function createContact(contact: ParsedContact): Promise<void> {
  const token = await getAccessToken();

  const hasAddress = contact.street || contact.city;

  const body = omitEmpty({
    givenName: contact.firstName,
    surname: contact.lastName,
    emailAddresses: contact.email
      ? [{ address: contact.email, name: `${contact.firstName} ${contact.lastName}`.trim() }]
      : undefined,
    businessPhones: contact.businessPhone ? [contact.businessPhone] : undefined,
    mobilePhone: contact.mobilePhone || undefined,
    companyName: contact.company || undefined,
    jobTitle: contact.jobTitle || undefined,
    businessHomePage: contact.website || undefined,
    businessAddress: hasAddress
      ? omitEmpty({
          street: contact.street,
          city: contact.city,
          state: contact.state,
          postalCode: contact.zip,
          countryOrRegion: contact.country,
        })
      : undefined,
  });

  const response = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    const message = (error as { error?: { message?: string } }).error?.message
      ?? `Graph API error: ${response.status}`;
    throw new Error(message);
  }
}
