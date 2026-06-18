import {
  type IPublicClientApplication,
  createNestablePublicClientApplication,
} from "@azure/msal-browser";

const CLIENT_ID = import.meta.env.VITE_CLIENT_ID as string;
const TENANT_ID = import.meta.env.VITE_TENANT_ID as string;

let msalInstance: IPublicClientApplication | null = null;

async function getInstance(): Promise<IPublicClientApplication> {
  if (msalInstance) return msalInstance;

  msalInstance = await createNestablePublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
    system: {
      allowNativeBroker: false,
    },
  });

  return msalInstance;
}

const SCOPES = [
  "openid",
  "profile",
  "email",
  `api://youvaibr.github.io/${CLIENT_ID}/access_as_user`,
];

export async function acquireToken(): Promise<string> {
  const msal = await getInstance();
  const account = msal.getAllAccounts()[0];

  try {
    const result = await msal.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch {
    const result = await msal.acquireTokenPopup({ scopes: SCOPES });
    return result.accessToken;
  }
}
