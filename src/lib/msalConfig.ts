import type { Configuration } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

export const AZURE_CLIENT_ID_STORAGE_KEY = 'pa_azure_client_id';
export const AZURE_TENANT_ID_STORAGE_KEY = 'pa_azure_tenant_id';

export const DEFAULT_CLIENT_ID =
  process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || 'abae4ffc-842f-45cb-aea5-8394d2de9a82';
export const DEFAULT_TENANT_ID =
  process.env.NEXT_PUBLIC_AZURE_TENANT_ID || '70fc3e15-9b2c-4e05-8f43-9cebcc9bfada';

export const getStoredClientId = (): string => {
  if (typeof window === 'undefined') return DEFAULT_CLIENT_ID;
  return localStorage.getItem(AZURE_CLIENT_ID_STORAGE_KEY) || DEFAULT_CLIENT_ID;
};

export const getStoredTenantId = (): string => {
  if (typeof window === 'undefined') return DEFAULT_TENANT_ID;
  const stored = localStorage.getItem(AZURE_TENANT_ID_STORAGE_KEY);
  if (!stored || stored === 'common') {
    return DEFAULT_TENANT_ID;
  }
  return stored;
};

export const setStoredAzureConfig = (
  clientId: string,
  tenantId: string = DEFAULT_TENANT_ID
) => {
  if (typeof window === 'undefined') return;
  localStorage.setItem(AZURE_CLIENT_ID_STORAGE_KEY, clientId.trim() || DEFAULT_CLIENT_ID);
  localStorage.setItem(AZURE_TENANT_ID_STORAGE_KEY, tenantId.trim() || DEFAULT_TENANT_ID);
};

export const createMsalConfig = (
  clientId: string,
  tenantId: string = DEFAULT_TENANT_ID
): Configuration => {
  const origin =
    typeof window !== 'undefined' ? window.location.origin : 'http://localhost:3000';
  return {
    auth: {
      clientId: clientId || DEFAULT_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${tenantId || DEFAULT_TENANT_ID}`,
      redirectUri: origin,
      postLogoutRedirectUri: origin,
    },
    cache: {
      cacheLocation: 'localStorage',
    },
    system: {
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (containsPii) return;
          if (level === LogLevel.Error) {
            console.error('[MSAL Error]:', message);
          }
        },
        logLevel: LogLevel.Warning,
      },
    },
  };
};

// User-consentable Microsoft Graph scopes
export const GRAPH_SCOPES = [
  'User.Read',
  'Tasks.ReadWrite',
  'Calendars.ReadWrite',
];

export const OPTIONAL_MAIL_SCOPES = [
  'Mail.Send',
  'Mail.ReadWrite',
];
