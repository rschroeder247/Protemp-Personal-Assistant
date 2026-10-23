export const AZURE_CLIENT_ID_STORAGE_KEY = 'pa_azure_client_id';
export const AZURE_TENANT_ID_STORAGE_KEY = 'pa_azure_tenant_id';

// Microsoft Office native public client (supports Device Code Flow without client_secret / AADSTS7000218)
export const DEFAULT_CLIENT_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';
export const DEFAULT_TENANT_ID =
  process.env.NEXT_PUBLIC_AZURE_TENANT_ID || '70fc3e15-9b2c-4e05-8f43-9cebcc9bfada';

export const getStoredClientId = (): string => {
  if (typeof window === 'undefined') return DEFAULT_CLIENT_ID;
  const stored = localStorage.getItem(AZURE_CLIENT_ID_STORAGE_KEY);
  if (!stored || stored === 'abae4ffc-842f-45cb-aea5-8394d2de9a82') {
    return DEFAULT_CLIENT_ID;
  }
  return stored;
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
  const cleanId =
    !clientId.trim() || clientId.trim() === 'abae4ffc-842f-45cb-aea5-8394d2de9a82'
      ? DEFAULT_CLIENT_ID
      : clientId.trim();
  localStorage.setItem(AZURE_CLIENT_ID_STORAGE_KEY, cleanId);
  localStorage.setItem(AZURE_TENANT_ID_STORAGE_KEY, tenantId.trim() || DEFAULT_TENANT_ID);
};
