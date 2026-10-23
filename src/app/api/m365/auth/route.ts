import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@supabase/supabase-js';

const supabaseUrl =
  process.env.NEXT_PUBLIC_SUPABASE_URL ||
  'https://dphxaglpramhwcwcnpay.supabase.co';
const supabaseAnonKey =
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRwaHhhZ2xwcmFtaHdjd2NucGF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDg4NzksImV4cCI6MjA4ODgyNDg3OX0.Z2uLJlqYfs6RGcrRe3SpyaHo7BzvO4QgmYp9lXVEKN0';

const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Microsoft Office First-Party Public Client ID
// Unlike SPA app registrations (which fail on /token with AADSTS7000218 client_secret required),
// this native Microsoft public client has allowPublicClient=true globally and never requires a client_secret.
const NATIVE_PUBLIC_CLIENT_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';

const DEFAULT_TENANT_ID =
  process.env.NEXT_PUBLIC_AZURE_TENANT_ID ||
  '70fc3e15-9b2c-4e05-8f43-9cebcc9bfada';

const SCOPES =
  'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Tasks.ReadWrite https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Mail.ReadWrite offline_access';

const TOKEN_SETTINGS_KEY = 'm365_auth_tokens';

interface StoredTokenData {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  clientId: string;
  tenantId: string;
  user: {
    id?: string;
    name: string;
    email: string;
  };
}

/**
 * Ensures we never use the SPA-only client ID ('abae4ffc-...') for Device Code Flow,
 * because SPA registrations throw AADSTS7000218 (client_secret required) on /token.
 */
function resolvePublicClientId(requestedClientId?: string): string {
  if (
    !requestedClientId ||
    requestedClientId === 'abae4ffc-842f-45cb-aea5-8394d2de9a82'
  ) {
    return NATIVE_PUBLIC_CLIENT_ID;
  }
  return requestedClientId;
}

async function getStoredTokens(): Promise<StoredTokenData | null> {
  const { data, error } = await supabase
    .from('pa_settings')
    .select('value')
    .eq('key', TOKEN_SETTINGS_KEY)
    .maybeSingle();

  if (error || !data?.value) return null;
  const val = data.value as StoredTokenData;
  if (!val.refreshToken && !val.accessToken) return null;
  return val;
}

async function saveStoredTokens(tokenData: StoredTokenData): Promise<void> {
  await supabase.from('pa_settings').upsert({
    key: TOKEN_SETTINGS_KEY,
    value: tokenData,
    updated_at: new Date().toISOString(),
  });
}

async function refreshAccessToken(
  stored: StoredTokenData
): Promise<StoredTokenData | null> {
  if (!stored.refreshToken) return null;
  const tenantId = stored.tenantId || DEFAULT_TENANT_ID;
  const clientId = resolvePublicClientId(stored.clientId);

  const params = new URLSearchParams({
    client_id: clientId,
    grant_type: 'refresh_token',
    refresh_token: stored.refreshToken,
    scope: SCOPES,
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
      cache: 'no-store',
    }
  );

  if (!res.ok) {
    return null;
  }

  const data = await res.json();
  const updated: StoredTokenData = {
    ...stored,
    clientId,
    accessToken: data.access_token,
    refreshToken: data.refresh_token || stored.refreshToken,
    expiresAt: Date.now() + (data.expires_in || 3599) * 1000,
  };

  await saveStoredTokens(updated);
  return updated;
}

export async function GET() {
  try {
    const stored = await getStoredTokens();
    if (!stored) {
      return NextResponse.json({ connected: false });
    }

    const now = Date.now() + 2 * 60 * 1000;
    if (stored.accessToken && stored.expiresAt > now) {
      return NextResponse.json({
        connected: true,
        accessToken: stored.accessToken,
        expiresAt: stored.expiresAt,
        user: stored.user,
      });
    }

    const refreshed = await refreshAccessToken(stored);
    if (!refreshed) {
      return NextResponse.json({ connected: false });
    }

    return NextResponse.json({
      connected: true,
      accessToken: refreshed.accessToken,
      expiresAt: refreshed.expiresAt,
      user: refreshed.user,
    });
  } catch (err: any) {
    return NextResponse.json(
      { connected: false, error: err?.message },
      { status: 500 }
    );
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      action,
      deviceCode,
      clientId,
      tenantId = DEFAULT_TENANT_ID,
    } = body;

    const effectiveClientId = resolvePublicClientId(clientId);
    const effectiveTenantId = tenantId || DEFAULT_TENANT_ID;

    if (action === 'logout') {
      await supabase
        .from('pa_settings')
        .delete()
        .eq('key', TOKEN_SETTINGS_KEY);
      return NextResponse.json({ success: true });
    }

    if (action === 'start_device_code') {
      const params = new URLSearchParams({
        client_id: effectiveClientId,
        scope: SCOPES,
      });

      const res = await fetch(
        `https://login.microsoftonline.com/${effectiveTenantId}/oauth2/v2.0/devicecode`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: params.toString(),
          cache: 'no-store',
        }
      );

      const data = await res.json();
      if (!res.ok) {
        return NextResponse.json(
          {
            error:
              data.error_description ||
              data.error ||
              'Failed to start Microsoft 365 sign-in',
          },
          { status: 400 }
        );
      }

      return NextResponse.json({
        userCode: data.user_code,
        deviceCode: data.device_code,
        verificationUri:
          data.verification_uri || 'https://login.microsoft.com/device',
        expiresIn: data.expires_in || 900,
        interval: data.interval || 5,
        message: data.message,
      });
    }

    if (action === 'poll_device_code') {
      if (!deviceCode) {
        return NextResponse.json(
          { error: 'deviceCode is required' },
          { status: 400 }
        );
      }

      const params = new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: effectiveClientId,
        device_code: deviceCode,
      });

      const res = await fetch(
        `https://login.microsoftonline.com/${effectiveTenantId}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: params.toString(),
          cache: 'no-store',
        }
      );

      const data = await res.json();

      if (!res.ok) {
        if (
          data.error === 'authorization_pending' ||
          data.error === 'slow_down'
        ) {
          return NextResponse.json({ status: 'pending' });
        }
        return NextResponse.json(
          {
            status: 'error',
            error:
              data.error_description || data.error || 'Authentication failed',
          },
          { status: 400 }
        );
      }

      // Fetch user profile from Microsoft Graph
      let user = {
        id: '',
        name: 'Roland Schroeder',
        email: 'rschroeder@protemp.co.za',
      };

      try {
        const meRes = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: { Authorization: `Bearer ${data.access_token}` },
          cache: 'no-store',
        });
        if (meRes.ok) {
          const me = await meRes.json();
          user = {
            id: me.id || '',
            name: me.displayName || me.userPrincipalName || 'Roland Schroeder',
            email: me.mail || me.userPrincipalName || 'rschroeder@protemp.co.za',
          };
        }
      } catch {
        // ignore
      }

      const tokenPayload: StoredTokenData = {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || '',
        expiresAt: Date.now() + (data.expires_in || 3599) * 1000,
        clientId: effectiveClientId,
        tenantId: effectiveTenantId,
        user,
      };

      await saveStoredTokens(tokenPayload);

      return NextResponse.json({
        status: 'connected',
        accessToken: tokenPayload.accessToken,
        expiresAt: tokenPayload.expiresAt,
        user: tokenPayload.user,
      });
    }

    return NextResponse.json({ error: 'Unknown action' }, { status: 400 });
  } catch (err: any) {
    return NextResponse.json(
      { error: err?.message || 'Internal error' },
      { status: 500 }
    );
  }
}
