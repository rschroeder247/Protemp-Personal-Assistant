import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@supabase/supabase-js';
import crypto from 'crypto';

const supabaseUrl =
  process.env.NEXT_PUBLIC_SUPABASE_URL ||
  'https://dphxaglpramhwcwcnpay.supabase.co';
const supabaseAnonKey =
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRwaHhhZ2xwcmFtaHdjd2NucGF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDg4NzksImV4cCI6MjA4ODgyNDg3OX0.Z2uLJlqYfs6RGcrRe3SpyaHo7BzvO4QgmYp9lXVEKN0';

const supabase = createClient(supabaseUrl, supabaseAnonKey);

const SPA_CLIENT_ID = 'abae4ffc-842f-45cb-aea5-8394d2de9a82';
const TENANT_ID = '70fc3e15-9b2c-4e05-8f43-9cebcc9bfada';
const REGISTERED_SPA_ORIGIN = 'https://eisenhower-matrix-flow.vercel.app';
const SCOPES = 'User.Read Calendars.ReadWrite Tasks.ReadWrite offline_access';

const PKCE_SETTINGS_KEY = 'm365_pkce_verifier';
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

function base64UrlEncode(buf: Buffer): string {
  return buf
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
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

  const params = new URLSearchParams({
    client_id: SPA_CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: stored.refreshToken,
    scope: SCOPES,
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Origin: REGISTERED_SPA_ORIGIN,
      },
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
    clientId: SPA_CLIENT_ID,
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
    const { action, origin = 'https://protemp-personal-assistant.vercel.app' } =
      body;

    if (action === 'logout') {
      await supabase
        .from('pa_settings')
        .delete()
        .eq('key', TOKEN_SETTINGS_KEY);
      return NextResponse.json({ success: true });
    }

    if (action === 'start_pkce_login') {
      const codeVerifier = base64UrlEncode(crypto.randomBytes(32));
      const codeChallenge = base64UrlEncode(
        crypto.createHash('sha256').update(codeVerifier).digest()
      );

      await supabase.from('pa_settings').upsert({
        key: PKCE_SETTINGS_KEY,
        value: {
          codeVerifier,
          createdAt: Date.now(),
        },
        updated_at: new Date().toISOString(),
      });

      const encodedOrigin = Buffer.from(origin).toString('base64');
      const state = `pa_bridge_${encodedOrigin}`;

      const authUrl = new URL(
        `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`
      );
      authUrl.searchParams.set('client_id', SPA_CLIENT_ID);
      authUrl.searchParams.set('response_type', 'code');
      authUrl.searchParams.set('redirect_uri', REGISTERED_SPA_ORIGIN);
      authUrl.searchParams.set('response_mode', 'query');
      authUrl.searchParams.set('scope', SCOPES);
      authUrl.searchParams.set('state', state);
      authUrl.searchParams.set('code_challenge', codeChallenge);
      authUrl.searchParams.set('code_challenge_method', 'S256');
      authUrl.searchParams.set('prompt', 'select_account');

      return NextResponse.json({
        authorizeUrl: authUrl.toString(),
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
