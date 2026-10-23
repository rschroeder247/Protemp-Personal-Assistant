import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@supabase/supabase-js';

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

export async function GET(req: NextRequest) {
  const url = new URL(req.url);
  const code = url.searchParams.get('code');
  const error = url.searchParams.get('error');
  const errorDesc = url.searchParams.get('error_description');

  const appHome = new URL('/', req.url);

  if (error || !code) {
    appHome.searchParams.set(
      'm365_error',
      errorDesc || error || 'Missing authorization code'
    );
    return NextResponse.redirect(appHome);
  }

  try {
    // 1. Load PKCE code_verifier from Supabase
    const { data: pkceRow } = await supabase
      .from('pa_settings')
      .select('value')
      .eq('key', PKCE_SETTINGS_KEY)
      .maybeSingle();

    const codeVerifier = (pkceRow?.value as any)?.codeVerifier;
    if (!codeVerifier) {
      appHome.searchParams.set(
        'm365_error',
        'PKCE session expired. Please click Connect Microsoft 365 again.'
      );
      return NextResponse.redirect(appHome);
    }

    // 2. Exchange authorization code for access_token & refresh_token using the registered SPA Origin header
    const params = new URLSearchParams({
      client_id: SPA_CLIENT_ID,
      grant_type: 'authorization_code',
      code,
      redirect_uri: REGISTERED_SPA_ORIGIN,
      code_verifier: codeVerifier,
      scope: SCOPES,
    });

    const tokenRes = await fetch(
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

    const tokenData = await tokenRes.json();
    if (!tokenRes.ok || tokenData.error) {
      appHome.searchParams.set(
        'm365_error',
        tokenData.error_description || tokenData.error || 'Token exchange failed'
      );
      return NextResponse.redirect(appHome);
    }

    // 3. Fetch signed-in user profile from Microsoft Graph
    let user = {
      id: '',
      name: 'Roland Schroeder',
      email: 'rschroeder@protemp.co.za',
    };

    try {
      const meRes = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
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

    // 4. Save access_token + refresh_token into Supabase
    await supabase.from('pa_settings').upsert({
      key: TOKEN_SETTINGS_KEY,
      value: {
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || '',
        expiresAt: Date.now() + (tokenData.expires_in || 3599) * 1000,
        clientId: SPA_CLIENT_ID,
        tenantId: TENANT_ID,
        user,
      },
      updated_at: new Date().toISOString(),
    });

    appHome.searchParams.set('m365_connected', '1');
    return NextResponse.redirect(appHome);
  } catch (err: any) {
    appHome.searchParams.set(
      'm365_error',
      err?.message || 'Unexpected callback error'
    );
    return NextResponse.redirect(appHome);
  }
}
