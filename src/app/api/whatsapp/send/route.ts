import { NextRequest, NextResponse } from 'next/server';

function normalizePhoneNumber(phone: string, defaultCountryCode = '+27'): string {
  let cleaned = phone.replace(/[^\d+]/g, '');
  if (!cleaned) return '';

  // If starts with 0 (e.g. 0821234567) and we have a country code (+27), replace leading 0
  if (cleaned.startsWith('0') && defaultCountryCode) {
    const code = defaultCountryCode.startsWith('+')
      ? defaultCountryCode
      : `+${defaultCountryCode}`;
    cleaned = `${code}${cleaned.slice(1)}`;
  } else if (!cleaned.startsWith('+')) {
    cleaned = `+${cleaned}`;
  }
  return cleaned;
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      to,
      message,
      provider = process.env.WHATSAPP_PROVIDER || 'meta',
      defaultCountryCode = '+27',
      // Credentials passed from Settings UI (or fallback to server env vars)
      metaAccessToken = process.env.META_WHATSAPP_TOKEN,
      metaPhoneNumberId = process.env.META_WHATSAPP_PHONE_NUMBER_ID,
      metaTemplateName = '',
      twilioAccountSid = process.env.TWILIO_ACCOUNT_SID,
      twilioAuthToken = process.env.TWILIO_AUTH_TOKEN,
      twilioFromNumber = process.env.TWILIO_WHATSAPP_FROM || 'whatsapp:+14155238886',
    } = body;

    if (!to || !message) {
      return NextResponse.json(
        { error: 'Recipient phone number and message body are required.' },
        { status: 400 }
      );
    }

    const normalizedE164 = normalizePhoneNumber(to, defaultCountryCode);
    if (!normalizedE164 || normalizedE164.length < 8) {
      return NextResponse.json(
        { error: `Invalid recipient phone number: "${to}". Please enter a valid number (e.g. +27821234567).` },
        { status: 400 }
      );
    }

    // ==========================================================
    // OPTION A: META WHATSAPP BUSINESS CLOUD API
    // ==========================================================
    if (provider === 'meta') {
      if (!metaAccessToken || !metaPhoneNumberId) {
        return NextResponse.json(
          {
            error:
              'Meta WhatsApp Business Cloud API credentials are not configured yet. Open Settings -> WhatsApp API and enter your Meta Access Token & Phone Number ID (or switch to Twilio).',
          },
          { status: 400 }
        );
      }

      // Meta API expects number without leading '+'
      const metaRecipient = normalizedE164.replace(/^\+/, '');
      const url = `https://graph.facebook.com/v21.0/${encodeURIComponent(
        metaPhoneNumberId.trim()
      )}/messages`;

      const payload: any = {
        messaging_product: 'whatsapp',
        recipient_type: 'individual',
        to: metaRecipient,
      };

      if (metaTemplateName && metaTemplateName.trim()) {
        payload.type = 'template';
        payload.template = {
          name: metaTemplateName.trim(),
          language: { code: 'en' },
        };
      } else {
        payload.type = 'text';
        payload.text = {
          preview_url: false,
          body: message,
        };
      }

      const metaRes = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${metaAccessToken.trim()}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      });

      const metaData = await metaRes.json();

      if (!metaRes.ok) {
        const apiErr =
          metaData?.error?.message ||
          JSON.stringify(metaData?.error || metaData);
        return NextResponse.json(
          { error: `Meta WhatsApp Cloud API Error: ${apiErr}` },
          { status: metaRes.status }
        );
      }

      const messageId = metaData?.messages?.[0]?.id || `meta-${Date.now()}`;
      return NextResponse.json({
        success: true,
        provider: 'meta',
        messageId,
        to: normalizedE164,
      });
    }

    // ==========================================================
    // OPTION B: TWILIO WHATSAPP API
    // ==========================================================
    if (provider === 'twilio') {
      if (!twilioAccountSid || !twilioAuthToken) {
        return NextResponse.json(
          {
            error:
              'Twilio WhatsApp credentials are not configured yet. Open Settings -> WhatsApp API and enter your Twilio Account SID & Auth Token.',
          },
          { status: 400 }
        );
      }

      const sid = twilioAccountSid.trim();
      const token = twilioAuthToken.trim();
      const fromFormatted = twilioFromNumber.trim().startsWith('whatsapp:')
        ? twilioFromNumber.trim()
        : `whatsapp:${twilioFromNumber.trim()}`;
      const toFormatted = `whatsapp:${normalizedE164}`;

      const twilioUrl = `https://api.twilio.com/2010-04-01/Accounts/${encodeURIComponent(
        sid
      )}/Messages.json`;

      const formParams = new URLSearchParams();
      formParams.append('From', fromFormatted);
      formParams.append('To', toFormatted);
      formParams.append('Body', message);

      const basicAuth = Buffer.from(`${sid}:${token}`).toString('base64');

      const twilioRes = await fetch(twilioUrl, {
        method: 'POST',
        headers: {
          Authorization: `Basic ${basicAuth}`,
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: formParams.toString(),
      });

      const twilioData = await twilioRes.json();

      if (!twilioRes.ok) {
        const apiErr =
          twilioData?.message || JSON.stringify(twilioData);
        return NextResponse.json(
          { error: `Twilio WhatsApp API Error: ${apiErr}` },
          { status: twilioRes.status }
        );
      }

      return NextResponse.json({
        success: true,
        provider: 'twilio',
        messageId: twilioData.sid || `twilio-${Date.now()}`,
        to: normalizedE164,
      });
    }

    return NextResponse.json(
      { error: `Unsupported WhatsApp provider: ${provider}` },
      { status: 400 }
    );
  } catch (error: any) {
    console.error('WhatsApp send API error:', error);
    return NextResponse.json(
      { error: error?.message || 'Internal server error sending WhatsApp message' },
      { status: 500 }
    );
  }
}
