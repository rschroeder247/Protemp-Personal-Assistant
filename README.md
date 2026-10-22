# Protemp Personal Assistant

Voice-powered personal assistant built with **Next.js**, **Supabase**, **Microsoft Graph (MSAL)**, and **WhatsApp Cloud API / Twilio**, deployed on **Vercel**.

## Capabilities
- **Voice Note Command Center**: Speak naturally (e.g., *"Send a WhatsApp to John saying the HVAC quote is ready, create a Planner task to order piping by Friday, and add an Outlook task to call the supplier tomorrow"*) and extract multiple actions at once.
- **Background WhatsApp API Sending**: Send WhatsApp messages via **Meta WhatsApp Business Cloud API** or **Twilio WhatsApp API** (`/api/whatsapp/send`).
- **Microsoft Outlook Tasks (To Do)**: Create and complete tasks in Microsoft To Do / Outlook with due dates and priorities.
- **Microsoft Planner Integration**: Select a Planner Plan & Bucket and generate tasks assigned to you with due dates, priority, and notes.
- **Outlook Email & Calendar**: Send or draft Outlook emails and schedule calendar events from voice commands.
- **Supabase Sync**: Persists contacts (`pa_contacts`), voice notes (`pa_voice_notes`), action history (`pa_actions`), and settings (`pa_settings`).
