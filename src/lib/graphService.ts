import { Client } from '@microsoft/microsoft-graph-client';
import { getStoredClientId, getStoredTenantId } from './msalConfig';
import {
  AssistantAction,
  Contact,
  PlannerBucket,
  PlannerPlan,
  TodoList,
} from './types';

export interface M365AccountProfile {
  id?: string;
  name: string;
  username: string;
}

export interface DeviceCodeSession {
  userCode: string;
  deviceCode: string;
  verificationUri: string;
  expiresIn: number;
  interval: number;
  message: string;
}

class GraphService {
  private currentAccount: M365AccountProfile | null = null;
  private cachedAccessToken: string | null = null;
  private tokenExpiresAt = 0;

  public async getAccount(): Promise<M365AccountProfile | null> {
    if (
      this.currentAccount &&
      this.cachedAccessToken &&
      this.tokenExpiresAt > Date.now() + 60000
    ) {
      return this.currentAccount;
    }

    try {
      const res = await fetch('/api/m365/auth', { cache: 'no-store' });
      if (!res.ok) return null;
      const data = await res.json();
      if (data.connected && data.accessToken) {
        this.cachedAccessToken = data.accessToken;
        this.tokenExpiresAt = data.expiresAt || Date.now() + 3000 * 1000;
        this.currentAccount = {
          id: data.user?.id,
          name: data.user?.name || 'Microsoft 365 User',
          username: data.user?.email || '',
        };
        return this.currentAccount;
      }
    } catch {
      // ignore
    }
    return null;
  }

  public async startDeviceCodeLogin(): Promise<DeviceCodeSession> {
    const res = await fetch('/api/m365/auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'start_device_code',
        clientId: getStoredClientId(),
        tenantId: getStoredTenantId(),
      }),
    });

    const data = await res.json();
    if (!res.ok || data.error) {
      throw new Error(data.error || 'Could not start Microsoft 365 sign-in');
    }
    return data as DeviceCodeSession;
  }

  public async pollDeviceCodeLogin(deviceCode: string): Promise<{
    status: 'pending' | 'connected' | 'error';
    account?: M365AccountProfile;
    error?: string;
  }> {
    const res = await fetch('/api/m365/auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'poll_device_code',
        deviceCode,
        clientId: getStoredClientId(),
        tenantId: getStoredTenantId(),
      }),
    });

    const data = await res.json();
    if (data.status === 'pending') {
      return { status: 'pending' };
    }
    if (!res.ok || data.status === 'error') {
      return {
        status: 'error',
        error: data.error || 'Sign-in expired or failed',
      };
    }

    this.cachedAccessToken = data.accessToken;
    this.tokenExpiresAt = data.expiresAt || Date.now() + 3500 * 1000;
    this.currentAccount = {
      id: data.user?.id,
      name: data.user?.name || 'Microsoft 365 User',
      username: data.user?.email || '',
    };

    return {
      status: 'connected',
      account: this.currentAccount,
    };
  }

  public async logout(): Promise<void> {
    this.currentAccount = null;
    this.cachedAccessToken = null;
    this.tokenExpiresAt = 0;
    await fetch('/api/m365/auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'logout' }),
    });
  }

  public async getAccessToken(): Promise<string> {
    const now = Date.now() + 60 * 1000;
    if (this.cachedAccessToken && this.tokenExpiresAt > now) {
      return this.cachedAccessToken;
    }

    const res = await fetch('/api/m365/auth', { cache: 'no-store' });
    const data = await res.json();
    if (!res.ok || !data.connected || !data.accessToken) {
      throw new Error(
        'Not connected to Microsoft 365 yet. Click "Connect Microsoft 365" in the top right header.'
      );
    }

    this.cachedAccessToken = data.accessToken;
    this.tokenExpiresAt = data.expiresAt || Date.now() + 3500 * 1000;
    this.currentAccount = {
      id: data.user?.id,
      name: data.user?.name || 'Microsoft 365 User',
      username: data.user?.email || '',
    };

    return data.accessToken;
  }

  public async getGraphClient(): Promise<Client> {
    const token = await this.getAccessToken();
    return Client.init({
      authProvider: (done) => {
        done(null, token);
      },
    });
  }

  // ==========================================================
  // 1. MICROSOFT TO DO / OUTLOOK TASKS
  // ==========================================================
  public async getTodoLists(): Promise<TodoList[]> {
    const client = await this.getGraphClient();
    const res = await client.api('/me/todo/lists').get();
    return (res.value || []).map((l: any) => ({
      id: l.id,
      displayName: l.displayName,
      wellknownListName: l.wellknownListName,
    }));
  }

  public async createOutlookTodoTask(
    action: AssistantAction,
    defaultListId?: string
  ): Promise<{ id: string; webUrl: string }> {
    const client = await this.getGraphClient();

    let targetListId = defaultListId;
    if (!targetListId) {
      const listsRes = await client
        .api('/me/todo/lists')
        .filter("wellknownListName eq 'defaultList'")
        .get();
      targetListId =
        listsRes.value?.[0]?.id ||
        (await client.api('/me/todo/lists').top(1).get()).value?.[0]?.id;
    }

    if (!targetListId) {
      throw new Error('Could not find a Microsoft To Do / Outlook Task list.');
    }

    const payload: any = {
      title: action.title,
      importance:
        action.priority === 'urgent' || action.priority === 'high'
          ? 'high'
          : action.priority === 'low'
          ? 'low'
          : 'normal',
      body: {
        contentType: 'text',
        content: action.description || `Created via Personal Voice Assistant`,
      },
    };

    if (action.due_date) {
      payload.dueDateTime = {
        dateTime: `${action.due_date.split('T')[0]}T17:00:00`,
        timeZone: 'UTC',
      };
    }

    const created = await client
      .api(`/me/todo/lists/${targetListId}/tasks`)
      .post(payload);

    return {
      id: created.id,
      webUrl: `https://to-do.office.com/tasks/id/${encodeURIComponent(created.id)}/details`,
    };
  }

  public async getLiveOutlookTasks(): Promise<any[]> {
    const client = await this.getGraphClient();
    const listsRes = await client.api('/me/todo/lists').top(10).get();
    const lists = listsRes.value || [];
    const results: any[] = [];

    await Promise.all(
      lists.slice(0, 5).map(async (list: any) => {
        try {
          const tRes = await client
            .api(`/me/todo/lists/${list.id}/tasks`)
            .filter("status ne 'completed'")
            .top(25)
            .get();
          for (const t of tRes.value || []) {
            results.push({
              ...t,
              listId: list.id,
              listName: list.displayName,
            });
          }
        } catch {
          // ignore list error
        }
      })
    );
    return results;
  }

  public async completeOutlookTask(listId: string, taskId: string): Promise<void> {
    const client = await this.getGraphClient();
    await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).patch({
      status: 'completed',
    });
  }

  // ==========================================================
  // 2. MICROSOFT PLANNER TASKS, PLANS & BUCKETS
  // ==========================================================
  public async getPlannerPlans(): Promise<PlannerPlan[]> {
    const client = await this.getGraphClient();
    const plansMap = new Map<string, PlannerPlan>();

    // Strategy A: Direct /me/planner/plans
    try {
      const directRes = await client.api('/me/planner/plans').get();
      for (const p of directRes.value || []) {
        plansMap.set(p.id, { id: p.id, title: p.title, owner: p.owner });
      }
    } catch {
      // Fallback to Strategy B
    }

    // Strategy B: Discover plans from user's existing Planner tasks (/me/planner/tasks only requires Tasks.ReadWrite!)
    try {
      const tasksRes = await client.api('/me/planner/tasks').top(50).get();
      const uniquePlanIds = Array.from(
        new Set(
          (tasksRes.value || [])
            .map((t: any) => t.planId)
            .filter(Boolean) as string[]
        )
      );

      await Promise.all(
        uniquePlanIds.map(async (planId) => {
          if (plansMap.has(planId)) return;
          try {
            const planDetail = await client.api(`/planner/plans/${planId}`).get();
            if (planDetail?.id) {
              plansMap.set(planDetail.id, {
                id: planDetail.id,
                title: planDetail.title || 'Team Plan',
                owner: planDetail.owner,
              });
            }
          } catch {
            plansMap.set(planId, {
              id: planId,
              title: `Planner Plan (${planId.slice(0, 6)})`,
            });
          }
        })
      );
    } catch (err) {
      console.warn('Planner task plan discovery notice:', err);
    }

    return Array.from(plansMap.values());
  }

  public async getPlannerBuckets(planId: string): Promise<PlannerBucket[]> {
    if (!planId) return [];
    const client = await this.getGraphClient();
    try {
      const res = await client.api(`/planner/plans/${planId}/buckets`).get();
      return (res.value || []).map((b: any) => ({
        id: b.id,
        name: b.name,
        planId: b.planId,
      }));
    } catch (err) {
      console.warn('Could not load buckets for plan:', err);
      return [];
    }
  }

  public async createPlannerTask(
    action: AssistantAction,
    fallbackPlanId?: string,
    fallbackBucketId?: string
  ): Promise<{ id: string; webUrl: string }> {
    const client = await this.getGraphClient();
    let planId = action.plan_id || fallbackPlanId;
    let bucketId = action.bucket_id || fallbackBucketId;

    if (!planId) {
      const plans = await this.getPlannerPlans();
      if (plans.length > 0) {
        planId = plans[0].id;
      }
    }

    if (!planId) {
      throw new Error(
        'No Microsoft Planner Plan selected. Please choose a Plan on the action card or in Settings.'
      );
    }

    if (!bucketId) {
      const buckets = await this.getPlannerBuckets(planId);
      if (buckets.length > 0) {
        bucketId = buckets[0].id;
      }
    }

    const priorityMap: Record<string, number> = {
      urgent: 1,
      high: 3,
      normal: 5,
      low: 9,
    };

    const payload: any = {
      planId,
      title: action.title,
      priority: priorityMap[action.priority] ?? 5,
    };

    if (bucketId) {
      payload.bucketId = bucketId;
    }

    if (action.due_date) {
      payload.dueDateTime = `${action.due_date.split('T')[0]}T17:00:00Z`;
    }

    try {
      const me = await client.api('/me').select('id').get();
      if (me?.id) {
        payload.assignments = {
          [me.id]: {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !',
          },
        };
      }
    } catch {
      // ignore
    }

    const created = await client.api('/planner/tasks').post(payload);

    if (action.description && created?.id) {
      try {
        const details = await client
          .api(`/planner/tasks/${created.id}/details`)
          .get();
        const etag = details['@odata.etag'];
        if (etag) {
          await client
            .api(`/planner/tasks/${created.id}/details`)
            .header('If-Match', etag)
            .patch({
              description: action.description,
              previewType: 'description',
            });
        }
      } catch (detailErr) {
        console.warn('Could not attach Planner task description:', detailErr);
      }
    }

    const webUrl = `https://tasks.office.com/protemp.co.za/Home/Task/?taskId=${encodeURIComponent(
      created.id
    )}&planId=${encodeURIComponent(planId)}`;

    return {
      id: created.id,
      webUrl,
    };
  }

  public async getLivePlannerTasks(): Promise<any[]> {
    const client = await this.getGraphClient();
    const res = await client.api('/me/planner/tasks').top(40).get();
    return (res.value || []).filter((t: any) => t.percentComplete !== 100);
  }

  public async completePlannerTask(taskId: string): Promise<void> {
    const client = await this.getGraphClient();
    const current = await client.api(`/planner/tasks/${taskId}`).get();
    const etag = current['@odata.etag'];
    await client
      .api(`/planner/tasks/${taskId}`)
      .header('If-Match', etag)
      .patch({ percentComplete: 100 });
  }

  // ==========================================================
  // 3. OUTLOOK EMAIL (GRAPH API + OUTLOOK WEB FALLBACK)
  // ==========================================================
  public async sendOrDraftOutlookEmail(
    action: AssistantAction,
    mode: 'send_immediately' | 'create_draft' = 'send_immediately'
  ): Promise<{ id: string; webUrl: string; method: string }> {
    const recipientEmail = action.recipient_email?.trim();
    if (!recipientEmail && mode === 'send_immediately') {
      throw new Error(
        `Recipient email is missing for "${
          action.recipient_name || action.title
        }". Please add their email on the card or in Contacts.`
      );
    }

    try {
      const client = await this.getGraphClient();

      const messagePayload: any = {
        subject: action.title,
        body: {
          contentType: 'HTML',
          content: `<p>${(action.description || action.title).replace(
            /\n/g,
            '<br/>'
          )}</p>`,
        },
      };

      if (recipientEmail) {
        messagePayload.toRecipients = [
          {
            emailAddress: {
              address: recipientEmail,
              name: action.recipient_name || recipientEmail,
            },
          },
        ];
      }

      if (mode === 'send_immediately') {
        await client.api('/me/sendMail').post({
          message: messagePayload,
          saveToSentItems: true,
        });
        return {
          id: `mail-${Date.now()}`,
          webUrl: 'https://outlook.office.com/mail/sentitems',
          method: 'graph_send',
        };
      } else {
        const draft = await client.api('/me/messages').post(messagePayload);
        return {
          id: draft.id,
          webUrl: draft.webLink || 'https://outlook.office.com/mail/drafts',
          method: 'graph_draft',
        };
      }
    } catch {
      const composeUrl = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(
        recipientEmail || ''
      )}&subject=${encodeURIComponent(action.title)}&body=${encodeURIComponent(
        action.description || action.title
      )}`;

      if (typeof window !== 'undefined') {
        window.open(composeUrl, '_blank', 'noopener,noreferrer');
      }

      return {
        id: `outlook-web-${Date.now()}`,
        webUrl: composeUrl,
        method: 'outlook_web_compose',
      };
    }
  }

  // ==========================================================
  // 4. OUTLOOK CALENDAR EVENTS
  // ==========================================================
  public async createOutlookCalendarEvent(
    action: AssistantAction
  ): Promise<{ id: string; webUrl: string }> {
    const client = await this.getGraphClient();

    const baseDate = action.due_date
      ? action.due_date.split('T')[0]
      : new Date().toISOString().split('T')[0];

    const startTimeStr = action.start_time || '09:00';
    const [startH, startM] = startTimeStr.split(':').map(Number);
    const endH = String(Math.min((startH || 9) + 1, 23)).padStart(2, '0');
    const endTimeStr =
      action.end_time || `${endH}:${String(startM || 0).padStart(2, '0')}`;

    const startLocal = new Date(`${baseDate}T${startTimeStr}:00`);
    const endLocal = new Date(`${baseDate}T${endTimeStr}:00`);

    const payload: any = {
      subject: action.title,
      body: {
        contentType: 'HTML',
        content: `<p>${
          action.description || 'Scheduled via Personal Voice Assistant'
        }</p>`,
      },
      start: {
        dateTime: startLocal.toISOString().replace(/Z$/, ''),
        timeZone: 'UTC',
      },
      end: {
        dateTime: endLocal.toISOString().replace(/Z$/, ''),
        timeZone: 'UTC',
      },
    };

    if (action.recipient_email) {
      payload.attendees = [
        {
          emailAddress: {
            address: action.recipient_email,
            name: action.recipient_name || action.recipient_email,
          },
          type: 'required',
        },
      ];
    }

    const created = await client.api('/me/events').post(payload);
    return {
      id: created.id,
      webUrl: created.webLink || 'https://outlook.office.com/calendar',
    };
  }

  // ==========================================================
  // 5. IMPORT CONTACTS FROM MICROSOFT 365
  // ==========================================================
  public async fetchMicrosoftPeople(): Promise<Partial<Contact>[]> {
    const client = await this.getGraphClient();
    const contacts: Partial<Contact>[] = [];

    try {
      const res = await client
        .api('/me/people')
        .top(50)
        .select('displayName,scoredEmailAddresses,phones,jobTitle,companyName')
        .get();

      for (const p of res.value || []) {
        const email = p.scoredEmailAddresses?.[0]?.address;
        const phone = p.phones?.[0]?.number;
        if (p.displayName && (email || phone)) {
          contacts.push({
            name: p.displayName,
            email: email || '',
            phone: phone || '',
            role: p.jobTitle || '',
            company: p.companyName || 'Protemp',
          });
        }
      }
    } catch {
      // ignore
    }

    return contacts;
  }
}

export const graphService = new GraphService();
