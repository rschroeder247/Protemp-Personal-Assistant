import { PublicClientApplication } from '@azure/msal-browser';
import type { AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  createMsalConfig,
  getStoredClientId,
  getStoredTenantId,
  GRAPH_SCOPES,
} from './msalConfig';
import {
  AssistantAction,
  Contact,
  PlannerBucket,
  PlannerPlan,
  TodoList,
} from './types';

const TOKEN_BACKUP_KEY = 'pa_ms_token_backup';

class GraphService {
  private currentAccount: AccountInfo | null = null;
  private cachedAccessToken: string | null = null;
  private tokenExpiresAt = 0;
  private tokenPromise: Promise<string> | null = null;
  private initPromise: Promise<PublicClientApplication | null> | null = null;
  public lastAuthError: string | null = null;

  private saveTokenResult(result: AuthenticationResult | null): void {
    if (!result?.accessToken) return;
    this.cachedAccessToken = result.accessToken;
    const expiresMs = result.expiresOn
      ? result.expiresOn.getTime()
      : Date.now() + 3500 * 1000;
    this.tokenExpiresAt = expiresMs;
    try {
      sessionStorage.setItem(
        TOKEN_BACKUP_KEY,
        JSON.stringify({
          accessToken: result.accessToken,
          expiresAt: expiresMs,
          username: result.account?.username || '',
        })
      );
    } catch {
      // ignore
    }
  }

  private getValidCachedToken(username?: string): string | null {
    const now = Date.now() + 60 * 1000;
    if (this.cachedAccessToken && this.tokenExpiresAt > now) {
      return this.cachedAccessToken;
    }
    if (typeof window === 'undefined') return null;
    try {
      const raw = sessionStorage.getItem(TOKEN_BACKUP_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (
          parsed.accessToken &&
          parsed.expiresAt > now &&
          (!username ||
            !parsed.username ||
            parsed.username.toLowerCase() === username.toLowerCase())
        ) {
          this.cachedAccessToken = parsed.accessToken;
          this.tokenExpiresAt = parsed.expiresAt;
          return parsed.accessToken;
        }
      }
    } catch {
      // ignore
    }
    return null;
  }

  public cleanupStaleState(): void {
    if (typeof window === 'undefined') return;
    try {
      const hash = window.location.hash;
      const search = window.location.search;
      if (
        hash.includes('code=') ||
        hash.includes('state=') ||
        hash.includes('error=') ||
        search.includes('code=') ||
        search.includes('state=') ||
        search.includes('error=')
      ) {
        window.history.replaceState({}, document.title, window.location.pathname);
      }

      for (let i = sessionStorage.length - 1; i >= 0; i--) {
        const key = sessionStorage.key(i);
        if (key && key.includes('interaction.status')) {
          sessionStorage.removeItem(key);
        }
      }
    } catch {
      // ignore
    }
  }

  public getMsal(): Promise<PublicClientApplication | null> {
    if (typeof window === 'undefined') return Promise.resolve(null);
    const clientId = getStoredClientId();
    if (!clientId) return Promise.resolve(null);

    if (!this.initPromise) {
      this.initPromise = (async () => {
        const config = createMsalConfig(clientId, getStoredTenantId());
        const instance = new PublicClientApplication(config);
        await instance.initialize();

        try {
          const redirectResponse = await instance.handleRedirectPromise();
          if (redirectResponse) {
            this.saveTokenResult(redirectResponse);
            if (redirectResponse.account) {
              instance.setActiveAccount(redirectResponse.account);
              this.currentAccount = redirectResponse.account;
            }
          }
        } catch (redirectErr: any) {
          const isCacheMiss =
            redirectErr?.errorCode === 'no_token_request_cache_error' ||
            redirectErr?.message?.includes('no_token_request_cache_error');
          if (!isCacheMiss) {
            console.warn('MSAL redirect notice:', redirectErr);
            this.lastAuthError = redirectErr?.message || String(redirectErr);
          }
        } finally {
          this.cleanupStaleState();
        }

        const accounts = instance.getAllAccounts();
        if (accounts.length > 0 && !this.currentAccount) {
          this.currentAccount = accounts[0];
          instance.setActiveAccount(accounts[0]);
        }

        return instance;
      })();
    }
    return this.initPromise;
  }

  public async resetMsal(): Promise<void> {
    this.cleanupStaleState();
    this.initPromise = null;
    this.currentAccount = null;
    this.cachedAccessToken = null;
    this.tokenExpiresAt = 0;
    try {
      sessionStorage.removeItem(TOKEN_BACKUP_KEY);
    } catch {
      // ignore
    }
  }

  public async getAccount(): Promise<AccountInfo | null> {
    const msal = await this.getMsal();
    if (!msal) return null;
    if (this.currentAccount) return this.currentAccount;
    const accounts = msal.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  public async login(): Promise<void> {
    this.cleanupStaleState();
    const msal = await this.getMsal();
    if (!msal) {
      throw new Error('Azure Client ID is not configured.');
    }
    await msal.loginRedirect({
      scopes: GRAPH_SCOPES,
      prompt: 'select_account',
    });
  }

  public async logout(): Promise<void> {
    const msal = await this.getMsal();
    if (!msal) return;
    const account = await this.getAccount();
    this.currentAccount = null;
    this.cachedAccessToken = null;
    this.tokenExpiresAt = 0;
    try {
      sessionStorage.removeItem(TOKEN_BACKUP_KEY);
    } catch {
      // ignore
    }
    this.cleanupStaleState();
    if (account) {
      await msal.logoutRedirect({ account });
    }
  }

  public async getAccessToken(extraScopes: string[] = []): Promise<string> {
    const msal = await this.getMsal();
    if (!msal) throw new Error('MSAL not initialized');

    const account = this.currentAccount || (await this.getAccount());
    if (!account) {
      throw new Error('Not signed in to Microsoft 365. Click "Connect Microsoft 365" in the header.');
    }

    if (extraScopes.length === 0) {
      const validToken = this.getValidCachedToken(account.username);
      if (validToken) return validToken;
    }

    const scopesToRequest = Array.from(new Set([...GRAPH_SCOPES, ...extraScopes]));

    if (!this.tokenPromise) {
      this.tokenPromise = (async () => {
        try {
          const response = await msal.acquireTokenSilent({
            scopes: scopesToRequest,
            account,
          });
          this.saveTokenResult(response);
          return response.accessToken;
        } catch (silentErr: any) {
          // If extraScopes caused consent failure, fallback to base GRAPH_SCOPES
          if (extraScopes.length > 0) {
            const fallbackResponse = await msal.acquireTokenSilent({
              scopes: GRAPH_SCOPES,
              account,
            });
            this.saveTokenResult(fallbackResponse);
            return fallbackResponse.accessToken;
          }
          throw silentErr;
        } finally {
          this.tokenPromise = null;
        }
      })();
    }

    return this.tokenPromise;
  }

  public async getGraphClient(extraScopes: string[] = []): Promise<Client> {
    const token = await this.getAccessToken(extraScopes);
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
      // May require Group.Read.All in some tenants; fallback to Strategy B
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

    // Auto-discover first available plan if none specified
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

    // Auto-discover first bucket in the plan if none specified
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

    // Assign to current signed-in user by default so it shows up in "My Tasks"
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
      // ignore assignment if /me fails
    }

    const created = await client.api('/planner/tasks').post(payload);

    // If description/notes exist, attach them to /planner/tasks/{id}/details
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
        `Recipient email is missing for "${action.recipient_name || action.title}". Please add their email on the card or in Contacts.`
      );
    }

    try {
      const client = await this.getGraphClient(
        mode === 'send_immediately' ? ['Mail.Send'] : ['Mail.ReadWrite']
      );

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
    } catch (err: any) {
      // If Mail.Send isn't consented on the Azure App Registration yet, open Outlook Web Compose pre-filled!
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
    const endTimeStr = action.end_time || `${endH}:${String(startM || 0).padStart(2, '0')}`;

    const startLocal = new Date(`${baseDate}T${startTimeStr}:00`);
    const endLocal = new Date(`${baseDate}T${endTimeStr}:00`);

    const payload: any = {
      subject: action.title,
      body: {
        contentType: 'HTML',
        content: `<p>${action.description || 'Scheduled via Personal Voice Assistant'}</p>`,
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
      // ignore if People.Read is not consented
    }

    return contacts;
  }
}

export const graphService = new GraphService();
