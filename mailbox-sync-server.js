import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import axios from 'axios';
import crypto from 'crypto';
import tls from 'tls';
import { createClient } from '@supabase/supabase-js';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import EmailService, {
  computeThreadKey,
  extractAttachmentMetadata,
  normalizeAddressList,
  normalizeMessageId,
  normalizeReferences,
} from './services/email-service.js';

const DEFAULT_SUPABASE_URL = '';
const DEFAULT_SUPABASE_ANON_KEY = '';
const DEFAULT_SERVICE_ROLE_KEY = '';
const DEFAULT_ALLOWED_ORIGINS = 'http://localhost:5173,http://localhost:8080,http://10.127.57.196:8080';
const DEFAULT_MAILBOX_PORT = 8787;
const DEFAULT_BYPASS_AUTH = 'true';
const DEFAULT_CRM_REDIRECT_URI = 'http://localhost:8080/dashboard?tab=integrations';
const DEFAULT_CRM_FRONTEND_URL = 'http://localhost:8080';
const DEFAULT_REPLY_INCLUDE_ORIGINAL_ATTACHMENTS = 'false';
const DEFAULT_REPLY_BLOCK_NOREPLY = 'false';
const DEFAULT_REPLY_NOREPLY_REGEX = '(^|[+._-])no-?reply|do-?not-?reply';
const DEFAULT_REPLY_REFERENCES_LIMIT = '20';
const DEFAULT_REPLY_REFERENCES_CHAR_LIMIT = '1900';
const DEFAULT_CHECK_LOOKBACK_DAYS = '7';
const DEFAULT_MAILBOX_ADMIN_SECRET = '';
const DEFAULT_AUTO_CHECK_REPLIES = 'false';
const DEFAULT_AUTO_CHECK_INTERVAL_MINUTES = '10';
const DEFAULT_AUTO_CHECK_LOOKBACK_DAYS = DEFAULT_CHECK_LOOKBACK_DAYS;
const DEFAULT_AUTO_CHECK_USE_DB_SCAN = 'false';
const DEFAULT_AUTO_CHECK_CONFIG_ID = '';
const DEFAULT_AUTO_SYNC_MAILBOXES = 'false';
const DEFAULT_AUTO_SYNC_INTERVAL_MINUTES = '10';
const DEFAULT_AUTO_SYNC_LIMIT = '50';
const DEFAULT_AUTO_SYNC_CONFIG_ID = '';
const DEFAULT_MAILBOX_OAUTH_SIMULATE = 'false';
const DEFAULT_MAILBOX_OAUTH_REDIRECT_URI = 'http://localhost:8080/dashboard?tab=settings';
const DEFAULT_GOOGLE_OAUTH_SCOPES = 'openid email profile https://mail.google.com/';
const DEFAULT_MICROSOFT_OAUTH_SCOPES =
  'openid profile email offline_access https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send https://graph.microsoft.com/User.Read';

const PORT = Number(process.env.PORT || process.env.MAILBOX_SERVER_PORT || DEFAULT_MAILBOX_PORT);
const normalizeOrigin = (origin) => String(origin || '').trim().replace(/\/+$/, '');
const ALLOWED_ORIGINS = (process.env.MAILBOX_ALLOWED_ORIGINS || DEFAULT_ALLOWED_ORIGINS)
  .split(',')
  .map((origin) => normalizeOrigin(origin))
  .filter(Boolean);

const SUPABASE_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || DEFAULT_SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || DEFAULT_SERVICE_ROLE_KEY;
const SUPABASE_ANON_KEY =
  process.env.SUPABASE_ANON_KEY ||
  process.env.VITE_SUPABASE_ANON_KEY ||
  process.env.SUPABASE_PUBLISHABLE_KEY ||
  DEFAULT_SUPABASE_ANON_KEY;
const MAILBOX_BYPASS_AUTH = (process.env.MAILBOX_BYPASS_AUTH || DEFAULT_BYPASS_AUTH).toLowerCase() === 'true';
const CRM_OAUTH_SIMULATE = (process.env.CRM_OAUTH_SIMULATE || 'true').toLowerCase() === 'true';
const CRM_REDIRECT_URI = process.env.CRM_OAUTH_REDIRECT_URI || DEFAULT_CRM_REDIRECT_URI;
const CRM_FRONTEND_URL = process.env.CRM_FRONTEND_URL || DEFAULT_CRM_FRONTEND_URL;
const HUBSPOT_CLIENT_ID = process.env.HUBSPOT_CLIENT_ID || '';
const HUBSPOT_CLIENT_SECRET = process.env.HUBSPOT_CLIENT_SECRET || '';
const SALESFORCE_CLIENT_ID = process.env.SALESFORCE_CLIENT_ID || '';
const SALESFORCE_CLIENT_SECRET = process.env.SALESFORCE_CLIENT_SECRET || '';
const SALESFORCE_LOGIN_URL = process.env.SALESFORCE_LOGIN_URL || 'https://login.salesforce.com';
const HUBSPOT_SCOPES =
  process.env.HUBSPOT_SCOPES ||
  'crm.objects.contacts.read crm.objects.contacts.write crm.objects.companies.read crm.objects.companies.write oauth';
const SALESFORCE_SCOPES =
  process.env.SALESFORCE_SCOPES ||
  'api refresh_token offline_access';
const MAILBOX_OAUTH_SIMULATE =
  (process.env.MAILBOX_OAUTH_SIMULATE || DEFAULT_MAILBOX_OAUTH_SIMULATE).toLowerCase() === 'true';
const MAILBOX_OAUTH_REDIRECT_URI =
  process.env.MAILBOX_OAUTH_REDIRECT_URI ||
  process.env.GOOGLE_OAUTH_REDIRECT_URI ||
  process.env.MICROSOFT_OAUTH_REDIRECT_URI ||
  DEFAULT_MAILBOX_OAUTH_REDIRECT_URI;
const GOOGLE_OAUTH_CLIENT_ID = process.env.GOOGLE_OAUTH_CLIENT_ID || '';
const GOOGLE_OAUTH_CLIENT_SECRET = process.env.GOOGLE_OAUTH_CLIENT_SECRET || '';
const GOOGLE_OAUTH_SCOPES = process.env.GOOGLE_OAUTH_SCOPES || DEFAULT_GOOGLE_OAUTH_SCOPES;
const MICROSOFT_OAUTH_CLIENT_ID = process.env.MICROSOFT_OAUTH_CLIENT_ID || '';
const MICROSOFT_OAUTH_CLIENT_SECRET = process.env.MICROSOFT_OAUTH_CLIENT_SECRET || '';
const MICROSOFT_OAUTH_TENANT_ID = process.env.MICROSOFT_OAUTH_TENANT_ID || 'common';
const MICROSOFT_OAUTH_SCOPES = process.env.MICROSOFT_OAUTH_SCOPES || DEFAULT_MICROSOFT_OAUTH_SCOPES;
const REPLY_INCLUDE_ORIGINAL_ATTACHMENTS =
  (process.env.MAILBOX_REPLY_INCLUDE_ORIGINAL_ATTACHMENTS || DEFAULT_REPLY_INCLUDE_ORIGINAL_ATTACHMENTS).toLowerCase() ===
  'true';
const REPLY_BLOCK_NOREPLY =
  (process.env.MAILBOX_REPLY_BLOCK_NOREPLY || DEFAULT_REPLY_BLOCK_NOREPLY).toLowerCase() === 'true';
const REPLY_REFERENCES_LIMIT = Number(process.env.MAILBOX_REPLY_REFERENCES_LIMIT || DEFAULT_REPLY_REFERENCES_LIMIT);
const REPLY_REFERENCES_CHAR_LIMIT = Number(
  process.env.MAILBOX_REPLY_REFERENCES_CHAR_LIMIT || DEFAULT_REPLY_REFERENCES_CHAR_LIMIT
);
const MAILBOX_ADMIN_SECRET = process.env.MAILBOX_ADMIN_SECRET || DEFAULT_MAILBOX_ADMIN_SECRET;
let replyNoreplyRegex = null;
try {
  replyNoreplyRegex = new RegExp(process.env.MAILBOX_REPLY_NOREPLY_REGEX || DEFAULT_REPLY_NOREPLY_REGEX, 'i');
} catch (error) {
  replyNoreplyRegex = new RegExp(DEFAULT_REPLY_NOREPLY_REGEX, 'i');
}

const isPlaceholderValue = (value) => {
  if (!value) return true;
  const normalized = value.toLowerCase();
  return (
    normalized.includes('redacted') ||
    normalized.includes('your-service-role-key') ||
    normalized.includes('your-anon-key')
  );
};

const hasSupabaseUrl = typeof SUPABASE_URL === 'string' && SUPABASE_URL.startsWith('http');
const hasServiceRoleKey = hasSupabaseUrl && !isPlaceholderValue(SUPABASE_SERVICE_ROLE_KEY);
const hasAnonKey = hasSupabaseUrl && !isPlaceholderValue(SUPABASE_ANON_KEY);
const supabaseAdmin = hasServiceRoleKey ? createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY) : null;

if (!hasSupabaseUrl) {
  console.warn('[mailbox] SUPABASE_URL is missing. Set it in .env.');
}
if (!hasServiceRoleKey) {
  console.warn('[mailbox] SUPABASE_SERVICE_ROLE_KEY is missing or placeholder. Falling back to anon key.');
}
if (!hasAnonKey) {
  console.warn('[mailbox] SUPABASE_ANON_KEY is missing or placeholder. Mailbox sync may fail.');
}
if (MAILBOX_BYPASS_AUTH && !hasServiceRoleKey) {
  console.warn('[mailbox] MAILBOX_BYPASS_AUTH requires a service role key. Bypass auth disabled.');
}

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(
  cors({
    origin(origin, callback) {
      const normalizedOrigin = normalizeOrigin(origin);
      if (!origin || ALLOWED_ORIGINS.includes('*') || ALLOWED_ORIGINS.includes(normalizedOrigin)) {
        callback(null, true);
      } else {
        callback(new Error('Not allowed by CORS'));
      }
    },
    credentials: true,
  })
);

const emailService = new EmailService({
  referencesLimit: Number.isFinite(REPLY_REFERENCES_LIMIT) ? REPLY_REFERENCES_LIMIT : undefined,
  referencesCharLimit: Number.isFinite(REPLY_REFERENCES_CHAR_LIMIT) ? REPLY_REFERENCES_CHAR_LIMIT : undefined,
  blockNoreply: REPLY_BLOCK_NOREPLY,
  noreplyRegex: replyNoreplyRegex,
  includeOriginalAttachments: REPLY_INCLUDE_ORIGINAL_ATTACHMENTS,
});

const parsedOutboundEmailCreditCost = Number.parseInt(process.env.OUTBOUND_EMAIL_CREDIT_COST || '1', 10);
const OUTBOUND_EMAIL_CREDIT_COST = Number.isFinite(parsedOutboundEmailCreditCost) && parsedOutboundEmailCreditCost > 0
  ? parsedOutboundEmailCreditCost
  : 1;

const consumeUserCredits = async (dbClient, userId, amount, eventType, referenceId, metadata = {}) => {
  const { data, error } = await dbClient.rpc('consume_user_credits', {
    p_amount: amount,
    p_event_type: eventType,
    p_reference_id: referenceId,
    p_metadata: metadata,
    p_user_id: userId,
  });

  if (error) {
    throw new Error(error?.message || 'Failed to consume credits');
  }

  const row = Array.isArray(data) ? data[0] : data;
  return {
    allowed: Boolean(row?.allowed),
    creditsRemaining: Number(row?.credits_remaining ?? 0),
    message: String(row?.message || ''),
  };
};

const refundUserCredits = async (dbClient, userId, amount, eventType, referenceId, metadata = {}) => {
  const { data, error } = await dbClient.rpc('refund_user_credits', {
    p_amount: amount,
    p_event_type: eventType,
    p_reference_id: referenceId,
    p_metadata: metadata,
    p_user_id: userId,
  });

  if (error) {
    console.error('[billing] Credit refund failed:', error?.message || error);
    return null;
  }

  const row = Array.isArray(data) ? data[0] : data;
  return Number(row?.credits_remaining ?? NaN);
};

const MAILBOX_OAUTH_REFRESH_SKEW_MS = 2 * 60 * 1000;

const resolveMailboxOAuthProvider = (value) => {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return null;
  if (normalized === 'gmail' || normalized === 'google') return 'gmail';
  if (['outlook', 'microsoft', 'office365', 'microsoft365'].includes(normalized)) return 'outlook';
  return null;
};

const isOAuthMailboxConfig = (config) => {
  const authType = String(config?.auth_type || '').trim().toLowerCase();
  const provider = resolveMailboxOAuthProvider(config?.oauth_provider);
  return authType === 'oauth' || Boolean(provider);
};

const hasMailboxTransportCredentials = (config) => {
  if (!config?.smtp_host || !config?.smtp_username) return false;
  if (!isOAuthMailboxConfig(config)) {
    return Boolean(String(config?.smtp_password || '').trim());
  }
  return Boolean(String(config?.oauth_access_token || '').trim() || String(config?.oauth_refresh_token || '').trim());
};

const resolveMailboxOAuthRedirectUri = (value) => {
  const candidate = String(value || '').trim();
  return candidate || MAILBOX_OAUTH_REDIRECT_URI;
};

const resolveMailboxPresetForProvider = (provider) => {
  if (provider === 'gmail') {
    return {
      smtpHost: 'smtp.gmail.com',
      smtpPort: 465,
      imapHost: 'imap.gmail.com',
      imapPort: 993,
      security: 'SSL',
    };
  }

  if (provider === 'outlook') {
    return {
      smtpHost: 'smtp.office365.com',
      smtpPort: 587,
      imapHost: 'outlook.office365.com',
      imapPort: 993,
      security: 'TLS',
    };
  }

  return null;
};

const buildImapAuthFromConfig = (config, loginMethod) => {
  if (!isOAuthMailboxConfig(config)) {
    const auth = {
      user: config.smtp_username,
      pass: config.smtp_password,
    };
    if (loginMethod) {
      auth.loginMethod = loginMethod;
    }
    return auth;
  }

  const accessToken = String(config?.oauth_access_token || '').trim();
  if (!accessToken) {
    throw new Error('OAuth access token is missing for this mailbox.');
  }

  return {
    user: config.smtp_username,
    accessToken,
  };
};

const needsMailboxOAuthRefresh = (config) => {
  if (!isOAuthMailboxConfig(config)) return false;

  const refreshToken = String(config?.oauth_refresh_token || '').trim();
  const accessToken = String(config?.oauth_access_token || '').trim();
  if (!accessToken && refreshToken) return true;
  if (!refreshToken) return false;

  const expiresAtRaw = String(config?.oauth_token_expires_at || '').trim();
  if (!expiresAtRaw) return false;

  const expiresAtMs = Date.parse(expiresAtRaw);
  if (!Number.isFinite(expiresAtMs)) return false;
  return expiresAtMs - MAILBOX_OAUTH_REFRESH_SKEW_MS <= Date.now();
};

const refreshMailboxOAuthToken = async (config, dbClient) => {
  const provider = resolveMailboxOAuthProvider(config?.oauth_provider);
  if (!provider) {
    throw new Error('Unsupported mailbox OAuth provider.');
  }

  const refreshToken = String(config?.oauth_refresh_token || '').trim();
  if (!refreshToken) {
    throw new Error('Refresh token is missing for this mailbox.');
  }

  let response;

  if (provider === 'gmail') {
    if (!GOOGLE_OAUTH_CLIENT_ID || !GOOGLE_OAUTH_CLIENT_SECRET) {
      throw new Error('Google OAuth credentials are not configured.');
    }
    const payload = new URLSearchParams({
      client_id: GOOGLE_OAUTH_CLIENT_ID,
      client_secret: GOOGLE_OAUTH_CLIENT_SECRET,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
    });
    response = await axios.post('https://oauth2.googleapis.com/token', payload.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
  } else if (provider === 'outlook') {
    if (!MICROSOFT_OAUTH_CLIENT_ID || !MICROSOFT_OAUTH_CLIENT_SECRET) {
      throw new Error('Microsoft OAuth credentials are not configured.');
    }
    const payload = new URLSearchParams({
      client_id: MICROSOFT_OAUTH_CLIENT_ID,
      client_secret: MICROSOFT_OAUTH_CLIENT_SECRET,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      scope: MICROSOFT_OAUTH_SCOPES,
    });
    response = await axios.post(
      `https://login.microsoftonline.com/${encodeURIComponent(MICROSOFT_OAUTH_TENANT_ID)}/oauth2/v2.0/token`,
      payload.toString(),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );
  } else {
    throw new Error('Unsupported mailbox OAuth provider.');
  }

  const payload = response?.data || {};
  const accessToken = String(payload.access_token || '').trim();
  if (!accessToken) {
    throw new Error('OAuth token refresh did not return an access token.');
  }

  const expiresIn = Number(payload.expires_in);
  const expiresAt = Number.isFinite(expiresIn) && expiresIn > 0
    ? new Date(Date.now() + expiresIn * 1000).toISOString()
    : null;

  const updates = {
    oauth_access_token: accessToken,
    oauth_refresh_token: String(payload.refresh_token || refreshToken || '').trim() || null,
    oauth_token_expires_at: expiresAt,
    oauth_scope: payload.scope ? String(payload.scope) : config.oauth_scope || null,
    oauth_token_type: payload.token_type ? String(payload.token_type) : config.oauth_token_type || 'Bearer',
  };

  if (dbClient && config?.id) {
    const { data, error } = await dbClient
      .from('email_configs')
      .update(updates)
      .eq('id', config.id)
      .select('*')
      .maybeSingle();
    if (error) throw error;
    if (data) return data;
  }

  return {
    ...config,
    ...updates,
  };
};

const ensureMailboxTransportConfig = async (config, dbClient, context = '') => {
  if (!isOAuthMailboxConfig(config)) return config;
  if (!needsMailboxOAuthRefresh(config)) return config;
  try {
    return await refreshMailboxOAuthToken(config, dbClient);
  } catch (error) {
    const details = error?.response?.data || error?.message || error;
    const contextLabel = context ? ` for ${context}` : '';
    throw new Error(`OAuth token refresh failed${contextLabel}: ${typeof details === 'string' ? details : JSON.stringify(details)}`);
  }
};

const buildImapClient = (config, loginMethod) => {
  const security = (config.security || 'SSL').toUpperCase();
  const secure = security !== 'NONE';
  const port = config.imap_port || (secure ? 993 : 143);

  return new ImapFlow({
    host: config.imap_host,
    port,
    secure,
    auth: buildImapAuthFromConfig(config, loginMethod),
    tls: {
      rejectUnauthorized: false,
    },
    logger: false,
  });
};

const mapMessageToRow = (config, message, parsed, mailboxPath) => {
  const parsedFrom = Array.isArray(parsed?.from?.value) ? parsed.from.value : [];
  const parsedTo = Array.isArray(parsed?.to?.value) ? parsed.to.value : [];
  const parsedCc = Array.isArray(parsed?.cc?.value) ? parsed.cc.value : [];
  const parsedReplyTo = Array.isArray(parsed?.replyTo?.value) ? parsed.replyTo.value : [];

  const fromAddresses = normalizeAddressList(parsedFrom.length ? parsedFrom : message.envelope?.from);
  const toAddresses = normalizeAddressList(parsedTo.length ? parsedTo : message.envelope?.to);
  const ccAddresses = normalizeAddressList(parsedCc.length ? parsedCc : message.envelope?.cc);
  const replyToAddresses = normalizeAddressList(parsedReplyTo.length ? parsedReplyTo : message.envelope?.replyTo);

  const fromAddress = fromAddresses?.[0] || message.envelope?.from?.[0]?.address || '';
  const toAddress = toAddresses?.[0] || config.smtp_username;
  const subject = parsed?.subject || message.envelope?.subject || '(No Subject)';
  const body = parsed?.html || parsed?.textAsHtml || parsed?.text || '';
  const messageDate = parsed?.date || message.internalDate || new Date();
  const messageId = normalizeMessageId(parsed?.messageId || message.envelope?.messageId || '');
  const rawInReplyTo = Array.isArray(parsed?.inReplyTo) ? parsed?.inReplyTo?.[0] : parsed?.inReplyTo;
  const inReplyTo = normalizeMessageId(rawInReplyTo || message.envelope?.inReplyTo || '');
  const references = normalizeReferences(parsed?.references || parsed?.headers?.get?.('references') || '');
  const attachments = extractAttachmentMetadata(message.bodyStructure, parsed?.attachments || []);
  const threadId = computeThreadKey({
    message_id: messageId,
    in_reply_to: inReplyTo,
    references,
    subject,
    from_email: fromAddress,
    date: messageDate,
  });

  let seen = false;
  if (Array.isArray(message.flags)) {
    seen = message.flags.includes('\\Seen');
  } else if (message.flags && typeof message.flags.has === 'function') {
    seen = message.flags.has('\\Seen');
  }

  return {
    user_id: config.user_id,
    config_id: config.id,
    uid: message.uid,
    from_email: fromAddress,
    to_email: toAddress,
    to_emails: toAddresses,
    cc_emails: ccAddresses,
    reply_to: replyToAddresses,
    subject,
    body,
    date: messageDate.toISOString(),
    folder: mailboxPath || message.mailbox || 'INBOX',
    read: seen,
    message_id: messageId || null,
    in_reply_to: inReplyTo || null,
    references: references.length ? references : null,
    attachments,
    thread_id: threadId || null,
    direction: 'inbound',
  };
};

const syncMailbox = async (config, limit = 50, dbClient) => {
  if (!dbClient) {
    throw new Error('Supabase client not configured');
  }
  const activeConfig = await ensureMailboxTransportConfig(
    config,
    dbClient,
    `mailbox sync (${config?.smtp_username || config?.id || 'unknown'})`
  );

  if (!hasMailboxTransportCredentials(activeConfig)) {
    throw new Error('Mailbox credentials are missing for this inbox.');
  }

  const client = buildImapClient(activeConfig);
  let processed = 0;
  const rows = [];

  await client.connect();

  try {
    const mailbox = await client.mailboxOpen('INBOX');
    if (!mailbox?.exists) {
      return { processed: 0, inserted: 0, skipped: 0 };
    }

    const startSeq = Math.max(mailbox.exists - limit + 1, 1);
    const range = `${startSeq}:*`;

    for await (const message of client.fetch(range, {
      uid: true,
      flags: true,
      envelope: true,
      source: true,
      internalDate: true,
      bodyStructure: true,
    })) {
      processed++;
      try {
        let parsed;

        try {
          if (message.source) {
            parsed = await simpleParser(message.source);
          } else {
            const download = await client.download(message.uid, { source: true });
            const chunks = [];
            for await (const chunk of download.content) {
              chunks.push(chunk);
            }
            const buffer = Buffer.concat(chunks);
            parsed = await simpleParser(buffer);
          }
        } catch (parseError) {
          console.warn(`[mailbox] Failed to parse message ${message.uid}:`, parseError?.message || parseError);
          parsed = {
            from: { value: message.envelope?.from || [] },
            to: { value: message.envelope?.to || [] },
            subject: message.envelope?.subject || '(No Subject)',
            text: '',
            html: '',
            date: message.internalDate ? new Date(message.internalDate) : new Date(),
          };
        }

        const row = mapMessageToRow(activeConfig, message, parsed, mailbox.path);
        rows.push(row);

        if (LOG_SYNC_SAMPLES && processed <= 3) {
          console.log('[mailbox] Sample row prepared:', {
            uid: row.uid,
            from: row.from_email,
            subject: row.subject,
            date: row.date,
          });
        }
      } catch (parseError) {
        console.error('[mailbox] Unexpected error preparing message', message.uid, parseError);
      }
    }

    if (rows.length === 0) {
      return { processed, inserted: 0, skipped: processed };
    }

    // Deduplicate within this batch
    const uniqueRowsMap = new Map();
    for (const row of rows) {
      if (row.uid == null) continue;
      if (!uniqueRowsMap.has(row.uid)) {
        uniqueRowsMap.set(row.uid, row);
      }
    }
    const uniqueRows = Array.from(uniqueRowsMap.values());

    // Fetch UIDs that already exist for this config
    let existingUids = new Set();
    if (uniqueRows.length > 0) {
      const { data: existingData, error: existingError } = await dbClient
        .from('email_messages')
        .select('uid')
        .eq('config_id', activeConfig.id)
        .in('uid', uniqueRows.map((row) => row.uid));

      if (existingError) {
        throw existingError;
      }

      existingUids = new Set((existingData ?? []).map((item) => item.uid));
    }

    const newRows = uniqueRows.filter((row) => !existingUids.has(row.uid));

    if (newRows.length > 0) {
      const { error: insertError } = await dbClient
        .from('email_messages')
        .insert(newRows);

      if (insertError) {
        throw insertError;
      }
    }

    return {
      processed,
      inserted: newRows.length,
      skipped: Math.max(processed - newRows.length, 0),
    };
  } finally {
    try {
      if (client?.mailboxLock) {
        await client.mailboxClose().catch(() => {});
      }
      await client.logout();
    } catch (closeError) {
      console.error('Error closing IMAP connection', closeError);
    }
  }
};

const AUTO_REPLY_SUBJECT_REGEX = /automatic reply|out of office|vacation|abwesend|auto-response|auto response/i;
const BOUNCE_SENDER_REGEX = /mailer-daemon|postmaster|mail delivery subsystem/i;
const BOUNCE_SUBJECT_REGEX = /delivery status notification|failure|failed|undelivered|undeliverable|returned|rejected/i;
const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp', 'bmp', 'ico']);

const parseBoolean = (value) => {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    return ['true', '1', 'yes', 'y', 'on'].includes(value.trim().toLowerCase());
  }
  return false;
};

const parseOptionalNumber = (value) => {
  if (value === null || value === undefined || value === '') return undefined;
  const num = Number(value);
  return Number.isFinite(num) ? num : undefined;
};

const parseHostCandidates = (value) => {
  if (Array.isArray(value)) {
    return value
      .map((entry) => (typeof entry === 'string' ? entry.trim() : ''))
      .filter(Boolean);
  }
  if (typeof value === 'string') {
    return value
      .split(',')
      .map((entry) => entry.trim())
      .filter(Boolean);
  }
  return undefined;
};

const AUTO_CHECK_REPLIES = parseBoolean(process.env.MAILBOX_AUTO_CHECK_REPLIES || DEFAULT_AUTO_CHECK_REPLIES);
const AUTO_CHECK_INTERVAL_MINUTES = (() => {
  const parsed = Number(process.env.MAILBOX_AUTO_CHECK_INTERVAL_MINUTES || DEFAULT_AUTO_CHECK_INTERVAL_MINUTES);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : Number(DEFAULT_AUTO_CHECK_INTERVAL_MINUTES);
})();
const AUTO_CHECK_LOOKBACK_DAYS = (() => {
  const parsed = Number(process.env.MAILBOX_AUTO_CHECK_LOOKBACK_DAYS || DEFAULT_AUTO_CHECK_LOOKBACK_DAYS);
  const safe = Number.isFinite(parsed) ? parsed : Number(DEFAULT_AUTO_CHECK_LOOKBACK_DAYS);
  return Math.max(1, Math.min(60, safe));
})();
const AUTO_CHECK_USE_DB_SCAN = parseBoolean(process.env.MAILBOX_AUTO_CHECK_USE_DB_SCAN || DEFAULT_AUTO_CHECK_USE_DB_SCAN);
const AUTO_CHECK_CONFIG_ID = (process.env.MAILBOX_AUTO_CHECK_CONFIG_ID || DEFAULT_AUTO_CHECK_CONFIG_ID || '').trim();
const LOG_SYNC_SAMPLES = parseBoolean(process.env.MAILBOX_LOG_SYNC_SAMPLES || 'false');
const AUTO_SYNC_MAILBOXES = parseBoolean(process.env.MAILBOX_AUTO_SYNC_MAILBOXES || DEFAULT_AUTO_SYNC_MAILBOXES);
const AUTO_SYNC_INTERVAL_MINUTES = (() => {
  const parsed = Number(process.env.MAILBOX_AUTO_SYNC_INTERVAL_MINUTES || DEFAULT_AUTO_SYNC_INTERVAL_MINUTES);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : Number(DEFAULT_AUTO_SYNC_INTERVAL_MINUTES);
})();
const AUTO_SYNC_LIMIT = (() => {
  const parsed = Number(process.env.MAILBOX_AUTO_SYNC_LIMIT || DEFAULT_AUTO_SYNC_LIMIT);
  const safe = Number.isFinite(parsed) ? parsed : Number(DEFAULT_AUTO_SYNC_LIMIT);
  return Math.max(1, Math.min(500, safe));
})();
const AUTO_SYNC_CONFIG_ID = (process.env.MAILBOX_AUTO_SYNC_CONFIG_ID || DEFAULT_AUTO_SYNC_CONFIG_ID || '').trim();

const buildCheckOverrides = (body) => {
  const nested = body?.overrides && typeof body.overrides === 'object' ? body.overrides : {};
  const source = { ...nested, ...body };

  return {
    force_legacy_hostinger: parseBoolean(source.force_legacy_hostinger),
    force_direct_tls: parseBoolean(source.force_direct_tls),
    force_starttls: parseBoolean(source.force_starttls),
    force_port: parseOptionalNumber(source.force_port),
    imap_host_override:
      typeof source.imap_host_override === 'string' ? source.imap_host_override.trim() : undefined,
    imap_host_candidates: parseHostCandidates(source.imap_host_candidates),
    max_attempts: parseOptionalNumber(source.max_attempts),
    connection_timeout_ms: parseOptionalNumber(source.connection_timeout_ms),
    greeting_timeout_ms: parseOptionalNumber(source.greeting_timeout_ms),
    socket_timeout_ms: parseOptionalNumber(source.socket_timeout_ms),
  };
};

const normalizeSecurity = (security) => {
  const value = (security || '').toUpperCase();
  if (value === 'TLS') return 'TLS';
  if (value === 'SSL') return 'SSL';
  return 'SSL';
};

const classifyMessage = (from, subject) => {
  const safeFrom = from || '';
  const safeSubject = subject || '';
  const isAutoReply = AUTO_REPLY_SUBJECT_REGEX.test(safeSubject);
  const isBounceSender = BOUNCE_SENDER_REGEX.test(safeFrom);
  const isBounceSubject = BOUNCE_SUBJECT_REGEX.test(safeSubject);
  return { isAutoReply, isBounce: isBounceSender || isBounceSubject };
};

const sanitizeBounceEmails = (emails, senderEmail) => {
  const sender = (senderEmail || '').toLowerCase();
  const unique = new Set();

  for (const email of emails) {
    const normalized = (email || '').toLowerCase();
    if (!normalized || normalized === sender) continue;
    const domain = normalized.split('@')[1] || '';
    const ext = domain.split('.').pop() || '';
    if (IMAGE_EXTENSIONS.has(ext)) continue;
    unique.add(normalized);
  }

  return [...unique];
};

const chunk = (items, size) => {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
};

const buildCaseInsensitiveEmailOrFilter = (emails) =>
  emails
    .map((email) => {
      const escaped = String(email)
        .replace(/\\/g, '\\\\')
        .replace(/,/g, '\\,')
        .replace(/%/g, '\\%')
        .replace(/_/g, '\\_');
      return `email.ilike.${escaped}`;
    })
    .join(',');

const fetchRecipientsByEmailCaseInsensitive = async (dbClient, emails, selectFields) => {
  const normalizedEmails = [...new Set(
    (emails || [])
      .map((email) => String(email || '').trim().toLowerCase())
      .filter(Boolean)
  )];

  if (normalizedEmails.length === 0) {
    return [];
  }

  const recipientsById = new Map();
  const emailChunks = chunk(normalizedEmails, 40);

  for (const emailChunk of emailChunks) {
    const orFilter = buildCaseInsensitiveEmailOrFilter(emailChunk);
    const { data, error } = await dbClient
      .from('recipients')
      .select(selectFields)
      .or(orFilter);

    if (error) {
      console.error('[mailbox] Failed case-insensitive recipient lookup:', error);
      continue;
    }

    (data || []).forEach((recipient) => {
      if (recipient?.id) recipientsById.set(recipient.id, recipient);
    });
  }

  return Array.from(recipientsById.values());
};

const getImapHostCandidates = (config) => {
  const candidates = [];
  const add = (host) => {
    const cleaned = (host || '').trim();
    if (cleaned && !candidates.includes(cleaned)) {
      candidates.push(cleaned);
    }
  };

  const imapHost = (config.imap_host || '').trim();
  const smtpHost = (config.smtp_host || '').toLowerCase();
  const imapLower = imapHost.toLowerCase();

  const prefersTitan = smtpHost.includes('titan.email');
  const isHostinger = smtpHost.includes('hostinger.com') || imapLower.includes('hostinger.com');

  if (prefersTitan || isHostinger) {
    add('imap.titan.email');
  }

  add(imapHost);

  if (isHostinger) {
    add('imap.hostinger.com');
    add('mail.hostinger.com');
  }

  return candidates;
};

const buildConnectionProfiles = (config, security, overrides) => {
  const profiles = [];
  const add = (port, secure, doStartTls, label) => {
    if (!profiles.some((p) => p.port === port && p.secure === secure && p.doStartTls === doStartTls)) {
      profiles.push({ port, secure, doStartTls, label });
    }
  };

  if (
    overrides?.force_legacy_hostinger &&
    !overrides?.force_starttls &&
    !overrides?.force_direct_tls &&
    !Number.isFinite(overrides?.force_port)
  ) {
    add(993, true, false, 'direct-tls');
    return profiles;
  }

  if (overrides?.force_direct_tls) {
    add(993, true, false, 'direct-tls');
    return profiles;
  }

  if (overrides?.force_starttls) {
    add(143, false, true, 'starttls');
    return profiles;
  }

  if (Number.isFinite(overrides?.force_port)) {
    const port = Number(overrides.force_port);
    if (port === 993) add(993, true, false, 'direct-tls');
    else if (port === 143) add(143, false, true, 'starttls');
    else add(port, port === 993, port === 143 || security === 'TLS', 'custom');
    return profiles;
  }

  const parsed = Number(config.imap_port);
  if (Number.isFinite(parsed) && parsed > 0) {
    if (parsed === 993) {
      add(993, true, false, 'direct-tls');
    } else if (parsed === 143) {
      add(143, false, true, 'starttls');
    } else {
      add(parsed, parsed === 993, parsed === 143 || security === 'TLS', 'custom');
    }
  }

  add(993, true, false, 'direct-tls');
  add(143, false, true, 'starttls');

  return profiles;
};

const resolveHostCandidates = (config, overrides) => {
  if (overrides?.imap_host_candidates && Array.isArray(overrides.imap_host_candidates)) {
    const unique = overrides.imap_host_candidates
      .map((host) => (typeof host === 'string' ? host.trim() : ''))
      .filter(Boolean);
    if (unique.length > 0) return Array.from(new Set(unique));
  }

  if (overrides?.imap_host_override && typeof overrides.imap_host_override === 'string') {
    return [overrides.imap_host_override.trim()].filter(Boolean);
  }

  if (overrides?.force_legacy_hostinger) {
    return ['imap.hostinger.com'];
  }

  return getImapHostCandidates(config);
};

const updateCampaignBounceCount = async (dbClient, campaignId) => {
  if (!campaignId) return;

  const { error: rpcError } = await dbClient.rpc('increment_bounced_count', { campaign_id: campaignId });
  if (!rpcError) return;

  console.error(`[mailbox] increment_bounced_count failed for campaign ${campaignId}:`, rpcError);

  const { count, error: countError } = await dbClient
    .from('recipients')
    .select('id', { count: 'exact', head: true })
    .eq('campaign_id', campaignId)
    .eq('bounced', true);

  if (countError) {
    console.error(`[mailbox] Failed to recount bounces for campaign ${campaignId}:`, countError);
    return;
  }

  const { error: updateError } = await dbClient
    .from('campaigns')
    .update({ bounced_count: count ?? 0, updated_at: new Date().toISOString() })
    .eq('id', campaignId);

  if (updateError) {
    console.error(`[mailbox] Failed to update bounced_count for campaign ${campaignId}:`, updateError);
  }
};

const processDbEmails = async (dbClient, config, lookbackDays) => {
  console.log(`[mailbox] [DB Mode] Processing emails for ${config.smtp_username} (Lookback: ${lookbackDays} days)...`);

  const searchDate = new Date();
  searchDate.setDate(searchDate.getDate() - lookbackDays);

  const { data: messages, error } = await dbClient
    .from('email_messages')
    .select('*')
    .eq('config_id', config.id)
    .gte('date', searchDate.toISOString())
    .order('date', { ascending: false })
    .limit(2000);

  if (error) {
    console.error('[mailbox] [DB Mode] Error fetching messages:', error);
    return { error: error.message };
  }

  const safeMessages = messages ?? [];
  console.log(`[mailbox] [DB Mode] Found ${safeMessages.length} messages in DB.`);

  let updatedCount = 0;
  let bouncedCount = 0;

  const bounceCandidates = [];
  const autoReplyCandidates = [];

  for (const msg of safeMessages) {
    const from = msg.from_email || '';
    const subject = msg.subject || '';
    const { isAutoReply, isBounce } = classifyMessage(from, subject);

    if (isBounce) {
      bounceCandidates.push(msg);
    } else if (isAutoReply) {
      autoReplyCandidates.push(msg);
    }
  }

  if (autoReplyCandidates.length > 0) {
    console.log(`[mailbox] [DB Mode] Found ${autoReplyCandidates.length} auto-replies (skipping for bounce count).`);
  }

  const excludedMessageIds = new Set([
    ...bounceCandidates.map((msg) => msg.id),
    ...autoReplyCandidates.map((msg) => msg.id),
  ]);

  if (bounceCandidates.length > 0) {
    console.log(`[mailbox] [DB Mode] Found ${bounceCandidates.length} potential bounce messages.`);
    const allFoundEmails = [];

    for (const msg of bounceCandidates) {
      const body = msg.body || '';
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi;

      const bodyMatches = body.match(emailRegex) || [];
      const subjectMatches = (msg.subject || '').match(emailRegex) || [];

      allFoundEmails.push(...bodyMatches);
      allFoundEmails.push(...subjectMatches);
    }

    const senderEmail = (config.smtp_username || '').toLowerCase();
    const uniqueBounceEmails = sanitizeBounceEmails(allFoundEmails, senderEmail);

    if (uniqueBounceEmails.length > 0) {
      const potentialBounces = await fetchRecipientsByEmailCaseInsensitive(
        dbClient,
        uniqueBounceEmails,
        'id, email, campaign_id, status, bounced'
      );

      if (potentialBounces && potentialBounces.length > 0) {
        const newBounces = potentialBounces.filter((recipient) => !recipient.bounced);
        console.log(
          `[mailbox] [DB Mode] Found ${newBounces.length} new bounces out of ${potentialBounces.length} matches.`
        );

        const bounceIdsToUpdate = [];
        const campaignCounts = {};

        for (const recipient of newBounces) {
          bounceIdsToUpdate.push(recipient.id);
          campaignCounts[recipient.campaign_id] = (campaignCounts[recipient.campaign_id] || 0) + 1;
          bouncedCount++;
        }

        if (bounceIdsToUpdate.length > 0) {
          const { error: updateError } = await dbClient
            .from('recipients')
            .update({
              bounced: true,
              bounced_at: new Date().toISOString(),
              status: 'bounced',
            })
            .in('id', bounceIdsToUpdate);

          if (updateError) {
            console.error('[mailbox] [DB Mode] Error updating bounced recipients:', updateError);
          } else {
            console.log(`[mailbox] [DB Mode] Updated ${bounceIdsToUpdate.length} recipients as bounced.`);
          }
        }

        for (const campaignId of Object.keys(campaignCounts)) {
          await updateCampaignBounceCount(dbClient, campaignId);
        }
      } else {
        console.log(
          `[mailbox] [DB Mode] No matching recipients found for ${uniqueBounceEmails.length} extracted emails.`
        );
      }
    }
  }

  const replyMessages = safeMessages.filter((msg) => !excludedMessageIds.has(msg.id));
  const senderEmails = replyMessages
    .map((msg) => msg.from_email)
    .filter((email) => email && email.includes('@'))
    .map((email) => email.toLowerCase());

  const uniqueSenders = [...new Set(senderEmails)];

  if (uniqueSenders.length > 0) {
    const potentialReplies = (await fetchRecipientsByEmailCaseInsensitive(
      dbClient,
      uniqueSenders,
      'id, email, campaign_id, last_email_sent_at, updated_at, replied'
    )).filter((recipient) => !recipient.replied);

    if (potentialReplies && potentialReplies.length > 0) {
      console.log(`[mailbox] [DB Mode] Found ${potentialReplies.length} potential replies to process.`);

      const senderLastMsgDate = new Map();
      replyMessages.forEach((msg) => {
        if (msg.from_email && msg.date) {
          const email = msg.from_email.toLowerCase();
          const date = new Date(msg.date);
          if (!senderLastMsgDate.has(email) || date > senderLastMsgDate.get(email)) {
            senderLastMsgDate.set(email, date);
          }
        }
      });

      const campaignCounts = {};
      const recipientIdsToUpdate = [];

      for (const recipient of potentialReplies) {
        const replyDate = senderLastMsgDate.get(recipient.email.toLowerCase());
        const sentDateStr = recipient.last_email_sent_at || recipient.updated_at;

        if (replyDate && sentDateStr) {
          const sentDate = new Date(sentDateStr);
          if (replyDate.getTime() > sentDate.getTime() + 60000) {
            recipientIdsToUpdate.push(recipient.id);
            campaignCounts[recipient.campaign_id] = (campaignCounts[recipient.campaign_id] || 0) + 1;
            updatedCount++;
          } else {
            console.log(
              `[mailbox] [DB Mode] Skipping reply from ${recipient.email}: reply date (${replyDate.toISOString()}) is before or too close to sent date (${sentDate.toISOString()})`
            );
          }
        }
      }

      if (recipientIdsToUpdate.length > 0) {
        await dbClient
          .from('recipients')
          .update({ replied: true, updated_at: new Date().toISOString() })
          .in('id', recipientIdsToUpdate);
      }

      for (const campaignId of Object.keys(campaignCounts)) {
        await dbClient.rpc('increment_replied_count', { campaign_id: campaignId });
      }
    }
  }

  return { processed: safeMessages.length, replies: updatedCount, bounces: bouncedCount };
};

const checkRepliesAndBouncesForConfig = async (dbClient, config, lookbackDays = 7, overrides) => {
  let lastError;
  const activeConfig = await ensureMailboxTransportConfig(
    config,
    dbClient,
    `reply check (${config?.smtp_username || config?.id || 'unknown'})`
  );

  if (!hasMailboxTransportCredentials(activeConfig)) {
    throw new Error(`Mailbox credentials missing for ${activeConfig.smtp_username || activeConfig.id}`);
  }

  const hostCandidates = resolveHostCandidates(activeConfig, overrides);
  if (!hostCandidates.length) {
    throw new Error(`IMAP host missing for ${activeConfig.smtp_username}`);
  }

  const security = normalizeSecurity(activeConfig.security);
  const connectionProfiles = buildConnectionProfiles(activeConfig, security, overrides);
  const enableTlsDebug = (process.env.IMAP_DEBUG_TLS || '').toLowerCase() === 'true';
  const maxAttempts = Number.isFinite(overrides?.max_attempts)
    ? Math.max(1, Math.min(3, Number(overrides.max_attempts)))
    : 3;
  const defaultConnectionTimeout = overrides?.force_legacy_hostinger ? 10000 : 20000;
  const connectionTimeout = Number.isFinite(overrides?.connection_timeout_ms)
    ? Math.max(3000, Number(overrides.connection_timeout_ms))
    : defaultConnectionTimeout;
  const greetingTimeout = Number.isFinite(overrides?.greeting_timeout_ms)
    ? Math.max(3000, Number(overrides.greeting_timeout_ms))
    : connectionTimeout;
  const socketTimeout = Number.isFinite(overrides?.socket_timeout_ms)
    ? Math.max(10000, Number(overrides.socket_timeout_ms))
    : 60000;

  const debugTlsConnect = (host, port) =>
    new Promise((resolve) => {
      const socket = tls.connect({ host, port, servername: host, rejectUnauthorized: false }, () => {
        console.log(`[mailbox] [Debug] Raw TLS connection successful to ${host}:${port}.`);
        socket.end();
        resolve();
      });

      socket.setTimeout(5000, () => {
        console.error(`[mailbox] [Debug] Raw TLS connection timed out to ${host}:${port}.`);
        socket.destroy();
        resolve();
      });

      socket.on('error', (err) => {
        console.error(`[mailbox] [Debug] Raw TLS connection failed to ${host}:${port}:`, err);
        resolve();
      });
    });

  for (const host of hostCandidates) {
    for (const profile of connectionProfiles) {
      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        console.log(
          `[mailbox] Connecting to ${host}:${profile.port} (${profile.label}) for ${activeConfig.smtp_username} (Attempt ${attempt})...`
        );

        if (enableTlsDebug && profile.secure) {
          await debugTlsConnect(host, profile.port);
        }

        const tlsOptions = { rejectUnauthorized: false };

        if (attempt === 1) {
          tlsOptions.servername = host;
          tlsOptions.minVersion = 'TLSv1.2';
        } else if (attempt === 2) {
          tlsOptions.servername = host;
        }

        const loginMethod = isOAuthMailboxConfig(activeConfig)
          ? undefined
          : attempt === 1
            ? 'AUTH=PLAIN'
            : attempt === 2
              ? 'AUTH=LOGIN'
              : 'LOGIN';
        const auth = buildImapAuthFromConfig(activeConfig, loginMethod);

        const client = new ImapFlow({
          host,
          port: profile.port,
          secure: profile.secure,
          doSTARTTLS: profile.doStartTls,
          auth,
          tls: tlsOptions,
          logger: {
            debug: (obj) => console.log(`[IMAP DEBUG] ${obj?.msg || JSON.stringify(obj)}`),
            info: (obj) => console.log(`[IMAP INFO] ${obj?.msg || JSON.stringify(obj)}`),
            warn: (obj) => console.warn(`[IMAP WARN] ${obj?.msg || JSON.stringify(obj)}`),
            error: (obj) => console.error(`[IMAP ERROR] ${obj?.msg || JSON.stringify(obj)}`),
          },
          clientInfo: {
            name: 'EmailBridge',
            version: '1.0.0',
          },
          disableAutoIdle: true,
          disableAutoEnable: true,
          disableCompression: true,
          connectionTimeout,
          greetingTimeout,
          socketTimeout,
        });

        client.on('error', (err) => {
          console.error(
            `[mailbox] IMAP Client Error for ${activeConfig.smtp_username} (${host}:${profile.port}) (Attempt ${attempt}):`,
            err
          );
        });

        let lock = null;
        try {
          await client.connect();
          lock = await client.getMailboxLock('INBOX');

          try {
            const searchDate = new Date();
            searchDate.setDate(searchDate.getDate() - lookbackDays);
            console.log(`[mailbox] Searching emails since ${searchDate.toISOString()}...`);

            const searchResult = await client.search({ since: searchDate });

            let processedCount = 0;
            let updatedCount = 0;
            let bouncedCount = 0;

            if (searchResult.length > 0) {
              const messagesToCheck = searchResult.slice(-1000);
              console.log(
                `[mailbox] Found ${searchResult.length} messages, checking headers for last ${messagesToCheck.length}...`
              );

              const bounceSequenceNumbers = [];

              for await (const message of client.fetch(messagesToCheck, { envelope: true })) {
                const { envelope, seq } = message;
                if (!envelope) continue;

                const from = envelope.from?.[0]?.address || '';
                const subject = envelope.subject || '';
                const inReplyTo = envelope.inReplyTo;

                const { isAutoReply, isBounce } = classifyMessage(from, subject);

                if (inReplyTo && !isBounce && !isAutoReply) {
                  const normalizeId = (id) => {
                    if (typeof id !== 'string') return '';
                    return id.replace(/[<>]/g, '').trim();
                  };

                  const rawIds = [];
                  if (typeof inReplyTo === 'string') {
                    rawIds.push(inReplyTo);
                  } else if (Array.isArray(inReplyTo)) {
                    inReplyTo.forEach((id) => {
                      if (typeof id === 'string') rawIds.push(id);
                    });
                  }

                  const idsToCheck = new Set();
                  rawIds.forEach((id) => {
                    if (!id) return;
                    const clean = normalizeId(id);
                    if (clean) {
                      idsToCheck.add(clean);
                      idsToCheck.add(`<${clean}>`);
                    }
                    idsToCheck.add(id.trim());
                  });

                  const messageIdsToCheck = Array.from(idsToCheck);

                  if (messageIdsToCheck.length > 0) {
                    const { data: recipientsByMessageId } = await dbClient
                      .from('recipients')
                      .select('id, email, campaign_id')
                      .in('message_id', messageIdsToCheck)
                      .eq('replied', false);

                    const { data: recipientsByThreadId } = await dbClient
                      .from('recipients')
                      .select('id, email, campaign_id')
                      .in('thread_id', messageIdsToCheck)
                      .eq('replied', false);

                    const allRecipients = [...(recipientsByMessageId || []), ...(recipientsByThreadId || [])];
                    const recipients = Array.from(new Map(allRecipients.map((r) => [r.id, r])).values());

                    if (recipients && recipients.length > 0) {
                      for (const recipient of recipients) {
                        console.log(
                          `[mailbox] Detected reply from ${recipient.email} (Campaign ${recipient.campaign_id})`
                        );
                        await dbClient
                          .from('recipients')
                          .update({ replied: true, updated_at: new Date().toISOString() })
                          .eq('id', recipient.id);

                        await dbClient.rpc('increment_replied_count', { campaign_id: recipient.campaign_id });

                        updatedCount++;
                      }
                    }
                  }
                }

                if (isBounce) {
                  bounceSequenceNumbers.push(seq);
                }
              }

              console.log(
                `[mailbox] Found ${bounceSequenceNumbers.length} potential bounces. Fetching full content for max 20...`
              );

              const bouncesToProcess = bounceSequenceNumbers.slice(0, 20);

              if (bouncesToProcess.length > 0) {
                for await (const message of client.fetch(bouncesToProcess, { source: true })) {
                  try {
                    const source = message.source;
                    const sourceBuffer = Buffer.isBuffer(source) ? source : Buffer.from(source || '');
                    const parsed = await simpleParser(sourceBuffer);

                    const from = parsed.from?.text || '';
                    const subject = parsed.subject || '';
                    console.log(`[mailbox] Processing bounce: ${subject} from ${from}`);

                    const body = parsed.text || parsed.html || '';
                    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi;
                    const foundEmails = body.match(emailRegex) || [];

                    console.log(
                      `[mailbox] Extracted ${foundEmails.length} emails from bounce body: ${foundEmails
                        .slice(0, 5)
                        .join(', ')}${foundEmails.length > 5 ? '...' : ''}`
                    );

                    const failedHeader = parsed.headers.get('x-failed-recipients');
                    if (failedHeader) {
                      if (typeof failedHeader === 'string') foundEmails.push(failedHeader);
                      else if (Array.isArray(failedHeader)) foundEmails.push(...failedHeader);
                    }

                    const senderEmail = (activeConfig.smtp_username || '').toLowerCase();
                    const uniqueEmails = sanitizeBounceEmails(foundEmails, senderEmail);

                    if (uniqueEmails.length > 0) {
                      const recipients = await fetchRecipientsByEmailCaseInsensitive(
                        dbClient,
                        uniqueEmails,
                        'id, email, campaign_id, bounced, last_email_sent_at'
                      );

                      if (recipients && recipients.length > 0) {
                        const recipientsToUpdate = recipients
                          .filter((recipient) => !recipient.bounced)
                          .sort((a, b) => {
                            const timeA = a.last_email_sent_at ? new Date(a.last_email_sent_at).getTime() : 0;
                            const timeB = b.last_email_sent_at ? new Date(b.last_email_sent_at).getTime() : 0;
                            return timeB - timeA;
                          });
                        const recipientIds = recipientsToUpdate.map((recipient) => recipient.id);
                        const campaignIds = new Set(recipientsToUpdate.map((recipient) => recipient.campaign_id));

                        if (recipientIds.length > 0) {
                          const { error: updateError } = await dbClient
                            .from('recipients')
                            .update({
                              bounced: true,
                              bounced_at: new Date().toISOString(),
                              status: 'bounced',
                            })
                            .in('id', recipientIds);

                          if (updateError) {
                            console.error('[mailbox] Error updating bounced recipients:', updateError);
                          } else {
                            recipientsToUpdate.forEach((recipient) => {
                              console.log(
                                `[mailbox] Confirmed bounce for ${recipient.email} (Campaign ${recipient.campaign_id})`
                              );
                              bouncedCount++;
                            });
                          }
                        }

                        for (const campaignId of campaignIds) {
                          await updateCampaignBounceCount(dbClient, campaignId);
                        }
                      }
                    }
                    processedCount++;
                  } catch (msgError) {
                    console.error(`[mailbox] Error processing message ${message.uid}:`, msgError);
                  }
                }
              }
            } else {
              console.log('[mailbox] No recent messages found.');
            }

            return { processed: processedCount, replies: updatedCount, bounces: bouncedCount };
          } finally {
            try {
              lock?.release();
            } catch (releaseError) {
              console.warn(`[mailbox] Mailbox lock release failed for ${activeConfig.smtp_username}:`, releaseError);
            }
          }
        } catch (err) {
          lastError = err;
          console.error(
            `[mailbox] Error checking emails for ${activeConfig.smtp_username} (${host}:${profile.port}) (Attempt ${attempt}):`,
            err
          );

          if (attempt < maxAttempts) {
            console.log(`[mailbox] Retrying ${activeConfig.smtp_username} in 2s...`);
            await new Promise((resolve) => setTimeout(resolve, 2000));
          }
        } finally {
          try {
            await client.logout();
          } catch (logoutError) {
            // ignore
          }
        }
      }
    }
  }

  throw lastError;
};
const BYPASS_USER = {
  id: 'mailbox-bypass-user',
  email: 'bypass@local.dev',
};

const buildUserSupabaseClient = (authHeader) => {
  if (!hasAnonKey) return null;
  const headers = authHeader ? { Authorization: authHeader } : {};
  return createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    global: { headers },
    auth: { persistSession: false },
  });
};

const authenticateRequest = async (authHeader) => {
  const canBypassAuth = MAILBOX_BYPASS_AUTH && hasServiceRoleKey;

  if (!authHeader?.startsWith('Bearer ')) {
    if (canBypassAuth) {
      console.warn('[mailbox] No Authorization header supplied - bypassing auth (development mode).');
      return BYPASS_USER;
    }
    console.warn('[mailbox] Missing Authorization header.');
    return null;
  }

  const token = authHeader.replace('Bearer ', '').trim();
  if (!token) {
    console.warn('[mailbox] Authorization header present but token empty.');
    return canBypassAuth ? BYPASS_USER : null;
  }

  const authClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
  if (!authClient) {
    console.warn('[mailbox] Supabase keys not configured for auth.');
    return null;
  }

  const { data, error } = await authClient.auth.getUser(token);
  if (error || !data?.user) {
    console.warn('[mailbox] Supabase token validation failed:', error?.message || 'unknown error');
    return canBypassAuth ? BYPASS_USER : null;
  }
  return data.user;
};

const matchesAdminSecret = (req) => {
  if (MAILBOX_ADMIN_SECRET) {
    const authHeader = req.headers.authorization || '';
    const bearer = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
    if (bearer && bearer === MAILBOX_ADMIN_SECRET) return true;
    const headerSecret = req.headers['x-mailbox-admin-secret'];
    if (typeof headerSecret === 'string' && headerSecret.trim() === MAILBOX_ADMIN_SECRET) return true;
    return false;
  }

  return MAILBOX_BYPASS_AUTH && hasServiceRoleKey;
};

const upsertCrmCredential = async (dbClient, payload) => {
  const { data, error } = await dbClient
    .from('crm_credentials')
    .upsert(payload, { onConflict: 'user_id,provider' })
    .select('provider, display_name, owner_id, instance_url')
    .maybeSingle();

  if (error) throw error;
  return data;
};
const loadMessageForUser = async (dbClient, user, messageId, canBypassAuth) => {
  let query = dbClient.from('email_messages').select('*').eq('id', messageId);
  if (!canBypassAuth) {
    query = query.eq('user_id', user.id);
  }
  const { data, error } = await query.maybeSingle();
  if (error) throw error;
  return data;
};

const loadConfigForUser = async (dbClient, user, configId, canBypassAuth) => {
  const normalizedConfigId = String(configId || '').trim();
  if (!normalizedConfigId) return null;

  const runQuery = async (client) => {
    let query = client.from('email_configs').select('*').eq('id', normalizedConfigId);
    if (!canBypassAuth) {
      query = query.eq('user_id', user.id);
    }
    return query.maybeSingle();
  };

  const primary = await runQuery(dbClient);
  if (!primary.error) return primary.data;

  if (supabaseAdmin && dbClient !== supabaseAdmin) {
    const fallback = await runQuery(supabaseAdmin);
    if (!fallback.error) return fallback.data;
    throw fallback.error;
  }

  throw primary.error;
};

const runReplyCheck = async ({ configId, lookbackDays, useDbScan, overrides, source = 'manual' } = {}) => {
  if (!supabaseAdmin) {
    throw new Error('Supabase service role key is required for reply checks.');
  }

  const safeConfigId = configId ? String(configId).trim() : '';
  const parsedLookback = Number(lookbackDays ?? DEFAULT_CHECK_LOOKBACK_DAYS);
  const fallbackLookback = Number(DEFAULT_CHECK_LOOKBACK_DAYS);
  const safeLookbackDays = Math.max(
    1,
    Math.min(60, Number.isFinite(parsedLookback) ? parsedLookback : fallbackLookback)
  );
  const safeUseDbScan = parseBoolean(useDbScan);
  const safeOverrides = overrides && typeof overrides === 'object' ? overrides : buildCheckOverrides({});

  let query = supabaseAdmin.from('email_configs').select('*');
  if (!safeUseDbScan) {
    query = query.not('imap_host', 'is', null);
  }
  if (safeConfigId) {
    query = query.eq('id', safeConfigId);
  }

  const { data: configs, error } = await query;
  if (error) {
    console.error('[mailbox] Failed to load email configs for reply check:', error);
    throw new Error('Failed to load email configurations.');
  }

  const results = [];
  const safeConfigs = configs ?? [];

  for (const config of safeConfigs) {
    try {
      const result = safeUseDbScan
        ? await processDbEmails(supabaseAdmin, config, safeLookbackDays)
        : await checkRepliesAndBouncesForConfig(supabaseAdmin, config, safeLookbackDays, safeOverrides);
      results.push({ email: config.smtp_username, result });
      await new Promise((resolve) => setTimeout(resolve, 1000));
    } catch (err) {
      console.error(`[mailbox] Error processing ${config.smtp_username}:`, err);
      results.push({ email: config.smtp_username, error: err?.message || 'Reply check failed' });
    }
  }

  return {
    success: true,
    results,
    lookbackDays: safeLookbackDays,
    configCount: safeConfigs.length,
    useDbScan: safeUseDbScan,
    source,
  };
};

let autoCheckInFlight = false;
let autoCheckTimer = null;
let autoSyncInFlight = false;
let autoSyncTimer = null;

const runMailboxSync = async ({ configId, limit, source = 'manual' } = {}) => {
  if (!supabaseAdmin) {
    throw new Error('Supabase service role key is required for mailbox sync.');
  }

  const safeConfigId = configId ? String(configId).trim() : '';
  const parsedLimit = Number(limit ?? 50);
  const safeLimit = Math.max(1, Math.min(500, Number.isFinite(parsedLimit) ? parsedLimit : 50));

  let query = supabaseAdmin
    .from('email_configs')
    .select('*')
    .not('imap_host', 'is', null)
    .not('imap_port', 'is', null)
    .not('smtp_username', 'is', null);

  if (safeConfigId) {
    query = query.eq('id', safeConfigId);
  }

  const { data: configs, error } = await query;
  if (error) {
    console.error('[mailbox] Failed to load email configs for mailbox sync:', error);
    throw new Error('Failed to load email configurations.');
  }

  const results = [];
  const safeConfigs = configs ?? [];
  let totalProcessed = 0;
  let totalInserted = 0;
  let totalSkipped = 0;
  let failed = 0;

  for (const config of safeConfigs) {
    try {
      const stats = await syncMailbox(config, safeLimit, supabaseAdmin);
      totalProcessed += Number(stats?.processed || 0);
      totalInserted += Number(stats?.inserted || 0);
      totalSkipped += Number(stats?.skipped || 0);
      results.push({
        config_id: config.id,
        email: config.smtp_username,
        ...stats,
      });
      await new Promise((resolve) => setTimeout(resolve, 1000));
    } catch (err) {
      failed++;
      console.error(`[mailbox] Error syncing ${config.smtp_username}:`, err);
      results.push({
        config_id: config.id,
        email: config.smtp_username,
        error: err?.message || 'Mailbox sync failed',
      });
    }
  }

  return {
    success: true,
    results,
    limit: safeLimit,
    configCount: safeConfigs.length,
    totals: {
      processed: totalProcessed,
      inserted: totalInserted,
      skipped: totalSkipped,
      failed,
    },
    source,
  };
};

const startAutoReplyChecks = () => {
  if (!AUTO_CHECK_REPLIES) return;
  if (!supabaseAdmin) {
    console.warn('[mailbox] Auto reply checks disabled: SUPABASE_SERVICE_ROLE_KEY is required.');
    return;
  }

  const intervalMs = Math.max(1, AUTO_CHECK_INTERVAL_MINUTES) * 60 * 1000;
  const overrides = buildCheckOverrides({});

  const run = async () => {
    if (autoCheckInFlight) {
      console.log('[mailbox] Auto reply check skipped (previous run still in progress).');
      return;
    }
    autoCheckInFlight = true;
    const startedAt = Date.now();
    try {
      const payload = await runReplyCheck({
        configId: AUTO_CHECK_CONFIG_ID || undefined,
        lookbackDays: AUTO_CHECK_LOOKBACK_DAYS,
        useDbScan: AUTO_CHECK_USE_DB_SCAN,
        overrides,
        source: 'auto',
      });
      const durationSeconds = Math.round((Date.now() - startedAt) / 1000);
      console.log(
        `[mailbox] Auto reply check complete in ${durationSeconds}s (${payload.configCount} mailbox${
          payload.configCount === 1 ? '' : 'es'
        }).`
      );
    } catch (error) {
      console.error('[mailbox] Auto reply check failed:', error?.message || error);
    } finally {
      autoCheckInFlight = false;
    }
  };

  console.log(
    `[mailbox] Auto reply checks enabled. Interval: ${AUTO_CHECK_INTERVAL_MINUTES} min, lookback: ${AUTO_CHECK_LOOKBACK_DAYS} days${
      AUTO_CHECK_CONFIG_ID ? `, configId: ${AUTO_CHECK_CONFIG_ID}` : ''
    }.`
  );
  run();
  autoCheckTimer = setInterval(run, intervalMs);
  if (typeof autoCheckTimer?.unref === 'function') {
    autoCheckTimer.unref();
  }
};

const startAutoMailboxSync = () => {
  if (!AUTO_SYNC_MAILBOXES) return;
  if (!supabaseAdmin) {
    console.warn('[mailbox] Auto mailbox sync disabled: SUPABASE_SERVICE_ROLE_KEY is required.');
    return;
  }

  const intervalMs = Math.max(1, AUTO_SYNC_INTERVAL_MINUTES) * 60 * 1000;

  const run = async () => {
    if (autoSyncInFlight) {
      console.log('[mailbox] Auto mailbox sync skipped (previous run still in progress).');
      return;
    }
    autoSyncInFlight = true;
    const startedAt = Date.now();
    try {
      const payload = await runMailboxSync({
        configId: AUTO_SYNC_CONFIG_ID || undefined,
        limit: AUTO_SYNC_LIMIT,
        source: 'auto',
      });
      const durationSeconds = Math.round((Date.now() - startedAt) / 1000);
      console.log(
        `[mailbox] Auto mailbox sync complete in ${durationSeconds}s (${payload.configCount} mailbox${
          payload.configCount === 1 ? '' : 'es'
        }, ${payload.totals?.inserted || 0} inserted, ${payload.totals?.failed || 0} failed).`
      );
    } catch (error) {
      console.error('[mailbox] Auto mailbox sync failed:', error?.message || error);
    } finally {
      autoSyncInFlight = false;
    }
  };

  console.log(
    `[mailbox] Auto mailbox sync enabled. Interval: ${AUTO_SYNC_INTERVAL_MINUTES} min, limit: ${AUTO_SYNC_LIMIT}${
      AUTO_SYNC_CONFIG_ID ? `, configId: ${AUTO_SYNC_CONFIG_ID}` : ''
    }.`
  );
  run();
  autoSyncTimer = setInterval(run, intervalMs);
  if (typeof autoSyncTimer?.unref === 'function') {
    autoSyncTimer.unref();
  }
};

const decodeJwtPayload = (token) => {
  if (!token || typeof token !== 'string') return {};
  const parts = token.split('.');
  if (parts.length < 2) return {};

  try {
    const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = base64.padEnd(base64.length + ((4 - (base64.length % 4)) % 4), '=');
    const raw = Buffer.from(padded, 'base64').toString('utf8');
    return JSON.parse(raw);
  } catch {
    return {};
  }
};

const fetchGoogleMailboxIdentity = async (accessToken) => {
  try {
    const response = await axios.get('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    return {
      email: String(response?.data?.email || '').trim(),
      displayName: String(response?.data?.name || '').trim(),
    };
  } catch (error) {
    console.warn('[mailbox] Google userinfo lookup failed:', error?.response?.data || error?.message || error);
    return { email: '', displayName: '' };
  }
};

const fetchMicrosoftMailboxIdentity = async (accessToken) => {
  try {
    const response = await axios.get('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    return {
      email: String(response?.data?.mail || response?.data?.userPrincipalName || '').trim(),
      displayName: String(response?.data?.displayName || '').trim(),
    };
  } catch (error) {
    console.warn('[mailbox] Microsoft Graph user lookup failed:', error?.response?.data || error?.message || error);
    return { email: '', displayName: '' };
  }
};

const buildMailboxOAuthAuthorizationUrl = ({ provider, state, redirectUri }) => {
  if (provider === 'gmail') {
    if (!GOOGLE_OAUTH_CLIENT_ID) {
      throw new Error('GOOGLE_OAUTH_CLIENT_ID is not configured.');
    }
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: GOOGLE_OAUTH_CLIENT_ID,
      redirect_uri: redirectUri,
      scope: GOOGLE_OAUTH_SCOPES,
      access_type: 'offline',
      include_granted_scopes: 'true',
      prompt: 'consent',
      state,
    });
    return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
  }

  if (provider === 'outlook') {
    if (!MICROSOFT_OAUTH_CLIENT_ID) {
      throw new Error('MICROSOFT_OAUTH_CLIENT_ID is not configured.');
    }
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: MICROSOFT_OAUTH_CLIENT_ID,
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: MICROSOFT_OAUTH_SCOPES,
      prompt: 'select_account',
      state,
    });
    return `https://login.microsoftonline.com/${encodeURIComponent(MICROSOFT_OAUTH_TENANT_ID)}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  throw new Error('Unsupported mailbox OAuth provider.');
};

const exchangeMailboxOAuthCode = async ({ provider, code, redirectUri }) => {
  let tokenResp;
  if (provider === 'gmail') {
    if (!GOOGLE_OAUTH_CLIENT_ID || !GOOGLE_OAUTH_CLIENT_SECRET) {
      throw new Error('Google OAuth credentials are not configured.');
    }

    const payload = new URLSearchParams({
      code,
      client_id: GOOGLE_OAUTH_CLIENT_ID,
      client_secret: GOOGLE_OAUTH_CLIENT_SECRET,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
    });
    tokenResp = await axios.post('https://oauth2.googleapis.com/token', payload.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
  } else if (provider === 'outlook') {
    if (!MICROSOFT_OAUTH_CLIENT_ID || !MICROSOFT_OAUTH_CLIENT_SECRET) {
      throw new Error('Microsoft OAuth credentials are not configured.');
    }
    const payload = new URLSearchParams({
      code,
      client_id: MICROSOFT_OAUTH_CLIENT_ID,
      client_secret: MICROSOFT_OAUTH_CLIENT_SECRET,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
      scope: MICROSOFT_OAUTH_SCOPES,
    });
    tokenResp = await axios.post(
      `https://login.microsoftonline.com/${encodeURIComponent(MICROSOFT_OAUTH_TENANT_ID)}/oauth2/v2.0/token`,
      payload.toString(),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );
  } else {
    throw new Error('Unsupported mailbox OAuth provider.');
  }

  const payload = tokenResp?.data || {};
  const accessToken = String(payload.access_token || '').trim();
  const refreshToken = String(payload.refresh_token || '').trim();
  const tokenType = String(payload.token_type || 'Bearer').trim();
  const scope = String(payload.scope || '').trim();
  const expiresIn = Number(payload.expires_in);
  const expiresAt = Number.isFinite(expiresIn) && expiresIn > 0
    ? new Date(Date.now() + expiresIn * 1000).toISOString()
    : null;

  if (!accessToken) {
    throw new Error('OAuth exchange did not return an access token.');
  }

  const jwtPayload = decodeJwtPayload(payload.id_token || '');
  const profile = provider === 'gmail'
    ? await fetchGoogleMailboxIdentity(accessToken)
    : await fetchMicrosoftMailboxIdentity(accessToken);

  const email =
    String(profile.email || jwtPayload.email || jwtPayload.preferred_username || jwtPayload.upn || '').trim();
  const displayName =
    String(profile.displayName || jwtPayload.name || jwtPayload.given_name || '').trim();

  if (!email) {
    throw new Error('OAuth exchange succeeded but mailbox email could not be resolved.');
  }

  const preset = resolveMailboxPresetForProvider(provider);
  return {
    provider,
    accessToken,
    refreshToken: refreshToken || null,
    tokenType,
    scope: scope || null,
    expiresIn: Number.isFinite(expiresIn) ? expiresIn : null,
    expiresAt,
    email,
    displayName: displayName || email,
    ...(preset || {}),
  };
};

app.get('/healthz', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.post('/mailbox/:provider/oauth/start', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const provider = resolveMailboxOAuthProvider(req.params.provider);
    if (!provider) {
      return res.status(400).json({ error: 'Unsupported mailbox OAuth provider.' });
    }

    const redirectUri = resolveMailboxOAuthRedirectUri(req.body?.redirectUri);
    const state = `${provider}:${crypto.randomUUID()}`;

    if (MAILBOX_OAUTH_SIMULATE) {
      const simulatedCode = `sim-${provider}-${Date.now()}`;
      const simulatedUrl = new URL(redirectUri);
      simulatedUrl.searchParams.set('code', simulatedCode);
      simulatedUrl.searchParams.set('state', state);
      return res.json({
        mode: 'simulate',
        provider,
        state,
        authUrl: simulatedUrl.toString(),
        simulatedCode,
      });
    }

    const authUrl = buildMailboxOAuthAuthorizationUrl({ provider, state, redirectUri });
    return res.json({
      mode: 'live',
      provider,
      state,
      authUrl,
    });
  } catch (error) {
    console.error('[mailbox] OAuth start failed:', error?.response?.data || error?.message || error);
    return res.status(500).json({ error: error?.message || 'Unable to start mailbox OAuth flow.' });
  }
});

app.post('/mailbox/:provider/oauth/exchange', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const provider = resolveMailboxOAuthProvider(req.params.provider);
    if (!provider) {
      return res.status(400).json({ error: 'Unsupported mailbox OAuth provider.' });
    }

    const code = String(req.body?.code || '').trim();
    if (!code) {
      return res.status(400).json({ error: 'OAuth code is required.' });
    }

    const redirectUri = resolveMailboxOAuthRedirectUri(req.body?.redirectUri);

    if (MAILBOX_OAUTH_SIMULATE || code.startsWith('sim-')) {
      const preset = resolveMailboxPresetForProvider(provider);
      const simulatedEmail = `sim.${provider}.${user.id.slice(0, 8)}@example.com`;
      return res.json({
        provider,
        accessToken: `sim-access-${provider}-${Date.now()}`,
        refreshToken: `sim-refresh-${provider}-${Date.now()}`,
        tokenType: 'Bearer',
        scope: provider === 'gmail' ? GOOGLE_OAUTH_SCOPES : MICROSOFT_OAUTH_SCOPES,
        expiresIn: 3600,
        expiresAt: new Date(Date.now() + 3600 * 1000).toISOString(),
        email: simulatedEmail,
        displayName: `Simulated ${provider === 'gmail' ? 'Gmail' : 'Outlook'} mailbox`,
        ...(preset || {}),
      });
    }

    const payload = await exchangeMailboxOAuthCode({
      provider,
      code,
      redirectUri,
    });

    return res.json(payload);
  } catch (error) {
    console.error('[mailbox] OAuth exchange failed:', error?.response?.data || error?.message || error);
    return res.status(500).json({ error: error?.message || 'Unable to complete mailbox OAuth exchange.' });
  }
});

app.post('/crm/:provider/oauth/start', (req, res) => {
  const provider = String(req.params.provider || '').toLowerCase();
  const state = crypto.randomUUID();

  if (CRM_OAUTH_SIMULATE) {
    return res.json({
      mode: 'simulate',
      authUrl: `${CRM_FRONTEND_URL}/oauth/simulate?provider=${provider}&state=${state}`,
      simulatedCode: `sim-${provider}-${Date.now()}`
    });
  }

  if (provider === 'hubspot') {
    if (!HUBSPOT_CLIENT_ID) {
      return res.status(400).send('HUBSPOT_CLIENT_ID is not configured.');
    }
    const params = new URLSearchParams({
      client_id: HUBSPOT_CLIENT_ID,
      redirect_uri: CRM_REDIRECT_URI,
      scope: HUBSPOT_SCOPES,
      state
    });
    return res.json({
      mode: 'live',
      authUrl: `https://app.hubspot.com/oauth/authorize?${params.toString()}`
    });
  }

  if (provider === 'salesforce') {
    if (!SALESFORCE_CLIENT_ID) {
      return res.status(400).send('SALESFORCE_CLIENT_ID is not configured.');
    }
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: SALESFORCE_CLIENT_ID,
      redirect_uri: CRM_REDIRECT_URI,
      scope: SALESFORCE_SCOPES,
      state
    });
    return res.json({
      mode: 'live',
      authUrl: `${SALESFORCE_LOGIN_URL}/services/oauth2/authorize?${params.toString()}`
    });
  }

  return res.status(400).send('Unsupported CRM provider.');
});

app.post('/crm/:provider/oauth/exchange', async (req, res) => {
  const provider = String(req.params.provider || '').toLowerCase();
  const { code } = req.body || {};
  if (!code) {
    return res.status(400).send('OAuth code is required.');
  }

  if (CRM_OAUTH_SIMULATE) {
    return res.json({
      accessToken: `sim-access-${provider}`,
      refreshToken: `sim-refresh-${provider}`,
      expiresIn: 3600,
      accountLabel: 'Sandbox workspace'
    });
  }

  try {
    if (provider === 'hubspot') {
      if (!HUBSPOT_CLIENT_ID || !HUBSPOT_CLIENT_SECRET) {
        return res.status(400).send('HubSpot OAuth credentials are not configured.');
      }
      const payload = new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: HUBSPOT_CLIENT_ID,
        client_secret: HUBSPOT_CLIENT_SECRET,
        redirect_uri: CRM_REDIRECT_URI,
        code
      });
      const tokenResp = await axios.post('https://api.hubapi.com/oauth/v1/token', payload.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      });
      return res.json({
        accessToken: tokenResp.data.access_token,
        refreshToken: tokenResp.data.refresh_token,
        expiresIn: tokenResp.data.expires_in,
        accountLabel: tokenResp.data.hub_id ? `Hub ID ${tokenResp.data.hub_id}` : 'HubSpot workspace'
      });
    }

    if (provider === 'salesforce') {
      if (!SALESFORCE_CLIENT_ID || !SALESFORCE_CLIENT_SECRET) {
        return res.status(400).send('Salesforce OAuth credentials are not configured.');
      }
      const payload = new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: SALESFORCE_CLIENT_ID,
        client_secret: SALESFORCE_CLIENT_SECRET,
        redirect_uri: CRM_REDIRECT_URI,
        code
      });
      const tokenResp = await axios.post(
        `${SALESFORCE_LOGIN_URL}/services/oauth2/token`,
        payload.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      return res.json({
        accessToken: tokenResp.data.access_token,
        refreshToken: tokenResp.data.refresh_token,
        expiresIn: tokenResp.data.expires_in,
        accountLabel: tokenResp.data.instance_url || 'Salesforce org'
      });
    }

    return res.status(400).send('Unsupported CRM provider.');
  } catch (error) {
    console.error('[crm] OAuth exchange failed:', error?.response?.data || error?.message || error);
    return res.status(500).send('OAuth exchange failed.');
  }
});

app.post('/crm/:provider/sync', async (req, res) => {
  const provider = String(req.params.provider || '').toLowerCase();
  if (!['hubspot', 'salesforce'].includes(provider)) {
    return res.status(400).send('Unsupported CRM provider.');
  }

  const synced = 120 + Math.floor(Math.random() * 240);
  const updated = 30 + Math.floor(Math.random() * 60);
  const warnings = Math.random() > 0.75 ? Math.floor(Math.random() * 6) : 0;

  return res.json({ synced, updated, warnings });
});

app.post(['/credentials/hubspot/add', '/credentials/hubspot/add/'], async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({ error: 'Supabase keys are not configured.' });
    }

    const { owner_id, access_token, display_name } = req.body || {};
    if (!owner_id || !access_token) {
      return res.status(400).json({ error: 'owner_id and access_token are required.' });
    }

    const payload = {
      user_id: user.id,
      provider: 'hubspot',
      owner_id: String(owner_id).trim(),
      access_token: String(access_token).trim(),
      display_name: typeof display_name === 'string' ? display_name.trim() : null,
    };

    const data = await upsertCrmCredential(dbClient, payload);

    return res.status(201).json({
      status: 201,
      provider: 'hubspot',
      display_name: data?.display_name || payload.display_name,
      owner_id: data?.owner_id || payload.owner_id,
    });
  } catch (error) {
    console.error('[crm] HubSpot credential save failed:', error?.message || error);
    return res.status(500).json({ error: 'Failed to save HubSpot credential.' });
  }
});

app.post(['/salesforce/callback', '/salesforce/callback/'], async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({ error: 'Supabase keys are not configured.' });
    }

    const { code, SF_CLIENT_ID, SF_CLIENT_SECRET, SF_REDIRECT_URI, display_name } = req.body || {};
    if (!code || !SF_CLIENT_ID || !SF_CLIENT_SECRET || !SF_REDIRECT_URI) {
      return res.status(400).json({
        error: 'code, SF_CLIENT_ID, SF_CLIENT_SECRET, and SF_REDIRECT_URI are required.'
      });
    }

    const payload = new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: SF_CLIENT_ID,
      client_secret: SF_CLIENT_SECRET,
      redirect_uri: SF_REDIRECT_URI,
      code
    });

    const tokenResp = await axios.post(
      `${SALESFORCE_LOGIN_URL}/services/oauth2/token`,
      payload.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const accessToken = tokenResp.data?.access_token || '';
    const refreshToken = tokenResp.data?.refresh_token || '';
    const instanceUrl = tokenResp.data?.instance_url || '';

    const data = await upsertCrmCredential(dbClient, {
      user_id: user.id,
      provider: 'salesforce',
      access_token: accessToken,
      refresh_token: refreshToken,
      instance_url: instanceUrl,
      display_name: typeof display_name === 'string' ? display_name.trim() : null,
    });

    return res.status(200).json({
      status: 200,
      provider: 'salesforce',
      instance_url: data?.instance_url || instanceUrl,
      display_name: data?.display_name || display_name || 'Salesforce org',
    });
  } catch (error) {
    const details = error?.response?.data || error?.message || error;
    console.error('[crm] Salesforce callback failed:', details);
    return res.status(500).json({ error: 'Salesforce token exchange failed.' });
  }
});

app.get('/api/inbox/messages/:id/reply-draft', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const canBypassAuth = MAILBOX_BYPASS_AUTH && hasServiceRoleKey;
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({ error: 'Supabase keys are not configured.' });
    }

    const messageId = req.params.id;
    const message = await loadMessageForUser(dbClient, user, messageId, canBypassAuth);
    if (!message) {
      return res.status(404).json({ error: 'Message not found' });
    }

    const config = await loadConfigForUser(dbClient, user, message.config_id, canBypassAuth);
    if (!config) {
      return res.status(404).json({ error: 'Email configuration not found' });
    }

    const mode = req.query.mode === 'replyAll' ? 'replyAll' : 'reply';
    const draft = emailService.buildReplyDraft({
      original: message,
      mode,
      mailboxAddress: config.smtp_username,
    });

    const attachments = Array.isArray(message.attachments) ? message.attachments : [];
    const includeOriginalAttachmentsAvailable = REPLY_INCLUDE_ORIGINAL_ATTACHMENTS && attachments.length > 0;

    return res.json({
      ...draft,
      mode,
      includeOriginalAttachmentsAvailable,
      originalAttachments: includeOriginalAttachmentsAvailable ? attachments : [],
      threadingLimited: !message.message_id,
    });
  } catch (error) {
    console.error('[mailbox] Reply draft failed:', error?.message || error);
    return res.status(500).json({ error: 'Failed to build reply draft' });
  }
});

app.post('/api/inbox/messages/:id/reply', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const canBypassAuth = MAILBOX_BYPASS_AUTH && hasServiceRoleKey;
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({ error: 'Supabase keys are not configured.' });
    }

    const messageId = req.params.id;
    const message = await loadMessageForUser(dbClient, user, messageId, canBypassAuth);
    if (!message) {
      return res.status(404).json({ error: 'Message not found' });
    }

    const config = await loadConfigForUser(dbClient, user, message.config_id, canBypassAuth);
    if (!config) {
      return res.status(404).json({ error: 'Email configuration not found' });
    }

    const activeConfig = await ensureMailboxTransportConfig(
      config,
      dbClient,
      `inbox reply (${config?.smtp_username || config?.id || 'unknown'})`
    );

    if (!hasMailboxTransportCredentials(activeConfig)) {
      return res.status(400).json({ error: 'SMTP credentials are missing for this inbox.' });
    }

    const payload = req.body || {};
    const mode = payload.mode === 'replyAll' ? 'replyAll' : 'reply';

    if (!payload.text && !payload.html) {
      return res.status(400).json({ error: 'Reply body is required.' });
    }

    const creditReferenceId = `inbox-reply:${message.id}:${Date.now()}`;
    const creditResult = await consumeUserCredits(
      dbClient,
      user.id,
      OUTBOUND_EMAIL_CREDIT_COST,
      'inbox_reply_send',
      creditReferenceId,
      {
        source: 'inbox_reply',
        message_id: message.id,
        config_id: message.config_id,
      }
    );

    if (!creditResult.allowed) {
      return res.status(402).json({
        error: creditResult.message || 'Insufficient credits',
        creditsRemaining: creditResult.creditsRemaining,
      });
    }

    let result;
    try {
      result = await emailService.sendReply({
        config: activeConfig,
        original: message,
        payload: {
          ...payload,
          mode,
        },
      });
    } catch (sendError) {
      await refundUserCredits(
        dbClient,
        user.id,
        OUTBOUND_EMAIL_CREDIT_COST,
        'inbox_reply_refund',
        creditReferenceId,
        {
          source: 'inbox_reply',
          message_id: message.id,
          config_id: message.config_id,
          reason: 'send_failure',
          error: sendError?.message || String(sendError),
        }
      );
      throw sendError;
    }

    const sentAt = new Date().toISOString();
    const sentRow = {
      user_id: message.user_id,
      config_id: message.config_id,
      uid: null,
      from_email: activeConfig.smtp_username,
      to_email: result.to?.[0] || '',
      to_emails: result.to,
      cc_emails: result.cc,
      subject: result.subject,
      body: result.html || result.text || '',
      date: sentAt,
      folder: 'Sent',
      read: true,
      message_id: result.messageId,
      in_reply_to: result.inReplyTo,
      references: result.references,
      attachments: result.attachmentsMeta || [],
      thread_id: result.threadId,
      direction: 'outbound',
      reply_to_message_id: message.id,
    };

    const { error: insertError } = await dbClient.from('email_messages').insert(sentRow);
    if (insertError) {
      console.error('[mailbox] Failed to store sent reply:', insertError?.message || insertError);
    }

    return res.json({
      success: true,
      messageId: result.messageId,
      threadId: result.threadId,
      threadingLimited: !message.message_id,
    });
  } catch (error) {
    console.error('[mailbox] Reply send failed:', error?.message || error);
    return res.status(500).json({ error: 'Failed to send reply' });
  }
});

app.post('/api/inbox/compose', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const canBypassAuth = MAILBOX_BYPASS_AUTH && hasServiceRoleKey;
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({ error: 'Supabase keys are not configured.' });
    }

    const payload = req.body || {};
    const configId = payload.configId || payload.config_id;
    if (!configId) {
      return res.status(400).json({ error: 'configId is required.' });
    }

    const config = await loadConfigForUser(dbClient, user, configId, canBypassAuth);
    if (!config) {
      return res.status(404).json({ error: 'Email configuration not found' });
    }

    const activeConfig = await ensureMailboxTransportConfig(
      config,
      dbClient,
      `inbox compose (${config?.smtp_username || config?.id || 'unknown'})`
    );

    if (!hasMailboxTransportCredentials(activeConfig)) {
      return res.status(400).json({ error: 'SMTP credentials are missing for this inbox.' });
    }

    if (!payload.text && !payload.html) {
      return res.status(400).json({ error: 'Message body is required.' });
    }

    const creditReferenceId = `inbox-compose:${config.id}:${Date.now()}`;
    const creditResult = await consumeUserCredits(
      dbClient,
      user.id,
      OUTBOUND_EMAIL_CREDIT_COST,
      'inbox_compose_send',
      creditReferenceId,
      {
        source: 'inbox_compose',
        config_id: config.id,
        subject: payload.subject || null,
      }
    );

    if (!creditResult.allowed) {
      return res.status(402).json({
        error: creditResult.message || 'Insufficient credits',
        creditsRemaining: creditResult.creditsRemaining,
      });
    }

    let result;
    try {
      result = await emailService.sendCompose({
        config: activeConfig,
        payload: {
          subject: payload.subject,
          text: payload.text,
          html: payload.html,
          to: payload.to,
          cc: payload.cc,
          bcc: payload.bcc,
          attachments: payload.attachments,
        },
      });
    } catch (sendError) {
      await refundUserCredits(
        dbClient,
        user.id,
        OUTBOUND_EMAIL_CREDIT_COST,
        'inbox_compose_refund',
        creditReferenceId,
        {
          source: 'inbox_compose',
          config_id: config.id,
          reason: 'send_failure',
          error: sendError?.message || String(sendError),
        }
      );
      throw sendError;
    }

    const sentAt = new Date().toISOString();
    const sentRow = {
      user_id: config.user_id || user.id,
      config_id: config.id,
      uid: null,
      from_email: activeConfig.smtp_username,
      to_email: result.to?.[0] || '',
      to_emails: result.to,
      cc_emails: result.cc,
      subject: result.subject,
      body: result.html || result.text || '',
      date: sentAt,
      folder: 'Sent',
      read: true,
      message_id: result.messageId,
      in_reply_to: null,
      references: null,
      attachments: result.attachmentsMeta || [],
      thread_id: result.threadId,
      direction: 'outbound',
    };

    const { error: insertError } = await dbClient.from('email_messages').insert(sentRow);
    if (insertError) {
      console.error('[mailbox] Failed to store sent compose message:', insertError?.message || insertError);
    }

    return res.json({
      success: true,
      messageId: result.messageId,
      threadId: result.threadId,
    });
  } catch (error) {
    console.error('[mailbox] Compose send failed:', error?.message || error);
    return res.status(500).json({ error: error?.message || 'Failed to send message' });
  }
});

app.post('/sync-mailbox', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const canBypassAuth = MAILBOX_BYPASS_AUTH && hasServiceRoleKey;
    const user = await authenticateRequest(authHeader);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const dbClient = supabaseAdmin || buildUserSupabaseClient(authHeader);
    if (!dbClient) {
      return res.status(500).json({
        error: 'Supabase keys are not configured. Set SUPABASE_SERVICE_ROLE_KEY or SUPABASE_ANON_KEY.',
      });
    }

    const body = req.body || {};
    const configId = body.configId || body.config_id || body.mailboxId || null;
    const limit = body.limit ?? 50;
    if (!configId) {
      return res.status(400).json({ error: 'configId is required' });
    }

    let config = null;
    try {
      config = await loadConfigForUser(dbClient, user, configId, canBypassAuth);
    } catch (configError) {
      console.error('[mailbox] Failed to load email configuration:', configError);
      return res.status(500).json({
        error: 'Failed to load email configuration',
        details: configError?.message || String(configError),
      });
    }

    if (!config) {
      return res.status(404).json({ error: 'Email configuration not found' });
    }

    if (!config.imap_host || !config.imap_port) {
      return res.status(400).json({ error: 'IMAP host/port missing on configuration' });
    }

    const stats = await syncMailbox(config, limit, dbClient);
    res.json({ success: true, ...stats });
  } catch (error) {
    console.error('Mailbox sync error', error);
    res.status(500).json({ error: error.message || 'Mailbox sync failed' });
  }
});

app.post('/check-email-replies', async (req, res) => {
  try {
    if (!matchesAdminSecret(req)) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    const body = req.body || {};
    const configId = body.config_id || body.configId || body.mailboxId || null;
    const lookbackValue = body.lookback_days ?? body.lookbackDays ?? DEFAULT_CHECK_LOOKBACK_DAYS;
    const useDbScan = parseBoolean(body.use_db_scan ?? body.useDbScan);
    const overrides = buildCheckOverrides(body);

    const payload = await runReplyCheck({
      configId,
      lookbackDays: lookbackValue,
      useDbScan,
      overrides,
      source: 'manual',
    });

    return res.json(payload);
  } catch (error) {
    console.error('[mailbox] Reply check error:', error);
    return res.status(500).json({ error: error?.message || 'Reply check failed' });
  }
});
app.listen(PORT, () => {
  console.log(`Mailbox sync server listening on port ${PORT}`);
  console.log('Allowed origins:', ALLOWED_ORIGINS.join(', ') || '*');
  startAutoMailboxSync();
  startAutoReplyChecks();
});
