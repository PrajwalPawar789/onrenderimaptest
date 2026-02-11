import crypto from 'crypto';
import nodemailer from 'nodemailer';
import { ImapFlow } from 'imapflow';

const DEFAULT_REFERENCES_LIMIT = 20;
const DEFAULT_REFERENCES_CHAR_LIMIT = 1900;
const DEFAULT_NOREPLY_REGEX = /(^|[+._-])no-?reply|do-?not-?reply/i;
const HTML_TAG_REGEX = /<\s*(html|head|body|div|p|br|table|tbody|tr|td|th|span|img|a|style|meta|link|!doctype)\b/i;

const normalizeText = (value) => (value || '').toString().trim();

export const looksLikeHtml = (value) => HTML_TAG_REGEX.test(value || '');

export const stripHtml = (value) => {
  if (!value) return '';
  return value
    .replace(/<\s*br\s*\/?>/gi, '\n')
    .replace(/<\s*\/p\s*>/gi, '\n')
    .replace(/<\s*\/div\s*>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n\n')
    .trim();
};

export const escapeHtml = (value) =>
  (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

export const normalizeAddressList = (value) => {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value
      .flatMap((entry) => {
        if (!entry) return [];
        if (typeof entry === 'string') return [entry];
        if (typeof entry === 'object' && entry.address) return [entry.address];
        return [];
      })
      .map((entry) => normalizeText(entry))
      .filter(Boolean);
  }
  if (typeof value === 'string') {
    return value
      .split(/[,;]+/)
      .map((entry) => normalizeText(entry))
      .filter(Boolean);
  }
  return [];
};

export const normalizeEmail = (value) => normalizeText(value).toLowerCase();

export const dedupeAddresses = (addresses) => {
  const seen = new Map();
  for (const address of addresses) {
    const normalized = normalizeEmail(address);
    if (!normalized) continue;
    if (!seen.has(normalized)) {
      seen.set(normalized, address);
    }
  }
  return Array.from(seen.values());
};

export const normalizeMessageId = (value) => {
  const raw = normalizeText(value);
  if (!raw) return '';
  const cleaned = raw.replace(/[<>]/g, '').trim();
  if (!cleaned) return '';
  return `<${cleaned}>`;
};

export const normalizeReferences = (value) => {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeMessageId(entry))
      .filter(Boolean);
  }
  if (typeof value === 'string') {
    return value
      .split(/\s+/)
      .map((entry) => normalizeMessageId(entry))
      .filter(Boolean);
  }
  return [];
};

export const truncateReferences = (references, limit = DEFAULT_REFERENCES_LIMIT, charLimit = DEFAULT_REFERENCES_CHAR_LIMIT) => {
  const trimmed = references.slice(-limit);
  let result = [...trimmed];
  while (result.length > 1 && result.join(' ').length > charLimit) {
    result.shift();
  }
  return result;
};

export const buildReplySubject = (subject) => {
  const base = normalizeText(subject) || '(No Subject)';
  return /^re:/i.test(base) ? base : `Re: ${base}`;
};

export const buildQuotedText = (message) => {
  const body = message?.body || '';
  const plain = looksLikeHtml(body) ? stripHtml(body) : body;
  const clipped = plain.trim();
  const quoted = clipped
    ? clipped
        .split(/\r?\n/)
        .map((line) => `> ${line}`)
        .join('\n')
    : '';
  const dateLabel = message?.date ? new Date(message.date).toLocaleString() : '';
  const fromLabel = message?.from_email || message?.from || 'someone';
  if (!quoted) {
    return `\n\nOn ${dateLabel}, ${fromLabel} wrote:`;
  }
  return `\n\nOn ${dateLabel}, ${fromLabel} wrote:\n${quoted}`;
};

export const buildQuotedHtml = (message) => {
  const body = message?.body || '';
  const dateLabel = message?.date ? new Date(message.date).toLocaleString() : '';
  const fromLabel = escapeHtml(message?.from_email || message?.from || 'someone');
  const quoteBody = looksLikeHtml(body)
    ? body
    : escapeHtml(body).replace(/\r?\n/g, '<br />');
  return `\n<br />\n<br />\n<div style="color:#64748b;font-size:12px;">On ${escapeHtml(dateLabel)}, ${fromLabel} wrote:</div>\n<blockquote style="margin:8px 0 0 12px;padding-left:12px;border-left:2px solid #e2e8f0;">${quoteBody}</blockquote>`;
};

export const computeThreadKey = (message) => {
  const references = normalizeReferences(message?.references);
  const root = references[0] || normalizeMessageId(message?.in_reply_to) || normalizeMessageId(message?.message_id);
  if (root) return root;
  const subject = normalizeText(message?.subject).replace(/^(re|fw|fwd)\s*:/gi, '').trim().toLowerCase();
  const from = normalizeEmail(message?.from_email || message?.from || '');
  const bucket = message?.date ? new Date(message.date).toISOString().slice(0, 10) : 'unknown';
  const seed = [subject, from, bucket].filter(Boolean).join('|') || crypto.randomUUID();
  return `fallback:${crypto.createHash('sha1').update(seed).digest('hex')}`;
};

export const extractAttachmentMetadata = (bodyStructure, parsedAttachments = []) => {
  if (!bodyStructure) return [];

  const parsedByFilename = new Map();
  const parsedByCid = new Map();

  (parsedAttachments || []).forEach((attachment) => {
    if (!attachment) return;
    if (attachment.filename) parsedByFilename.set(attachment.filename, attachment);
    if (attachment.cid) parsedByCid.set(attachment.cid, attachment);
    if (attachment.contentId) parsedByCid.set(attachment.contentId, attachment);
  });

  const results = [];

  const walk = (node) => {
    if (!node) return;
    if (Array.isArray(node.childNodes)) {
      node.childNodes.forEach(walk);
    }

    const filename =
      node?.dispositionParameters?.filename ||
      node?.parameters?.filename ||
      node?.parameters?.name ||
      '';
    const disposition = normalizeText(node?.disposition).toLowerCase();
    const isMultipart = normalizeText(node?.type).toLowerCase().startsWith('multipart/');
    const isAttachment = disposition === 'attachment' || (disposition === 'inline' && filename);

    if (!node.part || isMultipart || (!isAttachment && !filename)) {
      return;
    }

    const parsed = parsedByFilename.get(filename) || parsedByCid.get(node?.id);

    results.push({
      partId: node.part,
      filename: filename || parsed?.filename || 'attachment',
      contentType: node?.type || parsed?.contentType || 'application/octet-stream',
      size: node?.size || parsed?.size || null,
      disposition: disposition || parsed?.contentDisposition || 'attachment',
      cid: node?.id || parsed?.cid || parsed?.contentId || null,
    });
  };

  walk(bodyStructure);
  return results;
};

const normalizeSecurity = (security) => {
  const value = normalizeText(security).toUpperCase();
  return value === 'TLS' ? 'TLS' : 'SSL';
};

const buildImapClient = (config) => {
  const security = normalizeSecurity(config.security);
  const secure = security === 'SSL';
  const port = Number(config.imap_port) || (secure ? 993 : 143);

  return new ImapFlow({
    host: config.imap_host,
    port,
    secure,
    doSTARTTLS: security === 'TLS',
    auth: {
      user: config.smtp_username,
      pass: config.smtp_password,
    },
    tls: { rejectUnauthorized: false },
    logger: false,
    disableAutoIdle: true,
  });
};

export class EmailService {
  constructor(options = {}) {
    this.referencesLimit = Number.isFinite(options.referencesLimit)
      ? Math.max(5, Math.min(100, options.referencesLimit))
      : DEFAULT_REFERENCES_LIMIT;
    this.referencesCharLimit = Number.isFinite(options.referencesCharLimit)
      ? Math.max(200, options.referencesCharLimit)
      : DEFAULT_REFERENCES_CHAR_LIMIT;
    this.blockNoreply = Boolean(options.blockNoreply);
    this.noreplyRegex = options.noreplyRegex instanceof RegExp ? options.noreplyRegex : DEFAULT_NOREPLY_REGEX;
    this.includeOriginalAttachments = Boolean(options.includeOriginalAttachments);
    this.transportFactory = options.transportFactory;
  }

  computeRecipients(original, mode, mailboxAddress) {
    const replyTo = normalizeAddressList(original?.reply_to || original?.reply_to_emails);
    const from = normalizeAddressList(original?.from_email || original?.from || original?.from_emails);

    const replyTargets = replyTo.length > 0 ? replyTo : from;

    let to = replyTargets;
    let cc = [];

    if (mode === 'replyAll') {
      const toCandidates = normalizeAddressList(original?.to_emails || original?.to_email || original?.to);
      const ccCandidates = normalizeAddressList(original?.cc_emails || original?.cc || original?.cc_email);
      to = replyTargets.concat(toCandidates);
      cc = ccCandidates;
    }

    const filteredTo = this.filterRecipients(to, mailboxAddress);
    const filteredCc = this.filterRecipients(cc, mailboxAddress);

    const dedupedTo = dedupeAddresses(filteredTo);
    const toSet = new Set(dedupedTo.map((addr) => normalizeEmail(addr)));
    const dedupedCc = dedupeAddresses(filteredCc).filter((addr) => !toSet.has(normalizeEmail(addr)));

    return { to: dedupedTo, cc: dedupedCc };
  }

  filterRecipients(addresses, mailboxAddress) {
    const mailbox = normalizeEmail(mailboxAddress);
    return (addresses || [])
      .filter((address) => {
        const normalized = normalizeEmail(address);
        if (!normalized) return false;
        if (mailbox && normalized === mailbox) return false;
        if (this.blockNoreply && this.noreplyRegex.test(normalized)) return false;
        return true;
      });
  }

  buildThreadHeaders(original) {
    const originalMessageId = normalizeMessageId(original?.message_id || original?.messageId);
    const references = normalizeReferences(original?.references);
    let chain = [...references];

    if (originalMessageId && !chain.includes(originalMessageId)) {
      chain.push(originalMessageId);
    }

    if (!originalMessageId && original?.in_reply_to) {
      const fallback = normalizeMessageId(original.in_reply_to);
      if (fallback && !chain.includes(fallback)) {
        chain.push(fallback);
      }
    }

    chain = dedupeAddresses(chain);
    chain = truncateReferences(chain, this.referencesLimit, this.referencesCharLimit);

    return {
      inReplyTo: originalMessageId || null,
      references: chain,
      threadId: computeThreadKey({
        message_id: originalMessageId,
        references: chain,
        in_reply_to: original?.in_reply_to,
        subject: original?.subject,
        from_email: original?.from_email,
        date: original?.date,
      }),
    };
  }

  buildReplyDraft({ original, mode, mailboxAddress }) {
    const recipients = this.computeRecipients(original, mode, mailboxAddress);
    return {
      to: recipients.to,
      cc: recipients.cc,
      subject: buildReplySubject(original?.subject),
      text: buildQuotedText(original),
      html: buildQuotedHtml(original),
    };
  }

  normalizeBody(text, html) {
    let normalizedText = text ? text.toString() : '';
    let normalizedHtml = html ? html.toString() : '';

    if (!normalizedText && normalizedHtml) {
      normalizedText = stripHtml(normalizedHtml);
    }

    if (!normalizedHtml && normalizedText) {
      normalizedHtml = escapeHtml(normalizedText).replace(/\r?\n/g, '<br />');
    }

    return { text: normalizedText, html: normalizedHtml };
  }

  buildReplyMessage({ config, original, payload }) {
    const mode = payload?.mode === 'replyAll' ? 'replyAll' : 'reply';
    const draft = this.buildReplyDraft({
      original,
      mode,
      mailboxAddress: config.smtp_username,
    });

    const to = draft.to;
    const ccCandidates = payload?.ccOverride ? normalizeAddressList(payload.ccOverride) : draft.cc;
    let cc = dedupeAddresses(this.filterRecipients(ccCandidates, config.smtp_username));
    const toSet = new Set(to.map((addr) => normalizeEmail(addr)));
    cc = cc.filter((addr) => !toSet.has(normalizeEmail(addr)));

    const bccCandidates = payload?.bcc ? normalizeAddressList(payload.bcc) : [];
    let bcc = dedupeAddresses(this.filterRecipients(bccCandidates, config.smtp_username));
    const ccSet = new Set(cc.map((addr) => normalizeEmail(addr)));
    bcc = bcc.filter((addr) => !toSet.has(normalizeEmail(addr)) && !ccSet.has(normalizeEmail(addr)));

    if (!to.length) {
      throw new Error('No reply recipients available.');
    }

    const { text, html } = this.normalizeBody(payload?.text || draft.text, payload?.html || draft.html);
    const threadHeaders = this.buildThreadHeaders(original);

    const fromName = normalizeText(config.sender_name);
    const from = fromName ? `${fromName} <${config.smtp_username}>` : config.smtp_username;

    const headers = {};
    if (threadHeaders.inReplyTo) {
      headers['In-Reply-To'] = threadHeaders.inReplyTo;
    }
    if (threadHeaders.references.length) {
      headers['References'] = threadHeaders.references.join(' ');
    }

    return {
      mailOptions: {
        from,
        to: to.join(', '),
        cc: cc.length ? cc.join(', ') : undefined,
        bcc: bcc.length ? bcc.join(', ') : undefined,
        subject: draft.subject,
        text,
        html,
        headers,
      },
      recipients: { to, cc, bcc },
      threadHeaders,
      subject: draft.subject,
      text,
      html,
    };
  }

  async sendReply({ config, original, payload }) {
    const { mailOptions, recipients, threadHeaders, subject, text, html } = this.buildReplyMessage({
      config,
      original,
      payload,
    });

    const attachments = await this.resolveAttachments({
      config,
      original,
      payload,
    });

    const transport = this.transportFactory
      ? this.transportFactory(config)
      : nodemailer.createTransport({
          host: config.smtp_host,
          port: Number(config.smtp_port) || 465,
          secure: normalizeSecurity(config.security) === 'SSL' || Number(config.smtp_port) === 465,
          requireTLS: normalizeSecurity(config.security) === 'TLS',
          auth: {
            user: config.smtp_username,
            pass: config.smtp_password,
          },
          tls: { rejectUnauthorized: false },
        });

    const info = await transport.sendMail({
      ...mailOptions,
      attachments,
    });

    return {
      info,
      to: recipients.to,
      cc: recipients.cc,
      bcc: recipients.bcc,
      subject,
      text,
      html,
      inReplyTo: threadHeaders.inReplyTo,
      references: threadHeaders.references,
      threadId: threadHeaders.threadId,
      messageId: info?.messageId || null,
      attachmentsMeta: this.buildOutgoingAttachmentMetadata(payload, attachments, original),
    };
  }

  async resolveAttachments({ config, original, payload }) {
    const attachments = [];

    const uploadAttachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
    for (const attachment of uploadAttachments) {
      if (!attachment?.content) continue;
      const raw = attachment.content.toString();
      const base64 = raw.includes('base64,') ? raw.split('base64,').pop() : raw;
      const content = Buffer.from(base64, 'base64');
      attachments.push({
        filename: attachment.filename || 'attachment',
        content,
        contentType: attachment.contentType || 'application/octet-stream',
        cid: attachment.cid || undefined,
        contentDisposition: attachment.disposition || 'attachment',
      });
    }

    if (this.includeOriginalAttachments && payload?.includeOriginalAttachments) {
      const originalAttachments = await this.fetchOriginalAttachments(config, original);
      attachments.push(...originalAttachments);
    }

    return attachments;
  }

  buildOutgoingAttachmentMetadata(payload, attachments, original) {
    const uploadAttachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
    const metadata = [];

    uploadAttachments.forEach((attachment) => {
      if (!attachment) return;
      metadata.push({
        filename: attachment.filename || 'attachment',
        contentType: attachment.contentType || 'application/octet-stream',
        size: attachment.size || null,
        disposition: attachment.disposition || 'attachment',
        source: 'upload',
      });
    });

    if (this.includeOriginalAttachments && payload?.includeOriginalAttachments) {
      const originalMetadata = Array.isArray(original?.attachments) ? original.attachments : [];
      originalMetadata.forEach((attachment) => {
        metadata.push({
          filename: attachment.filename || 'attachment',
          contentType: attachment.contentType || 'application/octet-stream',
          size: attachment.size || null,
          disposition: attachment.disposition || 'attachment',
          source: 'original',
          partId: attachment.partId,
        });
      });
    }

    if (metadata.length === 0 && attachments.length > 0) {
      attachments.forEach((attachment) => {
        metadata.push({
          filename: attachment.filename || 'attachment',
          contentType: attachment.contentType || 'application/octet-stream',
          size: attachment.content?.length || null,
          disposition: attachment.contentDisposition || 'attachment',
          source: 'upload',
        });
      });
    }

    return metadata;
  }

  async fetchOriginalAttachments(config, original) {
    const attachments = Array.isArray(original?.attachments) ? original.attachments : [];
    const partIds = attachments.map((att) => att?.partId).filter(Boolean);
    if (!partIds.length || !original?.uid) return [];

    const client = buildImapClient(config);
    let lock = null;

    try {
      await client.connect();
      lock = await client.getMailboxLock(original.folder || 'INBOX');
      const downloads = await client.downloadMany(original.uid, partIds, { uid: true });
      return partIds
        .map((partId) => {
          const entry = downloads?.[partId];
          if (!entry?.content) return null;
          const meta = entry.meta || {};
          const attachmentMeta = attachments.find((att) => att.partId === partId) || {};
          return {
            filename: attachmentMeta.filename || meta.filename || 'attachment',
            content: entry.content,
            contentType: attachmentMeta.contentType || meta.contentType || 'application/octet-stream',
            cid: attachmentMeta.cid || undefined,
            contentDisposition: attachmentMeta.disposition || meta.disposition || 'attachment',
          };
        })
        .filter(Boolean);
    } finally {
      try {
        lock?.release();
      } catch (error) {
        // ignore
      }
      try {
        await client.logout();
      } catch (error) {
        // ignore
      }
    }
  }
}

export default EmailService;
