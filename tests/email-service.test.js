import test from 'node:test';
import assert from 'node:assert/strict';
import EmailService from '../services/email-service.js';

const baseConfig = {
  smtp_username: 'me@example.com',
  smtp_password: 'secret',
  smtp_host: 'smtp.example.com',
  smtp_port: 587,
  imap_host: 'imap.example.com',
  imap_port: 993,
  sender_name: 'Me',
  security: 'TLS',
};

test('computeRecipients uses reply-to for reply', () => {
  const service = new EmailService();
  const original = {
    reply_to: ['reply@example.com'],
    from_email: 'from@example.com',
    to_emails: ['to@example.com'],
  };

  const { to, cc } = service.computeRecipients(original, 'reply', 'me@example.com');

  assert.deepEqual(to, ['reply@example.com']);
  assert.deepEqual(cc, []);
});

test('computeRecipients replyAll dedupes and excludes mailbox address', () => {
  const service = new EmailService();
  const original = {
    from_email: 'sender@example.com',
    to_emails: ['me@example.com', 'Friend@example.com', 'friend@example.com'],
    cc_emails: ['cc@example.com', 'me@example.com'],
  };

  const { to, cc } = service.computeRecipients(original, 'replyAll', 'me@example.com');

  assert.deepEqual(to.sort(), ['sender@example.com', 'Friend@example.com'].sort());
  assert.deepEqual(cc, ['cc@example.com']);
});

test('computeRecipients replyAll respects reply-to', () => {
  const service = new EmailService();
  const original = {
    reply_to: ['support@example.com'],
    from_email: 'from@example.com',
    to_emails: ['team@example.com'],
    cc_emails: [],
  };

  const { to, cc } = service.computeRecipients(original, 'replyAll', 'me@example.com');

  assert.deepEqual(to.sort(), ['support@example.com', 'team@example.com'].sort());
  assert.deepEqual(cc, []);
});

test('buildThreadHeaders truncates references chain', () => {
  const service = new EmailService({ referencesLimit: 3 });
  const original = {
    message_id: '<msg-4@example.com>',
    references: ['<msg-1@example.com>', '<msg-2@example.com>', '<msg-3@example.com>'],
  };

  const { references } = service.buildThreadHeaders(original);

  assert.deepEqual(references, ['<msg-2@example.com>', '<msg-3@example.com>', '<msg-4@example.com>']);
});

test('sendReply uses injected transport', async () => {
  const sent = [];
  const transportFactory = () => ({
    sendMail: async (options) => {
      sent.push(options);
      return { messageId: '<test-message-id@example.com>' };
    },
  });

  const service = new EmailService({ transportFactory });
  const result = await service.sendReply({
    config: baseConfig,
    original: {
      from_email: 'sender@example.com',
      subject: 'Hello',
      body: 'Original content',
      date: new Date().toISOString(),
      message_id: '<original@example.com>',
    },
    payload: {
      mode: 'reply',
      text: 'Thanks!\n\n> Original content',
      html: '<p>Thanks!</p>',
    },
  });

  assert.equal(result.messageId, '<test-message-id@example.com>');
  assert.equal(sent.length, 1);
  assert.ok(sent[0].headers['In-Reply-To']);
});
