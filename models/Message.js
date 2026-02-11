import mongoose from 'mongoose';

const AttachmentSchema = new mongoose.Schema(
  {
    filename: { type: String },
    contentType: { type: String },
    size: { type: Number },
    disposition: { type: String },
    partId: { type: String },
    cid: { type: String },
    source: { type: String, enum: ['upload', 'original'] },
  },
  { _id: false }
);

const MessageSchema = new mongoose.Schema(
  {
    userId: { type: String, required: true, index: true },
    accountId: { type: String, required: true, index: true },
    threadId: { type: String, index: true },
    messageId: { type: String, index: true },
    inReplyTo: { type: String },
    references: { type: [String], default: [] },
    from: { type: String },
    to: { type: [String], default: [] },
    cc: { type: [String], default: [] },
    replyTo: { type: [String], default: [] },
    subject: { type: String },
    bodyHtml: { type: String },
    bodyText: { type: String },
    date: { type: Date },
    direction: { type: String, enum: ['inbound', 'outbound'], default: 'inbound' },
    folder: { type: String, default: 'INBOX' },
    attachments: { type: [AttachmentSchema], default: [] },
    replyToMessageId: { type: mongoose.Schema.Types.ObjectId, ref: 'Message' },
  },
  { timestamps: true }
);

MessageSchema.index({ accountId: 1, messageId: 1 });
MessageSchema.index({ threadId: 1, date: -1 });

export default mongoose.models.Message || mongoose.model('Message', MessageSchema);
