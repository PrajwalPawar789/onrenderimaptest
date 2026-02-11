import mongoose from 'mongoose';

const ThreadSchema = new mongoose.Schema(
  {
    userId: { type: String, required: true, index: true },
    accountId: { type: String, required: true, index: true },
    threadId: { type: String, required: true, index: true },
    subject: { type: String },
    participants: { type: [String], default: [] },
    lastMessageAt: { type: Date },
    messageCount: { type: Number, default: 0 },
  },
  { timestamps: true }
);

ThreadSchema.index({ userId: 1, accountId: 1, threadId: 1 }, { unique: true });

export default mongoose.models.Thread || mongoose.model('Thread', ThreadSchema);
