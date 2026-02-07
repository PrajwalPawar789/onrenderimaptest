import { ImapFlow } from "imapflow";

const client = new ImapFlow({
  host: process.env.IMAP_HOST,
  port: Number(process.env.IMAP_PORT || 993),
  secure: true,
  auth: {
    user: process.env.IMAP_USER,
    pass: process.env.IMAP_PASS
  }
});

(async () => {
  await client.connect();
  console.log("✅ Connected to IMAP");

  const boxes = await client.list();
  console.log("📬 Mailboxes:", boxes.map(b => b.path));

  await client.logout();
  console.log("👋 Logged out");
  process.exit(0);
})().catch(err => {
  console.error("❌ IMAP failed:", err);
  process.exit(1);
});
