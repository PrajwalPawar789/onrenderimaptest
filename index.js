import { ImapFlow } from "imapflow";

/**
 * ⚠️ TEMP CONFIG FOR TESTING
 * Replace these values locally.
 * Do NOT commit real credentials.
 */
const IMAP_CONFIG = {
  host: "imap.hostinger.com",   // or "imap.titan.email"
  port: 993,
  secure: true,
  user: "daniel.clark@theciovision.com",
  pass: "Prajwal@78910"
};

console.log("IMAP HOST:", IMAP_CONFIG.host);
console.log("IMAP PORT:", IMAP_CONFIG.port);

const client = new ImapFlow({
  host: IMAP_CONFIG.host,
  port: IMAP_CONFIG.port,
  secure: IMAP_CONFIG.secure,
  auth: {
    user: IMAP_CONFIG.user,
    pass: IMAP_CONFIG.pass
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
