import type { VercelRequest, VercelResponse } from "@vercel/node";
import Anthropic from "@anthropic-ai/sdk";

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { subject, body, sender, senderEmail, replyTo } = req.body || {};

  if (!body && !subject) {
    return res.status(400).json({ error: "Email content required" });
  }

  const prompt = `You are an email security analyst. Analyze the following email for spam, phishing, or scam indicators. Your goal is accurate verdicts — avoid both false positives on legitimate business email and false negatives on real threats.

Email details:
- From display name: ${sender || "Unknown"}
- From email address: ${senderEmail || "Unknown"}
- Reply-To: ${replyTo || "(same as sender)"}
- Subject: ${subject || "(no subject)"}
- Body:
${(body || "(no body)").slice(0, 3000)}

Respond with ONLY valid JSON in this exact format, no extra text:
{
  "verdict": "SAFE",
  "confidence": 95,
  "summary": "One clear sentence explaining your verdict.",
  "flags": ["specific red flag 1", "specific red flag 2"]
}

Rules:
- verdict must be one of: SAFE, SUSPICIOUS, or SPAM
- confidence is 0-100
- flags should be an empty array [] if none are found
- Keep flags concise and specific
- Only flag something if it is genuinely suspicious in context — not just because it matches a surface-level pattern

Verdicts:
- SAFE: Legitimate email. Use this when the sender domain is established, content matches a normal business purpose, and any concerns are easily explained by normal business practice.
- SUSPICIOUS: Genuine uncertainty. Use this only when there are specific, concrete indicators that cannot be explained by normal business practice.
- SPAM: Clear spam or phishing with multiple strong indicators.

Important context to apply:
- Link protection/rewriting services (linkprotect.cudasvc.com, urldefense.com, safelinks.protection.outlook.com, proofpoint.com redirects, etc.) are legitimate email security tools used by businesses — do NOT flag these as suspicious redirects.
- Transactional emails (invoices, payment receipts, shipping notifications) from domains matching the sender's company name are normal business email — generic greetings like "Valued Customer" are common and not a red flag on their own.
- A physical address, phone number, and matching sender domain together are strong legitimacy signals.
- Business acquisitions and account transfers (e.g. "transferred from Company X") are normal and not indicators of spoofing.
- Weigh ALL available signals together. A single surface-level pattern match is not enough for SUSPICIOUS — require multiple concrete concerns.

Consider: sender domain legitimacy, urgency or pressure tactics, requests for credentials or money, suspicious links (excluding known link protection services), grammar/spelling issues, mismatched reply-to, spoofed display names.`;

  try {
    const message = await client.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 400,
      messages: [{ role: "user", content: prompt }],
    });

    const text = message.content[0].type === "text" ? message.content[0].text : "";

    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error("Invalid response format");

    const result = JSON.parse(jsonMatch[0]);

    if (!["SAFE", "SUSPICIOUS", "SPAM"].includes(result.verdict)) {
      throw new Error("Invalid verdict in response");
    }

    return res.status(200).json(result);
  } catch (err: any) {
    console.error("Analyze error:", err.message);
    return res.status(500).json({ error: "Analysis failed. Please try again." });
  }
}
