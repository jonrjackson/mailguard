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

  const prompt = `You are an email security analyst. Analyze the following email for spam, phishing, or scam indicators.

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
- Keep flags concise and specific (e.g. "Reply-To differs from sender domain")

Consider: sender domain legitimacy, urgency or pressure tactics, requests for credentials or money, suspicious links, grammar/spelling issues, mismatched reply-to, spoofed display names.`;

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
