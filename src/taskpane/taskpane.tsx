import * as React from "react";
import * as ReactDOM from "react-dom/client";

/* global Office */

interface AnalysisResult {
  verdict: "SAFE" | "SUSPICIOUS" | "SPAM";
  confidence: number;
  summary: string;
  flags: string[];
}

interface EmailAttachment {
  id: string;
  name: string;
}

const VERDICT_STYLES = {
  SAFE:       { color: "#107c10", bg: "#e8f5e9", icon: "✓", label: "Safe" },
  SUSPICIOUS: { color: "#b45309", bg: "#fffbeb", icon: "⚠", label: "Suspicious" },
  SPAM:       { color: "#b91c1c", bg: "#fef2f2", icon: "✗", label: "Spam / Phishing" },
};

function parseEml(emlText: string) {
  const result = { sender: "", senderEmail: "", subject: "", replyTo: "", body: "" };

  const splitIdx = emlText.indexOf("\n\n");
  if (splitIdx === -1) return result;

  const headerSection = emlText.substring(0, splitIdx);
  const bodySection   = emlText.substring(splitIdx + 2);

  // Parse headers, handling folded (multi-line) values
  const headers: Record<string, string> = {};
  let curKey = "";
  let curVal = "";
  for (const line of headerSection.split("\n")) {
    if (/^\s+/.test(line) && curKey) {
      curVal += " " + line.trim();
    } else {
      if (curKey) headers[curKey.toLowerCase()] = curVal;
      const ci = line.indexOf(":");
      if (ci > 0) { curKey = line.substring(0, ci).trim(); curVal = line.substring(ci + 1).trim(); }
      else { curKey = ""; curVal = ""; }
    }
  }
  if (curKey) headers[curKey.toLowerCase()] = curVal;

  // From
  const from = headers["from"] || "";
  const fromMatch = from.match(/^(.*?)\s*<(.+?)>$/);
  if (fromMatch) {
    result.sender      = fromMatch[1].trim().replace(/^"|"$/g, "");
    result.senderEmail = fromMatch[2].trim();
  } else {
    result.senderEmail = from.trim();
  }

  // Subject (strip encoded-word encoding like =?UTF-8?Q?...?= as best-effort)
  result.subject = (headers["subject"] || "").replace(/=\?[^?]+\?[BQ]\?[^?]+\?=/gi, "").trim();

  // Reply-To
  const rt = headers["reply-to"] || "";
  const rtMatch = rt.match(/<(.+?)>/);
  result.replyTo = rtMatch ? rtMatch[1] : rt.trim();

  // Body — extract text/plain from multipart, otherwise decode directly
  const ct = headers["content-type"] || "";
  if (ct.toLowerCase().includes("multipart")) {
    const bm = ct.match(/boundary="?([^";\s\r]+)"?/);
    if (bm) {
      const parts = bodySection.split("--" + bm[1]);
      for (const part of parts) {
        if (/content-type:\s*text\/plain/i.test(part)) {
          const ps = part.indexOf("\n\n");
          if (ps !== -1) { result.body = part.substring(ps + 2).trim(); break; }
        }
      }
    }
    if (!result.body) result.body = bodySection.substring(0, 3000);
  } else {
    const enc = (headers["content-transfer-encoding"] || "").toLowerCase();
    if (enc === "base64") {
      try { result.body = atob(bodySection.replace(/\s/g, "")); } catch { result.body = bodySection; }
    } else if (enc === "quoted-printable") {
      result.body = bodySection
        .replace(/=\r?\n/g, "")
        .replace(/=([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
    } else {
      result.body = bodySection;
    }
  }

  result.body = result.body.substring(0, 3000);
  return result;
}

function App() {
  const [loading, setLoading]             = React.useState(false);
  const [result, setResult]               = React.useState<AnalysisResult | null>(null);
  const [error, setError]                 = React.useState<string | null>(null);
  const [analyzedLabel, setAnalyzedLabel] = React.useState("This email");
  const [attachments, setAttachments]     = React.useState<EmailAttachment[]>([]);

  const callApi = async (data: object): Promise<AnalysisResult> => {
    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    if (!response.ok) throw new Error("Analysis service unavailable");
    return response.json();
  };

  const analyzeEmail = async () => {
    setLoading(true);
    setError(null);
    setResult(null);
    setAnalyzedLabel("This email");

    try {
      const item = (Office as any).context.mailbox.item;

      const sender      = item.from?.displayName  || "";
      const senderEmail = item.from?.emailAddress || "";
      const subject     = item.subject            || "";
      const replyTo     = item.replyTo?.length > 0 ? item.replyTo[0].emailAddress : "";

      // Detect email-type attachments
      setAttachments(
        (item.attachments || [])
          .filter((a: any) => a.attachmentType === "item")
          .map((a: any) => ({ id: a.id, name: a.name || "Attached email" }))
      );

      const body = await new Promise<string>((resolve, reject) => {
        item.body.getAsync(
          (Office as any).CoercionType.Text,
          (r: any) => r.status === (Office as any).AsyncResultStatus.Succeeded
            ? resolve(r.value)
            : reject(new Error("Could not read email body"))
        );
      });

      setResult(await callApi({ subject, body, sender, senderEmail, replyTo }));
    } catch (err: any) {
      setError(err.message || "Analysis failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const analyzeAttachment = async (att: EmailAttachment) => {
    setLoading(true);
    setError(null);
    setResult(null);
    setAnalyzedLabel(`Attached: "${att.name}"`);

    try {
      const item = (Office as any).context.mailbox.item;

      const { content, format } = await new Promise<{ content: string; format: string }>((resolve, reject) => {
        item.getAttachmentContentAsync(att.id, (r: any) =>
          r.status === (Office as any).AsyncResultStatus.Succeeded
            ? resolve({ content: r.value.content, format: r.value.format })
            : reject(new Error("Could not read attached email"))
        );
      });

      // Item attachments return EML text directly; file attachments return base64
      const emlText = (format === "base64" || format === "Base64")
        ? atob(content.replace(/\s/g, ""))
        : content;

      const parsed = parseEml(emlText);
      if (!parsed.senderEmail && !parsed.subject && !parsed.body) {
        throw new Error("Could not parse attached email content");
      }

      setResult(await callApi(parsed));
    } catch (err: any) {
      setError(err.message || "Could not analyze attached email.");
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => { analyzeEmail(); }, []);

  const vstyle = result ? VERDICT_STYLES[result.verdict] : null;

  return (
    <div style={{ padding: "16px", maxWidth: "380px", margin: "0 auto" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", marginBottom: "20px", borderBottom: "1px solid #e5e7eb", paddingBottom: "12px" }}>
        <span style={{ fontSize: "18px", fontWeight: 700, color: "#0078d4" }}>MailGuard</span>
        <span style={{ marginLeft: "8px", fontSize: "11px", color: "#6b7280", background: "#f3f4f6", padding: "2px 8px", borderRadius: "999px" }}>AI Spam Detector</span>
      </div>

      {/* What was analyzed */}
      {!loading && (result || error) && (
        <div style={{ fontSize: "11px", color: "#6b7280", marginBottom: "10px" }}>
          Analyzed: {analyzedLabel}
        </div>
      )}

      {/* Loading */}
      {loading && (
        <div style={{ textAlign: "center", padding: "40px 0", color: "#6b7280" }}>
          <div style={{ fontSize: "28px", marginBottom: "10px" }}>🔍</div>
          <div style={{ fontSize: "14px" }}>Analyzing {analyzedLabel.toLowerCase()}...</div>
        </div>
      )}

      {/* Error */}
      {error && !loading && (
        <div style={{ padding: "12px", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: "8px", color: "#b91c1c", fontSize: "13px", marginBottom: "12px" }}>
          {error}
        </div>
      )}

      {/* Result */}
      {result && !loading && vstyle && (
        <div>
          <div style={{ background: vstyle.bg, border: `1px solid ${vstyle.color}30`, borderRadius: "10px", padding: "20px", textAlign: "center", marginBottom: "16px" }}>
            <div style={{ fontSize: "36px", marginBottom: "4px" }}>{vstyle.icon}</div>
            <div style={{ fontSize: "20px", fontWeight: 700, color: vstyle.color }}>{vstyle.label}</div>
            <div style={{ fontSize: "12px", color: "#6b7280", marginTop: "4px" }}>{result.confidence}% confidence</div>
          </div>

          <div style={{ fontSize: "13px", lineHeight: "1.6", color: "#374151", marginBottom: "16px", padding: "12px", background: "#fff", borderRadius: "8px", border: "1px solid #e5e7eb" }}>
            {result.summary}
          </div>

          {result.flags?.length > 0 ? (
            <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: "8px", padding: "12px", marginBottom: "16px" }}>
              <div style={{ fontWeight: 600, fontSize: "12px", color: "#374151", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Red Flags</div>
              <ul style={{ paddingLeft: "18px", margin: 0 }}>
                {result.flags.map((flag, i) => (
                  <li key={i} style={{ fontSize: "13px", color: "#4b5563", marginBottom: "4px", lineHeight: "1.5" }}>{flag}</li>
                ))}
              </ul>
            </div>
          ) : (
            <div style={{ fontSize: "13px", color: "#6b7280", textAlign: "center", marginBottom: "16px" }}>
              No specific red flags detected.
            </div>
          )}
        </div>
      )}

      {/* Re-analyze button */}
      {!loading && (
        <button
          onClick={analyzeEmail}
          style={{ width: "100%", padding: "10px", background: "#0078d4", color: "#fff", border: "none", borderRadius: "6px", fontSize: "14px", fontFamily: "Segoe UI, sans-serif", cursor: "pointer", fontWeight: 500 }}
        >
          {result || error ? "Re-analyze this email" : "Analyze Email"}
        </button>
      )}

      {/* Attached emails section */}
      {attachments.length > 0 && !loading && (
        <div style={{ marginTop: "16px", padding: "12px", background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: "8px" }}>
          <div style={{ fontWeight: 600, fontSize: "12px", color: "#0369a1", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>
            Attached Email{attachments.length > 1 ? "s" : ""}
          </div>
          {attachments.map((att) => (
            <div key={att.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "6px" }}>
              <span style={{ fontSize: "12px", color: "#374151", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1, marginRight: "8px" }}>
                {att.name}
              </span>
              <button
                onClick={() => analyzeAttachment(att)}
                style={{ flexShrink: 0, padding: "4px 10px", background: "#0369a1", color: "#fff", border: "none", borderRadius: "4px", fontSize: "12px", cursor: "pointer", fontFamily: "Segoe UI, sans-serif" }}
              >
                Analyze
              </button>
            </div>
          ))}
        </div>
      )}

    </div>
  );
}

Office.onReady(() => {
  ReactDOM.createRoot(document.getElementById("root")!).render(<App />);
});
