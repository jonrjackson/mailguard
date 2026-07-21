import * as React from "react";
import * as ReactDOM from "react-dom/client";

/* global Office */

interface AnalysisResult {
  verdict: "SAFE" | "SUSPICIOUS" | "SPAM";
  confidence: number;
  summary: string;
  flags: string[];
}

const VERDICT_STYLES = {
  SAFE:       { color: "#107c10", bg: "#e8f5e9", icon: "✓", label: "Safe" },
  SUSPICIOUS: { color: "#b45309", bg: "#fffbeb", icon: "⚠", label: "Suspicious" },
  SPAM:       { color: "#b91c1c", bg: "#fef2f2", icon: "✗", label: "Spam / Phishing" },
};

function App() {
  const [loading, setLoading]   = React.useState(false);
  const [result, setResult]     = React.useState<AnalysisResult | null>(null);
  const [error, setError]       = React.useState<string | null>(null);

  const analyzeEmail = async () => {
    setLoading(true);
    setError(null);
    setResult(null);

    try {
      const item = Office.context.mailbox.item as any;

      const sender      = item.from?.displayName  || "";
      const senderEmail = item.from?.emailAddress || "";
      const subject     = item.subject            || "";
      const replyTo     = item.replyTo?.length > 0 ? item.replyTo[0].emailAddress : "";

      const body = await new Promise<string>((resolve, reject) => {
        item.body.getAsync(
          (Office as any).CoercionType.Text,
          (asyncResult: any) => {
            if (asyncResult.status === (Office as any).AsyncResultStatus.Succeeded) {
              resolve(asyncResult.value);
            } else {
              reject(new Error("Could not read email body"));
            }
          }
        );
      });

      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ subject, body, sender, senderEmail, replyTo }),
      });

      if (!response.ok) throw new Error("Analysis service unavailable");

      const data: AnalysisResult = await response.json();
      setResult(data);
    } catch (err: any) {
      setError(err.message || "Analysis failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  // Auto-analyze when the task pane opens
  React.useEffect(() => { analyzeEmail(); }, []);

  const styles = result ? VERDICT_STYLES[result.verdict] : null;

  return (
    <div style={{ padding: "16px", maxWidth: "380px", margin: "0 auto" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", marginBottom: "20px", borderBottom: "1px solid #e5e7eb", paddingBottom: "12px" }}>
        <span style={{ fontSize: "18px", fontWeight: 700, color: "#0078d4" }}>MailGuard</span>
        <span style={{ marginLeft: "8px", fontSize: "11px", color: "#6b7280", background: "#f3f4f6", padding: "2px 8px", borderRadius: "999px" }}>AI Spam Detector</span>
      </div>

      {/* Loading state */}
      {loading && (
        <div style={{ textAlign: "center", padding: "40px 0", color: "#6b7280" }}>
          <div style={{ fontSize: "28px", marginBottom: "10px" }}>🔍</div>
          <div style={{ fontSize: "14px" }}>Analyzing this email...</div>
        </div>
      )}

      {/* Error state */}
      {error && !loading && (
        <div style={{ padding: "12px", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: "8px", color: "#b91c1c", fontSize: "13px", marginBottom: "12px" }}>
          {error}
        </div>
      )}

      {/* Result */}
      {result && !loading && styles && (
        <div>
          {/* Verdict card */}
          <div style={{ background: styles.bg, border: `1px solid ${styles.color}30`, borderRadius: "10px", padding: "20px", textAlign: "center", marginBottom: "16px" }}>
            <div style={{ fontSize: "36px", marginBottom: "4px" }}>{styles.icon}</div>
            <div style={{ fontSize: "20px", fontWeight: 700, color: styles.color }}>{styles.label}</div>
            <div style={{ fontSize: "12px", color: "#6b7280", marginTop: "4px" }}>{result.confidence}% confidence</div>
          </div>

          {/* Summary */}
          <div style={{ fontSize: "13px", lineHeight: "1.6", color: "#374151", marginBottom: "16px", padding: "12px", background: "#fff", borderRadius: "8px", border: "1px solid #e5e7eb" }}>
            {result.summary}
          </div>

          {/* Red flags */}
          {result.flags && result.flags.length > 0 && (
            <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: "8px", padding: "12px", marginBottom: "16px" }}>
              <div style={{ fontWeight: 600, fontSize: "12px", color: "#374151", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                Red Flags
              </div>
              <ul style={{ paddingLeft: "18px", margin: 0 }}>
                {result.flags.map((flag, i) => (
                  <li key={i} style={{ fontSize: "13px", color: "#4b5563", marginBottom: "4px", lineHeight: "1.5" }}>
                    {flag}
                  </li>
                ))}
              </ul>
            </div>
          )}

          {result.flags && result.flags.length === 0 && (
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
          style={{
            width: "100%",
            padding: "10px",
            background: "#0078d4",
            color: "#fff",
            border: "none",
            borderRadius: "6px",
            fontSize: "14px",
            fontFamily: "Segoe UI, sans-serif",
            cursor: "pointer",
            fontWeight: 500,
          }}
        >
          {result || error ? "Re-analyze" : "Analyze Email"}
        </button>
      )}
    </div>
  );
}

Office.onReady(() => {
  const container = document.getElementById("root")!;
  ReactDOM.createRoot(container).render(<App />);
});
