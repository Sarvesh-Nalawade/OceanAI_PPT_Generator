import { useState } from "react";
import PdfViewer from "../components/PDFviewer";
import ErrorDisplay from "../components/ErrorDisplay";


type JsonResponse = {
  pdfUrl?: string;
  pptUrl?: string;
  [k: string]: unknown;
};

export default function Home(): JSX.Element {
  const [topic, setTopic] = useState("");
  const [loading, setLoading] = useState(false);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [pptUrl, setPptUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const BACKEND_URL = process.env.NEXT_PUBLIC_BACKEND_URL || "";
  const endpoint = BACKEND_URL ? `${BACKEND_URL}/generate` : "/api/generate";

  async function handleGenerate() {
    setError(null);
    setPdfUrl(null);
    setPptUrl(null);

    if (!topic.trim()) {
      setError("Please enter a topic.");
      return;
    }

    setLoading(true);
    try {
      const fd = new FormData();
      fd.append("topic", topic);

      const res = await fetch(endpoint, {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Server error: ${res.status} ${text}`);
      }

      const contentType = res.headers.get("content-type") || "";

      // JSON response with URLs
      if (contentType.includes("application/json")) {
        const json = (await res.json()) as JsonResponse;
        if (json.pdfUrl) setPdfUrl(json.pdfUrl);
        if (json.pptUrl) setPptUrl(json.pptUrl);
        if (!json.pdfUrl && !json.pptUrl) setError("Backend returned JSON but no file URLs were found.");
      } else {
        // Binary response (pdf or pptx)
        const buffer = await res.arrayBuffer();

        if (contentType.includes("application/pdf")) {
          const blob = new Blob([buffer], { type: "application/pdf" });
          setPdfUrl(URL.createObjectURL(blob));
        } else if (
          contentType.includes("presentation") ||
          contentType.includes("powerpoint") ||
          contentType.includes(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
          )
        ) {
          const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          });
          setPptUrl(URL.createObjectURL(blob));
          setError("Received PPTX. If you want a PDF preview, have the backend also return a PDF or pdfUrl.");
        } else {
          // fallback — make blob, offer download
          const blob = new Blob([buffer]);
          setPptUrl(URL.createObjectURL(blob));
          setError("Received unknown file type. Provided a download link.");
        }
      }
    } catch (err: any) {
      console.error(err);
      setError(err?.message ?? "Unknown error");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="container">
      <main className="card">
        <h1>OceanAI PPT Generator</h1>

        <label className="label">Topic</label>
        <input
          className="input"
          placeholder="Enter a topic..."
          value={topic}
          onChange={(e) => setTopic(e.target.value)}
        />

        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          <button className="btn" onClick={handleGenerate} disabled={loading}>
            {loading ? "Generating..." : "Generate PPT"}
          </button>
          <button
            className="btn ghost"
            onClick={() => {
              setTopic("");
              setPdfUrl(null);
              setPptUrl(null);
              setError(null);
            }}
          >
            Reset
          </button>
        </div>

        {error && <ErrorDisplay message={error} />}

        {pptUrl && (
          <div className="section">
            <a className="link" href={pptUrl} download={`presentation-${Date.now()}.pptx`}>
              Download PPTX
            </a>
          </div>
        )}

        {pdfUrl && (
          <div className="section">
            <h3>Preview PDF</h3>
            <PdfViewer pdfUrl={pdfUrl} />
            <div style={{ marginTop: 8 }}>
              <a className="link" href={pdfUrl} target="_blank" rel="noopener noreferrer">
                Open PDF in new tab
              </a>
              {" · "}
              <a className="link" href={pdfUrl} download={`presentation-${Date.now()}.pdf`}>
                Download PDF
              </a>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
