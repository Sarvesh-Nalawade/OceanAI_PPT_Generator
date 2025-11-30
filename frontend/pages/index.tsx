import { useState, useEffect, useRef } from "react";
import PdfViewer from "../components/PDFviewer";
import ErrorDisplay from "../components/ErrorDisplay";

type Message = {
  role: "user" | "assistant";
  content: string;
};

type BackendResponse = {
  status: boolean;
  content: string;
  file_path?: string;
  pdfUrl?: string;
  pptUrl?: string;
};

export default function Home(): JSX.Element {
  const [messages, setMessages] = useState<Message[]>([]);
  const [currentInput, setCurrentInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [pptUrl, setPptUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isComplete, setIsComplete] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const BACKEND_URL = process.env.NEXT_PUBLIC_BACKEND_URL || "/api";

  // Auto-scroll to bottom when messages change
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  async function handleSendMessage() {
    if (!currentInput.trim()) {
      setError("Please enter a message.");
      return;
    }

    const userMessage: Message = {
      role: "user",
      content: currentInput,
    };

    setMessages((prev) => [...prev, userMessage]);
    const userInput = currentInput; // Store before clearing
    setCurrentInput("");
    setLoading(true);
    setError(null);

    // Don't add generating message - we'll show loading spinner instead

    try {
      const fd = new FormData();
      fd.append("topic", userInput);

      const res = await fetch(`${BACKEND_URL}/generate`, {
        method: "POST",
        body: fd,
      });

      const contentType = res.headers.get("content-type") || "";

      // Check if response is a file (PPTX) - when ppt_generated is True
      if (
        contentType.includes("presentation") ||
        contentType.includes("powerpoint") ||
        contentType.includes("application/vnd.openxmlformats-officedocument.presentationml.presentation")
      ) {
        const successMessage: Message = {
          role: "assistant",
          content: "✅ Your presentation has been generated successfully! You can download it below.",
        };
        setMessages((prev) => [...prev, successMessage]);

        // Convert response to blob - use response.blob() instead of arrayBuffer for better compatibility
        const blob = await res.blob();
        
        // Verify blob size
        console.log("PPTX file size:", blob.size, "bytes");
        if (blob.size === 0) {
          throw new Error("Received empty file from server");
        }
        
        setPptUrl(URL.createObjectURL(blob));
        setIsComplete(true);
      } else if (contentType.includes("application/json")) {
        // Handle JSON response - when status is false (follow-up questions)
        const data: BackendResponse = await res.json();
        
        if (!data.status) {
          // This is a follow-up question from the agent
          const assistantMessage: Message = {
            role: "assistant",
            content: data.content,
          };
          setMessages((prev) => [...prev, assistantMessage]);
          // Don't set isComplete - user can continue the conversation
        } else {
          // Status is true but we got JSON instead of file (shouldn't happen normally)
          throw new Error("Unexpected response: status is true but no file received");
        }
      } else if (!res.ok) {
        // Handle HTTP error responses
        const text = await res.text();
        let errorMessage = "";
        
        try {
          const errorData = JSON.parse(text);
          errorMessage = errorData.detail || "Something went wrong. Please try again.";
        } catch {
          errorMessage = text || "Something went wrong. Please try again.";
        }

        const assistantMessage: Message = {
          role: "assistant",
          content: errorMessage,
        };
        setMessages((prev) => [...prev, assistantMessage]);
      } else {
        // Unexpected response type
        throw new Error("Unexpected response type from server");
      }
    } catch (err: any) {
      console.error(err);
      setError(err?.message ?? "Unknown error");
    } finally {
      setLoading(false);
    }
  }

  function handleReset() {
    setMessages([]);
    setCurrentInput("");
    setPdfUrl(null);
    setPptUrl(null);
    setError(null);
    setIsComplete(false);
  }

  return (
    <div className="container">
      <main className="card">
        <h1 style={{ marginBottom: "1.5rem", fontSize: "2rem", fontWeight: "700" }}>
          🌊 OceanAI PPT Generator
        </h1>

        {/* Chat Messages */}
        <div 
          className="section" 
          style={{ 
            maxHeight: "550px", 
            overflowY: "auto", 
            marginBottom: "1rem",
            backgroundColor: "#343541",
            borderRadius: "12px",
            position: "relative",
          }}
        >
          {messages.length === 0 ? (
            <div style={{ 
              color: "#9ca3af", 
              textAlign: "center", 
              padding: "4rem 2rem",
              fontSize: "1.05rem"
            }}>
              <div style={{ fontSize: "3rem", marginBottom: "1rem" }}>💬</div>
              <p style={{ fontWeight: "500", marginBottom: "0.5rem" }}>Welcome to OceanAI PPT Generator</p>
              <p style={{ fontSize: "0.9rem", color: "#6b7280" }}>
                Enter a topic to generate your presentation
              </p>
            </div>
          ) : (
            <>
              {messages.map((msg, idx) => (
                <div
                  key={idx}
                  style={{
                    padding: "1.5rem 1.25rem",
                    backgroundColor: msg.role === "user" ? "#343541" : "#444654",
                    borderBottom: idx < messages.length - 1 ? "1px solid #4a4a5e" : "none",
                  }}
                >
                  <div style={{ 
                    maxWidth: "48rem",
                    margin: "0 auto",
                    display: "flex",
                    gap: "1.5rem",
                    alignItems: "flex-start"
                  }}>
                    {/* Avatar */}
                    <div style={{
                      width: "32px",
                      height: "32px",
                      borderRadius: "6px",
                      backgroundColor: msg.role === "user" ? "#5b21b6" : "#059669",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      flexShrink: 0,
                      fontSize: "0.875rem",
                      fontWeight: "700",
                      color: "#ffffff",
                      boxShadow: "0 2px 4px rgba(0,0,0,0.2)"
                    }}>
                      {msg.role === "user" ? "👤" : "🤖"}
                    </div>
                    
                    {/* Message Content */}
                    <div style={{ flex: 1 }}>
                      <div style={{ 
                        fontSize: "0.8rem", 
                        fontWeight: "600",
                        marginBottom: "0.5rem",
                        color: msg.role === "user" ? "#a78bfa" : "#34d399",
                        letterSpacing: "0.3px"
                      }}>
                        {msg.role === "user" ? "Your Response" : "OceanAI PPT Generator"}
                      </div>
                      <div style={{ 
                        color: "#ececf1",
                        lineHeight: "1.75",
                        fontSize: "0.98rem",
                        whiteSpace: "pre-wrap"
                      }}>
                        {msg.content}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
              
              {/* Loading Indicator */}
              {loading && (
                <div
                  style={{
                    padding: "1.5rem 1.25rem",
                    backgroundColor: "#444654",
                  }}
                >
                  <div style={{ 
                    maxWidth: "48rem",
                    margin: "0 auto",
                    display: "flex",
                    gap: "1.5rem",
                    alignItems: "flex-start"
                  }}>
                    <div style={{
                      width: "32px",
                      height: "32px",
                      borderRadius: "6px",
                      backgroundColor: "#059669",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      flexShrink: 0,
                      fontSize: "0.875rem",
                      fontWeight: "700",
                      color: "#ffffff",
                      boxShadow: "0 2px 4px rgba(0,0,0,0.2)"
                    }}>
                      🤖
                    </div>
                    
                    <div style={{ flex: 1 }}>
                      <div style={{ 
                        fontSize: "0.8rem", 
                        fontWeight: "600",
                        marginBottom: "0.5rem",
                        color: "#34d399",
                        letterSpacing: "0.3px"
                      }}>
                        OceanAI PPT Generator
                      </div>
                      <div style={{ 
                        color: "#ececf1",
                        lineHeight: "1.75",
                        fontSize: "0.98rem",
                        display: "flex",
                        alignItems: "center",
                        gap: "0.75rem"
                      }}>
                        <div style={{
                          width: "20px",
                          height: "20px",
                          border: "3px solid #059669",
                          borderTopColor: "transparent",
                          borderRadius: "50%",
                          animation: "spin 1s linear infinite"
                        }} />
                        <span>Generating your presentation...</span>
                      </div>
                    </div>
                  </div>
                </div>
              )}
              <div ref={messagesEndRef} />
            </>
          )}
        </div>

        {error && <ErrorDisplay message={error} />}

        {/* Input Area */}
        {!isComplete && (
          <>
            <label className="label">
              {messages.length === 0 ? "Topic" : "Your response"}
            </label>
            <input
              className="input"
              placeholder={messages.length === 0 ? "Enter a topic..." : "Type your answer..."}
              value={currentInput}
              onChange={(e) => setCurrentInput(e.target.value)}
              onKeyPress={(e) => {
                if (e.key === "Enter" && !loading) {
                  handleSendMessage();
                }
              }}
            />

            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button className="btn" onClick={handleSendMessage} disabled={loading}>
                {loading ? "Sending..." : "Send"}
              </button>
              <button className="btn ghost" onClick={handleReset}>
                Reset
              </button>
            </div>
          </>
        )}

        {/* Download Section */}
        {isComplete && (
          <div className="section" style={{ marginTop: "1rem" }}>
            <h3 style={{ color: "#388e3c" }}>✓ Presentation Ready!</h3>
            
            {pptUrl && (
              <div style={{ marginBottom: "1rem" }}>
                <a className="link" href={pptUrl} download={`presentation-${Date.now()}.pptx`}>
                  Download PPTX
                </a>
              </div>
            )}

            {pdfUrl && (
              <div>
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

            <button className="btn" onClick={handleReset} style={{ marginTop: "1rem" }}>
              Start New Presentation
            </button>
          </div>
        )}
      </main>
    </div>
  );
}
