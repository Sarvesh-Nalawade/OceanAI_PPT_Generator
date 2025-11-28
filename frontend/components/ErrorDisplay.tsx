// frontend/components/ErrorDisplay.tsx
import React from "react";

interface ErrorDisplayProps {
  message: string;
}

const ErrorDisplay: React.FC<ErrorDisplayProps> = ({ message }) => {
  let title = "An Error Occurred";
  let body = message;
  let isInfo = false;

  if (message.includes("Received PPTX")) {
    title = "PPTX Ready for Download";
    body = "Your presentation has been generated. You can download it using the link provided.";
    isInfo = true;
  } else if (message.startsWith("Server error:")) {
    try {
      const jsonString = message.substring(message.indexOf("{"));
      const errorObj = JSON.parse(jsonString);
      if (errorObj.detail) {
        title = "Server Error";
        // Further parsing the detail for the specific error message
        const detail = errorObj.detail;
        const agentResponsePrefix = "Agent response: ";
        const agentResponseIndex = detail.indexOf(agentResponsePrefix);
        if (agentResponseIndex !== -1) {
            body = detail.substring(agentResponseIndex + agentResponsePrefix.length);
        } else {
            body = detail;
        }
      }
    } catch (e) {
      // Not a JSON error, show the raw message
      body = message;
    }
  }

  return (
    <div className={`error-display ${isInfo ? "info" : "error"}`}>
      <div className="error-content">
        {/* <div className="error-title">{title}</div> */}
        <div className="error-body">{body}</div>
      </div>
    </div>
  );
};

export default ErrorDisplay;
