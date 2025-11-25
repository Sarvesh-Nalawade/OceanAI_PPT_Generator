import React from "react";

type Props = {
  pdfUrl: string;
};

const PdfViewer: React.FC<Props> = ({ pdfUrl }) => {
  return (
    <div style={{ border: "1px solid rgba(255,255,255,0.06)", borderRadius: 8, overflow: "hidden", marginTop: 8 }}>
      <iframe src={pdfUrl} title="PDF Preview" style={{ width: "100%", height: 600, border: "none" }} />
    </div>
  );
};

export default PdfViewer;
