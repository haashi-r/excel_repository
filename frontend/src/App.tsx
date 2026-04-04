import { useState, useCallback, useRef } from "react";

const API = "";

type Stage = "idle" | "uploading" | "analyzed" | "generating" | "done" | "error";

interface Contract {
  tipo_contratto: string;
  nome_cliente_segnalatore: string;
  data_contratto: string;
  importo_contratto: string;
  // new backend fields
  commissioni_annuale?: string;
  frequenza_pagamento?: string;
  percentuale_per_periodo?: string;
  durata_mesi?: string;
  // old backend field (fallback)
  commissioni?: string;
  piano: string;
  payment_schedule: Array<{
    mese: string;
    data_pagamento_commissioni: string;
    management_mensile: string;
    _amount?: number;
  }>;
}

interface JobResult {
  job_id: string;
  filename: string;
  contracts_found: number;
  document_number: string;
  extraction_notes: string;
  data: { contracts: Contract[]; document_number: string };
}

// Works with both old (commissioni) and new (commissioni_annuale) backend
const getComm = (c: Contract) =>
  c.commissioni_annuale ?? c.commissioni ?? "—";

export default function App() {
  const [stage, setStage] = useState<Stage>("idle");
  const [dragOver, setDragOver] = useState(false);
  const [job, setJob] = useState<JobResult | null>(null);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState(0);
  const [generatedFiles, setGeneratedFiles] = useState<{ docx: boolean; excel: boolean }>({ docx: false, excel: false });
  const fileRef = useRef<HTMLInputElement>(null);

  const reset = () => {
    setStage("idle");
    setJob(null);
    setError("");
    setProgress(0);
    setGeneratedFiles({ docx: false, excel: false });
  };

  const uploadFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".pdf")) {
      setError("Only PDF files are accepted.");
      setStage("error");
      return;
    }
    setStage("uploading");
    setProgress(0);
    const formData = new FormData();
    formData.append("file", file);
    const prog = setInterval(() => setProgress(p => Math.min(p + 8, 85)), 200);
    try {
      const res = await fetch(`${API}/api/upload`, { method: "POST", body: formData });
      clearInterval(prog);
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || err.error || "Upload failed");
      }
      const data: JobResult = await res.json();
      setProgress(100);
      setJob(data);
      setTimeout(() => setStage("analyzed"), 400);
    } catch (e: unknown) {
      clearInterval(prog);
      setError(e instanceof Error ? e.message : "Upload failed");
      setStage("error");
    }
  }, []);

  const generateBoth = async () => {
    if (!job) return;
    setStage("generating");
    setProgress(0);
    const prog = setInterval(() => setProgress(p => Math.min(p + 12, 88)), 180);
    try {
      const res = await fetch(`${API}/api/generate/both`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ job_id: job.job_id, data: job.data }),
      });
      clearInterval(prog);
      if (!res.ok) throw new Error("Generation failed");
      const result = await res.json();
      setProgress(100);
      setGeneratedFiles({ docx: result.docx_ready, excel: result.excel_ready });
      setStage("done");
    } catch (e: unknown) {
      clearInterval(prog);
      setError(e instanceof Error ? e.message : "Generation failed");
      setStage("error");
    }
  };

  const download = (fmt: "docx" | "excel") => {
    if (!job) return;
    window.open(`${API}/api/download/${job.job_id}/${fmt}`, "_blank");
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) uploadFile(file);
  };

  return (
    <div className="app">
      <div className="bg-grid" />
      <div className="bg-glow" />

      <header className="header">
        <div className="logo">
          <span className="logo-icon">⬡</span>
          <div>
            <div className="logo-title">MetalloDoc</div>
            <div className="logo-sub">Conto Metallo Deluxe · Commission Manager</div>
          </div>
        </div>
      </header>

      <main className="main">
        {stage === "idle" && (
          <div className="card center-card animate-in">
            <div className="card-eyebrow">PDF Processor</div>
            <h1 className="card-title">Upload your contract PDF</h1>
            <p className="card-desc">Drag & drop or click to select. We'll extract contract data and generate both Excel and DOCX files.</p>
            <div
              className={`drop-zone ${dragOver ? "drag-active" : ""}`}
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current?.click()}
            >
              <div className="drop-icon">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                  <polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                  <line x1="12" y1="18" x2="12" y2="12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                  <polyline points="9 15 12 12 15 15" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <div className="drop-text">Drop PDF here</div>
              <div className="drop-hint">or click to browse · max 50 MB</div>
              <input ref={fileRef} type="file" accept=".pdf" style={{ display: "none" }}
                onChange={e => { if (e.target.files?.[0]) uploadFile(e.target.files[0]); }} />
            </div>
            <div className="feature-row">
              {["AI Extraction", "Excel .xlsx", "Word .docx"].map(f => (
                <div className="feature-pill" key={f}><span className="pill-check">✓</span>{f}</div>
              ))}
            </div>
          </div>
        )}

        {stage === "uploading" && (
          <div className="card center-card animate-in">
            <div className="spinner-ring" />
            <h2 className="status-title">Analysing PDF…</h2>
            <p className="status-sub">Extracting contract data with AI</p>
            <div className="progress-bar">
              <div className="progress-fill" style={{ width: `${progress}%` }} />
            </div>
            <div className="progress-label">{progress}%</div>
          </div>
        )}

        {stage === "analyzed" && job && (
          <div className="animate-in wide-card">
            <div className="result-header">
              <div>
                <div className="card-eyebrow">Extraction Complete</div>
                <h2 className="result-title">{job.filename}</h2>
                <p className="result-meta">
                  Doc № {job.document_number} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""} found
                </p>
              </div>
              <button className="btn-ghost" onClick={reset}>↩ New PDF</button>
            </div>

            <div className="contracts-grid">
              {job.data.contracts.map((c, i) => (
                <div className="contract-card" key={i}>
                  <div className="contract-tag">Contract {i + 1} · {c.piano}</div>
                  <div className="contract-name">{c.nome_cliente_segnalatore}</div>

                  <div className="contract-row">
                    <span className="cl">Date</span>
                    <span className="cv">{c.data_contratto}</span>
                  </div>
                  <div className="contract-row">
                    <span className="cl">Amount</span>
                    <span className="cv amt">
                      € {Number(c.importo_contratto || 0).toLocaleString("it-IT", { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                  <div className="contract-row">
                    <span className="cl">Commission</span>
                    <span className="cv">{getComm(c)}% p.a.</span>
                  </div>
                  {c.frequenza_pagamento && (
                    <div className="contract-row">
                      <span className="cl">Frequency</span>
                      <span className="cv">{c.frequenza_pagamento}</span>
                    </div>
                  )}
                  {c.durata_mesi && (
                    <div className="contract-row">
                      <span className="cl">Duration</span>
                      <span className="cv">{c.durata_mesi} months</span>
                    </div>
                  )}

                  <div className="sched-table">
                    <div className="sched-head">
                      <span>Month</span><span>Payment Date</span><span>Amount</span>
                    </div>
                    {(c.payment_schedule || []).slice(0, 12).map((s, j) => (
                      <div className={`sched-row ${j % 2 === 0 ? "even" : ""}`} key={j}>
                        <span>{s.mese}</span>
                        <span>{s.data_pagamento_commissioni || "—"}</span>
                        <span className={s.management_mensile ? "pay-amt" : "pay-empty"}>
                          {s.management_mensile
                            ? `€ ${Number(s.management_mensile).toLocaleString("it-IT", { minimumFractionDigits: 2 })}`
                            : "—"}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>

            <div className="action-bar">
              <button className="btn-primary" onClick={generateBoth}>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                  <polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                  <line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                </svg>
                Generate Excel + DOCX
              </button>
            </div>
          </div>
        )}

        {stage === "generating" && (
          <div className="card center-card animate-in">
            <div className="spinner-ring gold" />
            <h2 className="status-title">Generating Files…</h2>
            <p className="status-sub">Building Excel spreadsheet and Word document</p>
            <div className="progress-bar">
              <div className="progress-fill gold" style={{ width: `${progress}%` }} />
            </div>
            <div className="progress-label">{progress}%</div>
          </div>
        )}

        {stage === "done" && job && (
          <div className="card center-card animate-in">
            <div className="done-icon">✓</div>
            <h2 className="status-title">Files Ready!</h2>
            <p className="status-sub">
              Job ID: {job.job_id} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""}
            </p>
            <div className="download-grid">
              {generatedFiles.excel && (
                <button className="dl-card excel" onClick={() => download("excel")}>
                  <div className="dl-icon">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
                      <rect x="3" y="3" width="18" height="18" rx="2" stroke="currentColor" strokeWidth="1.5"/>
                      <path d="M3 9h18M9 3v18" stroke="currentColor" strokeWidth="1.5"/>
                    </svg>
                  </div>
                  <div className="dl-name">Excel File</div>
                  <div className="dl-ext">.xlsx · Commissioni_{job.job_id}</div>
                  <div className="dl-btn">Download ↓</div>
                </button>
              )}
              {generatedFiles.docx && (
                <button className="dl-card docx" onClick={() => download("docx")}>
                  <div className="dl-icon">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5"/>
                      <polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5"/>
                      <line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" strokeWidth="1.5"/>
                      <line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" strokeWidth="1.5"/>
                    </svg>
                  </div>
                  <div className="dl-name">Word Document</div>
                  <div className="dl-ext">.docx · Commissioni_{job.job_id}</div>
                  <div className="dl-btn">Download ↓</div>
                </button>
              )}
            </div>
            <button className="btn-ghost" onClick={reset}>↩ Process another PDF</button>
          </div>
        )}

        {stage === "error" && (
          <div className="card center-card animate-in">
            <div className="error-icon">!</div>
            <h2 className="status-title">Something went wrong</h2>
            <p className="error-msg">{error}</p>
            <button className="btn-primary" onClick={reset}>Try Again</button>
          </div>
        )}
      </main>
    </div>
  );
}

















// import { useState, useCallback, useRef } from "react";

// const API = import.meta.env.VITE_API_URL || "";
// type Stage = "idle" | "uploading" | "analyzed" | "generating" | "done" | "error";

// interface Contract {
//   tipo_contratto: string;
//   nome_cliente_segnalatore: string;
//   data_contratto: string;
//   importo_contratto: string;
//   commissioni: string;
//   piano: string;
//   payment_schedule: Array<{
//     mese: string;
//     data_pagamento_commissioni: string;
//     management_mensile: string;
//   }>;
// }

// interface JobResult {
//   job_id: string;
//   filename: string;
//   contracts_found: number;
//   document_number: string;
//   extraction_notes: string;
//   data: { contracts: Contract[]; document_number: string };
// }

// export default function App() {
//   const [stage, setStage] = useState<Stage>("idle");
//   const [dragOver, setDragOver] = useState(false);
//   const [job, setJob] = useState<JobResult | null>(null);
//   const [error, setError] = useState("");
//   const [progress, setProgress] = useState(0);
//   const [generatedFiles, setGeneratedFiles] = useState<{ docx: boolean; excel: boolean }>({ docx: false, excel: false });
//   const fileRef = useRef<HTMLInputElement>(null);

//   const reset = () => {
//     setStage("idle");
//     setJob(null);
//     setError("");
//     setProgress(0);
//     setGeneratedFiles({ docx: false, excel: false });
//   };



//   const uploadFile = useCallback(async (file: File) => {
//     if (!file.name.toLowerCase().endsWith(".pdf")) {
//       setError("Only PDF files are accepted.");
//       setStage("error");
//       return;
//     }
//     setStage("uploading");
//     setProgress(0);
//     const formData = new FormData();
//     formData.append("file", file);
//     const prog = setInterval(() => setProgress(p => Math.min(p + 8, 85)), 200);
//     try {
//       const res = await fetch(`${API}/api/upload`, { method: "POST", body: formData });
//       clearInterval(prog);
//       if (!res.ok) {
//         const err = await res.json();
//         throw new Error(err.error || "Upload failed");
//       }
//       const data: JobResult = await res.json();
//       setProgress(100);
//       setJob(data);
//       setTimeout(() => setStage("analyzed"), 400);
//     } catch (e: unknown) {
//       clearInterval(prog);
//       setError(e instanceof Error ? e.message : "Upload failed");
//       setStage("error");
//     }
//   }, []);

//   const generateBoth = async () => {
//     if (!job) return;
//     setStage("generating");
//     setProgress(0);
//     const prog = setInterval(() => setProgress(p => Math.min(p + 12, 88)), 180);
//     try {
//       const res = await fetch(`${API}/api/generate/both`, {
//         method: "POST",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ job_id: job.job_id, data: job.data }),
//       });
//       clearInterval(prog);
//       if (!res.ok) throw new Error("Generation failed");
//       const result = await res.json();
//       setProgress(100);
//       setGeneratedFiles({ docx: result.docx_ready, excel: result.excel_ready });
//       // Auto-download Excel immediately
//       if (result.excel_ready) {
//         window.open(`${API}/api/download/${job.job_id}/excel`, "_blank");
//       }
//       setStage("analyzed"); // stay on preview, show download buttons
//     } catch (e: unknown) {
//       clearInterval(prog);
//       setError(e instanceof Error ? e.message : "Generation failed");
//       setStage("error");
//     }
//   };
//   const download = (fmt: "docx" | "excel") => {
//     if (!job) return;
//     window.open(`${API}/api/download/${job.job_id}/${fmt}`, "_blank");
//   };

//   const onDrop = (e: React.DragEvent) => {
//     e.preventDefault();
//     setDragOver(false);
//     const file = e.dataTransfer.files[0];
//     if (file) uploadFile(file);
//   };

//   return (
//     <div className="app">
//       <div className="bg-grid" />
//       <div className="bg-glow" />

//       <header className="header">
//         <div className="logo">
//           <span className="logo-icon">⬡</span>
//           <div>
//             <div className="logo-title">MetalloDoc</div>
//             <div className="logo-sub">Conto Metallo Deluxe · Commission Manager</div>
//           </div>
//         </div>
//       </header>

//       <main className="main">
//         {stage === "idle" && (
//           <div className="card center-card animate-in">
//             <div className="card-eyebrow">PDF Processor</div>
//             <h1 className="card-title">Upload your contract PDF</h1>
//             <p className="card-desc">Drag & drop or click to select. We'll extract contract data and generate both Excel and DOCX files.</p>
//             <div
//               className={`drop-zone ${dragOver ? "drag-active" : ""}`}
//               onDragOver={e => { e.preventDefault(); setDragOver(true); }}
//               onDragLeave={() => setDragOver(false)}
//               onDrop={onDrop}
//               onClick={() => fileRef.current?.click()}
//             >
//               <div className="drop-icon">
//                 <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
//                   <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                   <polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                   <line x1="12" y1="18" x2="12" y2="12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
//                   <polyline points="9 15 12 12 15 15" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                 </svg>
//               </div>
//               <div className="drop-text">Drop PDF here</div>
//               <div className="drop-hint">or click to browse · max 50 MB</div>
//               <input ref={fileRef} type="file" accept=".pdf" style={{ display: "none" }} onChange={e => { if (e.target.files?.[0]) uploadFile(e.target.files[0]); }} />
//             </div>
//             <div className="feature-row">
//               {["AI Extraction", "Excel .xlsx", "Word .docx"].map(f => (
//                 <div className="feature-pill" key={f}><span className="pill-check">✓</span>{f}</div>
//               ))}
//             </div>
//           </div>
//         )}

//         {stage === "uploading" && (
//           <div className="card center-card animate-in">
//             <div className="spinner-ring" />
//             <h2 className="status-title">Analysing PDF…</h2>
//             <p className="status-sub">Extracting contract data with AI</p>
//             <div className="progress-bar">
//               <div className="progress-fill" style={{ width: `${progress}%` }} />
//             </div>
//             <div className="progress-label">{progress}%</div>
//           </div>
//         )}

//         {stage === "analyzed" && job && (
//           <div className="animate-in wide-card">
//             <div className="result-header">
//               <div>
//                 <div className="card-eyebrow">Extraction Complete</div>
//                 <h2 className="result-title">{job.filename}</h2>
//                 <p className="result-meta">Doc № {job.document_number} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""} found</p>
//               </div>
//               <button className="btn-ghost" onClick={reset}>↩ New PDF</button>
//             </div>
//             <div className="contracts-grid">
//               {job.data.contracts.map((c, i) => (
//                 <div className="contract-card" key={i}>
//                   <div className="contract-tag">Contract {i + 1} · {c.piano}</div>
//                   <div className="contract-name">{c.nome_cliente_segnalatore}</div>
//                   <div className="contract-row">
//                     <span className="cl">Date</span>
//                     <span className="cv">{c.data_contratto}</span>
//                   </div>
//                   <div className="contract-row">
//                     <span className="cl">Amount</span>
//                     <span className="cv amt">€ {Number(c.importo_contratto || 0).toLocaleString("it-IT", { minimumFractionDigits: 2 })}</span>
//                   </div>
//                   <div className="contract-row">
//                     <span className="cl">Commission</span>
//                     <span className="cv">{c.commissioni}%</span>
//                   </div>
//                   <div className="sched-table">
//                     <div className="sched-head">
//                       <span>Month</span><span>Payment Date</span><span>Amount</span>
//                     </div>
//                     {c.payment_schedule.slice(0, 12).map((s, j) => (
//                       <div className={`sched-row ${j % 2 === 0 ? "even" : ""}`} key={j}>
//                         <span>{s.mese}</span>
//                         <span>{s.data_pagamento_commissioni || "—"}</span>
//                         <span className={s.management_mensile ? "pay-amt" : "pay-empty"}>
//                           {s.management_mensile ? `€ ${Number(s.management_mensile).toLocaleString("it-IT", { minimumFractionDigits: 2 })}` : "—"}
//                         </span>
//                       </div>
//                     ))}
//                   </div>
//                 </div>
//               ))}
//             </div>
//             <div className="action-bar">
//               <button className="btn-primary" onClick={generateBoth}>
//                 <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/><polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>
//                 Generate Excel + DOCX
//               </button>
//             </div>
//           </div>
//         )}

//         {stage === "generating" && (
//           <div className="card center-card animate-in">
//             <div className="spinner-ring gold" />
//             <h2 className="status-title">Generating Files…</h2>
//             <p className="status-sub">Building Excel spreadsheet and Word document</p>
//             <div className="progress-bar">
//               <div className="progress-fill gold" style={{ width: `${progress}%` }} />
//             </div>
//             <div className="progress-label">{progress}%</div>
//           </div>
//         )}

//         {stage === "done" && job && (
//           <div className="card center-card animate-in">
//             <div className="done-icon">✓</div>
//             <h2 className="status-title">Files Ready!</h2>
//             <p className="status-sub">Job ID: {job.job_id} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""}</p>
//             <div className="download-grid">
//               {generatedFiles.excel && (
//                 <button className="dl-card excel" onClick={() => download("excel")}>
//                   <div className="dl-icon">
//                     <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="18" height="18" rx="2" stroke="currentColor" strokeWidth="1.5"/><path d="M3 9h18M9 3v18" stroke="currentColor" strokeWidth="1.5"/></svg>
//                   </div>
//                   <div className="dl-name">Excel File</div>
//                   <div className="dl-ext">.xlsx · Commissioni_{job.job_id}</div>
//                   <div className="dl-btn">Download ↓</div>
//                 </button>
//               )}
//               {generatedFiles.docx && (
//                 <button className="dl-card docx" onClick={() => download("docx")}>
//                   <div className="dl-icon">
//                     <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5"/><polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5"/><line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" strokeWidth="1.5"/><line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" strokeWidth="1.5"/></svg>
//                   </div>
//                   <div className="dl-name">Word Document</div>
//                   <div className="dl-ext">.docx · Commissioni_{job.job_id}</div>
//                   <div className="dl-btn">Download ↓</div>
//                 </button>
//               )}
//             </div>
//             <button className="btn-ghost" onClick={reset}>↩ Process another PDF</button>
//           </div>
//         )}

//         {stage === "error" && (
//           <div className="card center-card animate-in">
//             <div className="error-icon">!</div>
//             <h2 className="status-title">Something went wrong</h2>
//             <p className="error-msg">{error}</p>
//             <button className="btn-primary" onClick={reset}>Try Again</button>
//           </div>
//         )}
//       </main>
//     </div>
//   );
// }













// import { useState, useCallback, useRef } from "react";

// const API = "http://localhost:8000";

// type Stage = "idle" | "uploading" | "analyzed" | "generating" | "done" | "error";

// interface Contract {
//   tipo_contratto: string;
//   nome_cliente_segnalatore: string;
//   data_contratto: string;
//   importo_contratto: string;
//   commissioni: string;
//   piano: string;
//   payment_schedule: Array<{
//     mese: string;
//     data_pagamento_commissioni: string;
//     management_mensile: string;
//   }>;
// }

// interface JobResult {
//   job_id: string;
//   filename: string;
//   contracts_found: number;
//   document_number: string;
//   extraction_notes: string;
//   data: { contracts: Contract[]; document_number: string };
// }

// export default function App() {
//   const [stage, setStage] = useState<Stage>("idle");
//   const [dragOver, setDragOver] = useState(false);
//   const [job, setJob] = useState<JobResult | null>(null);
//   const [error, setError] = useState("");
//   const [progress, setProgress] = useState(0);
//   const [generatedFiles, setGeneratedFiles] = useState<{ docx: boolean; excel: boolean }>({ docx: false, excel: false });
//   const fileRef = useRef<HTMLInputElement>(null);

//   const reset = () => {
//     setStage("idle");
//     setJob(null);
//     setError("");
//     setProgress(0);
//     setGeneratedFiles({ docx: false, excel: false });
//   };

//   const uploadFile = useCallback(async (file: File) => {
//     if (!file.name.toLowerCase().endsWith(".pdf")) {
//       setError("Only PDF files are accepted.");
//       setStage("error");
//       return;
//     }

//     setStage("uploading");
//     setProgress(0);

//     const formData = new FormData();
//     formData.append("file", file);

//     // Simulate progress
//     const prog = setInterval(() => setProgress(p => Math.min(p + 8, 85)), 200);

//     try {
//       const res = await fetch(`${API}/api/upload`, { method: "POST", body: formData });
//       clearInterval(prog);
//       if (!res.ok) {
//         const err = await res.json();
//         throw new Error(err.error || "Upload failed");
//       }
//       const data: JobResult = await res.json();
//       setProgress(100);
//       setJob(data);
//       setTimeout(() => setStage("analyzed"), 400);
//     } catch (e: unknown) {
//       clearInterval(prog);
//       setError(e instanceof Error ? e.message : "Upload failed");
//       setStage("error");
//     }
//   }, []);

//   const generateBoth = async () => {
//     if (!job) return;
//     setStage("generating");
//     setProgress(0);
//     const prog = setInterval(() => setProgress(p => Math.min(p + 12, 88)), 180);

//     try {
//       const res = await fetch(`${API}/api/generate/both`, {
//         method: "POST",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ job_id: job.job_id, data: job.data }),
//       });
//       clearInterval(prog);
//       if (!res.ok) throw new Error("Generation failed");
//       const result = await res.json();
//       setProgress(100);
//       setGeneratedFiles({ docx: result.docx_ready, excel: result.excel_ready });
//       setStage("done");
//     } catch (e: unknown) {
//       clearInterval(prog);
//       setError(e instanceof Error ? e.message : "Generation failed");
//       setStage("error");
//     }
//   };

//   const download = (fmt: "docx" | "excel") => {
//     if (!job) return;
//     window.open(`${API}/api/download/${job.job_id}/${fmt}`, "_blank");
//   };

//   const onDrop = (e: React.DragEvent) => {
//     e.preventDefault();
//     setDragOver(false);
//     const file = e.dataTransfer.files[0];
//     if (file) uploadFile(file);
//   };

//   return (
//     <div className="app">
//       <div className="bg-grid" />
//       <div className="bg-glow" />

//       <header className="header">
//         <div className="logo">
//           <span className="logo-icon">⬡</span>
//           <div>
//             <div className="logo-title">MetalloDoc</div>
//             <div className="logo-sub">Conto Metallo Deluxe · Commission Manager</div>
//           </div>
//         </div>
//       </header>

//       <main className="main">

//         {/* ── IDLE ── */}
//         {stage === "idle" && (
//           <div className="card center-card animate-in">
//             <div className="card-eyebrow">PDF Processor</div>
//             <h1 className="card-title">Upload your contract PDF</h1>
//             <p className="card-desc">Drag & drop or click to select. We'll extract contract data and generate both Excel and DOCX files.</p>

//             <div
//               className={`drop-zone ${dragOver ? "drag-active" : ""}`}
//               onDragOver={e => { e.preventDefault(); setDragOver(true); }}
//               onDragLeave={() => setDragOver(false)}
//               onDrop={onDrop}
//               onClick={() => fileRef.current?.click()}
//             >
//               <div className="drop-icon">
//                 <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
//                   <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                   <polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                   <line x1="12" y1="18" x2="12" y2="12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
//                   <polyline points="9 15 12 12 15 15" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
//                 </svg>
//               </div>
//               <div className="drop-text">Drop PDF here</div>
//               <div className="drop-hint">or click to browse · max 50 MB</div>
//               <input ref={fileRef} type="file" accept=".pdf" style={{ display: "none" }} onChange={e => { if (e.target.files?.[0]) uploadFile(e.target.files[0]); }} />
//             </div>

//             <div className="feature-row">
//               {["AI Extraction", "Excel .xlsx", "Word .docx"].map(f => (
//                 <div className="feature-pill" key={f}><span className="pill-check">✓</span>{f}</div>
//               ))}
//             </div>
//           </div>
//         )}

//         {/* ── UPLOADING ── */}
//         {stage === "uploading" && (
//           <div className="card center-card animate-in">
//             <div className="spinner-ring" />
//             <h2 className="status-title">Analysing PDF…</h2>
//             <p className="status-sub">Extracting contract data with AI</p>
//             <div className="progress-bar">
//               <div className="progress-fill" style={{ width: `${progress}%` }} />
//             </div>
//             <div className="progress-label">{progress}%</div>
//           </div>
//         )}

//         {/* ── ANALYZED ── */}
//         {stage === "analyzed" && job && (
//           <div className="animate-in wide-card">
//             <div className="result-header">
//               <div>
//                 <div className="card-eyebrow">Extraction Complete</div>
//                 <h2 className="result-title">{job.filename}</h2>
//                 <p className="result-meta">Doc № {job.document_number} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""} found</p>
//               </div>
//               <button className="btn-ghost" onClick={reset}>↩ New PDF</button>
//             </div>

//             {/* Contracts preview */}
//             <div className="contracts-grid">
//               {job.data.contracts.map((c, i) => (
//                 <div className="contract-card" key={i}>
//                   <div className="contract-tag">Contract {i + 1} · {c.piano}</div>
//                   <div className="contract-name">{c.nome_cliente_segnalatore}</div>
//                   <div className="contract-row">
//                     <span className="cl">Date</span>
//                     <span className="cv">{c.data_contratto}</span>
//                   </div>
//                   <div className="contract-row">
//                     <span className="cl">Amount</span>
//                     <span className="cv amt">€ {Number(c.importo_contratto || 0).toLocaleString("it-IT", { minimumFractionDigits: 2 })}</span>
//                   </div>
//                   <div className="contract-row">
//                     <span className="cl">Commission</span>
//                     <span className="cv">{c.commissioni}%</span>
//                   </div>

//                   {/* Schedule mini table */}
//                   <div className="sched-table">
//                     <div className="sched-head">
//                       <span>Month</span><span>Payment Date</span><span>Amount</span>
//                     </div>
//                     {c.payment_schedule.slice(0, 12).map((s, j) => (
//                       <div className={`sched-row ${j % 2 === 0 ? "even" : ""}`} key={j}>
//                         <span>{s.mese}</span>
//                         <span>{s.data_pagamento_commissioni || "—"}</span>
//                         <span className={s.management_mensile ? "pay-amt" : "pay-empty"}>
//                           {s.management_mensile ? `€ ${Number(s.management_mensile).toLocaleString("it-IT", { minimumFractionDigits: 2 })}` : "—"}
//                         </span>
//                       </div>
//                     ))}
//                   </div>
//                 </div>
//               ))}
//             </div>

//             <div className="action-bar">
//               <button className="btn-primary" onClick={generateBoth}>
//                 <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/><polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>
//                 Generate Excel + DOCX
//               </button>
//             </div>
//           </div>
//         )}

//         {/* ── GENERATING ── */}
//         {stage === "generating" && (
//           <div className="card center-card animate-in">
//             <div className="spinner-ring gold" />
//             <h2 className="status-title">Generating Files…</h2>
//             <p className="status-sub">Building Excel spreadsheet and Word document</p>
//             <div className="progress-bar">
//               <div className="progress-fill gold" style={{ width: `${progress}%` }} />
//             </div>
//             <div className="progress-label">{progress}%</div>
//           </div>
//         )}

//         {/* ── DONE ── */}
//         {stage === "done" && job && (
//           <div className="card center-card animate-in">
//             <div className="done-icon">✓</div>
//             <h2 className="status-title">Files Ready!</h2>
//             <p className="status-sub">Job ID: {job.job_id} · {job.contracts_found} contract{job.contracts_found !== 1 ? "s" : ""}</p>

//             <div className="download-grid">
//               {generatedFiles.excel && (
//                 <button className="dl-card excel" onClick={() => download("excel")}>
//                   <div className="dl-icon">
//                     <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="18" height="18" rx="2" stroke="currentColor" strokeWidth="1.5"/><path d="M3 9h18M9 3v18" stroke="currentColor" strokeWidth="1.5"/></svg>
//                   </div>
//                   <div className="dl-name">Excel File</div>
//                   <div className="dl-ext">.xlsx · Commissioni_{job.job_id}</div>
//                   <div className="dl-btn">Download ↓</div>
//                 </button>
//               )}
//               {generatedFiles.docx && (
//                 <button className="dl-card docx" onClick={() => download("docx")}>
//                   <div className="dl-icon">
//                     <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.5"/><polyline points="14 2 14 8 20 8" stroke="currentColor" strokeWidth="1.5"/><line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" strokeWidth="1.5"/><line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" strokeWidth="1.5"/></svg>
//                   </div>
//                   <div className="dl-name">Word Document</div>
//                   <div className="dl-ext">.docx · Commissioni_{job.job_id}</div>
//                   <div className="dl-btn">Download ↓</div>
//                 </button>
//               )}
//             </div>

//             <button className="btn-ghost" onClick={reset}>↩ Process another PDF</button>
//           </div>
//         )}

//         {/* ── ERROR ── */}
//         {stage === "error" && (
//           <div className="card center-card animate-in">
//             <div className="error-icon">!</div>
//             <h2 className="status-title">Something went wrong</h2>
//             <p className="error-msg">{error}</p>
//             <button className="btn-primary" onClick={reset}>Try Again</button>
//           </div>
//         )}
//       </main>


//     </div>
//   );
// }