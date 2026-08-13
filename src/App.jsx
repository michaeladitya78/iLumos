import { useState, useEffect, useRef, useCallback } from 'react';
import { saveAs } from 'file-saver';
import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, WidthType,
} from 'docx';

// ═══════════════════════════════════════════════════════════════════════════════
// INITIAL MOCK DATA
// ═══════════════════════════════════════════════════════════════════════════════

const initialChart = [
  {
    id: 1,
    claimElement: 'A temperature control device with a wireless communication module',
    accusedFeature:
      'Acme Thermostat product page states: "WiFi-enabled smart thermostat connects to your home network"',
    aiReasoning:
      'The Acme device has WiFi capability which satisfies the wireless communication module requirement.',
    weakEvidence: false,
  },
  {
    id: 2,
    claimElement: 'A motion sensor for detecting occupancy',
    accusedFeature:
      'Acme technical specifications document shows: "Built-in motion sensor detects when people are home"',
    aiReasoning:
      'Motion sensor explicitly mentioned in specs directly maps to the claim element for occupancy detection.',
    weakEvidence: false,
  },
  {
    id: 3,
    claimElement: 'Machine learning algorithm that learns user temperature preferences over time',
    accusedFeature: 'Acme marketing materials claim: "Auto-Schedule learns your preferred temperatures"',
    aiReasoning:
      'The learning behavior described suggests an ML algorithm, though technical implementation details are not disclosed. May need stronger technical evidence.',
    weakEvidence: true,
  },
];

// ═══════════════════════════════════════════════════════════════════════════════
// AI LOGIC — deterministic, clearly separated per spec §11
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Confidence gate simulation.
 * "temperature sensor array" → { confident: false } → Edge Case C.
 * All other messages → { confident: true } → happy path.
 */
function checkRetrievalConfidence(message) {
  if (message.toLowerCase().includes('temperature sensor array')) {
    return { confident: false };
  }
  return { confident: true };
}

/**
 * Generate a scripted proposal based on the user's message.
 * Returns { rowId, field, proposalText, evidenceQuote, sourceDocument }.
 */
function generateProposal(message, _currentChart, extraContext = '') {
  const lower = message.toLowerCase();

  const correction = extraContext ? `Thanks for the correction — revising based on your input: ` : '';

  if (lower.includes('machine learning') || lower.includes(' ml ') || lower.includes('element 3') ||
      lower.includes('learning') || lower.includes('algorithm') || lower.includes('auto-schedule')) {
    const body = extraContext
      ? `${extraContext} — additionally, the technical specifications document confirms the thermostat employs a local federated learning model for schedule prediction without transmitting raw telemetry.`
      : 'I analyzed the technical specifications document. I found a reference stating: "The thermostat utilizes a local federated learning model to predict scheduling behavior without transmitting raw telemetry to the cloud." This constitutes direct technical evidence of an ML algorithm satisfying Claim Element 3.';
    return {
      rowId: 3, field: 'aiReasoning',
      proposalText: correction + body,
      evidenceQuote: '"The thermostat utilizes a local federated learning model to predict scheduling behavior without transmitting raw telemetry to the cloud."',
      sourceDocument: 'Acme Technical Specifications v2.3, §4.2 – Adaptive Scheduling',
    };
  }

  if (lower.includes('wifi') || lower.includes('wireless') || lower.includes('element 1') ||
      lower.includes('fcc') || lower.includes('communication module')) {
    const body = extraContext
      ? `${extraContext} — the FCC filing additionally confirms IEEE 802.11 b/g/n compliance directly satisfying the wireless communication module limitation.`
      : 'I analyzed the Acme Thermostat product page and FCC filing documents. I found a reference stating: "The Acme thermostat is certified under FCC ID: ABCD-12345 for IEEE 802.11 b/g/n wireless operation." This provides stronger technical evidence for the wireless communication module claim element.';
    return {
      rowId: 1, field: 'aiReasoning',
      proposalText: correction + body,
      evidenceQuote: '"Acme thermostat is certified under FCC ID: ABCD-12345 for IEEE 802.11 b/g/n wireless operation."',
      sourceDocument: 'FCC Equipment Authorization Database, Filing ID: ABCD-12345',
    };
  }

  if (lower.includes('motion') || lower.includes('occupancy') || lower.includes('element 2') ||
      lower.includes('pir') || lower.includes('infrared')) {
    const body = extraContext
      ? `${extraContext} — the hardware schematic confirms a dual-element PIR sensor meeting the occupancy detection limitation.`
      : 'I analyzed the Acme hardware schematic. I found a reference stating: "The device incorporates a dual-element pyroelectric infrared (PIR) sensor operating at 5–14 μm wavelength range for presence detection." This is strong technical evidence for the motion sensor claim element.';
    return {
      rowId: 2, field: 'aiReasoning',
      proposalText: correction + body,
      evidenceQuote: '"dual-element pyroelectric infrared (PIR) sensor operating at 5–14 μm wavelength range"',
      sourceDocument: 'Acme Hardware Schematic Rev D, Sheet 3 – Sensor Array',
    };
  }

  if (lower.includes('sensor array') || lower.includes('schematic') || lower.includes('hardware') ||
      lower.includes('thermistor') || lower.includes('thermocouple')) {
    const body = extraContext
      ? `${extraContext} — the newly indexed schematic additionally reveals a 4-element NTC thermistor array providing ±0.1°C accuracy across a −20°C to 60°C range.`
      : 'After indexing the temperature sensor array schematic, I found: "The Acme device incorporates a 4-element NTC thermistor array providing ±0.1°C accuracy across a −20°C to 60°C range." This directly supports the temperature sensing claim elements with high confidence.';
    return {
      rowId: 1, field: 'aiReasoning',
      proposalText: correction + body,
      evidenceQuote: '"4-element NTC thermistor array providing ±0.1°C accuracy across a −20°C to 60°C range"',
      sourceDocument: 'temperature_sensor_array_schematic.pdf, Section 2 – Sensing Elements',
    };
  }

  // Generic fallback
  const body = extraContext
    ? `${extraContext} — updated analysis incorporated.`
    : 'I reviewed the indexed documents and found supporting evidence. The technical specifications provide corroborating data that strengthens the infringement mapping for the identified claim elements.';
  return {
    rowId: 3, field: 'aiReasoning',
    proposalText: correction + body,
    evidenceQuote: 'See indexed documents for full citation chain.',
    sourceDocument: 'Supporting Evidence Documents',
  };
}

/**
 * Apply a proposal to the current chart rows (immutable update).
 */
function applyProposalToChart(currentRows, proposal) {
  return currentRows.map((row) =>
    row.id === proposal.rowId
      ? { ...row, [proposal.field]: proposal.proposalText, weakEvidence: false }
      : row
  );
}

/**
 * Push new chart state onto history stack (FIFO, max 20).
 * Index 0 = current. New states unshift onto front; oldest popped when > 20.
 */
function pushHistory(chartHistory, newChartState) {
  const next = [newChartState, ...chartHistory];
  if (next.length > 20) next.pop();
  return next;
}

/**
 * Undo by discarding current state (index 0). Caller must check length > 1.
 */
function undoLastChange(chartHistory) {
  if (chartHistory.length <= 1) return chartHistory;
  return chartHistory.slice(1);
}

/**
 * Export chart: generates a true Microsoft Word (.docx) document using the `docx` library.
 * 
 * CRITICAL FIX FOR FILENAME:
 * Chrome/Edge strip the `download="iLumos_EoU_Chart.docx"` attribute and fallback to a
 * random UUID filename if `URL.revokeObjectURL(url)` is called synchronously after `click()`.
 * We keep the ObjectURL valid for 60 seconds before cleanup.
 */
async function exportChart(currentRows) {
  const filename = 'iLumos_EoU_Chart.docx';

  try {
    // Header cells styling (Blue background, white bold text)
    const headerCells = ['Patent Claim Element', 'Accused Product Feature', 'AI Reasoning (Evidence)'].map(
      (txt) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: txt, bold: true, size: 20, color: 'FFFFFF' })],
            }),
          ],
          shading: { fill: '1D4ED8' },
        })
    );

    // Body rows styling
    const bodyRows = currentRows.map(
      (row) =>
        new TableRow({
          children: [
            row.claimElement || '',
            row.accusedFeature || '',
            row.aiReasoning || '',
          ].map(
            (txt) =>
              new TableCell({
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: String(txt), size: 18, color: '1E293B' })],
                  }),
                ],
                width: { size: 33, type: WidthType.PERCENTAGE },
              })
          ),
        })
    );

    // Create Word Document
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: 'Rule 11-Compliant Evidence of Use (EoU) Claim Chart',
              heading: HeadingLevel.HEADING_1,
              spacing: { after: 120 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `Patent US123456 · Generated by iLumos AI · ${new Date().toLocaleDateString()}`,
                  italics: true,
                  size: 18,
                  color: '64748B',
                }),
              ],
              spacing: { after: 300 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({ children: headerCells, tableHeader: true }),
                ...bodyRows,
              ],
            }),
          ],
        },
      ],
    });

    // Generate Blob
    const blob = await Packer.toBlob(doc);
    
    // Download trigger with deferred URL revocation
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();

    // Delay cleanup so Chrome background process retains the filename!
    setTimeout(() => {
      try {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      } catch (_) {}
    }, 60000);

    return true;

  } catch (docxErr) {
    console.error('[iLumos] .docx export error, using Word HTML fallback:', docxErr);

    // Fallback: Word-compatible HTML file saved as .doc (opens natively in MS Word)
    try {
      const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>iLumos EoU Claim Chart</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; color: #1e293b; }
            h1 { color: #0f172a; font-size: 20px; }
            .meta { color: #64748b; font-style: italic; font-size: 12px; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 10px; }
            th { background-color: #1d4ed8; color: #ffffff; font-weight: bold; text-align: left; padding: 10px; border: 1px solid #cbd5e1; }
            td { padding: 10px; border: 1px solid #cbd5e1; vertical-align: top; font-size: 13px; }
            tr:nth-child(even) { background-color: #f8fafc; }
          </style>
        </head>
        <body>
          <h1>Rule 11-Compliant Evidence of Use (EoU) Claim Chart</h1>
          <div class="meta">Patent US123456 &middot; Generated by iLumos AI &middot; ${new Date().toLocaleDateString()}</div>
          <table>
            <thead>
              <tr>
                <th>Patent Claim Element</th>
                <th>Accused Product Feature</th>
                <th>AI Reasoning (Evidence)</th>
              </tr>
            </thead>
            <tbody>
              ${currentRows.map(r => `
                <tr>
                  <td>${r.claimElement || ''}</td>
                  <td>${r.accusedFeature || ''}</td>
                  <td>${r.aiReasoning || ''}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </body>
        </html>
      `;

      const fallbackBlob = new Blob(['\ufeff' + htmlContent], { type: 'application/msword' });
      const fallbackUrl = URL.createObjectURL(fallbackBlob);
      const fallbackLink = document.createElement('a');
      fallbackLink.href = fallbackUrl;
      fallbackLink.download = 'iLumos_EoU_Chart.doc';
      fallbackLink.setAttribute('download', 'iLumos_EoU_Chart.doc');
      document.body.appendChild(fallbackLink);
      fallbackLink.click();

      setTimeout(() => {
        try {
          document.body.removeChild(fallbackLink);
          URL.revokeObjectURL(fallbackUrl);
        } catch (_) {}
      }, 60000);

    } catch (fallbackErr) {
      console.error('[iLumos] Fallback export failed:', fallbackErr);
    }
    return true;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SMALL SHARED UI COMPONENTS
// ═══════════════════════════════════════════════════════════════════════════════

function TypingIndicator() {
  return (
    <div className="flex items-center gap-1 px-4 py-3 bg-white border border-slate-200 rounded-2xl rounded-tl-sm w-fit shadow-sm message-slide-in">
      <div className="w-2 h-2 bg-slate-400 rounded-full typing-dot" />
      <div className="w-2 h-2 bg-slate-400 rounded-full typing-dot" />
      <div className="w-2 h-2 bg-slate-400 rounded-full typing-dot" />
    </div>
  );
}

function Toast({ message, type, onDismiss }) {
  useEffect(() => {
    const t = setTimeout(onDismiss, 4000);
    return () => clearTimeout(t);
  }, [onDismiss]);

  const bg = type === 'warning' ? 'bg-orange-500' : 'bg-emerald-600';
  return (
    <div className={`fixed bottom-6 right-6 z-50 flex items-center gap-2 px-4 py-3 rounded-xl text-white text-sm font-medium shadow-2xl toast-slide-up max-w-sm ${bg}`}>
      {type !== 'warning' && (
        <svg className="w-4 h-4 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
        </svg>
      )}
      {type === 'warning' && (
        <svg className="w-4 h-4 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" />
        </svg>
      )}
      {message}
    </div>
  );
}

function FileChip({ name, onRemove }) {
  return (
    <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-blue-50 border border-blue-200 rounded-full text-xs text-blue-700 font-medium">
      <svg className="w-3 h-3 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
      </svg>
      {name}
      <button onClick={onRemove} className="ml-0.5 text-blue-400 hover:text-blue-700 font-bold leading-none">×</button>
    </span>
  );
}

function AiAvatar() {
  return (
    <div className="w-7 h-7 bg-gradient-to-br from-blue-500 to-blue-700 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5 shadow-sm">
      <svg className="w-3.5 h-3.5 text-white" fill="currentColor" viewBox="0 0 20 20">
        <path d="M10 2a1 1 0 011 1v1.323l3.954 1.582 1.599-.8a1 1 0 01.894 1.79l-1.233.616 1.738 5.42a1 1 0 01-.285 1.05A3.989 3.989 0 0115 15a3.989 3.989 0 01-2.667-1.019 1 1 0 01-.285-1.05l1.715-5.349L11 6.477V16h2a1 1 0 110 2H7a1 1 0 110-2h2V6.477L6.237 7.582l1.715 5.349a1 1 0 01-.285 1.05A3.989 3.989 0 015 15a3.989 3.989 0 01-2.667-1.019 1 1 0 01-.285-1.05l1.738-5.42-1.233-.617a1 1 0 01.894-1.788l1.599.799L9 4.323V3a1 1 0 011-1z" />
      </svg>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SETUP SCREEN
// ═══════════════════════════════════════════════════════════════════════════════

function SetupScreen({ onStart }) {
  const [claimFile, setClaimFile]         = useState(null);
  const [evidenceFile, setEvidenceFile]   = useState(null);
  const [instructions, setInstructions]   = useState('');

  const canStart = !!(claimFile && evidenceFile);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 flex items-center justify-center p-6 relative overflow-hidden">
      {/* decorative blobs */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-blue-600/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-1/4 right-1/4 w-80 h-80 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none" />

      <div className="relative w-full max-w-2xl">
        {/* logo + headline */}
        <div className="text-center mb-10">
          <div className="inline-flex items-center gap-3 mb-5">
            <div className="w-12 h-12 bg-blue-600 rounded-2xl flex items-center justify-center shadow-xl shadow-blue-900/40">
              <svg className="w-7 h-7 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
              </svg>
            </div>
            <span className="text-4xl font-bold text-white tracking-tight">iLumos</span>
          </div>
          <h1 className="text-xl font-semibold text-slate-300">AI Patent Claim Chart Refinement</h1>
          <p className="text-sm text-slate-500 mt-1 max-w-md mx-auto">
            Upload your baseline chart and evidence documents to begin an AI-assisted refinement session
          </p>
        </div>

        {/* card */}
        <div className="bg-white rounded-2xl shadow-2xl shadow-black/30 p-8 space-y-6">

          {/* upload row */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
            {/* Baseline Claim Chart */}
            <div className="space-y-2">
              <label className="block text-sm font-semibold text-slate-700">
                Baseline Claim Chart <span className="font-normal text-slate-400 text-xs">(CSV / Word)</span>
              </label>
              {claimFile ? (
                <div className="flex flex-col gap-2 p-4 border border-slate-200 rounded-xl bg-slate-50">
                  <FileChip name={claimFile} onRemove={() => setClaimFile(null)} />
                  <span className="text-xs text-emerald-600 flex items-center gap-1">
                    <svg className="w-3.5 h-3.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                    Ready to index
                  </span>
                </div>
              ) : (
                <button
                  onClick={() => setClaimFile('thermostat_claim_chart.csv')}
                  className="w-full border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 rounded-xl p-6 text-center transition-all duration-200 group"
                >
                  <div className="w-10 h-10 mx-auto mb-2 bg-slate-100 group-hover:bg-blue-100 rounded-lg flex items-center justify-center transition-colors">
                    <svg className="w-5 h-5 text-slate-400 group-hover:text-blue-500 transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                    </svg>
                  </div>
                  <span className="text-sm font-medium text-slate-500 group-hover:text-blue-600 transition-colors">Click to attach</span>
                  <p className="text-xs text-slate-400 mt-0.5">CSV, DOCX up to 10 MB</p>
                </button>
              )}
            </div>

            {/* Supporting Evidence */}
            <div className="space-y-2">
              <label className="block text-sm font-semibold text-slate-700">
                Supporting Evidence Documents <span className="font-normal text-slate-400 text-xs">(PDF / URL)</span>
              </label>
              {evidenceFile ? (
                <div className="flex flex-col gap-2 p-4 border border-slate-200 rounded-xl bg-slate-50">
                  <FileChip name={evidenceFile} onRemove={() => setEvidenceFile(null)} />
                  <span className="text-xs text-emerald-600 flex items-center gap-1">
                    <svg className="w-3.5 h-3.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                    Ready to index
                  </span>
                </div>
              ) : (
                <button
                  onClick={() => setEvidenceFile('acme_tech_specs.pdf')}
                  className="w-full border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 rounded-xl p-6 text-center transition-all duration-200 group"
                >
                  <div className="w-10 h-10 mx-auto mb-2 bg-slate-100 group-hover:bg-blue-100 rounded-lg flex items-center justify-center transition-colors">
                    <svg className="w-5 h-5 text-slate-400 group-hover:text-blue-500 transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                  </div>
                  <span className="text-sm font-medium text-slate-500 group-hover:text-blue-600 transition-colors">Click to attach</span>
                  <p className="text-xs text-slate-400 mt-0.5">PDF, URL up to 50 MB</p>
                </button>
              )}
            </div>
          </div>

          {/* System instructions */}
          <div className="space-y-2">
            <label className="block text-sm font-semibold text-slate-700">
              System Instructions <span className="font-normal text-slate-400 text-xs">(optional)</span>
            </label>
            <input
              type="text"
              value={instructions}
              onChange={(e) => setInstructions(e.target.value)}
              placeholder="e.g. Focus on literal infringement"
              className="w-full px-4 py-3 text-sm border border-slate-200 rounded-xl bg-slate-50 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent placeholder:text-slate-400 text-slate-700 transition-all"
            />
          </div>

          {/* Start button */}
          <button
            onClick={() => canStart && onStart({ claimFile, evidenceFile, instructions })}
            disabled={!canStart}
            className={`w-full py-3.5 rounded-xl text-sm font-bold tracking-wide transition-all duration-200 ${
              canStart
                ? 'bg-blue-600 hover:bg-blue-700 text-white shadow-lg shadow-blue-200 hover:shadow-blue-300 active:scale-[0.99]'
                : 'bg-slate-100 text-slate-400 cursor-not-allowed'
            }`}
          >
            {canStart ? 'Start Refinement Session →' : 'Attach both documents to continue'}
          </button>
          {!canStart && (
            <p className="text-xs text-center text-slate-400">
              A claim chart and at least one evidence document are required.
            </p>
          )}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MESSAGE BUBBLE
// ═══════════════════════════════════════════════════════════════════════════════

function MessageBubble({ msg, onApply, onReject, onFlag, onSubmitCorrection, onSchematicUpload, onCorrectionChange }) {
  const isUser = msg.role === 'user';

  // ── User bubble ────────────────────────────────────────────────────────────
  if (isUser) {
    return (
      <div className="flex justify-end message-slide-in">
        <div className="max-w-[85%] space-y-1">
          {msg.variant === 'correction' && (
            <div className="flex justify-end mb-1">
              <span className="text-xs font-semibold text-orange-600 bg-orange-50 px-2 py-0.5 rounded-full border border-orange-200">
                ⚑ Correction
              </span>
            </div>
          )}
          <div className="px-4 py-3 bg-blue-600 text-white rounded-2xl rounded-tr-sm text-sm leading-relaxed shadow-sm">
            {msg.content}
          </div>
          <p className="text-right text-xs text-slate-400 pr-1">
            {msg.ts.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </p>
        </div>
      </div>
    );
  }

  // ── AI: system / info bubble ───────────────────────────────────────────────
  if (msg.variant === 'system') {
    return (
      <div className="flex justify-start message-slide-in">
        <div className="max-w-[90%] space-y-1">
          <div className="flex items-start gap-2">
            <div className="w-7 h-7 bg-emerald-100 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5">
              <svg className="w-3.5 h-3.5 text-emerald-600" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
              </svg>
            </div>
            <div className="px-4 py-3 bg-emerald-50 border border-emerald-200 rounded-2xl rounded-tl-sm text-sm text-emerald-800 leading-relaxed">
              {msg.content}
            </div>
          </div>
          <p className="pl-9 text-xs text-slate-400">
            {msg.ts.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </p>
        </div>
      </div>
    );
  }

  // ── AI: missing evidence (Edge Case C) ────────────────────────────────────
  if (msg.variant === 'missing-evidence') {
    return (
      <div className="flex justify-start message-slide-in">
        <div className="max-w-[92%] w-full space-y-1">
          <div className="flex items-start gap-2">
            <AiAvatar />
            <div className="flex-1 px-4 py-3 bg-white border border-slate-200 rounded-2xl rounded-tl-sm shadow-sm space-y-3">
              {/* orange edge-case label */}
              <span className="inline-block text-xs font-semibold text-orange-600 bg-orange-50 px-2 py-0.5 rounded-full border border-orange-200">
                ⚠ Evidence Missing
              </span>
              <p className="text-sm text-slate-700 leading-relaxed">{msg.content}</p>

              {/* orange dashed upload zone */}
              {!msg.uploadedFile ? (
                <div className="border-2 border-dashed border-orange-300 bg-orange-50 rounded-xl p-4 text-center space-y-2 transition-all">
                  <div className="w-9 h-9 mx-auto bg-orange-100 rounded-lg flex items-center justify-center">
                    <svg className="w-4.5 h-4.5 text-orange-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                    </svg>
                  </div>
                  <p className="text-xs text-orange-700 font-medium">Upload sensor hardware schematic to proceed</p>
                  <button
                    onClick={() => onSchematicUpload(msg.id)}
                    className="px-4 py-1.5 bg-orange-500 hover:bg-orange-600 text-white text-xs font-semibold rounded-lg transition-all active:scale-95 shadow-sm"
                  >
                    Choose File
                  </button>
                </div>
              ) : (
                <div className="flex items-center gap-2 px-3 py-2 bg-emerald-50 border border-emerald-200 rounded-lg">
                  <svg className="w-4 h-4 text-emerald-600 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                  <span className="text-xs text-emerald-700 font-medium">{msg.uploadedFile}</span>
                  <span className="text-xs text-emerald-500 ml-auto">Indexed ✓</span>
                </div>
              )}
            </div>
          </div>
          <p className="pl-9 text-xs text-slate-400">
            {msg.ts.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </p>
        </div>
      </div>
    );
  }

  // ── AI: proposal bubble ────────────────────────────────────────────────────
  if (msg.variant === 'proposal') {
    const { status, proposal, evidenceQuote, sourceDocument, isRevised, correctionText = '' } = msg;

    return (
      <div className="flex justify-start message-slide-in">
        <div className="max-w-[92%] w-full space-y-1">
          <div className="flex items-start gap-2">
            <AiAvatar />
            <div className="flex-1 px-4 py-3 bg-white border border-slate-200 rounded-2xl rounded-tl-sm shadow-sm space-y-3">
              {isRevised && (
                <span className="inline-block text-xs font-semibold text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200">
                  ✦ Revised Proposal
                </span>
              )}

              <p className="text-sm text-slate-700 leading-relaxed">{msg.content}</p>

              {/* evidence blockquote */}
              {evidenceQuote && (
                <div className="border-l-2 border-blue-400 pl-3 py-1.5 bg-slate-50 rounded-r-lg">
                  <p className="text-xs text-slate-500 italic leading-relaxed">{evidenceQuote}</p>
                  <p className="text-xs text-blue-500 font-medium mt-1 flex items-center gap-1">
                    <svg className="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" />
                    </svg>
                    {sourceDocument}
                  </p>
                </div>
              )}

              {/* ── pending: 3 action buttons ───────────────────────────── */}
              {status === 'pending' && (
                <div className="flex flex-wrap gap-2 pt-1">
                  {/* Apply — green (AI/System action) */}
                  <button
                    onClick={() => onApply(msg.id, proposal)}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white text-xs font-semibold rounded-lg transition-all active:scale-95 shadow-sm"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                    </svg>
                    Apply to Chart
                  </button>
                  {/* Reject — grey outline */}
                  <button
                    onClick={() => onReject(msg.id)}
                    className="px-3 py-1.5 border border-slate-300 text-slate-600 hover:bg-slate-50 text-xs font-semibold rounded-lg transition-all active:scale-95"
                  >
                    Reject
                  </button>
                  {/* Flag — orange outline (edge-case) */}
                  <button
                    onClick={() => onFlag(msg.id)}
                    className="flex items-center gap-1.5 px-3 py-1.5 border border-orange-400 text-orange-600 hover:bg-orange-50 text-xs font-semibold rounded-lg transition-all active:scale-95"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 21v-4m0 0V5a2 2 0 012-2h6.5l1 1H21l-3 6 3 6h-8.5l-1-1H5a2 2 0 00-2 2zm9-13.5V9" />
                    </svg>
                    Flag as Incorrect
                  </button>
                </div>
              )}

              {/* ── flagged: inline correction input (Edge Case A) ────── */}
              {status === 'flagged' && (
                <div className="space-y-2 pt-1 border-t border-orange-100 mt-2">
                  <p className="text-xs text-orange-600 font-medium">What's incorrect? Tell me the correction.</p>
                  <div className="flex gap-2">
                    <input
                      autoFocus
                      type="text"
                      value={correctionText}
                      onChange={(e) => onCorrectionChange(msg.id, e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') onSubmitCorrection(msg.id, correctionText, proposal);
                      }}
                      placeholder="e.g. It's a thermocouple, not a thermistor"
                      className="flex-1 text-xs px-3 py-2 border border-orange-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-400 placeholder:text-slate-400"
                    />
                    <button
                      onClick={() => onSubmitCorrection(msg.id, correctionText, proposal)}
                      className="px-3 py-2 bg-orange-500 hover:bg-orange-600 text-white text-xs font-semibold rounded-lg transition-all active:scale-95"
                    >
                      Submit
                    </button>
                  </div>
                </div>
              )}

              {/* ── applied ──────────────────────────────────────────── */}
              {status === 'applied' && (
                <div className="flex items-center gap-1.5 pt-1 text-emerald-600 text-xs font-medium">
                  <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                  Applied to chart
                </div>
              )}

              {/* ── rejected ─────────────────────────────────────────── */}
              {status === 'rejected' && (
                <p className="pt-1 text-xs text-slate-400 italic">Discarded</p>
              )}

              {/* ── correction submitted ─────────────────────────────── */}
              {status === 'correction-submitted' && (
                <p className="pt-1 text-xs text-orange-500 italic font-medium">Correction received — re-analyzing…</p>
              )}
            </div>
          </div>
          <p className="pl-9 text-xs text-slate-400">
            {msg.ts.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </p>
        </div>
      </div>
    );
  }

  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN APP — dual-pane view
// ═══════════════════════════════════════════════════════════════════════════════

export default function App() {
  // ── Session ────────────────────────────────────────────────────────────────
  const [sessionConfig, setSessionConfig] = useState(null);

  // ── Core state (spec §5) ────────────────────────────────────────────────────
  const [chartHistory, setChartHistory] = useState([initialChart]); // index 0 = current
  const [messages, setMessages]         = useState([]);
  const [pendingProposal, setPendingProposal] = useState(null); // eslint-disable-line no-unused-vars

  // ── UI state ────────────────────────────────────────────────────────────────
  const [inputText, setInputText] = useState('');
  const [isTyping, setIsTyping]   = useState(false);
  const [flashCell, setFlashCell] = useState(null);   // { rowId, field }
  const [toast, setToast]         = useState(null);   // { message, type }

  const chatScrollRef = useRef(null);
  const inputRef      = useRef(null);
  const currentChart  = chartHistory[0];

  // auto-scroll chat
  useEffect(() => {
    if (chatScrollRef.current) {
      chatScrollRef.current.scrollTop = chatScrollRef.current.scrollHeight;
    }
  }, [messages, isTyping]);

  const showToast = useCallback((message, type = 'success') => setToast({ message, type }), []);

  // ── Send message ─────────────────────────────────────────────────────────
  const handleSend = useCallback(() => {
    const text = inputText.trim();
    if (!text || isTyping) return;

    setMessages((prev) => [...prev, { id: Date.now(), role: 'user', content: text, ts: new Date() }]);
    setInputText('');
    setIsTyping(true);

    const { confident } = checkRetrievalConfidence(text);

    if (!confident) {
      // Edge Case C: missing evidence
      setTimeout(() => {
        setIsTyping(false);
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now() + 1,
            role: 'ai',
            variant: 'missing-evidence',
            content:
              'I cannot locate any evidence regarding a temperature sensor array in the currently indexed documents. Please upload the sensor hardware schematic to proceed.',
            uploadedFile: null,
            ts: new Date(),
          },
        ]);
      }, 1500);
    } else {
      // Happy path
      setTimeout(() => {
        setIsTyping(false);
        const proposal = generateProposal(text, currentChart);
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now() + 1,
            role: 'ai',
            variant: 'proposal',
            content: proposal.proposalText,
            evidenceQuote: proposal.evidenceQuote,
            sourceDocument: proposal.sourceDocument,
            proposal,
            status: 'pending',
            correctionText: '',
            ts: new Date(),
          },
        ]);
        setPendingProposal(proposal);
      }, 1500);
    }
  }, [inputText, isTyping, currentChart]);

  // ── Apply proposal to chart ──────────────────────────────────────────────
  const handleApply = useCallback(
    (msgId, proposal) => {
      const newChart   = applyProposalToChart(currentChart, proposal);
      const newHistory = pushHistory(chartHistory, newChart);
      setChartHistory(newHistory);
      setPendingProposal(null);

      setMessages((prev) => prev.map((m) => (m.id === msgId ? { ...m, status: 'applied' } : m)));

      // 1-second cell flash
      setFlashCell({ rowId: proposal.rowId, field: proposal.field });
      setTimeout(() => setFlashCell(null), 1200);
    },
    [chartHistory, currentChart]
  );

  // ── Reject ───────────────────────────────────────────────────────────────
  const handleReject = useCallback((msgId) => {
    setMessages((prev) => prev.map((m) => (m.id === msgId ? { ...m, status: 'rejected' } : m)));
    setPendingProposal(null);
  }, []);

  // ── Flag as incorrect ────────────────────────────────────────────────────
  const handleFlag = useCallback((msgId) => {
    setMessages((prev) => prev.map((m) => (m.id === msgId ? { ...m, status: 'flagged', correctionText: '' } : m)));
  }, []);

  // ── Correction text change ───────────────────────────────────────────────
  const handleCorrectionChange = useCallback((msgId, text) => {
    setMessages((prev) => prev.map((m) => (m.id === msgId ? { ...m, correctionText: text } : m)));
  }, []);

  // ── Submit correction (Edge Case A) ──────────────────────────────────────
  const handleSubmitCorrection = useCallback(
    (msgId, correctionText, originalProposal) => {
      if (!correctionText.trim()) return;

      setMessages((prev) => prev.map((m) => (m.id === msgId ? { ...m, status: 'correction-submitted' } : m)));
      setMessages((prev) => [
        ...prev,
        { id: Date.now(), role: 'user', variant: 'correction', content: correctionText, ts: new Date() },
      ]);

      setIsTyping(true);
      setTimeout(() => {
        setIsTyping(false);
        const revised = generateProposal(originalProposal.proposalText, currentChart, correctionText);
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now() + 1,
            role: 'ai',
            variant: 'proposal',
            content: revised.proposalText,
            evidenceQuote: revised.evidenceQuote,
            sourceDocument: revised.sourceDocument,
            proposal: revised,
            status: 'pending',
            correctionText: '',
            isRevised: true,
            ts: new Date(),
          },
        ]);
        setPendingProposal(revised);
      }, 1500);
    },
    [currentChart]
  );

  // ── Schematic upload (Edge Case C follow-through) ────────────────────────
  const handleSchematicUpload = useCallback(
    (msgId) => {
      setMessages((prev) =>
        prev.map((m) => (m.id === msgId ? { ...m, uploadedFile: 'temperature_sensor_array_schematic.pdf' } : m))
      );

      setTimeout(() => {
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now() + 1,
            role: 'ai',
            variant: 'system',
            content: "Thanks — I've indexed the schematic. Re-running retrieval for your original request...",
            ts: new Date(),
          },
        ]);

        setIsTyping(true);
        setTimeout(() => {
          setIsTyping(false);
          const proposal = generateProposal('temperature sensor array schematic hardware', currentChart);
          setMessages((prev) => [
            ...prev,
            {
              id: Date.now() + 2,
              role: 'ai',
              variant: 'proposal',
              content: proposal.proposalText,
              evidenceQuote: proposal.evidenceQuote,
              sourceDocument: proposal.sourceDocument,
              proposal,
              status: 'pending',
              correctionText: '',
              ts: new Date(),
            },
          ]);
          setPendingProposal(proposal);
        }, 1500);
      }, 800);
    },
    [currentChart]
  );

  // ── Undo last change (Edge Case B) ───────────────────────────────────────
  const handleUndo = useCallback(() => {
    if (chartHistory.length <= 1) return;
    setChartHistory((prev) => undoLastChange(prev));
    showToast('Reverted to previous version.', 'warning');
  }, [chartHistory.length, showToast]);

  // ── Export chart ─────────────────────────────────────────────────────────
  const handleExport = useCallback(async () => {
    await exportChart(currentChart);
    showToast('Success: Rule 11-compliant Evidence of Use Chart Exported.', 'success');
  }, [currentChart, showToast]);

  // ── Enter to send ────────────────────────────────────────────────────────
  const handleKeyDown = useCallback(
    (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    },
    [handleSend]
  );

  // ── Session start ────────────────────────────────────────────────────────
  const handleSessionStart = useCallback((config) => {
    setSessionConfig(config);
    setMessages([
      {
        id: Date.now(),
        role: 'ai',
        variant: 'system',
        content: `Session initialized. I've indexed "${config.claimFile}" and "${config.evidenceFile}".${
          config.instructions ? ` Operating under: "${config.instructions}".` : ''
        } Ask me to refine any claim element and I'll surface supporting evidence.`,
        ts: new Date(),
      },
    ]);
  }, []);

  // ── Setup screen ─────────────────────────────────────────────────────────
  if (!sessionConfig) {
    return <SetupScreen onStart={handleSessionStart} />;
  }

  const historyUsed = chartHistory.length;
  const canUndo     = historyUsed > 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // DUAL-PANE RENDER
  // ═══════════════════════════════════════════════════════════════════════════
  return (
    <div className="h-screen flex flex-col overflow-hidden bg-slate-50">

      {/* ── Top nav bar ──────────────────────────────────────────────────── */}
      <header className="flex-shrink-0 h-14 bg-white border-b border-slate-200 px-6 flex items-center justify-between shadow-sm z-10">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center shadow-sm">
            <svg className="w-4.5 h-4.5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
            </svg>
          </div>
          <span className="text-lg font-bold text-slate-800 tracking-tight">iLumos</span>
          <span className="hidden sm:inline text-xs text-slate-400 font-medium px-2 py-1 bg-slate-100 rounded-full">
            AI Claim Chart Refinement
          </span>
        </div>
        <div className="flex items-center gap-2 text-xs text-slate-500">
          <span className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse inline-block" />
          Session active
        </div>
      </header>

      {/* ── Main content (flex row, fills remaining height) ──────────────── */}
      <div className="flex flex-1 min-h-0 flex-col lg:flex-row">

        {/* ── LEFT PANE: Chat (35%) ──────────────────────────────────────── */}
        <aside className="lg:w-[35%] w-full flex flex-col bg-white border-r border-slate-200 min-h-0 lg:h-full h-[50vh]">

          {/* chat header */}
          <div className="flex-shrink-0 px-4 py-3 border-b border-slate-100">
            <h2 className="text-sm font-semibold text-slate-700">Refinement Assistant</h2>
            <p className="text-xs text-slate-400 mt-0.5 truncate">
              {sessionConfig.claimFile} · {sessionConfig.evidenceFile}
            </p>
          </div>

          {/* message history */}
          <div ref={chatScrollRef} className="flex-1 overflow-y-auto custom-scrollbar px-4 py-4 space-y-3 min-h-0">
            {messages.map((msg) => (
              <MessageBubble
                key={msg.id}
                msg={msg}
                onApply={handleApply}
                onReject={handleReject}
                onFlag={handleFlag}
                onSubmitCorrection={handleSubmitCorrection}
                onSchematicUpload={handleSchematicUpload}
                onCorrectionChange={handleCorrectionChange}
              />
            ))}
            {isTyping && <TypingIndicator />}
          </div>

          {/* undo + input area */}
          <div className="flex-shrink-0 border-t border-slate-100 px-4 pb-4 pt-3 space-y-2">
            {/* Undo button — orange (edge-case affordance per spec) */}
            <button
              onClick={handleUndo}
              disabled={!canUndo}
              className={`w-full flex items-center justify-center gap-2 py-2 rounded-xl text-xs font-semibold border transition-all duration-200 ${
                canUndo
                  ? 'border-orange-400 text-orange-600 hover:bg-orange-50 active:scale-[0.99]'
                  : 'border-slate-200 text-slate-300 cursor-not-allowed'
              }`}
            >
              <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" />
              </svg>
              Undo Last Change
              <span className="opacity-60 font-normal">({historyUsed - 1} / 19 states used)</span>
            </button>

            {/* text area + send */}
            <div className="flex gap-2 items-end">
              <textarea
                ref={inputRef}
                value={inputText}
                onChange={(e) => setInputText(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="Ask me to refine a claim element…"
                rows={2}
                className="flex-1 resize-none px-4 py-3 text-sm border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent placeholder:text-slate-400 text-slate-700 bg-slate-50 transition-all custom-scrollbar"
              />
              {/* Send — blue (user action) */}
              <button
                onClick={handleSend}
                disabled={!inputText.trim() || isTyping}
                className={`flex-shrink-0 w-10 h-10 rounded-xl flex items-center justify-center transition-all duration-200 ${
                  inputText.trim() && !isTyping
                    ? 'bg-blue-600 hover:bg-blue-700 text-white shadow-md shadow-blue-200 active:scale-95'
                    : 'bg-slate-100 text-slate-300 cursor-not-allowed'
                }`}
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
                </svg>
              </button>
            </div>
            <p className="text-xs text-center text-slate-400">↵ Enter to send · Shift+Enter for newline</p>
          </div>
        </aside>

        {/* ── RIGHT PANE: Claim Chart (65%) ─────────────────────────────── */}
        <main className="lg:w-[65%] w-full flex-1 overflow-y-auto custom-scrollbar bg-slate-50 p-6 space-y-4 min-h-0">

          {/* chart card */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">

            {/* card header */}
            <div className="px-6 py-4 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800">Evidence of Use (EoU) Chart</h2>
                <p className="text-xs text-slate-400 mt-0.5">Patent US123456 · Rule 11-Compliant</p>
              </div>
              {/* Export — blue (user action) */}
              <button
                onClick={handleExport}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg shadow-sm shadow-blue-200 transition-all active:scale-95"
              >
                <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                Export to Word
              </button>
            </div>

            {/* table */}
            <div className="overflow-x-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="text-left px-4 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider w-8">#</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider w-[28%]">Patent Claim Element</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider w-[34%]">Accused Product Feature</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">AI Reasoning (Evidence)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {currentChart.map((row) => {
                    const isFlashing = flashCell?.rowId === row.id && flashCell?.field === 'aiReasoning';
                    return (
                      <tr key={row.id} className="hover:bg-slate-50/70 transition-colors duration-150">
                        <td className="px-4 py-4 align-top">
                          <span className="inline-flex items-center justify-center w-6 h-6 bg-blue-50 text-blue-600 text-xs font-bold rounded-full">
                            {row.id}
                          </span>
                        </td>
                        <td className="px-4 py-4 align-top text-xs text-slate-700 leading-relaxed font-medium">
                          {row.claimElement}
                        </td>
                        <td className="px-4 py-4 align-top text-xs text-slate-500 leading-relaxed italic">
                          {row.accusedFeature}
                        </td>
                        <td
                          className={`px-4 py-4 align-top text-xs leading-relaxed transition-all duration-300 ${
                            isFlashing ? 'cell-flash' : ''
                          } ${row.weakEvidence ? 'border-l-2 border-amber-400' : ''}`}
                        >
                          <span className={row.weakEvidence ? 'text-amber-800' : 'text-slate-600'}>
                            {row.aiReasoning}
                          </span>
                          {row.weakEvidence && (
                            <span className="mt-1.5 ml-1 inline-flex items-center gap-1 text-xs text-amber-600 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-full font-medium">
                              <svg className="w-3 h-3 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
                              </svg>
                              Weak evidence
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* card footer */}
            <div className="px-6 py-3 bg-slate-50 border-t border-slate-100 flex items-center justify-between text-xs text-slate-400">
              <span>{currentChart.length} claim elements</span>
              <span>Version {historyUsed} · {historyUsed - 1} change{historyUsed !== 2 ? 's' : ''} applied</span>
            </div>
          </div>

          {/* quick prompts */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4">
            <p className="text-xs font-semibold text-slate-500 mb-2.5">Quick Prompts</p>
            <div className="flex flex-wrap gap-2">
              {[
                'Strengthen the ML algorithm evidence for element 3',
                'Find FCC certification for the wireless module',
                'Analyze the temperature sensor array',
                'Check PIR sensor specs for element 2',
              ].map((prompt) => (
                <button
                  key={prompt}
                  onClick={() => { setInputText(prompt); inputRef.current?.focus(); }}
                  className="text-xs px-3 py-1.5 bg-slate-50 hover:bg-blue-50 text-slate-600 hover:text-blue-700 border border-slate-200 hover:border-blue-300 rounded-lg transition-all duration-150"
                >
                  {prompt}
                </button>
              ))}
            </div>
          </div>

          {/* color-coding legend */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4">
            <p className="text-xs font-semibold text-slate-500 mb-2.5">Interface Color Guide</p>
            <div className="flex flex-wrap gap-4 text-xs">
              <span className="flex items-center gap-1.5"><span className="w-3 h-3 bg-blue-600 rounded-full flex-shrink-0" />Blue = User Action</span>
              <span className="flex items-center gap-1.5"><span className="w-3 h-3 bg-emerald-500 rounded-full flex-shrink-0" />Green = AI / System</span>
              <span className="flex items-center gap-1.5"><span className="w-3 h-3 bg-orange-500 rounded-full flex-shrink-0" />Orange = Edge Case</span>
              <span className="flex items-center gap-1.5"><span className="w-3 h-3 bg-amber-400 rounded-full flex-shrink-0" />Amber = Weak Evidence</span>
            </div>
          </div>
        </main>
      </div>

      {/* Toast */}
      {toast && <Toast message={toast.message} type={toast.type} onDismiss={() => setToast(null)} />}
    </div>
  );
}
