import { useMemo, useState } from 'react';
import { ComingSoonMemoTab } from './components/ComingSoonMemoTab';
import { ICMemoPreview } from './components/ICMemoPreview';
import { LandingScreen } from './components/LandingScreen';
import { ManualInputsTab } from './components/ManualInputsTab';
import { resolveICMemo } from './engines/icMemoEngine';
import { MEMO_ENGINES } from './engines/memoRegistry';
import { generateICMemoDocxBlob } from './exporters/icMemoDocx';
import {
  detectDocumentKeyFromFile,
  parseAccountStatementPdf,
  parseBranchProposalFile,
  parseCibReportPdf,
  parseCostOfFundFile,
  parseOutstandingPdf,
  parseOverduePdf,
  parsePartyPerformancePdf,
} from './parsers/documentParsers';
import type {
  AiSectionsData,
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  MemoTabKey,
} from './types/memo';

const INITIAL_DOC_METAS: Record<DocumentKey, DocumentUploadMeta> = {
  branchProposal: {
    key: 'branchProposal',
    label: 'Branch Proposal',
    acceptedFormats: '.pdf,image/*',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  outstanding: {
    key: 'outstanding',
    label: 'Outstanding',
    acceptedFormats: '.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  partyPerformance: {
    key: 'partyPerformance',
    label: 'Party Performance',
    acceptedFormats: '.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  overdue: {
    key: 'overdue',
    label: 'Overdue',
    acceptedFormats: '.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  accountStatement: {
    key: 'accountStatement',
    label: 'Account Statement',
    acceptedFormats: '.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  cibReport: {
    key: 'cibReport',
    label: 'CIB Report',
    acceptedFormats: '.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  costOfFund: {
    key: 'costOfFund',
    label: 'Cost of Fund',
    acceptedFormats: '.xlsx,.xls,.csv,.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
};

const INITIAL_EXTRACTED_DOCS: ExtractedDocumentsState = {
  branchProposal: null,
  outstanding: null,
  partyPerformance: null,
  overdue: null,
  accountStatement: null,
  cibReport: null,
  costOfFund: null,
};

const INITIAL_MANUAL_INPUTS: ManualInputsData = {
  proposalDate: '',
  waiverAmount: '',
  settlementAmount: '',
  settlementPeriod: '',
  waiverRecClientBranch: '',
  waiverRecHO: '',
  sourceOfRepayment: '',
};

const INITIAL_AI_SECTIONS: AiSectionsData = {
  justification: null,
  branchVisit: null,
  samdObservation: null,
};

const TABS: Array<{ id: MemoTabKey; label: string }> = [
  { id: 'manual', label: 'Manual Inputs' },
  { id: 'ic', label: 'IC Memo' },
  { id: 'mec', label: 'MEC Memo' },
  { id: 'board', label: 'Board Memo' },
  { id: 'sanction', label: 'Sanction' },
];

export default function App() {
  const [entered, setEntered] = useState(false);
  const [activeTab, setActiveTab] = useState<MemoTabKey>('manual');

  const [docMetas, setDocMetas] = useState<Record<DocumentKey, DocumentUploadMeta>>(INITIAL_DOC_METAS);
  const [extractedDocs, setExtractedDocs] = useState<ExtractedDocumentsState>(INITIAL_EXTRACTED_DOCS);
  const [manualInputs, setManualInputs] = useState<ManualInputsData>(INITIAL_MANUAL_INPUTS);
  const [aiSections, setAiSections] = useState<AiSectionsData>(INITIAL_AI_SECTIONS);
  const [overrides, setOverrides] = useState<Record<string, string>>({});

  const [loadingSamples, setLoadingSamples] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{
    active: boolean;
    total: number;
    completed: number;
  }>({ active: false, total: 0, completed: 0 });
  const [generatingSection, setGeneratingSection] = useState<
    'justification' | 'branchVisit' | 'samdObservation' | null
  >(null);
  const [aiError, setAiError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);

  const resolvedICMemo = useMemo(
    () => resolveICMemo(extractedDocs, manualInputs, aiSections, overrides),
    [extractedDocs, manualInputs, aiSections, overrides]
  );

  const updateDocStage = (key: DocumentKey, stage: string) => {
    setDocMetas((prev) => ({
      ...prev,
      [key]: {
        ...prev[key],
        stage,
      },
    }));
  };

  const handleFileUpload = async (key: DocumentKey, file: File) => {
    setDocMetas((prev) => ({
      ...prev,
      [key]: {
        ...prev[key],
        parsing: true,
        stage: 'Starting...',
        error: null,
        fileName: file.name,
      },
    }));

    const onStage = (stage: string) => updateDocStage(key, stage);

    try {
      if (key === 'branchProposal') {
        const bpData = await parseBranchProposalFile(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, branchProposal: bpData }));
      } else if (key === 'outstanding') {
        const res = await parseOutstandingPdf(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, outstanding: res.data }));
      } else if (key === 'partyPerformance') {
        const res = await parsePartyPerformancePdf(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, partyPerformance: res.data }));
      } else if (key === 'overdue') {
        const res = await parseOverduePdf(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, overdue: res.data }));
      } else if (key === 'accountStatement') {
        const res = await parseAccountStatementPdf(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, accountStatement: res.data }));
      } else if (key === 'cibReport') {
        const res = await parseCibReportPdf(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, cibReport: res.data }));
      } else if (key === 'costOfFund') {
        const res = await parseCostOfFundFile(file, onStage);
        setExtractedDocs((prev) => ({ ...prev, costOfFund: res.data }));
      }

      setDocMetas((prev) => ({
        ...prev,
        [key]: {
          ...prev[key],
          uploaded: true,
          parsing: false,
          stage: null,
          error: null,
          fileName: file.name,
        },
      }));
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse document.';
      setDocMetas((prev) => ({
        ...prev,
        [key]: {
          ...prev[key],
          uploaded: false,
          parsing: false,
          stage: null,
          error: msg,
        },
      }));
    }
  };

  const handleLoadSampleDocuments = async () => {
    setLoadingSamples(true);
    const sampleFiles: Array<{ key: DocumentKey; url: string; name: string }> = [
      { key: 'outstanding', url: '/assets/aistudio/Outstanding_18.11.26_644e801e.pdf', name: 'Outstanding_18.11.26.pdf' },
      { key: 'partyPerformance', url: '/assets/aistudio/Party_Performance_55885ce3.pdf', name: 'Party_Performance.pdf' },
      { key: 'overdue', url: '/assets/aistudio/Overdue_a9f320c0.pdf', name: 'Overdue.pdf' },
      { key: 'accountStatement', url: '/assets/aistudio/Account_Statement_4197efed.pdf', name: 'Account_Statement.pdf' },
      { key: 'cibReport', url: '/assets/aistudio/CIB_Report_dec9ba1f.pdf', name: 'CIB_Report.pdf' },
      { key: 'costOfFund', url: '/assets/aistudio/Cost_of_Fund_3c8b0d9a.pdf', name: 'Cost_of_Fund.pdf' },
      { key: 'branchProposal', url: '/assets/aistudio/Branch_Proposal_a3692300.pdf', name: 'Branch_Proposal.pdf' },
    ];

    setBatchProgress({ active: true, total: sampleFiles.length, completed: 0 });

    try {
      await Promise.all(
        sampleFiles.map(async ({ key, url, name }) => {
          updateDocStage(key, 'Fetching sample...');
          try {
            const response = await fetch(url);
            if (!response.ok) throw new Error(`Sample file missing: ${name}`);
            const blob = await response.blob();
            const file = new File([blob], name, { type: blob.type || 'application/pdf' });
            await handleFileUpload(key, file);
          } finally {
            setBatchProgress((prev) => ({ ...prev, completed: prev.completed + 1 }));
          }
        })
      );
    } finally {
      setLoadingSamples(false);
      setBatchProgress({ active: false, total: 0, completed: 0 });
    }
  };

  const handleBatchUpload = async (files: FileList | File[]) => {
    const fileList = Array.from(files);
    const unmatched: string[] = [];

    setBatchProgress({ active: true, total: fileList.length, completed: 0 });

    try {
      await Promise.all(
        fileList.map(async (file) => {
          try {
            const key = await detectDocumentKeyFromFile(file);
            if (key) {
              await handleFileUpload(key, file);
            } else {
              unmatched.push(file.name);
            }
          } finally {
            setBatchProgress((prev) => ({ ...prev, completed: prev.completed + 1 }));
          }
        })
      );
    } finally {
      setBatchProgress({ active: false, total: 0, completed: 0 });
    }

    return unmatched;
  };

  const handleManualChange = (field: keyof ManualInputsData, value: string) => {
    setManualInputs((prev) => ({ ...prev, [field]: value }));
  };

  const handleOverrideChange = (code: string, value: string) => {
    setOverrides((prev) => ({ ...prev, [code]: value }));
  };

  const handleResetOverrides = () => {
    setOverrides({});
  };

  const handleGenerateAiSection = async (
    section: 'justification' | 'branchVisit' | 'samdObservation'
  ) => {
    setGeneratingSection(section);
    setAiError(null);

    try {
      const context = {
        placeholders: resolvedICMemo.placeholders,
        sourceOfRepayment: resolvedICMemo.sourceOfRepayment,
        branchProposal: extractedDocs.branchProposal,
      };

      const response = await fetch('/api/gemini/generate-section', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ section, context }),
      });

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Generation failed (${response.status}).`);
      }

      const data = await response.json();
      setAiSections((prev) => ({
        ...prev,
        [section]: data.text ?? '',
      }));
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate section.');
    } finally {
      setGeneratingSection(null);
    }
  };

  const handleUpdateAiSection = (
    section: 'justification' | 'branchVisit' | 'samdObservation',
    value: string | null
  ) => {
    setAiSections((prev) => ({ ...prev, [section]: value }));
  };

  const handleDownloadICDocx = async () => {
    setDownloading(true);
    try {
      const blob = await generateICMemoDocxBlob(resolvedICMemo);
      const acct = resolvedICMemo.placeholders['#2#'] || 'Draft';
      const safeName = acct.replace(/[^A-Za-z0-9_-]/g, '_');
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `IC_Memo_${safeName}.docx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } finally {
      setDownloading(false);
    }
  };

  if (!entered) {
    return <LandingScreen onEnter={() => setEntered(true)} />;
  }

  const mecEngine = MEMO_ENGINES.find((e) => e.id === 'mec')!;
  const boardEngine = MEMO_ENGINES.find((e) => e.id === 'board')!;
  const sanctionEngine = MEMO_ENGINES.find((e) => e.id === 'sanction')!;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Three-Zone Top Bar */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between gap-8">
          {/* Zone 1: Brand Title */}
          <button
            type="button"
            onClick={() => setEntered(false)}
            className="font-display text-lg font-bold tracking-tight text-slate-900 hover:text-emerald-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            Memo Engine by Tuhin
          </button>

          {/* Zone 2: 5 Required Tabs in Exact Order */}
          <nav
            aria-label="Memo Workspace Tabs"
            className="flex items-center gap-1 overflow-x-auto"
          >
            {TABS.map((tab) => {
              const active = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3.5 py-2 text-sm font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    active
                      ? 'bg-emerald-800 text-white'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </nav>

          {/* Zone 3: Primary Action */}
          <div className="hidden sm:flex items-center shrink-0">
            {activeTab === 'ic' ? (
              <button
                type="button"
                onClick={handleDownloadICDocx}
                disabled={downloading}
                className="px-4 py-2 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-60"
              >
                {downloading ? 'Building...' : 'Download .docx'}
              </button>
            ) : (
              <button
                type="button"
                onClick={() => setActiveTab('ic')}
                className="px-4 py-2 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                IC Memo
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Active Tab Content */}
      <main className="flex-1 px-4 sm:px-6 py-6">
        {activeTab === 'manual' && (
          <ManualInputsTab
            docMetas={docMetas}
            extractedDocs={extractedDocs}
            manualInputs={manualInputs}
            overrides={overrides}
            resolvedMemo={resolvedICMemo}
            onFileUpload={handleFileUpload}
            onBatchUpload={handleBatchUpload}
            onManualChange={handleManualChange}
            onOverrideChange={handleOverrideChange}
            onResetOverrides={handleResetOverrides}
            onLoadSampleDocuments={handleLoadSampleDocuments}
            loadingSamples={loadingSamples}
            batchProgress={batchProgress}
            onGoToICMemo={() => setActiveTab('ic')}
          />
        )}

        {activeTab === 'ic' && (
          <ICMemoPreview
            memo={resolvedICMemo}
            aiSections={aiSections}
            overrides={overrides}
            onOverrideChange={handleOverrideChange}
            onGenerateAiSection={handleGenerateAiSection}
            onUpdateAiSection={handleUpdateAiSection}
            generatingSection={generatingSection}
            aiError={aiError}
            onDownloadDocx={handleDownloadICDocx}
            downloading={downloading}
          />
        )}

        {activeTab === 'mec' && (
          <ComingSoonMemoTab
            engine={mecEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
          />
        )}

        {activeTab === 'board' && (
          <ComingSoonMemoTab
            engine={boardEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
          />
        )}

        {activeTab === 'sanction' && (
          <ComingSoonMemoTab
            engine={sanctionEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
          />
        )}
      </main>
    </div>
  );
}
