import { useEffect, useMemo, useRef, useState } from 'react';
import { CheckCircle2, RefreshCw } from 'lucide-react';
import { ComingSoonMemoTab } from './components/ComingSoonMemoTab';
import { ICMemoPreview } from './components/ICMemoPreview';
import { LandingScreen } from './components/LandingScreen';
import { ManualInputsTab } from './components/ManualInputsTab';
import { resolveICMemo } from './engines/icMemoEngine';
import { MEMO_ENGINES } from './engines/memoRegistry';
import { inspectDocxTemplate } from './exporters/docxTemplateEngine';
import { generateMemoPdfBlob, generateSynchronizedDocxBlob } from './exporters/pdfExporter';
import {
  detectDocumentKeyFromFile,
  parseAccountStatementPdf,
  parseBranchProposalFile,
  parseCibReportPdf,
  parseCostOfFundFile,
  parseOutstandingPdf,
  parseOverduePdf,
  parsePartyPerformancePdf,
  prepareBranchProposalSource,
  warmBranchProposalSource,
} from './parsers/documentParsers';
import type {
  AiSectionsData,
  AppliedSyncState,
  BatchScrutinyState,
  BranchProposalSource,
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  MemoDocType,
  MemoTabKey,
  MemoTemplateRecord,
  ScrutinySectionKey,
} from './types/memo';
import {
  deleteTemplateRecord,
  loadAllTemplateRecords,
  saveTemplateRecord,
} from './utils/templateStorage';

const INITIAL_DOC_METAS: Record<DocumentKey, DocumentUploadMeta> = {
  branchProposal: {
    key: 'branchProposal',
    label: 'Branch Proposal',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls,image/*',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  outstanding: {
    key: 'outstanding',
    label: 'Outstanding',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  partyPerformance: {
    key: 'partyPerformance',
    label: 'Party Performance',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  overdue: {
    key: 'overdue',
    label: 'Overdue',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  accountStatement: {
    key: 'accountStatement',
    label: 'Account Statement',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  cibReport: {
    key: 'cibReport',
    label: 'CIB Report',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  costOfFund: {
    key: 'costOfFund',
    label: 'Cost of Fund',
    acceptedFormats: '.xlsx,.xls,.csv,.docx,.pdf',
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
  sec1NatureOfBusiness: null,
  sec2LiabilityOurBank: null,
  sec3LiabilityOtherUnits: null,
  sec4CibStatus: null,
  sec5LegalAction: null,
  sec6AuditCompliance: null,
  sec7Security: null,
  sec8SourceOfRepayment: null,
  sec9HoSanctions: null,
  sec10ClientParticulars: null,
  sec11Financials: null,
  branchVisit: null,
  sec13BranchRecommendation: null,
  sec14IccdRecommendation: null,
  samdObservation: null,
};

const TABS: Array<{ id: MemoTabKey; label: string }> = [
  { id: 'manual', label: 'Manual Inputs' },
  { id: 'ic', label: 'IC Memo' },
  { id: 'mec', label: 'MEC Memo' },
  { id: 'board', label: 'Board Memo' },
  { id: 'sanction', label: 'Sanction' },
];

const MANUAL_FIELD_TO_CODE: Partial<Record<keyof ManualInputsData, string>> = {
  proposalDate: '#1#',
  waiverAmount: '#4#',
  settlementAmount: '#27#',
  settlementPeriod: '#28#',
  waiverRecClientBranch: '#45#',
  waiverRecHO: '#47#',
};

const ALL_SCRUTINY_SECTIONS: ScrutinySectionKey[] = [
  'justification',
  'sec1NatureOfBusiness',
  'sec2LiabilityOurBank',
  'sec3LiabilityOtherUnits',
  'sec4CibStatus',
  'sec5LegalAction',
  'sec6AuditCompliance',
  'sec7Security',
  'sec8SourceOfRepayment',
  'sec9HoSanctions',
  'sec10ClientParticulars',
  'sec11Financials',
  'branchVisit',
  'sec13BranchRecommendation',
  'sec14IccdRecommendation',
  'samdObservation',
];

const INITIAL_BATCH_SCRUTINY: BatchScrutinyState = {
  running: false,
  total: 0,
  targetSections: [],
  completedSections: [],
  retryingSections: [],
  failedSections: [],
  insufficientSections: [],
  skippedManualSections: [],
};

const DOC_FILE_PREFIX: Record<MemoDocType, string> = {
  ic: 'IC_Memo',
  mec: 'MEC_Memo',
  board: 'Board_Memo',
  sanction: 'Sanction',
};

export default function App() {
  const [entered, setEntered] = useState(false);
  const [activeTab, setActiveTab] = useState<MemoTabKey>('manual');

  const [docMetas, setDocMetas] = useState<Record<DocumentKey, DocumentUploadMeta>>(INITIAL_DOC_METAS);
  const [extractedDocs, setExtractedDocs] = useState<ExtractedDocumentsState>(INITIAL_EXTRACTED_DOCS);
  const [manualInputs, setManualInputs] = useState<ManualInputsData>(INITIAL_MANUAL_INPUTS);
  const [aiSections, setAiSections] = useState<AiSectionsData>(INITIAL_AI_SECTIONS);
  const [overrides, setOverrides] = useState<Record<string, string>>({});

  const [templates, setTemplates] = useState<Record<MemoDocType, MemoTemplateRecord | null>>({
    ic: null,
    mec: null,
    board: null,
    sanction: null,
  });

  const [loadingSamples, setLoadingSamples] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{
    active: boolean;
    total: number;
    completed: number;
  }>({ active: false, total: 0, completed: 0 });
  const [branchProposalFile, setBranchProposalFile] = useState<File | null>(null);
  const [scrutinizingSection, setScrutinizingSection] = useState<ScrutinySectionKey | null>(null);
  const [batchScrutiny, setBatchScrutiny] = useState<BatchScrutinyState>(INITIAL_BATCH_SCRUTINY);
  const [manualEditedSections, setManualEditedSections] = useState<
    Partial<Record<ScrutinySectionKey, boolean>>
  >({});
  const [aiError, setAiError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [downloadingPdf, setDownloadingPdf] = useState(false);

  const serverCachedSourceIdRef = useRef<string | null>(null);
  const batchAbortRef = useRef<AbortController | null>(null);

  const resolvedICMemo = useMemo(
    () => resolveICMemo(extractedDocs, manualInputs, aiSections, overrides),
    [extractedDocs, manualInputs, aiSections, overrides]
  );

  const [appliedSnapshot, setAppliedSnapshot] = useState<AppliedSyncState>(() => ({
    appliedAt: null,
    version: 0,
    memo: resolveICMemo(INITIAL_EXTRACTED_DOCS, INITIAL_MANUAL_INPUTS, INITIAL_AI_SECTIONS, {}),
  }));
  const [hasPendingChanges, setHasPendingChanges] = useState(false);
  const [applySignal, setApplySignal] = useState(0);
  const [applyNotice, setApplyNotice] = useState<string | null>(null);

  useEffect(() => {
    void loadAllTemplateRecords().then((loaded) => {
      setTemplates(loaded);
    });
  }, []);

  useEffect(() => {
    if (!applyNotice) return;
    const timer = window.setTimeout(() => setApplyNotice(null), 5000);
    return () => window.clearTimeout(timer);
  }, [applyNotice]);

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
        setBranchProposalFile(file);
        setExtractedDocs((prev) => ({ ...prev, branchProposal: bpData }));
        serverCachedSourceIdRef.current = null;
        setBatchScrutiny(INITIAL_BATCH_SCRUTINY);
        // Pre-warm client page images + register server cache silently so AI calls start immediately
        warmBranchProposalSource(file);
        void prepareBranchProposalSource(file)
          .then(async (src) => {
            if (!src.sourceId) return;
            const resp = await fetch('/api/gemini/cache-source', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ sourceId: src.sourceId, branchProposalSource: src }),
            });
            if (resp.ok) {
              serverCachedSourceIdRef.current = src.sourceId;
            }
          })
          .catch(() => {
            // Silent pre-warm fallback; on-demand scrutiny sends source directly if needed
          });
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
      setHasPendingChanges(true);
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
    if (field === 'sourceOfRepayment') {
      setManualEditedSections((prev) => ({
        ...prev,
        sec8SourceOfRepayment: value.trim() !== '',
      }));
    }
    const mappedCode = MANUAL_FIELD_TO_CODE[field];
    if (mappedCode) {
      setOverrides((prev) => {
        if (prev[mappedCode] === undefined) return prev;
        const next = { ...prev };
        delete next[mappedCode];
        return next;
      });
    }
    setHasPendingChanges(true);
  };

  const handleManualSourceOfRepaymentChange = (value: string) => {
    setManualInputs((prev) => ({ ...prev, sourceOfRepayment: value }));
    setAiSections((prev) => ({ ...prev, sec8SourceOfRepayment: value }));
    setManualEditedSections((prev) => ({ ...prev, sec8SourceOfRepayment: true }));
    setHasPendingChanges(true);
  };

  const handleOverrideChange = (code: string, value: string) => {
    setOverrides((prev) => ({ ...prev, [code]: value }));
    if (value.trim() !== '') {
      if (code === '#51#') {
        setManualEditedSections((prev) => ({ ...prev, justification: true }));
      } else if (code === '#54#') {
        setManualEditedSections((prev) => ({ ...prev, sec1NatureOfBusiness: true }));
      } else if (code === '#56#' || code === '#58#') {
        setManualEditedSections((prev) => ({ ...prev, sec2LiabilityOurBank: true }));
      }
    }
    setHasPendingChanges(true);
  };

  const handleResetOverrides = () => {
    setOverrides({});
    setHasPendingChanges(true);
  };

  const clearOverridesForSection = (section: ScrutinySectionKey) => {
    const codesToClear =
      section === 'justification'
        ? ['#51#']
        : section === 'sec1NatureOfBusiness'
          ? ['#54#']
          : section === 'sec2LiabilityOurBank'
            ? ['#56#', '#58#']
            : [];
    if (codesToClear.length === 0) return;
    setOverrides((prev) => {
      let changed = false;
      const next = { ...prev };
      for (const c of codesToClear) {
        if (next[c] !== undefined) {
          delete next[c];
          changed = true;
        }
      }
      return changed ? next : prev;
    });
  };

  const buildScrutinyContext = () => ({
    placeholders: resolvedICMemo.placeholders,
    sourceOfRepayment: resolvedICMemo.sourceOfRepayment,
    branchVisitReport: resolvedICMemo.branchVisitReport,
    justificationLines: resolvedICMemo.justificationLines,
    sec13BranchRecommendation: resolvedICMemo.sec13BranchRecommendation,
    sec14IccdRecommendation: resolvedICMemo.sec14IccdRecommendation,
    extractedSummaries: {
      outstanding: extractedDocs.outstanding,
      partyPerformance: extractedDocs.partyPerformance,
      overdue: extractedDocs.overdue,
      accountStatement: extractedDocs.accountStatement,
      cibReport: extractedDocs.cibReport,
      costOfFund: extractedDocs.costOfFund,
    },
    branchProposal: extractedDocs.branchProposal,
  });

  const applyScrutinizedSectionResult = (
    section: ScrutinySectionKey,
    result: Record<string, unknown>
  ) => {
    const isNarrative =
      section === 'justification' ||
      section === 'sec1NatureOfBusiness' ||
      section === 'sec8SourceOfRepayment' ||
      section === 'branchVisit' ||
      section === 'sec13BranchRecommendation' ||
      section === 'sec14IccdRecommendation' ||
      section === 'samdObservation';

    if (isNarrative) {
      const text = String(result.text ?? '').trim();
      setAiSections((prev) => ({
        ...prev,
        [section]: text,
      }));
      if (section === 'sec8SourceOfRepayment') {
        setManualInputs((prev) => ({ ...prev, sourceOfRepayment: '' }));
      }
    } else {
      setAiSections((prev) => ({
        ...prev,
        [section]: result,
      }));
    }

    clearOverridesForSection(section);
    setManualEditedSections((prev) => ({ ...prev, [section]: false }));
    setHasPendingChanges(true);
  };

  const fetchWithSourceCache = async (
    url: string,
    payload: Record<string, unknown>,
    source: BranchProposalSource,
    signal?: AbortSignal
  ): Promise<Response> => {
    const canReuseServerCache =
      Boolean(source.sourceId) && serverCachedSourceIdRef.current === source.sourceId;

    const firstResponse = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal,
      body: JSON.stringify(
        canReuseServerCache
          ? { ...payload, sourceId: source.sourceId }
          : { ...payload, sourceId: source.sourceId, branchProposalSource: source }
      ),
    });

    if (firstResponse.status === 404 && canReuseServerCache) {
      serverCachedSourceIdRef.current = null;
      return fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal,
        body: JSON.stringify({
          ...payload,
          sourceId: source.sourceId,
          branchProposalSource: source,
        }),
      });
    }

    return firstResponse;
  };

  const handleScrutinizeSection = async (
    section: ScrutinySectionKey,
    forceOverwrite = false
  ) => {
    if (batchScrutiny.running) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }

    const isProtected =
      Boolean(manualEditedSections[section]) ||
      (section === 'sec8SourceOfRepayment' && manualInputs.sourceOfRepayment.trim() !== '');
    if (isProtected && !forceOverwrite) {
      return;
    }

    setScrutinizingSection(section);
    setAiError(null);

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext();

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-section',
        { section, context },
        branchProposalSource
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `AI scrutiny failed (${response.status}).`);
      }

      const data = (await response.json()) as {
        result?: Record<string, unknown>;
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      };
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      } else if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      applyScrutinizedSectionResult(section, data.result ?? {});

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Relevant details for this section could not be found in the uploaded Branch Proposal.';
      setBatchScrutiny((prev) => {
        if (prev.total === 0 && prev.failedSections.length === 0 && !isInsufficient) {
          return prev;
        }
        const nextFailed = prev.failedSections.filter((f) => f.section !== section);
        const nextCompleted = prev.completedSections.includes(section)
          ? prev.completedSections
          : [...prev.completedSections, section];
        const nextInsufficient = isInsufficient
          ? [
              ...prev.insufficientSections.filter((item) => item.section !== section),
              { section, reason },
            ]
          : prev.insufficientSections.filter((item) => item.section !== section);
        return {
          ...prev,
          failedSections: nextFailed,
          completedSections: nextCompleted,
          insufficientSections: nextInsufficient,
        };
      });
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to scrutinize section.');
    } finally {
      setScrutinizingSection(null);
    }
  };

  const handleRetrySection = async (section: ScrutinySectionKey) => {
    if (batchScrutiny.running) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }
    if (batchScrutiny.retryingSections.includes(section)) return;

    setAiError(null);
    setBatchScrutiny((prev) => ({
      ...prev,
      retryingSections: prev.retryingSections.includes(section)
        ? prev.retryingSections
        : [...prev.retryingSections, section],
    }));

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext();

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-section',
        { section, context },
        branchProposalSource
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Retry failed (${response.status}).`);
      }

      const data = (await response.json()) as {
        result?: Record<string, unknown>;
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      };
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      } else if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      applyScrutinizedSectionResult(section, data.result ?? {});

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Relevant information is unavailable or insufficient in the uploaded Branch Proposal.';

      setBatchScrutiny((prev) => ({
        ...prev,
        failedSections: prev.failedSections.filter((f) => f.section !== section),
        completedSections: prev.completedSections.includes(section)
          ? prev.completedSections
          : [...prev.completedSections, section],
        insufficientSections: isInsufficient
          ? [
              ...prev.insufficientSections.filter((item) => item.section !== section),
              { section, reason },
            ]
          : prev.insufficientSections.filter((item) => item.section !== section),
      }));
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Retry failed for this section.';
      setBatchScrutiny((prev) => {
        const others = prev.failedSections.filter((f) => f.section !== section);
        return {
          ...prev,
          completedSections: prev.completedSections.filter((s) => s !== section),
          failedSections: [...others, { section, error: errMsg }],
        };
      });
    } finally {
      setBatchScrutiny((prev) => ({
        ...prev,
        retryingSections: prev.retryingSections.filter((s) => s !== section),
      }));
    }
  };

  const handleScrutinizeAll = async () => {
    if (batchScrutiny.running || scrutinizingSection !== null) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }

    batchAbortRef.current?.abort();
    const controller = new AbortController();
    batchAbortRef.current = controller;

    setAiError(null);

    const skippedManualSections: ScrutinySectionKey[] = [];
    const targetSections: ScrutinySectionKey[] = [];
    for (const s of ALL_SCRUTINY_SECTIONS) {
      const isProtected =
        Boolean(manualEditedSections[s]) ||
        (s === 'sec8SourceOfRepayment' && manualInputs.sourceOfRepayment.trim() !== '');
      if (isProtected) {
        skippedManualSections.push(s);
      } else {
        targetSections.push(s);
      }
    }

    if (targetSections.length === 0) {
      setBatchScrutiny({
        running: false,
        total: 0,
        targetSections: [],
        completedSections: [],
        retryingSections: [],
        failedSections: [],
        insufficientSections: [],
        skippedManualSections,
      });
      return;
    }

    setBatchScrutiny({
      running: true,
      total: targetSections.length,
      targetSections,
      completedSections: [],
      retryingSections: [],
      failedSections: [],
      insufficientSections: [],
      skippedManualSections,
    });

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext();

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-all',
        { sections: targetSections, context },
        branchProposalSource,
        controller.signal
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Batch AI scrutiny failed (${response.status}).`);
      }

      if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        throw new Error('Streaming response body unavailable from server.');
      }

      const decoder = new TextDecoder();
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const frames = buffer.split(/\r?\n\r?\n/);
        buffer = frames.pop() ?? '';

        for (const frame of frames) {
          const dataLines = frame
            .split(/\r?\n/)
            .filter((line) => line.startsWith('data:'))
            .map((line) => line.slice(5).trimStart());
          if (dataLines.length === 0) continue;

          const rawPayload = dataLines.join('\n');
          try {
            const event = JSON.parse(rawPayload) as {
              type: 'section' | 'error' | 'fatal' | 'done';
              section?: ScrutinySectionKey;
              result?: Record<string, unknown>;
              insufficient?: boolean;
              insufficientReason?: string | null;
              error?: string;
              sourceId?: string;
            };

            if (event.type === 'section' && event.section && event.result) {
              applyScrutinizedSectionResult(event.section, event.result);
              const sec = event.section;
              const isInsufficient = Boolean(event.insufficient);
              const reason =
                event.insufficientReason?.trim() ||
                'Relevant details for this section were not available or sufficient in the uploaded Branch Proposal.';
              setBatchScrutiny((prev) => ({
                ...prev,
                completedSections: prev.completedSections.includes(sec)
                  ? prev.completedSections
                  : [...prev.completedSections, sec],
                failedSections: prev.failedSections.filter((f) => f.section !== sec),
                insufficientSections: isInsufficient
                  ? [
                      ...prev.insufficientSections.filter((item) => item.section !== sec),
                      { section: sec, reason },
                    ]
                  : prev.insufficientSections.filter((item) => item.section !== sec),
              }));
            } else if (event.type === 'error' && event.section) {
              const sec = event.section;
              const errMsg = event.error || 'Failed to generate section.';
              setBatchScrutiny((prev) => ({
                ...prev,
                failedSections: [
                  ...prev.failedSections.filter((f) => f.section !== sec),
                  { section: sec, error: errMsg },
                ],
              }));
            } else if (event.type === 'fatal') {
              setAiError(event.error || 'Batch AI scrutiny failed.');
            } else if (event.type === 'done' && event.sourceId) {
              serverCachedSourceIdRef.current = event.sourceId;
            }
          } catch {
            // Ignore malformed SSE chunk
          }
        }
      }
    } catch (err) {
      if ((err as { name?: string })?.name !== 'AbortError') {
        setAiError(err instanceof Error ? err.message : 'Batch AI scrutiny failed.');
      }
    } finally {
      setBatchScrutiny((prev) => {
        const completedSet = new Set(prev.completedSections);
        const failedSet = new Set(prev.failedSections.map((f) => f.section));
        const uncompleted = prev.targetSections.filter(
          (sec) => !completedSet.has(sec) && !failedSet.has(sec)
        );
        const extraFailures = uncompleted.map((sec) => ({
          section: sec,
          error: 'Section did not finish generating — click Retry to scrutinize this section.',
        }));
        return {
          ...prev,
          running: false,
          failedSections: [...prev.failedSections, ...extraFailures],
        };
      });
      if (batchAbortRef.current === controller) {
        batchAbortRef.current = null;
      }
    }
  };

  const handleUpdateAiSections = (
    updater: (prev: AiSectionsData) => AiSectionsData,
    markManualSection?: ScrutinySectionKey
  ) => {
    setAiSections(updater);
    if (markManualSection) {
      setManualEditedSections((prev) => ({ ...prev, [markManualSection]: true }));
      clearOverridesForSection(markManualSection);
    }
    setHasPendingChanges(true);
  };

  const handleUploadTemplate = async (docType: MemoDocType, file: File) => {
    const record = await inspectDocxTemplate(file, docType);
    setTemplates((prev) => ({ ...prev, [docType]: record }));
    await saveTemplateRecord(record);
    setHasPendingChanges(true);
  };

  const handleRemoveTemplate = async (docType: MemoDocType) => {
    setTemplates((prev) => ({ ...prev, [docType]: null }));
    await deleteTemplateRecord(docType);
    setHasPendingChanges(true);
  };

  const handleApplyAllChanges = () => {
    const syncedMemo = resolveICMemo(extractedDocs, manualInputs, aiSections, overrides);
    const nowStr = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setApplySignal((s) => s + 1);
    setAppliedSnapshot((prev) => ({
      appliedAt: nowStr,
      version: prev.version + 1,
      memo: syncedMemo,
    }));
    setHasPendingChanges(false);
    setApplyNotice(
      `All changes applied at ${nowStr} — IC Memo, MEC Memo, Board Memo, and Sanction outputs synchronized.`
    );
  };

  const triggerBlobDownload = (blob: Blob, fileName: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const handleDownloadDocx = async (docType: MemoDocType = 'ic') => {
    const activeTemplate = templates[docType];
    if (docType !== 'ic' && !activeTemplate) {
      setAiError(
        `No default .docx template is uploaded for ${docType.toUpperCase()} Memo. Upload its own .docx template before exporting.`
      );
      return;
    }

    setDownloading(true);
    setAiError(null);
    try {
      const syncedMemo = resolveICMemo(extractedDocs, manualInputs, aiSections, overrides);
      const blob = await generateSynchronizedDocxBlob(docType, syncedMemo, activeTemplate);
      const acct = syncedMemo.placeholders['#2#'] || 'Draft';
      const safeName = acct.replace(/[^A-Za-z0-9_-]/g, '_');
      triggerBlobDownload(blob, `${DOC_FILE_PREFIX[docType]}_${safeName}.docx`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .docx document.');
    } finally {
      setDownloading(false);
    }
  };

  const handleDownloadPdf = async (docType: MemoDocType = 'ic') => {
    const activeTemplate = templates[docType];
    if (docType !== 'ic' && !activeTemplate) {
      setAiError(
        `No default .docx template is uploaded for ${docType.toUpperCase()} Memo. Upload its own .docx template before exporting.`
      );
      return;
    }

    setDownloadingPdf(true);
    setAiError(null);
    try {
      const syncedMemo = resolveICMemo(extractedDocs, manualInputs, aiSections, overrides);
      const blob = await generateMemoPdfBlob(docType, syncedMemo, activeTemplate);
      const acct = syncedMemo.placeholders['#2#'] || 'Draft';
      const safeName = acct.replace(/[^A-Za-z0-9_-]/g, '_');
      triggerBlobDownload(blob, `${DOC_FILE_PREFIX[docType]}_${safeName}.pdf`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .pdf document.');
    } finally {
      setDownloadingPdf(false);
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

          {/* Zone 3: Primary Actions */}
          <div className="hidden sm:flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={handleApplyAllChanges}
              className={`inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                hasPendingChanges
                  ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                  : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
              }`}
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Apply All Changes</span>
            </button>

            {activeTab === 'ic' ||
            (activeTab !== 'manual' && templates[activeTab as MemoDocType] !== null) ? (
              <>
                <button
                  type="button"
                  onClick={() => void handleDownloadDocx(activeTab as MemoDocType)}
                  disabled={downloading || downloadingPdf}
                  className="px-3.5 py-2 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-60"
                >
                  {downloading ? 'Building...' : 'Download .docx'}
                </button>
                <button
                  type="button"
                  onClick={() => void handleDownloadPdf(activeTab as MemoDocType)}
                  disabled={downloading || downloadingPdf}
                  className="px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-60"
                >
                  {downloadingPdf ? 'Rendering PDF...' : 'Download .pdf'}
                </button>
              </>
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

        {applyNotice && (
          <div className="bg-emerald-50 border-t border-emerald-200 px-4 sm:px-6 py-2">
            <div className="max-w-7xl mx-auto flex items-center gap-2 text-xs font-medium text-emerald-900">
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
              <span>{applyNotice}</span>
            </div>
          </div>
        )}
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
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
          />
        )}

        {activeTab === 'ic' && (
          <ICMemoPreview
            memo={resolvedICMemo}
            aiSections={aiSections}
            overrides={overrides}
            onOverrideChange={handleOverrideChange}
            onScrutinizeSection={handleScrutinizeSection}
            onRetrySection={handleRetrySection}
            onScrutinizeAll={handleScrutinizeAll}
            onDismissBatchSummary={() => setBatchScrutiny(INITIAL_BATCH_SCRUTINY)}
            batchScrutiny={batchScrutiny}
            onUpdateAiSections={handleUpdateAiSections}
            manualEditedSections={manualEditedSections}
            sourceOfRepaymentManual={manualInputs.sourceOfRepayment}
            onManualSourceOfRepaymentChange={handleManualSourceOfRepaymentChange}
            scrutinizingSection={scrutinizingSection}
            hasBranchProposal={branchProposalFile !== null}
            aiError={aiError}
            onDownloadDocx={() => handleDownloadDocx('ic')}
            onDownloadPdf={() => handleDownloadPdf('ic')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            applySignal={applySignal}
          />
        )}

        {activeTab === 'mec' && (
          <ComingSoonMemoTab
            engine={mecEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
            memo={resolvedICMemo}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={handleDownloadDocx}
            onDownloadPdf={handleDownloadPdf}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
          />
        )}

        {activeTab === 'board' && (
          <ComingSoonMemoTab
            engine={boardEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
            memo={resolvedICMemo}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={handleDownloadDocx}
            onDownloadPdf={handleDownloadPdf}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
          />
        )}

        {activeTab === 'sanction' && (
          <ComingSoonMemoTab
            engine={sanctionEngine}
            accountNo={resolvedICMemo.placeholders['#2#']}
            branchName={resolvedICMemo.placeholders['#3#']}
            proposalDate={resolvedICMemo.placeholders['#1#']}
            memo={resolvedICMemo}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={handleDownloadDocx}
            onDownloadPdf={handleDownloadPdf}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
          />
        )}
      </main>
    </div>
  );
}
