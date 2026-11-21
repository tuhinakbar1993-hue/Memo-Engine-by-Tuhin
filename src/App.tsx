import { useEffect, useMemo, useRef, useState } from 'react';
import { AlertTriangle, CheckCircle2, RefreshCw, Trash2 } from 'lucide-react';
import { ICMemoPreview } from './components/ICMemoPreview';
import { LandingScreen } from './components/LandingScreen';
import { ManualInputsTab } from './components/ManualInputsTab';
import { StageMemoTab } from './components/StageMemoTab';
import {
  INITIAL_STAGE_AI_SECTIONS,
  INITIAL_STAGE_INPUTS,
  resolveICMemo,
  resolveStageMemo,
  STAGE_SECTIONS_BY_DOC,
} from './engines/icMemoEngine';
import { inspectDocxTemplate, verifyPopulatedDocxBlob } from './exporters/docxTemplateEngine';
import { generatePdfFromDocxBlob, generateSynchronizedDocxBlob } from './exporters/pdfExporter';
import { createEmptyBranchProposal } from './parsers/branchProposalParser';
import {
  clearDocumentParseCache,
  detectDocumentKeyFromFile,
  parseAccountStatementPdf,
  parseBranchProposalFile,
  parseCibReportPdf,
  parseCostOfFundFile,
  parseOutstandingPdf,
  parseOverduePdf,
  parsePartyPerformancePdf,
  prepareBranchProposalSource,
  prepareGenericDocumentSource,
  warmBranchProposalSource,
} from './parsers/documentParsers';
import type {
  AiSectionsData,
  AppliedSyncState,
  BatchScrutinyState,
  BranchProposalData,
  BranchProposalSource,
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  MemoDocType,
  MemoTabKey,
  MemoTemplateRecord,
  ScrutinySectionKey,
  Section10BankDirectorRow,
  Section10DirectorRow,
  Section10Particulars,
  Section11Row,
  Section3Row,
  Section4Row,
  Section5Row,
  Section6Row,
  Section7PrimaryRow,
  Section7ValuationRow,
  Section9Row,
  StageAiSectionsData,
  StageBatchScrutinyState,
  StageMemoDocType,
  StageMemoInputs,
  StageScrutinySectionKey,
  StageUploadedSourceMeta,
} from './types/memo';
import {
  clearCentralCaseState,
  loadCentralCaseState,
  saveCentralCaseState,
} from './utils/caseStorage';
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
  clientPrayer: null,
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
  'clientPrayer',
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

const INITIAL_STAGE_BATCH_SCRUTINY: StageBatchScrutinyState = {
  running: false,
  docType: null,
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
  const [stageInputs, setStageInputs] = useState<StageMemoInputs>(INITIAL_STAGE_INPUTS);
  const [stageAiSections, setStageAiSections] = useState<StageAiSectionsData>(
    INITIAL_STAGE_AI_SECTIONS
  );
  const [stageManualEditedSections, setStageManualEditedSections] = useState<
    Partial<Record<StageScrutinySectionKey, boolean>>
  >({});
  const [scrutinizingStageSection, setScrutinizingStageSection] =
    useState<StageScrutinySectionKey | null>(null);
  const [stageBatchScrutiny, setStageBatchScrutiny] = useState<StageBatchScrutinyState>(
    INITIAL_STAGE_BATCH_SCRUTINY
  );
  const [stageUploadedSources, setStageUploadedSources] = useState<
    Partial<Record<StageMemoDocType, StageUploadedSourceMeta>>
  >({});
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [downloadingPdf, setDownloadingPdf] = useState(false);

  const serverCachedSourceIdRef = useRef<string | null>(null);
  const stageSourceFilesRef = useRef<Partial<Record<StageMemoDocType, File | null>>>({});
  const batchAbortRef = useRef<AbortController | null>(null);
  const stageBatchAbortRef = useRef<AbortController | null>(null);
  const caseDirtyRef = useRef(false);
  const persistTimerRef = useRef<number | null>(null);

  const caseRef = useRef<{
    extractedDocs: ExtractedDocumentsState;
    manualInputs: ManualInputsData;
    aiSections: AiSectionsData;
    overrides: Record<string, string>;
    manualEditedSections: Partial<Record<ScrutinySectionKey, boolean>>;
    stageInputs: StageMemoInputs;
    stageAiSections: StageAiSectionsData;
    stageManualEditedSections: Partial<Record<StageScrutinySectionKey, boolean>>;
    stageUploadedSources: Partial<Record<StageMemoDocType, StageUploadedSourceMeta>>;
  }>({
    extractedDocs: INITIAL_EXTRACTED_DOCS,
    manualInputs: INITIAL_MANUAL_INPUTS,
    aiSections: INITIAL_AI_SECTIONS,
    overrides: {},
    manualEditedSections: {},
    stageInputs: INITIAL_STAGE_INPUTS,
    stageAiSections: INITIAL_STAGE_AI_SECTIONS,
    stageManualEditedSections: {},
    stageUploadedSources: {},
  });
  const docMetasRef = useRef<Record<DocumentKey, DocumentUploadMeta>>(INITIAL_DOC_METAS);

  const resolvedICMemo = useMemo(
    () =>
      resolveICMemo(
        extractedDocs,
        manualInputs,
        aiSections,
        overrides,
        manualEditedSections
      ),
    [extractedDocs, manualInputs, aiSections, overrides, manualEditedSections]
  );

  const resolvedMecMemo = useMemo(
    () => resolveStageMemo('mec', resolvedICMemo, stageInputs, stageAiSections),
    [resolvedICMemo, stageInputs, stageAiSections]
  );

  const resolvedBoardMemo = useMemo(
    () => resolveStageMemo('board', resolvedICMemo, stageInputs, stageAiSections),
    [resolvedICMemo, stageInputs, stageAiSections]
  );

  const resolvedSanctionMemo = useMemo(
    () => resolveStageMemo('sanction', resolvedICMemo, stageInputs, stageAiSections),
    [resolvedICMemo, stageInputs, stageAiSections]
  );

  const [appliedSnapshot, setAppliedSnapshot] = useState<AppliedSyncState>(() => ({
    appliedAt: null,
    version: 0,
    memo: resolveICMemo(INITIAL_EXTRACTED_DOCS, INITIAL_MANUAL_INPUTS, INITIAL_AI_SECTIONS, {}, {}),
    stageInputs: INITIAL_STAGE_INPUTS,
    stageAiSections: INITIAL_STAGE_AI_SECTIONS,
  }));
  const [hasPendingChanges, setHasPendingChanges] = useState(false);
  const [applySignal, setApplySignal] = useState(0);
  const [applyNotice, setApplyNotice] = useState<string | null>(null);

  useEffect(() => {
    caseRef.current = {
      extractedDocs,
      manualInputs,
      aiSections,
      overrides,
      manualEditedSections,
      stageInputs,
      stageAiSections,
      stageManualEditedSections,
      stageUploadedSources,
    };
  }, [
    extractedDocs,
    manualInputs,
    aiSections,
    overrides,
    manualEditedSections,
    stageInputs,
    stageAiSections,
    stageManualEditedSections,
    stageUploadedSources,
  ]);

  useEffect(() => {
    docMetasRef.current = docMetas;
  }, [docMetas]);

  const persistCentralCaseNow = async (immediate = false) => {
    if (persistTimerRef.current !== null) {
      window.clearTimeout(persistTimerRef.current);
      persistTimerRef.current = null;
    }
    if (immediate) {
      await saveCentralCaseState(caseRef.current, docMetasRef.current);
      return;
    }
    persistTimerRef.current = window.setTimeout(() => {
      persistTimerRef.current = null;
      void saveCentralCaseState(caseRef.current, docMetasRef.current);
    }, 150);
  };

  const syncAppliedSnapshotFromRef = () => {
    const latest = caseRef.current;
    const syncedMemo = resolveICMemo(
      latest.extractedDocs,
      latest.manualInputs,
      latest.aiSections,
      latest.overrides,
      latest.manualEditedSections
    );
    const nowStr = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setAppliedSnapshot((prev) => ({
      appliedAt: nowStr,
      version: prev.version + 1,
      memo: syncedMemo,
      stageInputs: latest.stageInputs,
      stageAiSections: latest.stageAiSections,
    }));
    setHasPendingChanges(false);
    return syncedMemo;
  };

  const setDocMetasSync = (
    updater:
      | Record<DocumentKey, DocumentUploadMeta>
      | ((prev: Record<DocumentKey, DocumentUploadMeta>) => Record<DocumentKey, DocumentUploadMeta>)
  ) => {
    const next = typeof updater === 'function' ? updater(docMetasRef.current) : updater;
    docMetasRef.current = next;
    setDocMetas(next);
  };

  const setExtractedDocsSync = (
    updater:
      | ExtractedDocumentsState
      | ((prev: ExtractedDocumentsState) => ExtractedDocumentsState)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.extractedDocs) : updater;
    caseRef.current = { ...caseRef.current, extractedDocs: next };
    setExtractedDocs(next);
  };

  const setManualInputsSync = (
    updater: ManualInputsData | ((prev: ManualInputsData) => ManualInputsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.manualInputs) : updater;
    caseRef.current = { ...caseRef.current, manualInputs: next };
    setManualInputs(next);
  };

  const setAiSectionsSync = (
    updater: AiSectionsData | ((prev: AiSectionsData) => AiSectionsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.aiSections) : updater;
    caseRef.current = { ...caseRef.current, aiSections: next };
    setAiSections(next);
  };

  const setOverridesSync = (
    updater: Record<string, string> | ((prev: Record<string, string>) => Record<string, string>)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.overrides) : updater;
    caseRef.current = { ...caseRef.current, overrides: next };
    setOverrides(next);
  };

  const setManualEditedSectionsSync = (
    updater:
      | Partial<Record<ScrutinySectionKey, boolean>>
      | ((
          prev: Partial<Record<ScrutinySectionKey, boolean>>
        ) => Partial<Record<ScrutinySectionKey, boolean>>)
  ) => {
    caseDirtyRef.current = true;
    const next =
      typeof updater === 'function' ? updater(caseRef.current.manualEditedSections) : updater;
    caseRef.current = { ...caseRef.current, manualEditedSections: next };
    setManualEditedSections(next);
  };

  const setStageInputsSync = (
    updater: StageMemoInputs | ((prev: StageMemoInputs) => StageMemoInputs)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.stageInputs) : updater;
    caseRef.current = { ...caseRef.current, stageInputs: next };
    setStageInputs(next);
  };

  const setStageAiSectionsSync = (
    updater: StageAiSectionsData | ((prev: StageAiSectionsData) => StageAiSectionsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.stageAiSections) : updater;
    caseRef.current = { ...caseRef.current, stageAiSections: next };
    setStageAiSections(next);
  };

  const setStageManualEditedSectionsSync = (
    updater:
      | Partial<Record<StageScrutinySectionKey, boolean>>
      | ((
          prev: Partial<Record<StageScrutinySectionKey, boolean>>
        ) => Partial<Record<StageScrutinySectionKey, boolean>>)
  ) => {
    caseDirtyRef.current = true;
    const next =
      typeof updater === 'function' ? updater(caseRef.current.stageManualEditedSections) : updater;
    caseRef.current = { ...caseRef.current, stageManualEditedSections: next };
    setStageManualEditedSections(next);
  };

  const setStageUploadedSourcesSync = (
    updater:
      | Partial<Record<StageMemoDocType, StageUploadedSourceMeta>>
      | ((
          prev: Partial<Record<StageMemoDocType, StageUploadedSourceMeta>>
        ) => Partial<Record<StageMemoDocType, StageUploadedSourceMeta>>)
  ) => {
    caseDirtyRef.current = true;
    const next =
      typeof updater === 'function' ? updater(caseRef.current.stageUploadedSources) : updater;
    caseRef.current = { ...caseRef.current, stageUploadedSources: next };
    setStageUploadedSources(next);
  };

  useEffect(() => {
    void loadAllTemplateRecords().then((loaded) => {
      setTemplates(loaded);
    });
    void loadCentralCaseState().then((saved) => {
      if (!saved || caseDirtyRef.current) return;
      const cleanStageSources: Partial<Record<StageMemoDocType, StageUploadedSourceMeta>> = {};
      if (saved.stageUploadedSources) {
        for (const dt of ['mec', 'board', 'sanction'] as StageMemoDocType[]) {
          const s = saved.stageUploadedSources[dt];
          if (s) {
            cleanStageSources[dt] = { ...s, scrutinizing: false, stage: null };
          }
        }
      }
      const hydrated = {
        extractedDocs: { ...INITIAL_EXTRACTED_DOCS, ...(saved.extractedDocs || {}) },
        manualInputs: { ...INITIAL_MANUAL_INPUTS, ...(saved.manualInputs || {}) },
        aiSections: { ...INITIAL_AI_SECTIONS, ...(saved.aiSections || {}) },
        overrides: saved.overrides || {},
        manualEditedSections: saved.manualEditedSections || {},
        stageInputs: { ...INITIAL_STAGE_INPUTS, ...(saved.stageInputs || {}) },
        stageAiSections: { ...INITIAL_STAGE_AI_SECTIONS, ...(saved.stageAiSections || {}) },
        stageManualEditedSections: saved.stageManualEditedSections || {},
        stageUploadedSources: cleanStageSources,
      };
      caseRef.current = hydrated;
      setExtractedDocs(hydrated.extractedDocs);
      setManualInputs(hydrated.manualInputs);
      setAiSections(hydrated.aiSections);
      setOverrides(hydrated.overrides);
      setManualEditedSections(hydrated.manualEditedSections);
      setStageInputs(hydrated.stageInputs);
      setStageAiSections(hydrated.stageAiSections);
      setStageManualEditedSections(hydrated.stageManualEditedSections);
      setStageUploadedSources(hydrated.stageUploadedSources);

      if (saved.docUploadedFlags) {
        setDocMetasSync((prev) => {
          const next = { ...prev };
          for (const key of Object.keys(INITIAL_DOC_METAS) as DocumentKey[]) {
            const flag = saved.docUploadedFlags?.[key];
            if (flag) {
              next[key] = {
                ...next[key],
                uploaded: Boolean(flag.uploaded),
                fileName: flag.fileName || '',
              };
            }
          }
          return next;
        });
      }

      const hydratedMemo = resolveICMemo(
        hydrated.extractedDocs,
        hydrated.manualInputs,
        hydrated.aiSections,
        hydrated.overrides,
        hydrated.manualEditedSections
      );
      setAppliedSnapshot({
        appliedAt: saved.updatedAt
          ? new Date(saved.updatedAt).toLocaleTimeString([], {
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit',
            })
          : null,
        version: 1,
        memo: hydratedMemo,
        stageInputs: hydrated.stageInputs,
        stageAiSections: hydrated.stageAiSections,
      });
      setHasPendingChanges(false);
    });

    return () => {
      if (persistTimerRef.current !== null) {
        window.clearTimeout(persistTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!applyNotice) return;
    const timer = window.setTimeout(() => setApplyNotice(null), 5000);
    return () => window.clearTimeout(timer);
  }, [applyNotice]);

  const updateDocStage = (key: DocumentKey, stage: string) => {
    setDocMetasSync((prev) => ({
      ...prev,
      [key]: {
        ...prev[key],
        stage,
      },
    }));
  };

  const handleFileUpload = async (key: DocumentKey, file: File) => {
    setDocMetasSync((prev) => ({
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
        setExtractedDocsSync((prev) => ({ ...prev, branchProposal: bpData }));
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
        setExtractedDocsSync((prev) => ({ ...prev, outstanding: res.data }));
      } else if (key === 'partyPerformance') {
        const res = await parsePartyPerformancePdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, partyPerformance: res.data }));
      } else if (key === 'overdue') {
        const res = await parseOverduePdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, overdue: res.data }));
      } else if (key === 'accountStatement') {
        const res = await parseAccountStatementPdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, accountStatement: res.data }));
      } else if (key === 'cibReport') {
        const res = await parseCibReportPdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, cibReport: res.data }));
      } else if (key === 'costOfFund') {
        const res = await parseCostOfFundFile(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, costOfFund: res.data }));
      }

      setDocMetasSync((prev) => ({
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
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse document.';
      setDocMetasSync((prev) => ({
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
    setManualInputsSync((prev) => ({ ...prev, [field]: value }));
    if (field === 'sourceOfRepayment') {
      setManualEditedSectionsSync((prev) => ({
        ...prev,
        sec8SourceOfRepayment: value.trim() !== '',
      }));
    }
    const mappedCode = MANUAL_FIELD_TO_CODE[field];
    if (mappedCode) {
      setOverridesSync((prev) => {
        if (prev[mappedCode] === undefined) return prev;
        const next = { ...prev };
        delete next[mappedCode];
        return next;
      });
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleManualSourceOfRepaymentChange = (value: string) => {
    setManualInputsSync((prev) => ({ ...prev, sourceOfRepayment: value }));
    setAiSectionsSync((prev) => ({ ...prev, sec8SourceOfRepayment: value }));
    setManualEditedSectionsSync((prev) => ({ ...prev, sec8SourceOfRepayment: true }));
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleOverrideChange = (code: string, value: string) => {
    setOverridesSync((prev) => ({ ...prev, [code]: value }));
    if (value.trim() !== '') {
      if (code === '#51#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, justification: true }));
      } else if (code === '#54#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec1NatureOfBusiness: true }));
      } else if (code === '#56#' || code === '#58#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec2LiabilityOurBank: true }));
      } else if (code === '#62#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec4CibStatus: true }));
      }
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleResetOverrides = () => {
    setOverridesSync({});
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(true);
  };

  const clearOverridesForSection = (section: ScrutinySectionKey) => {
    const codesToClear =
      section === 'justification'
        ? ['#51#']
        : section === 'sec1NatureOfBusiness'
          ? ['#54#']
          : section === 'sec2LiabilityOurBank'
            ? ['#56#', '#58#']
            : section === 'sec4CibStatus'
              ? ['#62#']
              : [];
    if (codesToClear.length === 0) return;
    setOverridesSync((prev) => {
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

  const buildScrutinyContext = () => {
    const latest = caseRef.current;
    const latestMemo = resolveICMemo(
      latest.extractedDocs,
      latest.manualInputs,
      latest.aiSections,
      latest.overrides,
      latest.manualEditedSections
    );
    return {
      placeholders: latestMemo.placeholders,
      sourceOfRepayment: latestMemo.sourceOfRepayment,
      clientPrayer: latestMemo.clientPrayer,
      branchVisitReport: latestMemo.branchVisitReport,
      justificationLines: latestMemo.justificationLines,
      sec13BranchRecommendation: latestMemo.sec13BranchRecommendation,
      sec14IccdRecommendation: latestMemo.sec14IccdRecommendation,
      samdObservation: latestMemo.samdObservation,
      sec2NonFundedNote: latestMemo.sec2NonFundedNote,
      sec7OthersNote: latestMemo.sec7OthersNote,
      sectionSummaries: latestMemo.sectionSummaries,
      icMemoFinalized: latestMemo,
      icMemoAiSections: latest.aiSections,
      mecMemoFinalized: resolveStageMemo(
        'mec',
        latestMemo,
        latest.stageInputs,
        latest.stageAiSections
      ),
      boardMemoFinalized: resolveStageMemo(
        'board',
        latestMemo,
        latest.stageInputs,
        latest.stageAiSections
      ),
      extractedSummaries: {
        outstanding: latest.extractedDocs.outstanding,
        partyPerformance: latest.extractedDocs.partyPerformance,
        overdue: latest.extractedDocs.overdue,
        accountStatement: latest.extractedDocs.accountStatement,
        cibReport: latest.extractedDocs.cibReport,
        costOfFund: latest.extractedDocs.costOfFund,
      },
      branchProposal: latest.extractedDocs.branchProposal,
      stageInputs: latest.stageInputs,
      stageSections: latest.stageAiSections,
    };
  };

  const applyScrutinizedSectionResult = (
    section: ScrutinySectionKey,
    result: Record<string, unknown>
  ) => {
    const isNarrative =
      section === 'justification' ||
      section === 'sec1NatureOfBusiness' ||
      section === 'sec8SourceOfRepayment' ||
      section === 'clientPrayer' ||
      section === 'branchVisit' ||
      section === 'sec13BranchRecommendation' ||
      section === 'sec14IccdRecommendation' ||
      section === 'samdObservation';

    if (isNarrative) {
      const text = String(result.text ?? '').trim();
      setAiSectionsSync((prev) => ({
        ...prev,
        [section]: text,
      }));
      if (section === 'sec8SourceOfRepayment') {
        setManualInputsSync((prev) => ({ ...prev, sourceOfRepayment: '' }));
      }
    } else {
      setAiSectionsSync((prev) => ({
        ...prev,
        [section]: result,
      }));
    }

    clearOverridesForSection(section);
    setManualEditedSectionsSync((prev) => ({ ...prev, [section]: false }));
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(true);
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
    setAiSectionsSync(updater);
    if (markManualSection) {
      setManualEditedSectionsSync((prev) => ({ ...prev, [markManualSection]: true }));
      clearOverridesForSection(markManualSection);
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleStageInputChange = (field: keyof StageMemoInputs, value: string) => {
    setStageInputsSync((prev) => ({ ...prev, [field]: value }));
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleUpdateStageAiSection = (
    section: StageScrutinySectionKey,
    value: string,
    markManual = true
  ) => {
    setStageAiSectionsSync((prev) => ({ ...prev, [section]: value }));
    if (markManual) {
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: true }));
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const fetchStageEndpoint = async (
    docType: StageMemoDocType,
    url: string,
    payload: Record<string, unknown>,
    signal?: AbortSignal
  ): Promise<Response> => {
    const stageSourceFile = stageSourceFilesRef.current[docType];
    if (stageSourceFile) {
      const source = await prepareGenericDocumentSource(stageSourceFile, `stage_${docType}`);
      return fetchWithSourceCache(url, payload, source, signal);
    }
    if (branchProposalFile) {
      const source = await prepareBranchProposalSource(branchProposalFile);
      return fetchWithSourceCache(url, payload, source, signal);
    }
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal,
      body: JSON.stringify(payload),
    });
  };

  const handleClearStageSource = (docType: StageMemoDocType) => {
    stageSourceFilesRef.current[docType] = null;
    setStageUploadedSourcesSync((prev) => {
      const next = { ...prev };
      delete next[docType];
      return next;
    });
    void persistCentralCaseNow(true);
  };

  const handleUploadStageSource = async (docType: StageMemoDocType, file: File) => {
    stageSourceFilesRef.current[docType] = file;
    setAiError(null);
    setStageUploadedSourcesSync((prev) => ({
      ...prev,
      [docType]: {
        docType,
        uploaded: false,
        fileName: file.name,
        fileSize: file.size,
        uploadedAt: null,
        scrutinizing: true,
        stage: 'Preparing document...',
        error: null,
        missingOrInconsistentNotes: [],
      },
    }));

    try {
      const source = await prepareGenericDocumentSource(file, `stage_${docType}`, (stageText) => {
        setStageUploadedSourcesSync((prev) => ({
          ...prev,
          [docType]: {
            ...(prev[docType] || {
              docType,
              uploaded: false,
              fileName: file.name,
              fileSize: file.size,
              uploadedAt: null,
              scrutinizing: true,
              error: null,
              missingOrInconsistentNotes: [],
            }),
            stage: stageText,
          },
        }));
      });

      setStageUploadedSourcesSync((prev) => ({
        ...prev,
        [docType]: {
          ...prev[docType]!,
          stage: 'Scrutinizing and extracting all details through Gemini AI...',
        },
      }));

      const currentMemo = resolveICMemo(
        caseRef.current.extractedDocs,
        caseRef.current.manualInputs,
        caseRef.current.aiSections,
        caseRef.current.overrides,
        caseRef.current.manualEditedSections
      );

      const response = await fetch('/api/gemini/extract-stage-source', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          docType,
          source,
          existingStageInputs: caseRef.current.stageInputs,
          existingPlaceholders: currentMemo.placeholders,
        }),
      });

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(
          errBody.error || `Failed to scrutinize uploaded memo (${response.status}).`
        );
      }

      const data = (await response.json()) as {
        extracted?: Record<string, unknown>;
        sourceId?: string;
        fileName?: string;
      };
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      }

      const ex = (data.extracted || {}) as Record<string, unknown>;
      const exStageInputs = (ex.extractedStageInputs || {}) as Record<string, string>;
      const exPlaceholders = (ex.placeholders || {}) as Record<string, string>;
      const exNarratives = (ex.narratives || {}) as Record<string, string>;
      const exStageDrafts = (ex.stageDrafts || {}) as Record<string, string>;
      const missingNotes = Array.isArray(ex.missingOrInconsistentNotes)
        ? (ex.missingOrInconsistentNotes as unknown[])
            .map((x) => String(x || '').trim())
            .filter(Boolean)
        : [];

      // 1. Update Stage Inputs where currently blank
      setStageInputsSync((prev) => ({
        icMemoNumber: prev.icMemoNumber.trim() || String(exStageInputs.icMemoNumber ?? '').trim(),
        icMemoDate: prev.icMemoDate.trim() || String(exStageInputs.icMemoDate ?? '').trim(),
        mecMemoNumber:
          prev.mecMemoNumber.trim() || String(exStageInputs.mecMemoNumber ?? '').trim(),
        mecMemoDate: prev.mecMemoDate.trim() || String(exStageInputs.mecMemoDate ?? '').trim(),
        boardMemoNumber:
          prev.boardMemoNumber.trim() || String(exStageInputs.boardMemoNumber ?? '').trim(),
        boardMemoDate:
          prev.boardMemoDate.trim() || String(exStageInputs.boardMemoDate ?? '').trim(),
        sanctionRefNumber: prev.sanctionRefNumber,
        sanctionDate: prev.sanctionDate,
      }));

      // 2. Populate Manual Inputs & Overrides from extracted placeholders without overwriting user-entered values
      const manualKeyByCode: Record<string, keyof ManualInputsData> = {
        '#1#': 'proposalDate',
        '#4#': 'waiverAmount',
        '#27#': 'settlementAmount',
        '#28#': 'settlementPeriod',
        '#45#': 'waiverRecClientBranch',
        '#47#': 'waiverRecHO',
      };

      setManualInputsSync((prev) => {
        const next = { ...prev };
        for (const [code, field] of Object.entries(manualKeyByCode)) {
          const extractedVal = String(exPlaceholders[code] ?? '').trim();
          if (!next[field].trim() && extractedVal) {
            next[field] = extractedVal;
          }
        }
        const extractedSor = String(ex.sourceOfRepayment ?? '').trim();
        if (!next.sourceOfRepayment.trim() && extractedSor) {
          next.sourceOfRepayment = extractedSor;
        }
        return next;
      });

      setOverridesSync((prev) => {
        const next = { ...prev };
        for (const [code, rawVal] of Object.entries(exPlaceholders)) {
          if (code === '#32#' || code === '#42#' || code === '#44#') continue;
          if (manualKeyByCode[code]) continue;
          const val = String(rawVal ?? '').trim();
          if (!val) continue;
          const existingResolved = currentMemo.placeholders[code];
          if ((next[code] === undefined || next[code] === '') && !existingResolved) {
            next[code] = val;
          }
        }
        return next;
      });

      // 3. Merge structured Branch Proposal tables & details into extractedDocs.branchProposal
      setExtractedDocsSync((prev) => {
        const baseBp: BranchProposalData = prev.branchProposal
          ? {
              ...prev.branchProposal,
              section10Particulars: { ...prev.branchProposal.section10Particulars },
            }
          : createEmptyBranchProposal();

        const pickStr = (curr: string, incoming: unknown): string =>
          curr.trim() ? curr : String(incoming ?? '').trim();

        const exParticulars = (ex.section10Particulars || {}) as Partial<Section10Particulars>;
        const mergedParticulars: Section10Particulars = { ...baseBp.section10Particulars };
        for (const k of Object.keys(mergedParticulars) as Array<keyof Section10Particulars>) {
          mergedParticulars[k] = pickStr(mergedParticulars[k], exParticulars[k]);
        }

        const pickRows = <T,>(curr: T[], incoming: unknown): T[] =>
          curr.length > 0 ? curr : Array.isArray(incoming) ? (incoming as T[]) : curr;

        const nextBp: BranchProposalData = {
          ...baseBp,
          referenceNo: pickStr(baseBp.referenceNo, exPlaceholders['#64#']),
          proposalDate: pickStr(baseBp.proposalDate, exPlaceholders['#65#']),
          branchName: pickStr(baseBp.branchName, exPlaceholders['#3#']),
          accountNo: pickStr(baseBp.accountNo, exPlaceholders['#2#']),
          limitOriginal: pickStr(baseBp.limitOriginal, exPlaceholders['#6#']),
          limitRSD: pickStr(baseBp.limitRSD, ex.limitRSD),
          securityMV: pickStr(baseBp.securityMV, exPlaceholders['#33#']),
          securityFSV: pickStr(baseBp.securityFSV, exPlaceholders['#34#']),
          justificationText: pickStr(baseBp.justificationText, exPlaceholders['#51#']),
          previousBoardDecisions: pickStr(
            baseBp.previousBoardDecisions,
            ex.previousBoardDecisions
          ),
          natureOfBusiness: pickStr(baseBp.natureOfBusiness, exPlaceholders['#54#']),
          sec2Limit: pickStr(baseBp.sec2Limit, exPlaceholders['#56#']),
          sec2Validity: pickStr(baseBp.sec2Validity, exPlaceholders['#58#']),
          sec2Overdue: pickStr(baseBp.sec2Overdue, exPlaceholders['#59#']),
          sec2Recovery: pickStr(baseBp.sec2Recovery, exPlaceholders['#60#']),
          sec2Remarks: pickStr(baseBp.sec2Remarks, exNarratives.sec2Remarks),
          section3Rows: pickRows<Section3Row>(baseBp.section3Rows, ex.section3Rows),
          section4CibDate: pickStr(
            baseBp.section4CibDate,
            ex.section4CibDate || exPlaceholders['#62#']
          ),
          section4Rows: pickRows<Section4Row>(baseBp.section4Rows, ex.section4Rows),
          section5StatusNote: pickStr(baseBp.section5StatusNote, ex.section5StatusNote),
          section5Rows: pickRows<Section5Row>(baseBp.section5Rows, ex.section5Rows),
          section6StatusNote: pickStr(baseBp.section6StatusNote, ex.section6StatusNote),
          section6Rows: pickRows<Section6Row>(baseBp.section6Rows, ex.section6Rows),
          section7PrimaryNote: pickStr(baseBp.section7PrimaryNote, ex.section7PrimaryNote),
          section7PrimaryRows: pickRows<Section7PrimaryRow>(
            baseBp.section7PrimaryRows,
            ex.section7PrimaryRows
          ),
          section7PrimaryTotal: pickStr(baseBp.section7PrimaryTotal, ex.section7PrimaryTotal),
          section7CollateralNote: pickStr(
            baseBp.section7CollateralNote,
            ex.section7CollateralNote
          ),
          section7CollateralSchedules:
            baseBp.section7CollateralSchedules.length > 0
              ? baseBp.section7CollateralSchedules
              : Array.isArray(ex.section7CollateralSchedules)
                ? (ex.section7CollateralSchedules as string[]).filter(Boolean)
                : baseBp.section7CollateralSchedules,
          section7ValuationRows: pickRows<Section7ValuationRow>(
            baseBp.section7ValuationRows,
            ex.section7ValuationRows
          ),
          section7CoveragePct: pickStr(baseBp.section7CoveragePct, ex.section7CoveragePct),
          section7Others: pickStr(baseBp.section7Others, ex.section7Others),
          section8SourceOfRepayment: pickStr(
            baseBp.section8SourceOfRepayment,
            ex.sourceOfRepayment
          ),
          section9Rows: pickRows<Section9Row>(baseBp.section9Rows, ex.section9Rows),
          section10Particulars: mergedParticulars,
          section10Directors: pickRows<Section10DirectorRow>(
            baseBp.section10Directors,
            ex.section10Directors
          ),
          section10BankDirectors: pickRows<Section10BankDirectorRow>(
            baseBp.section10BankDirectors,
            ex.section10BankDirectors
          ),
          section11FinancialsNote: pickStr(
            baseBp.section11FinancialsNote,
            ex.section11FinancialsNote
          ),
          section11Rows: pickRows<Section11Row>(baseBp.section11Rows, ex.section11Rows),
        };

        return { ...prev, branchProposal: nextBp };
      });

      // 4. Populate carried-forward IC narratives where not manually edited
      setAiSectionsSync((prev) => {
        const next = { ...prev };
        const edited = caseRef.current.manualEditedSections;
        const setIfEmpty = (key: keyof AiSectionsData, val: string) => {
          const trimmed = val.trim();
          if (!trimmed) return;
          if (!edited[key as ScrutinySectionKey] && (!next[key] || String(next[key]).trim() === '')) {
            (next[key] as string | null) = trimmed;
          }
        };
        setIfEmpty('justification', String(exPlaceholders['#51#'] ?? ''));
        setIfEmpty('sec1NatureOfBusiness', String(exPlaceholders['#54#'] ?? ''));
        setIfEmpty('sec8SourceOfRepayment', String(ex.sourceOfRepayment ?? ''));
        setIfEmpty('clientPrayer', String(exNarratives.clientPrayer ?? ''));
        setIfEmpty('branchVisit', String(exNarratives.branchVisit ?? ''));
        setIfEmpty('sec13BranchRecommendation', String(exNarratives.sec13BranchRecommendation ?? ''));
        setIfEmpty('sec14IccdRecommendation', String(exNarratives.sec14IccdRecommendation ?? ''));
        setIfEmpty('samdObservation', String(exNarratives.samdObservation ?? ''));
        return next;
      });

      // 5. Populate Stage AI sections from extracted stageDrafts (respecting manual edits)
      setStageAiSectionsSync((prev) => {
        const next = { ...prev };
        const edited = caseRef.current.stageManualEditedSections;
        for (const k of Object.keys(INITIAL_STAGE_AI_SECTIONS) as StageScrutinySectionKey[]) {
          const draft = String(exStageDrafts[k] ?? '').trim();
          if (draft && !edited[k]) {
            next[k] = draft;
          }
        }
        return next;
      });

      const uploadedAt = new Date().toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
      });
      setStageUploadedSourcesSync((prev) => ({
        ...prev,
        [docType]: {
          docType,
          uploaded: true,
          fileName: file.name,
          fileSize: file.size,
          uploadedAt,
          scrutinizing: false,
          stage: null,
          error: null,
          missingOrInconsistentNotes: missingNotes,
        },
      }));

      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);
      setApplyNotice(
        `Extracted & scrutinized ${file.name} (${uploadedAt}) — data carried forward into ${docType.toUpperCase()} output.`
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to extract uploaded memo.';
      setStageUploadedSourcesSync((prev) => ({
        ...prev,
        [docType]: {
          docType,
          uploaded: false,
          fileName: file.name,
          fileSize: file.size,
          uploadedAt: null,
          scrutinizing: false,
          stage: null,
          error: msg,
          missingOrInconsistentNotes: [],
        },
      }));
    }
  };

  const handleClearAllData = () => {
    batchAbortRef.current?.abort();
    batchAbortRef.current = null;
    stageBatchAbortRef.current?.abort();
    stageBatchAbortRef.current = null;
    if (persistTimerRef.current !== null) {
      window.clearTimeout(persistTimerRef.current);
      persistTimerRef.current = null;
    }

    caseDirtyRef.current = true;
    clearDocumentParseCache();
    void clearCentralCaseState();

    stageSourceFilesRef.current = {};
    serverCachedSourceIdRef.current = null;
    setBranchProposalFile(null);

    docMetasRef.current = INITIAL_DOC_METAS;
    setDocMetas(INITIAL_DOC_METAS);

    const emptyCaseState = {
      extractedDocs: INITIAL_EXTRACTED_DOCS,
      manualInputs: INITIAL_MANUAL_INPUTS,
      aiSections: INITIAL_AI_SECTIONS,
      overrides: {},
      manualEditedSections: {},
      stageInputs: INITIAL_STAGE_INPUTS,
      stageAiSections: INITIAL_STAGE_AI_SECTIONS,
      stageManualEditedSections: {},
      stageUploadedSources: {},
    };
    caseRef.current = emptyCaseState;

    setExtractedDocs(INITIAL_EXTRACTED_DOCS);
    setManualInputs(INITIAL_MANUAL_INPUTS);
    setAiSections(INITIAL_AI_SECTIONS);
    setOverrides({});
    setManualEditedSections({});
    setStageInputs(INITIAL_STAGE_INPUTS);
    setStageAiSections(INITIAL_STAGE_AI_SECTIONS);
    setStageManualEditedSections({});
    setStageUploadedSources({});

    setScrutinizingSection(null);
    setScrutinizingStageSection(null);
    setBatchScrutiny(INITIAL_BATCH_SCRUTINY);
    setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY);
    setBatchProgress({ active: false, total: 0, completed: 0 });
    setAiError(null);

    const emptyMemo = resolveICMemo(
      INITIAL_EXTRACTED_DOCS,
      INITIAL_MANUAL_INPUTS,
      INITIAL_AI_SECTIONS,
      {},
      {}
    );
    setAppliedSnapshot((prev) => ({
      appliedAt: null,
      version: prev.version + 1,
      memo: emptyMemo,
      stageInputs: INITIAL_STAGE_INPUTS,
      stageAiSections: INITIAL_STAGE_AI_SECTIONS,
    }));
    setHasPendingChanges(false);
    setApplySignal((s) => s + 1);
    setShowClearConfirm(false);
    setApplyNotice(
      'All case data cleared across all tabs. Saved .docx templates and software configurations remain intact.'
    );
  };

  const handleScrutinizeStageSection = async (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey,
    forceOverwrite = false
  ) => {
    if (stageBatchScrutiny.running) return;
    if (
      !branchProposalFile &&
      !stageSourceFilesRef.current[docType] &&
      !stageUploadedSources[docType]?.uploaded &&
      !caseRef.current.extractedDocs.branchProposal &&
      !resolvedICMemo.placeholders['#2#']
    ) {
      setAiError(
        'Upload the Branch Proposal (on Manual Inputs) or the prior completed memo on this tab first so AI can scrutinize and populate this stage.'
      );
      return;
    }
    if (stageManualEditedSections[section] && !forceOverwrite) {
      return;
    }

    setScrutinizingStageSection(section);
    setAiError(null);

    try {
      const context = buildScrutinyContext();
      const response = await fetchStageEndpoint(
        docType,
        '/api/gemini/scrutinize-stage-section',
        {
          docType,
          section,
          context,
        }
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Stage AI scrutiny failed (${response.status}).`);
      }

      const data = (await response.json()) as {
        result?: { text?: string };
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      };
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      }

      const text = String(data.result?.text ?? '').trim();
      setStageAiSectionsSync((prev) => ({ ...prev, [section]: text }));
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: false }));
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Some reference details were incomplete in the current case documents.';
      setStageBatchScrutiny((prev) => {
        if (prev.docType !== docType && prev.total === 0) return prev;
        return {
          ...prev,
          docType,
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
        };
      });
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to scrutinize stage section.');
    } finally {
      setScrutinizingStageSection(null);
    }
  };

  const handleRetryStageSection = async (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey
  ) => {
    if (stageBatchScrutiny.running) return;
    if (stageBatchScrutiny.retryingSections.includes(section)) return;

    setAiError(null);
    setStageBatchScrutiny((prev) => ({
      ...prev,
      docType,
      retryingSections: prev.retryingSections.includes(section)
        ? prev.retryingSections
        : [...prev.retryingSections, section],
    }));

    try {
      const context = buildScrutinyContext();
      const response = await fetchStageEndpoint(
        docType,
        '/api/gemini/scrutinize-stage-section',
        {
          docType,
          section,
          context,
        }
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Retry failed (${response.status}).`);
      }

      const data = (await response.json()) as {
        result?: { text?: string };
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      };
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      }

      const text = String(data.result?.text ?? '').trim();
      setStageAiSectionsSync((prev) => ({ ...prev, [section]: text }));
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: false }));
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Some reference details were incomplete in the current case documents.';

      setStageBatchScrutiny((prev) => ({
        ...prev,
        docType,
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
      setStageBatchScrutiny((prev) => ({
        ...prev,
        docType,
        completedSections: prev.completedSections.filter((s) => s !== section),
        failedSections: [
          ...prev.failedSections.filter((f) => f.section !== section),
          { section, error: errMsg },
        ],
      }));
    } finally {
      setStageBatchScrutiny((prev) => ({
        ...prev,
        retryingSections: prev.retryingSections.filter((s) => s !== section),
      }));
    }
  };

  const handleScrutinizeStageAll = async (docType: StageMemoDocType) => {
    if (stageBatchScrutiny.running || scrutinizingStageSection !== null) return;
    if (
      !branchProposalFile &&
      !stageSourceFilesRef.current[docType] &&
      !stageUploadedSources[docType]?.uploaded &&
      !caseRef.current.extractedDocs.branchProposal &&
      !resolvedICMemo.placeholders['#2#']
    ) {
      setAiError(
        'Upload the Branch Proposal (on Manual Inputs) or the prior completed memo on this tab first so AI can scrutinize and populate this stage.'
      );
      return;
    }

    stageBatchAbortRef.current?.abort();
    const controller = new AbortController();
    stageBatchAbortRef.current = controller;
    setAiError(null);

    const allStageKeys = STAGE_SECTIONS_BY_DOC[docType];
    const skippedManualSections: StageScrutinySectionKey[] = [];
    const targetSections: StageScrutinySectionKey[] = [];
    for (const s of allStageKeys) {
      if (stageManualEditedSections[s]) {
        skippedManualSections.push(s);
      } else {
        targetSections.push(s);
      }
    }

    if (targetSections.length === 0) {
      setStageBatchScrutiny({
        running: false,
        docType,
        total: allStageKeys.length,
        targetSections: [],
        completedSections: [],
        retryingSections: [],
        failedSections: [],
        insufficientSections: [],
        skippedManualSections,
      });
      return;
    }

    setStageBatchScrutiny({
      running: true,
      docType,
      total: targetSections.length,
      targetSections,
      completedSections: [],
      retryingSections: [],
      failedSections: [],
      insufficientSections: [],
      skippedManualSections,
    });

    try {
      const context = buildScrutinyContext();
      const response = await fetchStageEndpoint(
        docType,
        '/api/gemini/scrutinize-stage-all',
        { docType, sections: targetSections, context },
        controller.signal
      );

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Stage batch scrutiny failed (${response.status}).`);
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

          try {
            const event = JSON.parse(dataLines.join('\n')) as {
              type: 'section' | 'error' | 'fatal' | 'done';
              section?: StageScrutinySectionKey;
              result?: { text?: string };
              insufficient?: boolean;
              insufficientReason?: string | null;
              error?: string;
              sourceId?: string;
            };

            if (event.type === 'section' && event.section && event.result) {
              const sec = event.section;
              const text = String(event.result.text ?? '').trim();
              setStageAiSectionsSync((prev) => ({ ...prev, [sec]: text }));
              setStageManualEditedSectionsSync((prev) => ({ ...prev, [sec]: false }));
              syncAppliedSnapshotFromRef();
              void persistCentralCaseNow(true);

              const isInsufficient = Boolean(event.insufficient);
              const reason =
                event.insufficientReason?.trim() ||
                'Some reference details were incomplete in the current case documents.';
              setStageBatchScrutiny((prev) => ({
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
              setStageBatchScrutiny((prev) => ({
                ...prev,
                failedSections: [
                  ...prev.failedSections.filter((f) => f.section !== sec),
                  { section: sec, error: errMsg },
                ],
              }));
            } else if (event.type === 'fatal') {
              setAiError(event.error || 'Stage batch AI scrutiny failed.');
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
        setAiError(err instanceof Error ? err.message : 'Stage batch AI scrutiny failed.');
      }
    } finally {
      setStageBatchScrutiny((prev) => {
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
      if (stageBatchAbortRef.current === controller) {
        stageBatchAbortRef.current = null;
      }
    }
  };

  const handleUploadTemplate = async (docType: MemoDocType, file: File) => {
    const record = await inspectDocxTemplate(file, docType);
    setTemplates((prev) => ({ ...prev, [docType]: record }));
    await saveTemplateRecord(record);
    syncAppliedSnapshotFromRef();
  };

  const handleRemoveTemplate = async (docType: MemoDocType) => {
    setTemplates((prev) => ({ ...prev, [docType]: null }));
    await deleteTemplateRecord(docType);
    syncAppliedSnapshotFromRef();
  };

  const handleApplyAllChanges = () => {
    void persistCentralCaseNow(true);
    syncAppliedSnapshotFromRef();
    const nowStr = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setApplySignal((s) => s + 1);
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

  const prepareVerifiedSynchronizedExport = async (docType: MemoDocType) => {
    const activeTemplate = templates[docType];

    // Flush any queued microtask/state updates and persist central case data immediately
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    await persistCentralCaseNow(true);
    const syncedMemo = syncAppliedSnapshotFromRef();
    const latest = caseRef.current;
    const syncedStageMemo =
      docType === 'ic'
        ? null
        : resolveStageMemo(docType, syncedMemo, latest.stageInputs, latest.stageAiSections);

    const docxBlob = await generateSynchronizedDocxBlob(
      docType,
      syncedMemo,
      activeTemplate,
      syncedStageMemo
    );
    const verification = await verifyPopulatedDocxBlob(
      docxBlob,
      syncedMemo,
      docType,
      activeTemplate,
      syncedStageMemo
    );
    if (!verification.verified) {
      throw new Error(
        `Download verification failed: populated content for ${verification.missingLabels.join(', ')} was not included in the generated document.`
      );
    }

    const acct = syncedMemo.placeholders['#2#'] || 'Draft';
    const safeName = acct.replace(/[^A-Za-z0-9_-]/g, '_');
    return { docxBlob, safeName };
  };

  const handleDownloadDocx = async (docType: MemoDocType = 'ic') => {
    setDownloading(true);
    setAiError(null);
    try {
      const { docxBlob, safeName } = await prepareVerifiedSynchronizedExport(docType);
      triggerBlobDownload(docxBlob, `${DOC_FILE_PREFIX[docType]}_${safeName}.docx`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .docx document.');
    } finally {
      setDownloading(false);
    }
  };

  const handleDownloadPdf = async (docType: MemoDocType = 'ic') => {
    setDownloadingPdf(true);
    setAiError(null);
    try {
      const { docxBlob, safeName } = await prepareVerifiedSynchronizedExport(docType);
      const pdfBlob = await generatePdfFromDocxBlob(docxBlob);
      triggerBlobDownload(pdfBlob, `${DOC_FILE_PREFIX[docType]}_${safeName}.pdf`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .pdf document.');
    } finally {
      setDownloadingPdf(false);
    }
  };

  if (!entered) {
    return <LandingScreen onEnter={() => setEntered(true)} />;
  }

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
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => setShowClearConfirm(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              title="Clear all case data across all tabs to start a new case"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Clear All Data</span>
              <span className="md:hidden">Clear</span>
            </button>

            <button
              type="button"
              onClick={handleApplyAllChanges}
              className={`hidden sm:inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                hasPendingChanges
                  ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                  : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
              }`}
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Apply All Changes</span>
            </button>

            {activeTab !== 'manual' ? (
              <div className="hidden sm:flex items-center gap-2">
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
              </div>
            ) : (
              <button
                type="button"
                onClick={() => setActiveTab('ic')}
                className="hidden sm:inline-flex px-4 py-2 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
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

        {aiError && activeTab !== 'ic' && (
          <div className="bg-red-50 border-t border-red-200 px-4 sm:px-6 py-2">
            <div className="max-w-7xl mx-auto flex items-center justify-between gap-2 text-xs font-medium text-red-800">
              <span>{aiError}</span>
              <button
                type="button"
                onClick={() => setAiError(null)}
                className="underline cursor-pointer shrink-0"
              >
                Dismiss
              </button>
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
            onOpenClearConfirm={() => setShowClearConfirm(true)}
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
          <StageMemoTab
            docType="mec"
            stageMemo={resolvedMecMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('mec')}
            onDownloadPdf={() => handleDownloadPdf('mec')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
            stageSourceMeta={stageUploadedSources.mec ?? null}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            onOpenClearConfirm={() => setShowClearConfirm(true)}
          />
        )}

        {activeTab === 'board' && (
          <StageMemoTab
            docType="board"
            stageMemo={resolvedBoardMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('board')}
            onDownloadPdf={() => handleDownloadPdf('board')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
            stageSourceMeta={stageUploadedSources.board ?? null}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            onOpenClearConfirm={() => setShowClearConfirm(true)}
          />
        )}

        {activeTab === 'sanction' && (
          <StageMemoTab
            docType="sanction"
            stageMemo={resolvedSanctionMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('sanction')}
            onDownloadPdf={() => handleDownloadPdf('sanction')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
            stageSourceMeta={stageUploadedSources.sanction ?? null}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            onOpenClearConfirm={() => setShowClearConfirm(true)}
          />
        )}
      </main>

      {/* Clear All Data Confirmation Modal */}
      {showClearConfirm && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 px-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="clear-all-modal-title"
        >
          <div className="bg-white border border-slate-200 rounded-xl shadow-xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start gap-3">
              <div className="p-2 rounded-lg bg-rose-50 text-rose-700 shrink-0">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div className="space-y-1.5">
                <h3
                  id="clear-all-modal-title"
                  className="font-display text-base font-bold text-slate-900"
                >
                  Clear All Case Data?
                </h3>
                <p className="text-xs text-slate-600 leading-relaxed">
                  This will reset all uploaded case documents, prior stage memo uploads, manual
                  inputs, field overrides, and AI-generated content across Manual Inputs, IC Memo,
                  MEC Memo, Board Memo, and Sanction so you can start a fresh case.
                </p>
                <p className="text-xs text-emerald-800 font-medium">
                  Your saved .docx document templates, default bank formats, and software
                  configurations will NOT be affected.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setShowClearConfirm(false)}
                className="px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleClearAllData}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-rose-700 hover:bg-rose-800 rounded-lg cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Yes, Clear All Data</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
