import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  ArrowRight,
  CheckCircle2,
  Download,
  FileCheck2,
  FileUp,
  Loader2,
  Pencil,
  RefreshCw,
  RotateCcw,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import {
  buildBoardMecRecommendationText,
  buildMecIcRecommendationText,
  formatChronologicOrderWithYear,
  STAGE_SECTION_LABELS,
  STAGE_SECTIONS_BY_DOC,
} from '../engines/icMemoEngine';
import { renderTemplatePreviewBlocks } from '../exporters/docxTemplateEngine';
import type {
  MemoDocType,
  MemoTemplateRecord,
  ResolvedStageMemo,
  StageAiSectionsData,
  StageBatchScrutinyState,
  StageMemoDocType,
  StageMemoInputs,
  StageScrutinySectionKey,
  StageUploadedSourceMeta,
} from '../types/memo';
import { CarriedICMemoBody } from './CarriedICMemoBody';
import { TemplateManagerPanel } from './TemplateManagerPanel';

interface StageMemoTabProps {
  docType: StageMemoDocType;
  stageMemo: ResolvedStageMemo;
  stageInputs: StageMemoInputs;
  stageAiSections: StageAiSectionsData;
  stageManualEditedSections: Partial<Record<StageScrutinySectionKey, boolean>>;
  scrutinizingStageSection: StageScrutinySectionKey | null;
  stageBatchScrutiny: StageBatchScrutinyState;
  hasBranchProposalFile: boolean;
  hasAnyCaseData: boolean;
  aiError: string | null;
  onStageInputChange: (field: keyof StageMemoInputs, value: string) => void;
  onUpdateStageAiSection: (
    section: StageScrutinySectionKey,
    value: string,
    markManual?: boolean
  ) => void;
  onScrutinizeStageSection: (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey,
    forceOverwrite?: boolean
  ) => Promise<void>;
  onRetryStageSection: (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey
  ) => Promise<void>;
  onScrutinizeStageAll: (docType: StageMemoDocType) => Promise<void>;
  onDismissStageBatchSummary: () => void;
  templates: Record<MemoDocType, MemoTemplateRecord | null>;
  onUploadTemplate: (docType: MemoDocType, file: File) => Promise<void>;
  onRemoveTemplate: (docType: MemoDocType) => Promise<void>;
  onDownloadDocx: () => Promise<void>;
  onDownloadPdf: () => Promise<void>;
  downloading: boolean;
  downloadingPdf: boolean;
  onApplyAllChanges: () => void;
  hasPendingChanges: boolean;
  lastAppliedAt: string | null;
  applySignal: number;
  stageSourceMeta: StageUploadedSourceMeta | null;
  onUploadStageSource: (docType: StageMemoDocType, file: File) => Promise<void>;
  onClearStageSource: (docType: StageMemoDocType) => void;
  onOpenClearConfirm?: () => void;
}

const STAGE_META: Record<
  StageMemoDocType,
  {
    title: string;
    subtitle: string;
    toLine: string;
    boxLabel: string;
    sectionHeader: string;
    signatories: string[];
  }
> = {
  mec: {
    title: 'MEC Memo',
    subtitle: 'Management Executive Committee (MEC) Appraisal & Waiver Memorandum',
    toLine: 'Management Executive Committee (MEC)',
    boxLabel: 'MEC Memorandum — Exit with Waiver',
    sectionHeader: 'MEC Specific Scrutiny, IC Reference & Recommendation',
    signatories: [
      'Officer / Appraiser (SAMD)',
      'Head of SAMD',
      'Chief Risk Officer (CRO)',
      'Managing Director & CEO',
    ],
  },
  board: {
    title: 'Board Memo',
    subtitle: 'Board of Directors Settlement & Waiver Approval Memorandum',
    toLine: 'Board of Directors',
    boxLabel: 'Board Memorandum — Exit with Waiver',
    sectionHeader: 'Board Specific Scrutiny, MEC Reference & Resolution Terms',
    signatories: [
      'Prepared By (SAMD)',
      'Head of SAMD',
      'Company Secretary',
      'Chairman / Board of Directors',
    ],
  },
  sanction: {
    title: 'Sanction',
    subtitle: 'Head Office Sanction Letter & Cross-Verified Operative Covenants',
    toLine: 'Branch Manager / Borrower',
    boxLabel: 'Sanction Letter — Exit with Waiver',
    sectionHeader: 'Sanction Authority, Cross-Verified Figures & Operative Conditions',
    signatories: [
      'Prepared & Verified By (SAMD)',
      'Checked By (Credit Admin)',
      'Head of SAMD',
      'Authorized Signatory',
    ],
  },
};

const PhVal: React.FC<{ ph: Record<string, string>; code: string; bold?: boolean }> = ({
  ph,
  code,
  bold = true,
}) => {
  const val = ph[code];
  if (!val || val.trim() === '') {
    return (
      <span className="bg-amber-100 text-amber-900 px-1 rounded font-mono text-[8.5pt] font-semibold">
        {code}
      </span>
    );
  }
  return <span className={bold ? 'font-bold' : ''}>{val}</span>;
};

const PRIOR_UPLOAD_CONFIG: Record<
  StageMemoDocType,
  {
    heading: string;
    sourceLabel: string;
    description: string;
  }
> = {
  mec: {
    heading: 'Upload Completed IC Memo (Optional Direct Source)',
    sourceLabel: 'Completed IC Memo',
    description:
      'All data from the IC Memo tab automatically carries forward. Optionally upload a finalized IC Memo (PDF, Word, Excel, or image) to extract, verify, and populate all client particulars, liabilities, justifications, and MEC sections.',
  },
  board: {
    heading: 'Upload Completed MEC Memo (Optional Direct Source)',
    sourceLabel: 'Completed MEC Memo',
    description:
      'All data from the IC & MEC Memo tabs automatically carries forward. Optionally upload a finalized MEC Memo (PDF, Word, Excel, or image) to extract, verify, and populate all particulars into the Board Memo format.',
  },
  sanction: {
    heading: 'Upload Completed Board Memo for AI Scrutiny (Optional Direct Source)',
    sourceLabel: 'Completed Board Memo',
    description:
      'All approved data from the Board Memo tab automatically carries forward. Optionally upload a completed Board Memo so AI reads, scrutinizes, cross-verifies every figure, flags missing/inconsistent items, and generates the formal Head Office Sanction Letter.',
  },
};

function renderHighlightedCommitteeRec(text: string): React.ReactNode {
  const lines = (text || '')
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length === 0) return '—';

  return lines.map((line, i) => {
    const match = line.match(
      /^(.*?\bin\s+its\s+)(.+?)(\s+meeting\s+held\s+on\s+)(.+?)(\.\s*The\s+Committee\s+recommended\s+the\s+proposal\s+as\s+presented\.?.*)$/i
    );
    if (match) {
      return (
        <span key={i} className="block">
          {match[1]}
          <strong className="font-bold">{match[2]}</strong>
          {match[3]}
          <strong className="font-bold">{match[4]}</strong>
          {match[5]}
        </span>
      );
    }
    return (
      <span key={i} className="block">
        {line}
      </span>
    );
  });
}

export const StageMemoTab: React.FC<StageMemoTabProps> = ({
  docType,
  stageMemo,
  stageInputs,
  stageAiSections,
  stageManualEditedSections,
  scrutinizingStageSection,
  stageBatchScrutiny,
  hasBranchProposalFile,
  hasAnyCaseData,
  aiError,
  onStageInputChange,
  onUpdateStageAiSection,
  onScrutinizeStageSection,
  onRetryStageSection,
  onScrutinizeStageAll,
  onDismissStageBatchSummary,
  templates,
  onUploadTemplate,
  onRemoveTemplate,
  onDownloadDocx,
  onDownloadPdf,
  downloading,
  downloadingPdf,
  onApplyAllChanges,
  hasPendingChanges,
  lastAppliedAt,
  applySignal,
  stageSourceMeta,
  onUploadStageSource,
  onClearStageSource,
  onOpenClearConfirm,
}) => {
  const meta = STAGE_META[docType];
  const uploadCfg = PRIOR_UPLOAD_CONFIG[docType];
  const sectionKeys = STAGE_SECTIONS_BY_DOC[docType];
  const ic = stageMemo.icMemo;
  const ph = ic.placeholders;
  const tpl = templates[docType];

  const [editingSections, setEditingSections] = useState<
    Partial<Record<StageScrutinySectionKey, boolean>>
  >({});
  const [confirmOverwriteSection, setConfirmOverwriteSection] =
    useState<StageScrutinySectionKey | null>(null);

  useEffect(() => {
    if (applySignal > 0) {
      setEditingSections({});
      setConfirmOverwriteSection(null);
    }
  }, [applySignal]);

  const isThisDocBatchRunning =
    stageBatchScrutiny.running && stageBatchScrutiny.docType === docType;
  const isAnyBatchRunning = stageBatchScrutiny.running;
  const showBatchPanel =
    stageBatchScrutiny.docType === docType && stageBatchScrutiny.total > 0;

  const toggleEdit = (key: StageScrutinySectionKey) => {
    setEditingSections((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const handleRequestScrutiny = (key: StageScrutinySectionKey) => {
    if (stageManualEditedSections[key]) {
      setConfirmOverwriteSection(key);
      return;
    }
    setConfirmOverwriteSection(null);
    void onScrutinizeStageSection(docType, key, false);
  };

  const buildFallbackDisplay = (key: StageScrutinySectionKey): string => {
    const val = stageMemo.stageSections[key];
    if (val && val.trim() !== '') return val;
    if (key === 'mecIcRecommendation') {
      return buildMecIcRecommendationText(stageInputs.icMemoNumber, stageInputs.icMemoDate);
    }
    if (key === 'boardMecRecommendation') {
      return buildBoardMecRecommendationText(stageInputs.mecMemoNumber, stageInputs.mecMemoDate);
    }
    return '';
  };

  const renderSectionControlBar = (key: StageScrutinySectionKey) => {
    const isEditing = Boolean(editingSections[key]);
    const isRetrying = stageBatchScrutiny.retryingSections.includes(key);
    const isRunning = scrutinizingStageSection === key || isRetrying;
    const anyRunning = scrutinizingStageSection !== null || isAnyBatchRunning;
    const failEntry =
      stageBatchScrutiny.docType === docType
        ? stageBatchScrutiny.failedSections.find((f) => f.section === key)
        : undefined;
    const insuffEntry =
      stageBatchScrutiny.docType === docType
        ? stageBatchScrutiny.insufficientSections.find((i) => i.section === key)
        : undefined;
    const isProtected = Boolean(stageManualEditedSections[key]);

    return (
      <div className="font-sans space-y-1.5">
        <div className="flex items-center gap-1.5 flex-wrap justify-end">
          <button
            type="button"
            onClick={() => toggleEdit(key)}
            className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium rounded border cursor-pointer transition-colors ${
              isEditing
                ? 'text-emerald-900 bg-emerald-100 hover:bg-emerald-200 border-emerald-400'
                : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
            }`}
          >
            <Pencil className="w-3 h-3" />
            <span>{isEditing ? 'Done Editing' : 'Edit'}</span>
          </button>

          {(failEntry || isRetrying) && (
            <button
              type="button"
              onClick={() => void onRetryStageSection(docType, key)}
              disabled={isRetrying || anyRunning}
              className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60 transition-colors whitespace-nowrap"
            >
              {isRetrying ? (
                <Loader2 className="w-3 h-3 animate-spin" />
              ) : (
                <RotateCcw className="w-3 h-3" />
              )}
              <span>{isRetrying ? 'Retrying...' : 'Retry'}</span>
            </button>
          )}

          <button
            type="button"
            onClick={() => handleRequestScrutiny(key)}
            disabled={anyRunning || isRetrying}
            title={
              isProtected
                ? 'This section has manual edits — clicking will ask before overwriting'
                : 'Scrutinize source documents and IC data with AI to populate this section'
            }
            className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded cursor-pointer disabled:opacity-60 transition-colors whitespace-nowrap"
          >
            {isRunning ? (
              <Loader2 className="w-3.5 h-3.5 animate-spin" />
            ) : (
              <Sparkles className="w-3.5 h-3.5" />
            )}
            <span>
              {isRunning ? 'Scrutinizing...' : 'Scrutinize and Fill Up Through AI'}
            </span>
          </button>
        </div>

        {failEntry && !isRetrying && (
          <div className="bg-red-50 border border-red-200 rounded px-2.5 py-1 text-[11px] text-red-800">
            {failEntry.error}
          </div>
        )}

        {!failEntry && insuffEntry && (
          <div className="bg-amber-50 border border-amber-200 rounded px-2.5 py-1 text-[11px] text-amber-900">
            {insuffEntry.reason}
          </div>
        )}

        {confirmOverwriteSection === key && (
          <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50 border border-amber-300 rounded px-2.5 py-1.5 text-[11px] text-amber-950">
            <div className="flex items-center gap-1.5 font-medium">
              <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" />
              <span>
                This section has manual edits. Overwrite it with AI scrutiny?
              </span>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setConfirmOverwriteSection(null);
                  void onScrutinizeStageSection(docType, key, true);
                }}
                disabled={anyRunning}
                className="px-2 py-0.5 font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60"
              >
                Yes, Overwrite
              </button>
              <button
                type="button"
                onClick={() => setConfirmOverwriteSection(null)}
                className="px-2 py-0.5 font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded cursor-pointer"
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="max-w-5xl mx-auto space-y-5 pb-12">
      {/* Top Action Bar */}
      <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="font-display text-lg font-semibold text-slate-900">
                {meta.title}
              </h2>
              <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold bg-emerald-50 text-emerald-900 border border-emerald-200 rounded">
                <FileCheck2 className="w-3.5 h-3.5" />
                <span>Synchronized Stage</span>
              </span>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">{meta.subtitle}</p>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => void onScrutinizeStageAll(docType)}
              disabled={isAnyBatchRunning || scrutinizingStageSection !== null}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap"
            >
              {isThisDocBatchRunning ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Sparkles className="w-4 h-4" />
              )}
              <span>
                {isThisDocBatchRunning
                  ? `Generating (${stageBatchScrutiny.completedSections.length}/${stageBatchScrutiny.total})...`
                  : 'Generate and Fill Up All Together'}
              </span>
            </button>

            <button
              type="button"
              onClick={onApplyAllChanges}
              className={`inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-lg border transition-colors cursor-pointer whitespace-nowrap ${
                hasPendingChanges
                  ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                  : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
              }`}
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Apply All Changes</span>
            </button>

            <button
              type="button"
              onClick={() => void onDownloadDocx()}
              disabled={downloading || downloadingPdf}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap"
            >
              <Download className="w-4 h-4" />
              <span>{downloading ? 'Building .docx...' : 'Download .docx'}</span>
            </button>

            <button
              type="button"
              onClick={() => void onDownloadPdf()}
              disabled={downloading || downloadingPdf}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap"
            >
              <Download className="w-4 h-4" />
              <span>{downloadingPdf ? 'Rendering PDF...' : 'Download .pdf'}</span>
            </button>

            {onOpenClearConfirm && (
              <button
                type="button"
                onClick={onOpenClearConfirm}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                title="Clear all case data across all tabs to start a new case"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Clear All Data</span>
              </button>
            )}
          </div>
        </div>

        {/* Chained Workflow Strip */}
        <div className="flex flex-wrap items-center gap-1.5 pt-2 border-t border-slate-100 text-[11px] text-slate-600">
          <span className="font-semibold text-slate-800">Workflow Chain:</span>
          <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-700">
            Branch Proposal &amp; Docs
          </span>
          <ArrowRight className="w-3 h-3 text-slate-400" />
          <span
            className={`px-2 py-0.5 rounded ${
              docType === 'mec' || docType === 'board' || docType === 'sanction'
                ? 'bg-emerald-50 text-emerald-900 font-medium'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            IC Memo
            {stageInputs.icMemoNumber ? ` (${stageInputs.icMemoNumber})` : ''}
          </span>
          <ArrowRight className="w-3 h-3 text-slate-400" />
          <span
            className={`px-2 py-0.5 rounded ${
              docType === 'mec'
                ? 'bg-emerald-800 text-white font-semibold'
                : stageInputs.mecMemoNumber
                  ? 'bg-emerald-50 text-emerald-900 font-medium'
                  : 'bg-slate-100 text-slate-600'
            }`}
          >
            MEC Memo
            {stageInputs.mecMemoNumber ? ` (${stageInputs.mecMemoNumber})` : ''}
          </span>
          <ArrowRight className="w-3 h-3 text-slate-400" />
          <span
            className={`px-2 py-0.5 rounded ${
              docType === 'board'
                ? 'bg-emerald-800 text-white font-semibold'
                : stageInputs.boardMemoNumber
                  ? 'bg-emerald-50 text-emerald-900 font-medium'
                  : 'bg-slate-100 text-slate-600'
            }`}
          >
            Board Memo
            {stageInputs.boardMemoNumber ? ` (${stageInputs.boardMemoNumber})` : ''}
          </span>
          <ArrowRight className="w-3 h-3 text-slate-400" />
          <span
            className={`px-2 py-0.5 rounded ${
              docType === 'sanction'
                ? 'bg-emerald-800 text-white font-semibold'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            Sanction
          </span>
          {lastAppliedAt && (
            <span className="ml-auto text-slate-500">
              Last synchronized: <span className="font-medium text-slate-800">{lastAppliedAt}</span>
            </span>
          )}
        </div>
      </div>

      {/* Batch Progress & Summary Panel */}
      {showBatchPanel && (
        <div className="bg-emerald-50/70 border border-emerald-200 rounded-lg p-3.5 text-xs space-y-2.5">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 font-semibold text-emerald-950">
              {isThisDocBatchRunning ? (
                <Loader2 className="w-4 h-4 animate-spin text-emerald-800" />
              ) : (
                <CheckCircle2 className="w-4 h-4 text-emerald-800" />
              )}
              <span>
                {isThisDocBatchRunning
                  ? `Scrutinizing and filling ${stageBatchScrutiny.completedSections.length} of ${stageBatchScrutiny.total} ${meta.title} sections...`
                  : `${meta.title} AI scrutiny finished: ${stageBatchScrutiny.completedSections.length} of ${stageBatchScrutiny.total} sections populated.`}
              </span>
            </div>
            {!isThisDocBatchRunning && (
              <button
                type="button"
                onClick={onDismissStageBatchSummary}
                className="text-slate-600 hover:text-slate-900 underline cursor-pointer"
              >
                Dismiss
              </button>
            )}
          </div>

          <div className="w-full h-1.5 bg-emerald-200/80 rounded-full overflow-hidden">
            <div
              className="h-full bg-emerald-800 transition-all duration-300"
              style={{
                width: `${
                  stageBatchScrutiny.total > 0
                    ? Math.round(
                        ((stageBatchScrutiny.completedSections.length +
                          stageBatchScrutiny.failedSections.length) /
                          stageBatchScrutiny.total) *
                          100
                      )
                    : 0
                }%`,
              }}
            />
          </div>

          {stageBatchScrutiny.failedSections.length > 0 && (
            <div className="bg-white/90 border border-red-200 rounded-md p-2.5 space-y-1.5">
              <p className="font-semibold text-[11px] text-red-900">
                Sections Requiring Retry ({stageBatchScrutiny.failedSections.length}):
              </p>
              <div className="divide-y divide-red-100">
                {stageBatchScrutiny.failedSections.map((f) => {
                  const isRetrying = stageBatchScrutiny.retryingSections.includes(f.section);
                  return (
                    <div
                      key={f.section}
                      className="py-1.5 first:pt-0 last:pb-0 flex flex-wrap items-center justify-between gap-2"
                    >
                      <div className="text-[11px] text-red-900">
                        <span className="font-semibold">{STAGE_SECTION_LABELS[f.section]}:</span>{' '}
                        <span>{f.error}</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => void onRetryStageSection(docType, f.section)}
                        disabled={isRetrying || isAnyBatchRunning}
                        className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
                      >
                        {isRetrying ? (
                          <Loader2 className="w-3 h-3 animate-spin" />
                        ) : (
                          <RotateCcw className="w-3 h-3" />
                        )}
                        <span>{isRetrying ? 'Retrying...' : 'Retry'}</span>
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {stageBatchScrutiny.insufficientSections.length > 0 && (
            <div className="bg-white/90 border border-amber-300 rounded-md p-2.5 space-y-1.5">
              <p className="font-semibold text-[11px] text-amber-950">
                Additional Reference Suggested ({stageBatchScrutiny.insufficientSections.length}):
              </p>
              {stageBatchScrutiny.insufficientSections.map((item) => (
                <div key={item.section} className="text-[11px] text-amber-950">
                  <span className="font-semibold">{STAGE_SECTION_LABELS[item.section]}:</span>{' '}
                  <span>{item.reason}</span>
                </div>
              ))}
            </div>
          )}

          {stageBatchScrutiny.skippedManualSections.length > 0 && (
            <p className="text-[11px] text-slate-700">
              <span className="font-semibold">
                Preserved manual edits ({stageBatchScrutiny.skippedManualSections.length}):
              </span>{' '}
              {stageBatchScrutiny.skippedManualSections
                .map((s) => STAGE_SECTION_LABELS[s])
                .join(', ')}
            </p>
          )}
        </div>
      )}

      {!hasAnyCaseData && !hasBranchProposalFile && !stageSourceMeta?.uploaded && (
        <div className="bg-amber-50 border border-amber-200 text-amber-900 rounded-lg px-4 py-2.5 text-xs flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" />
          <span>
            Upload the case documents on the Manual Inputs tab (for automatic carry-forward), or upload the completed {docType === 'mec' ? 'IC Memo' : docType === 'board' ? 'MEC Memo' : 'Board Memo'} directly below to populate {meta.title}.
          </span>
        </div>
      )}

      {/* Upload Completed Prior Stage Memo Card */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="space-y-1">
            <h3 className="font-display text-sm font-semibold text-slate-900">
              {uploadCfg.heading}
            </h3>
            <p className="text-xs text-slate-600 leading-relaxed max-w-3xl">
              {uploadCfg.description}
            </p>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <label className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg cursor-pointer transition-colors whitespace-nowrap">
              {stageSourceMeta?.scrutinizing ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <FileUp className="w-3.5 h-3.5" />
              )}
              <span>
                {stageSourceMeta?.scrutinizing
                  ? 'Reading & Scrutinizing...'
                  : stageSourceMeta?.uploaded
                    ? `Replace ${uploadCfg.sourceLabel}`
                    : `Upload ${uploadCfg.sourceLabel}`}
              </span>
              <input
                type="file"
                accept=".pdf,.docx,.xlsx,.xls,image/*"
                disabled={Boolean(stageSourceMeta?.scrutinizing)}
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) {
                    void onUploadStageSource(docType, file);
                  }
                  e.target.value = '';
                }}
              />
            </label>

            {stageSourceMeta?.uploaded && !stageSourceMeta?.scrutinizing && (
              <button
                type="button"
                onClick={() => onClearStageSource(docType)}
                className="inline-flex items-center gap-1 px-2.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg cursor-pointer whitespace-nowrap"
                title="Remove uploaded prior memo file reference"
              >
                <X className="w-3.5 h-3.5" />
                <span>Remove File</span>
              </button>
            )}
          </div>
        </div>

        {stageSourceMeta?.scrutinizing && (
          <div className="bg-emerald-50 border border-emerald-200 rounded-md px-3 py-2 flex items-center gap-2 text-xs font-medium text-emerald-900">
            <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-800 shrink-0" />
            <span>
              {stageSourceMeta.stage || `Extracting & scrutinizing ${stageSourceMeta.fileName}...`}
            </span>
          </div>
        )}

        {stageSourceMeta?.uploaded && !stageSourceMeta?.scrutinizing && (
          <div className="bg-emerald-50/70 border border-emerald-200 rounded-md px-3 py-2 flex flex-wrap items-center justify-between gap-2 text-xs text-emerald-950">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
              <span>
                Loaded &amp; Scrutinized: <span className="font-semibold">{stageSourceMeta.fileName}</span>
                {stageSourceMeta.uploadedAt ? ` (${stageSourceMeta.uploadedAt})` : ''}
              </span>
            </div>
            <span className="text-[11px] text-emerald-800 font-medium">
              Extracted data merged into {meta.title}
            </span>
          </div>
        )}

        {stageSourceMeta?.error && (
          <div className="bg-red-50 border border-red-200 rounded-md px-3 py-2 text-xs text-red-800 font-medium">
            {stageSourceMeta.error}
          </div>
        )}

        {stageSourceMeta?.missingOrInconsistentNotes &&
          stageSourceMeta.missingOrInconsistentNotes.length > 0 && (
            <div className="bg-amber-50 border border-amber-300 rounded-md p-3 space-y-1 text-xs text-amber-950">
              <div className="flex items-center gap-1.5 font-semibold">
                <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" />
                <span>AI Verification &amp; Missing/Inconsistent Data Notes:</span>
              </div>
              <ul className="list-disc pl-5 space-y-0.5 text-[11px]">
                {stageSourceMeta.missingOrInconsistentNotes.map((note, idx) => (
                  <li key={idx}>{note}</li>
                ))}
              </ul>
            </div>
          )}
      </div>

      {aiError && (
        <div className="bg-red-50 border border-red-200 text-red-800 rounded-lg px-4 py-3 text-sm">
          {aiError}
        </div>
      )}

      {/* Stage Inputs & Chained Reference Card */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-3">
          <div>
            <h3 className="font-display text-sm font-semibold text-slate-900">
              {docType === 'mec'
                ? 'IC Memo Reference Inputs & MEC Reference'
                : docType === 'board'
                  ? 'MEC Memo Reference Inputs & Board Reference'
                  : 'Board Memo Reference Inputs & Sanction Reference'}
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              {docType === 'mec'
                ? 'Enter the Chronologic Order of the IC Memo and IC Memo Date. All IC Memo data carries forward intact, and these references are incorporated into “Recommendation of the Investment Committee (IC)” and all applicable MEC Memo sections.'
                : docType === 'board'
                  ? 'Enter the Chronologic Order of the MEC Memo and MEC Memo Date. All IC Memo and MEC Memo data carry forward intact, and these references are incorporated into “Recommendation of the Management Executive Committee (MEC)” and all Board Memo sections.'
                  : 'Enter the Chronologic Order of the Board Memo and Board Memo Date. All approved IC, MEC, and Board details carry forward intact into Sanction cross-verification and operative conditions.'}
            </p>
          </div>
        </div>

        {docType === 'mec' && (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            <div>
              <label
                htmlFor="mec-ic-memo-number"
                className="block text-xs font-semibold text-slate-800 mb-1"
              >
                Chronologic Order of the IC Memo <span className="text-emerald-800">*</span>
              </label>
              <input
                id="mec-ic-memo-number"
                type="text"
                value={stageInputs.icMemoNumber}
                onChange={(e) => onStageInputChange('icMemoNumber', e.target.value)}
                placeholder="e.g., SAMD/IC/2026/114"
                className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
              />
              {stageInputs.icMemoNumber.trim() && (
                <p className="text-[11px] text-emerald-800 mt-1">
                  Sec. 16 format:{' '}
                  <span className="font-semibold">
                    {formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')}
                  </span>
                </p>
              )}
            </div>
            <div>
              <label
                htmlFor="mec-ic-memo-date"
                className="block text-xs font-semibold text-slate-800 mb-1"
              >
                IC Memo Date <span className="text-emerald-800">*</span>
              </label>
              <input
                id="mec-ic-memo-date"
                type="text"
                value={stageInputs.icMemoDate}
                onChange={(e) => onStageInputChange('icMemoDate', e.target.value)}
                placeholder="e.g., 12.11.2026"
                className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
              />
            </div>
            <div>
              <label
                htmlFor="mec-self-memo-number"
                className="block text-xs font-medium text-slate-600 mb-1"
              >
                Chronologic Order of the MEC Memo <span className="text-emerald-800">*</span>
              </label>
              <input
                id="mec-self-memo-number"
                type="text"
                value={stageInputs.mecMemoNumber}
                onChange={(e) => onStageInputChange('mecMemoNumber', e.target.value)}
                placeholder="e.g., SAMD/MEC/2026/058"
                className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
              />
            </div>
            <div>
              <label
                htmlFor="mec-self-memo-date"
                className="block text-xs font-medium text-slate-600 mb-1"
              >
                MEC Memo Date (carried to Board)
              </label>
              <input
                id="mec-self-memo-date"
                type="text"
                value={stageInputs.mecMemoDate}
                onChange={(e) => onStageInputChange('mecMemoDate', e.target.value)}
                placeholder="e.g., 16.11.2026"
                className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
              />
            </div>
          </div>
        )}

        {docType === 'board' && (
          <div className="space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              <div>
                <label
                  htmlFor="board-mec-memo-number"
                  className="block text-xs font-semibold text-slate-800 mb-1"
                >
                  Chronologic Order of the MEC Memo <span className="text-emerald-800">*</span>
                </label>
                <input
                  id="board-mec-memo-number"
                  type="text"
                  value={stageInputs.mecMemoNumber}
                  onChange={(e) => onStageInputChange('mecMemoNumber', e.target.value)}
                  placeholder="e.g., SAMD/MEC/2026/058"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
                {stageInputs.mecMemoNumber.trim() && (
                  <p className="text-[11px] text-emerald-800 mt-1">
                    Sec. 17 format:{' '}
                    <span className="font-semibold">
                      {formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')}
                    </span>
                  </p>
                )}
              </div>
              <div>
                <label
                  htmlFor="board-mec-memo-date"
                  className="block text-xs font-semibold text-slate-800 mb-1"
                >
                  MEC Memo Date <span className="text-emerald-800">*</span>
                </label>
                <input
                  id="board-mec-memo-date"
                  type="text"
                  value={stageInputs.mecMemoDate}
                  onChange={(e) => onStageInputChange('mecMemoDate', e.target.value)}
                  placeholder="e.g., 16.11.2026"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
              <div>
                <label
                  htmlFor="board-self-memo-number"
                  className="block text-xs font-medium text-slate-600 mb-1"
                >
                  Chronologic Order of the Board Memo <span className="text-emerald-800">*</span>
                </label>
                <input
                  id="board-self-memo-number"
                  type="text"
                  value={stageInputs.boardMemoNumber}
                  onChange={(e) => onStageInputChange('boardMemoNumber', e.target.value)}
                  placeholder="e.g., SAMD/BRD/2026/019"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
              <div>
                <label
                  htmlFor="board-self-memo-date"
                  className="block text-xs font-medium text-slate-600 mb-1"
                >
                  Board Memo Date (carried to Sanction)
                </label>
                <input
                  id="board-self-memo-date"
                  type="text"
                  value={stageInputs.boardMemoDate}
                  onChange={(e) => onStageInputChange('boardMemoDate', e.target.value)}
                  placeholder="e.g., 20.11.2026"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-md px-3 py-2 text-xs text-slate-700 flex flex-wrap gap-x-6 gap-y-1">
              <span>
                <span className="font-semibold">Carried Chronologic Order of the IC Memo:</span>{' '}
                {stageInputs.icMemoNumber || 'Not entered'}
              </span>
              <span>
                <span className="font-semibold">Carried IC Memo Date:</span>{' '}
                {stageInputs.icMemoDate || 'Not entered'}
              </span>
            </div>
          </div>
        )}

        {docType === 'sanction' && (
          <div className="space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              <div>
                <label
                  htmlFor="sanction-board-memo-number"
                  className="block text-xs font-semibold text-slate-800 mb-1"
                >
                  Chronologic Order of the Board Memo <span className="text-emerald-800">*</span>
                </label>
                <input
                  id="sanction-board-memo-number"
                  type="text"
                  value={stageInputs.boardMemoNumber}
                  onChange={(e) => onStageInputChange('boardMemoNumber', e.target.value)}
                  placeholder="e.g., SAMD/BRD/2026/019"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
              <div>
                <label
                  htmlFor="sanction-board-memo-date"
                  className="block text-xs font-semibold text-slate-800 mb-1"
                >
                  Board Memo Date <span className="text-emerald-800">*</span>
                </label>
                <input
                  id="sanction-board-memo-date"
                  type="text"
                  value={stageInputs.boardMemoDate}
                  onChange={(e) => onStageInputChange('boardMemoDate', e.target.value)}
                  placeholder="e.g., 20.11.2026"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
              <div>
                <label
                  htmlFor="sanction-ref-number"
                  className="block text-xs font-medium text-slate-600 mb-1"
                >
                  Sanction Reference No.
                </label>
                <input
                  id="sanction-ref-number"
                  type="text"
                  value={stageInputs.sanctionRefNumber}
                  onChange={(e) => onStageInputChange('sanctionRefNumber', e.target.value)}
                  placeholder="e.g., SIBL/HO/SAMD/2026/312"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
              <div>
                <label
                  htmlFor="sanction-date"
                  className="block text-xs font-medium text-slate-600 mb-1"
                >
                  Sanction Date
                </label>
                <input
                  id="sanction-date"
                  type="text"
                  value={stageInputs.sanctionDate}
                  onChange={(e) => onStageInputChange('sanctionDate', e.target.value)}
                  placeholder="e.g., 24.11.2026"
                  className="w-full px-3 py-2 text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700"
                />
              </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-md px-3 py-2 text-xs text-slate-700 flex flex-wrap gap-x-6 gap-y-1">
              <span>
                <span className="font-semibold">Carried Chronologic Order of the IC Memo:</span>{' '}
                {stageInputs.icMemoNumber || '—'}{' '}
                {stageInputs.icMemoDate ? `dated ${stageInputs.icMemoDate}` : ''}
              </span>
              <span>
                <span className="font-semibold">Carried Chronologic Order of the MEC Memo:</span>{' '}
                {stageInputs.mecMemoNumber || '—'}{' '}
                {stageInputs.mecMemoDate ? `dated ${stageInputs.mecMemoDate}` : ''}
              </span>
            </div>
          </div>
        )}
      </div>

      {/* Stage Template Upload Control */}
      <TemplateManagerPanel
        templates={templates}
        onUploadTemplate={onUploadTemplate}
        onRemoveTemplate={onRemoveTemplate}
        singleDocType={docType}
      />

      {/* Interactive Stage AI Sections Card */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 space-y-4">
        <div className="border-b border-slate-100 pb-2">
          <h3 className="font-display text-sm font-semibold text-slate-900">
            {meta.sectionHeader}
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Each section can be individually scrutinized and populated through Gemini AI, manually edited, or retried, and is automatically synchronized into both the built-in format and any uploaded {meta.title} .docx template.
          </p>
        </div>

        <div className="space-y-4">
          {sectionKeys.map((secKey) => {
            const label = STAGE_SECTION_LABELS[secKey];
            const isEditing = Boolean(editingSections[secKey]);
            const rawVal = stageAiSections[secKey] ?? '';
            const displayText = buildFallbackDisplay(secKey);

            return (
              <div
                key={secKey}
                className="border border-slate-200 rounded-lg p-3.5 bg-slate-50/40 space-y-2"
              >
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <h4 className="text-xs font-bold text-slate-900">{label}</h4>
                    {stageManualEditedSections[secKey] && (
                      <span className="px-1.5 py-0.5 text-[10px] font-medium bg-amber-50 text-amber-900 border border-amber-200 rounded">
                        Manually Edited
                      </span>
                    )}
                  </div>
                  {renderSectionControlBar(secKey)}
                </div>

                {isEditing ? (
                  <textarea
                    rows={5}
                    value={stageAiSections[secKey] !== null ? stageAiSections[secKey]! : displayText}
                    onChange={(e) => onUpdateStageAiSection(secKey, e.target.value, true)}
                    placeholder={`Enter or refine ${label}...`}
                    className="w-full px-3 py-2 text-xs text-slate-900 bg-white border border-emerald-400 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-700 leading-relaxed"
                  />
                ) : (
                  <div className="bg-white border border-slate-200 rounded-md px-3 py-2.5 text-xs text-slate-800 whitespace-pre-line leading-relaxed min-h-[54px]">
                    {displayText || (
                      <span className="text-slate-400 italic">
                        Not generated yet — click “Scrutinize and Fill Up Through AI” or “Edit” to populate this section.
                      </span>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Live Document Preview (Uploaded Template OR Built-In Bank Sheet) */}
      {tpl ? (
        <div className="bg-slate-200/70 p-4 sm:p-8 rounded-lg overflow-x-auto space-y-3">
          <div className="max-w-[794px] mx-auto bg-emerald-50 border border-emerald-200 rounded-md px-4 py-2 text-xs text-emerald-950 flex items-center justify-between gap-2">
            <span>
              <span className="font-semibold">Active Uploaded Template:</span> {tpl.fileName} — all IC Memo data, stage numbers/dates, and AI sections above are synchronized while preserving original margins, fonts, and table geometry.
            </span>
          </div>
          <div className="ic-memo-sheet max-w-[794px] mx-auto bg-white shadow-sm border border-slate-300 px-12 py-8 space-y-2.5 text-black text-[10pt]">
            {renderTemplatePreviewBlocks(tpl, ic, stageMemo).map((block, idx) => {
              if (block.type === 'paragraph') {
                return (
                  <p key={idx} className="whitespace-pre-line leading-snug">
                    {block.text}
                  </p>
                );
              }
              return (
                <table
                  key={idx}
                  className="w-full table-fixed border-collapse border border-black text-[9.5pt]"
                >
                  <tbody>
                    {(block.rows || []).map((row, rIdx) => (
                      <tr key={rIdx}>
                        {row.map((cell, cIdx) => (
                          <td
                            key={cIdx}
                            className="border border-black px-2 py-1 align-middle whitespace-pre-line"
                          >
                            {cell}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              );
            })}
          </div>
        </div>
      ) : (
        <div className="bg-slate-200/70 p-4 sm:p-8 rounded-lg overflow-x-auto">
          <div className="ic-memo-sheet max-w-[794px] mx-auto bg-white shadow-sm border border-slate-300 px-14 py-8 space-y-3 text-black">
            {/* Header */}
            <div className="-mx-6">
              <div className="flex flex-col items-center">
                <img
                  src="/assets/bismillah.jpg"
                  alt="Bismillahir Rahmanir Rahim"
                  referrerPolicy="no-referrer"
                  className="h-5 object-contain mb-1"
                />
              </div>

              <div className="flex items-start justify-between gap-3">
                <img
                  src="/assets/sib_logo.jpg"
                  alt="Sammilito Islami Bank PLC"
                  referrerPolicy="no-referrer"
                  className="w-14 h-auto object-contain shrink-0"
                />
                <div className="text-center flex-1 leading-tight">
                  <p className="text-[14pt] underline">SAMMILITO ISLAMI BANK PLC.</p>
                  <p className="underline">
                    <span className="text-[9pt]">(</span>
                    <span className="text-[13pt]">State-owned Shariah-based Bank</span>
                    <span className="text-[9pt]">)</span>
                  </p>
                  <p className="text-[11pt] underline">UNION BANK UNIT</p>
                  <p className="text-[9pt] underline">
                    Special Asset Management Division Head Office, Dhaka
                  </p>
                </div>
                <div className="border border-black px-2.5 py-1 text-[10pt] font-bold whitespace-nowrap shrink-0 mt-1 text-center">
                  {meta.boxLabel}
                </div>
              </div>
            </div>

            {/* Stage Metadata & Body Rendered Per Prescribed Format */}
            {docType === 'sanction' ? (
              <>
                <table className="w-[82%] table-fixed text-[10pt] border border-black">
                  <colgroup>
                    <col className="w-[38%]" />
                    <col className="w-[4%]" />
                    <col className="w-[58%]" />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Sanction Ref. No.</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {stageInputs.sanctionRefNumber || '[Sanction Reference Pending]'}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Sanction Date</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {stageInputs.sanctionDate || stageInputs.boardMemoDate || ph['#1#'] || '—'}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Board Memo Reference</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {formatChronologicOrderWithYear(
                          stageInputs.boardMemoNumber,
                          '[Chronologic Order of the Board Memo]'
                        )}{' '}
                        dated {stageInputs.boardMemoDate || '[Board Memo Date]'}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Prior Committee Chain</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5">
                        MEC:{' '}
                        {stageInputs.mecMemoNumber
                          ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
                          : '—'}{' '}
                        ({stageInputs.mecMemoDate || '—'}) | IC:{' '}
                        {stageInputs.icMemoNumber
                          ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
                          : '—'}{' '}
                        ({stageInputs.icMemoDate || '—'})
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Branch Proposal Ref.</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5">
                        <PhVal ph={ph} code="#64#" bold={false} /> dated{' '}
                        <PhVal ph={ph} code="#65#" bold={false} />
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">From</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        Special Asset Management Division (SAMD), Head Office
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">To</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        The Head of Branch, <PhVal ph={ph} code="#3#" /> Branch
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="text-[10pt] leading-snug pt-1 space-y-0.5">
                  <p className="font-bold">The Head of Branch,</p>
                  <p>Sammilito Islami Bank PLC.</p>
                  <p>
                    <PhVal ph={ph} code="#3#" bold={false} /> Branch
                  </p>
                  <p className="pt-1">Dear Sir / Madam,</p>
                </div>

                <div className="border border-black px-2.5 py-1.5 text-[10pt] font-bold leading-snug">
                  Subject: Head Office Sanction of Full &amp; Final Exit with Waiver Facility — A/C:{' '}
                  <PhVal ph={ph} code="#2#" /> (CIS: <PhVal ph={ph} code="#69#" />), Client:{' '}
                  {ic.branchProposal?.section10Particulars?.nameOfClient?.trim() || (
                    <PhVal ph={ph} code="#2#" />
                  )}
                  , Branch: <PhVal ph={ph} code="#3#" /> under BRPD Circular Letter No. 23 dated
                  29.06.2026.
                </div>

                <p className="text-right text-[10pt] font-bold">
                  [All Figures in this Sanction Letter are in Million BDT Unless Stated Otherwise]
                </p>

                {/* 1. Board Approval Reference & Sanction Basis */}
                <div className="space-y-1 pt-1">
                  <p className="font-bold underline text-[10pt]">
                    1. Board Approval Reference &amp; Sanction Basis:
                  </p>
                  <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                    {buildFallbackDisplay('sanctionBoardAuthority') || '—'}
                  </p>
                </div>

                {/* 2. Particulars of Sanctioned Facility, Client & Liability Matrix */}
                <div className="space-y-1 pt-1">
                  <p className="font-bold underline text-[10pt]">
                    2. Particulars of Sanctioned Facility, Client &amp; Liability Matrix:
                  </p>
                  <table className="w-full table-fixed text-[9.5pt] border border-black tabular-nums">
                    <colgroup>
                      <col className="w-[34%]" />
                      <col className="w-[66%]" />
                    </colgroup>
                    <thead>
                      <tr>
                        <th className="border border-black px-2 py-1 text-left">
                          Sanction Parameter
                        </th>
                        <th className="border border-black px-2 py-1 text-left">
                          Sanctioned / Cross-Verified Particular
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Name of the Client &amp; A/C No.
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          {ic.branchProposal?.section10Particulars?.nameOfClient?.trim() || (
                            <PhVal ph={ph} code="#2#" />
                          )}{' '}
                          | A/C: <PhVal ph={ph} code="#2#" /> (CIS: <PhVal ph={ph} code="#69#" />)
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Branch Name &amp; Nature of Business
                        </td>
                        <td className="border border-black px-2 py-0.5">
                          <PhVal ph={ph} code="#3#" bold={false} /> Branch |{' '}
                          <PhVal ph={ph} code="#54#" bold={false} />
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Mode / Product &amp; Original Limit
                        </td>
                        <td className="border border-black px-2 py-0.5">
                          <PhVal ph={ph} code="#5#" bold={false} /> | Original Limit: Tk.{' '}
                          <PhVal ph={ph} code="#6#" bold={false} /> M (Sec. 2 Limit: Tk.{' '}
                          <PhVal ph={ph} code="#56#" bold={false} /> M; Validity:{' '}
                          <PhVal ph={ph} code="#58#" bold={false} />)
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          NET Outstanding as on <PhVal ph={ph} code="#68#" />
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Tk. <PhVal ph={ph} code="#25#" /> Million (Principal: Tk.{' '}
                          <PhVal ph={ph} code="#7#" /> M; Charges: Tk. <PhVal ph={ph} code="#8#" /> M;
                          Charged Profit: Tk. <PhVal ph={ph} code="#10#" /> M; Comp.: Tk.{' '}
                          <PhVal ph={ph} code="#11#" /> M; Unapplied Comp.: Tk.{' '}
                          <PhVal ph={ph} code="#12#" /> M)
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Overdue &amp; Cumulative Recovery
                        </td>
                        <td className="border border-black px-2 py-0.5">
                          Overdue: Tk. <PhVal ph={ph} code="#59#" bold={false} /> Million | Recovery
                          Till Date: Tk. <PhVal ph={ph} code="#60#" bold={false} /> Million
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Cost of Fund &amp; Permissible Ceiling
                        </td>
                        <td className="border border-black px-2 py-0.5">
                          COF Surplus/Deficit: Tk. <PhVal ph={ph} code="#31#" bold={false} /> M |
                          Breakup Total (#7#+#8#+#31#): Tk.{' '}
                          <PhVal ph={ph} code="#32#" bold={false} /> M | Max Waiver Considering COF:
                          Tk. <PhVal ph={ph} code="#44#" bold={false} /> M
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Sanctioned Down Payment
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Tk. <PhVal ph={ph} code="#35#" /> Million deposited in A/C:{' '}
                          <PhVal ph={ph} code="#36#" /> on <PhVal ph={ph} code="#68#" />
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Sanctioned Waiver Amount
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Tk. <PhVal ph={ph} code="#4#" /> Million (plus accrued/unapplied profit up
                          to execution date)
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Sanctioned Settlement / Exit Amount
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Tk. <PhVal ph={ph} code="#27#" /> Million (Lump Sum without charging
                          further profit)
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Sanctioned Settlement Period
                        </td>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          <PhVal ph={ph} code="#28#" />
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Verified Source of Repayment
                        </td>
                        <td className="border border-black px-2 py-0.5">
                          {ic.sourceOfRepayment || 'As stipulated in the approved Board Memorandum'}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 3. AI Cross-Verified Facility & Liability Confirmation */}
                <div className="space-y-1 pt-1">
                  <p className="font-bold underline text-[10pt]">
                    3. AI Cross-Verified Facility &amp; Liability Confirmation:
                  </p>
                  <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                    {buildFallbackDisplay('sanctionCrossVerifiedSummary') || '—'}
                  </p>
                </div>

                {/* 4. Collateral Security, Legal & Regulatory Clearance Schedule */}
                <div className="space-y-1 pt-1">
                  <p className="font-bold underline text-[10pt]">
                    4. Collateral Security, Legal &amp; Regulatory Clearance Schedule:
                  </p>
                  <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                    {buildFallbackDisplay('sanctionSecurityAndCompliance') || '—'}
                  </p>
                </div>

                {/* 5. Operative Terms, Conditions & Redemption Covenants */}
                <div className="space-y-1 pt-1">
                  <p className="font-bold underline text-[10pt]">
                    5. Operative Terms, Conditions &amp; Redemption Covenants:
                  </p>
                  <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                    {buildFallbackDisplay('sanctionSpecialConditions') || '—'}
                  </p>
                  <div className="pl-4 space-y-1 text-[10pt] pt-1">
                    <p>
                      <span className="font-semibold">vii.</span> The waiver of profit/compensation
                      shall be adjusted from FAD, Head Office from the provision amount kept (if
                      any), and the remaining amount shall be adjusted gradually by the Branch upon
                      full realization of the sanctioned settlement amount within the approved
                      deadline.
                    </p>
                    <p>
                      <span className="font-semibold">viii.</span> Branch shall strictly comply with
                      Bangladesh Bank BRPD Circular No. 06/22, Circular Letter No. 18/22, 46/22, BRPD
                      Circular Letter No. 23 dated 29.06.2026 and all applicable Head Office
                      instructions.
                    </p>
                    <p className="font-bold">
                      ix. The Profit rate of the stipulated investment shall NOT be changed or reduced
                      to 0.00% until full and final adjustment of the Investment, as approved by Head
                      Office, and all legal/suit withdrawal charges shall be borne by the client.
                    </p>
                  </div>
                </div>
              </>
            ) : (
              <>
                <table className="w-[78%] table-fixed text-[10pt] border border-black">
                  <colgroup>
                    <col className="w-[44%]" />
                    <col className="w-[5%]" />
                    <col className="w-[51%]" />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">Dated</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'mec'
                          ? stageInputs.mecMemoDate || stageInputs.icMemoDate || ph['#1#'] || '—'
                          : stageInputs.boardMemoDate || stageInputs.mecMemoDate || ph['#1#'] || '—'}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'mec'
                          ? 'Chronologic Order of the MEC Memo'
                          : 'Chronologic Order of the Board Memo'}
                      </td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'mec'
                          ? stageInputs.mecMemoNumber
                            ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
                            : '—'
                          : stageInputs.boardMemoNumber
                            ? formatChronologicOrderWithYear(stageInputs.boardMemoNumber, '')
                            : '—'}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'mec'
                          ? 'Chronologic Order of the IC Memo & Date'
                          : 'Chronologic Order of the MEC Memo & Date'}
                      </td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'mec'
                          ? `${
                              stageInputs.icMemoNumber
                                ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
                                : '—'
                            } dated ${stageInputs.icMemoDate || '—'}`
                          : `${
                              stageInputs.mecMemoNumber
                                ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
                                : '—'
                            } dated ${stageInputs.mecMemoDate || '—'}`}
                      </td>
                    </tr>
                    {docType === 'board' && (
                      <tr>
                        <td className="border border-black px-2 py-0.5 font-bold">
                          Chronologic Order of the IC Memo &amp; Date
                        </td>
                        <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                        <td className="border border-black px-2 py-0.5">
                          {`${
                            stageInputs.icMemoNumber
                              ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
                              : '—'
                          } dated ${stageInputs.icMemoDate || '—'}`}
                        </td>
                      </tr>
                    )}
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">To</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">{meta.toLine}</td>
                    </tr>
                    <tr>
                      <td className="border border-black px-2 py-0.5 font-bold">From</td>
                      <td className="border border-black px-1 py-0.5 font-bold text-center">:</td>
                      <td className="border border-black px-2 py-0.5 font-bold">
                        {docType === 'board'
                          ? 'Management Executive Committee (MEC) / SAMD'
                          : 'SAMD'}
                      </td>
                    </tr>
                  </tbody>
                </table>

                {/* A/C & Branch + Subject */}
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[58%]" />
                    <col className="w-[42%]" />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
                        A/C: <PhVal ph={ph} code="#2#" /> (CIS: <PhVal ph={ph} code="#69#" />)
                      </td>
                      <td className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
                        Branch: <PhVal ph={ph} code="#3#" />
                      </td>
                    </tr>
                    <tr>
                      <td
                        colSpan={2}
                        className="border border-dotted border-black px-2 py-1 font-bold leading-snug"
                      >
                        Subject: The customer proposed for full and final settlement of their
                        liability by allowing waiver of Tk.
                        <PhVal ph={ph} code="#4#" /> along with profit/compensation to be accrued
                        up to the date of full and final settlement as follows:
                      </td>
                    </tr>
                  </tbody>
                </table>

                <p className="text-right text-[10pt] font-bold">
                  [All Figures in this Memo are in Million BDT Unless Stated Otherwise]
                </p>

                {/* Carried-Forward IC Memo Tables & Sections with Section 16 (IC) and Section 17 (MEC) in exact numbered order */}
                <CarriedICMemoBody
                  memo={ic}
                  facilityStartNumber={docType === 'mec' ? 17 : 18}
                  afterSamdSlot={
                    docType === 'mec' ? (
                      <div className="space-y-1 text-[10pt]">
                        <p className="font-bold underline">
                          16. Recommendation of the Investment Committee (IC):
                        </p>
                        <div className="pl-2 leading-snug">
                          {renderHighlightedCommitteeRec(buildFallbackDisplay('mecIcRecommendation'))}
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-2.5 text-[10pt]">
                        {(stageMemo.stageSections.mecIcRecommendation ||
                          stageInputs.icMemoNumber ||
                          stageInputs.icMemoDate) && (
                          <div className="space-y-1">
                            <p className="font-bold underline">
                              16. Recommendation of the Investment Committee (IC):
                            </p>
                            <div className="pl-2 leading-snug">
                              {renderHighlightedCommitteeRec(
                                stageMemo.stageSections.mecIcRecommendation ||
                                  buildMecIcRecommendationText(
                                    stageInputs.icMemoNumber,
                                    stageInputs.icMemoDate
                                  )
                              )}
                            </div>
                          </div>
                        )}
                        <div className="space-y-1">
                          <p className="font-bold underline">
                            17. Management Executive Committee (MEC) Recommendation:
                          </p>
                          <div className="pl-2 leading-snug">
                            {renderHighlightedCommitteeRec(
                              buildFallbackDisplay('boardMecRecommendation')
                            )}
                          </div>
                        </div>
                      </div>
                    )
                  }
                />

                {/* Additional Stage Appraisal / SAMD / Board Resolution Sections */}
                {docType === 'mec' ? (
                  <div className="space-y-2.5 pt-2">
                    <div className="space-y-1">
                      <p className="font-bold underline text-[10pt]">
                        19. {STAGE_SECTION_LABELS.mecExecutiveSummary}:
                      </p>
                      <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                        {buildFallbackDisplay('mecExecutiveSummary') || '—'}
                      </p>
                    </div>
                    <div className="space-y-1">
                      <p className="font-bold underline text-[10pt]">
                        20. {STAGE_SECTION_LABELS.mecSamdRecommendation}:
                      </p>
                      <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                        {buildFallbackDisplay('mecSamdRecommendation') || '—'}
                      </p>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-2.5 pt-2">
                    {stageMemo.stageSections.mecExecutiveSummary && (
                      <div className="space-y-1">
                        <p className="font-bold underline text-[10pt]">
                          Executive Appraisal &amp; Waiver Rationale (Carried Forward from MEC
                          Memo):
                        </p>
                        <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                          {stageMemo.stageSections.mecExecutiveSummary}
                        </p>
                      </div>
                    )}
                    {stageMemo.stageSections.mecSamdRecommendation && (
                      <div className="space-y-1">
                        <p className="font-bold underline text-[10pt]">
                          Observation and Recommendation of SAMD for MEC (Carried Forward):
                        </p>
                        <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                          {stageMemo.stageSections.mecSamdRecommendation}
                        </p>
                      </div>
                    )}
                    <div className="space-y-1">
                      <p className="font-bold underline text-[10pt]">
                        20. {STAGE_SECTION_LABELS.boardProposalRationale}:
                      </p>
                      <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                        {buildFallbackDisplay('boardProposalRationale') || '—'}
                      </p>
                    </div>
                    <div className="space-y-1">
                      <p className="font-bold underline text-[10pt]">
                        21. {STAGE_SECTION_LABELS.boardResolutionTerms}:
                      </p>
                      <p className="text-[10pt] whitespace-pre-line leading-snug pl-2">
                        {buildFallbackDisplay('boardResolutionTerms') || '—'}
                      </p>
                    </div>
                  </div>
                )}
              </>
            )}

            {/* Closing Placement Box & Signatories */}
            <div className="pt-4 space-y-6">
              <div className="border border-black p-2.5 text-center font-bold text-[10pt]">
                {docType === 'mec'
                  ? 'The proposal is hereby placed before the Management Executive Committee (MEC) for kind recommendation/approval.'
                  : docType === 'board'
                    ? 'Submitted before the Honourable Board of Directors for kind consideration and approval.'
                    : 'Issued under Head Office Sanction Authority. Strict compliance with all sanctioned terms, timelines, and BRPD directives is mandatory.'}
              </div>

              <div className="grid grid-cols-4 gap-3 pt-8 text-center text-[9pt]">
                {meta.signatories.map((sig) => (
                  <div key={sig} className="border-t border-black pt-1 font-bold">
                    {sig}
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
