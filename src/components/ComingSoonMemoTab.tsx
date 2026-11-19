import React, { useMemo } from 'react';
import { CheckCircle2, Clock, Download, FileText, Loader2, RefreshCw } from 'lucide-react';
import { renderTemplatePreviewBlocks } from '../exporters/docxTemplateEngine';
import type {
  MemoDocType,
  MemoEngineDefinition,
  MemoTemplateRecord,
  ResolvedICMemo,
} from '../types/memo';
import { TemplateManagerPanel } from './TemplateManagerPanel';

interface ComingSoonMemoTabProps {
  engine: MemoEngineDefinition;
  accountNo?: string;
  branchName?: string;
  proposalDate?: string;
  memo: ResolvedICMemo;
  templates: Record<MemoDocType, MemoTemplateRecord | null>;
  onUploadTemplate: (docType: MemoDocType, file: File) => Promise<void>;
  onRemoveTemplate: (docType: MemoDocType) => Promise<void>;
  onDownloadDocx: (docType: MemoDocType) => Promise<void>;
  downloading: boolean;
  onApplyAllChanges: () => void;
  hasPendingChanges: boolean;
  lastAppliedAt: string | null;
}

export const ComingSoonMemoTab: React.FC<ComingSoonMemoTabProps> = ({
  engine,
  accountNo,
  branchName,
  proposalDate,
  memo,
  templates,
  onUploadTemplate,
  onRemoveTemplate,
  onDownloadDocx,
  downloading,
  onApplyAllChanges,
  hasPendingChanges,
  lastAppliedAt,
}) => {
  const activeTemplate = templates[engine.id];

  const previewBlocks = useMemo(() => {
    if (!activeTemplate) return [];
    return renderTemplatePreviewBlocks(activeTemplate, memo);
  }, [activeTemplate, memo]);

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      {/* Toolbar */}
      <div className="bg-white border border-slate-200 rounded-lg p-4 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <FileText className="w-5 h-5 text-slate-500" />
          <div>
            <h2 className="font-display text-lg font-semibold text-slate-900">{engine.name}</h2>
            <p className="text-xs text-slate-500">{engine.description}</p>
          </div>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            type="button"
            onClick={onApplyAllChanges}
            className={`inline-flex items-center gap-2 px-3.5 py-2 text-xs font-semibold rounded-lg border transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
              hasPendingChanges
                ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
            }`}
          >
            <RefreshCw className="w-3.5 h-3.5" />
            <span>Apply All Changes</span>
          </button>

          {activeTemplate ? (
            <button
              type="button"
              onClick={() => void onDownloadDocx(engine.id)}
              disabled={downloading}
              className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
            >
              {downloading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Download className="w-4 h-4" />
              )}
              <span>{downloading ? 'Building...' : 'Download .docx'}</span>
            </button>
          ) : (
            <button
              type="button"
              disabled
              title="Rules coming soon — or upload a default .docx format below"
              className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-400 bg-slate-100 border border-slate-200 rounded-lg cursor-not-allowed whitespace-nowrap shrink-0"
            >
              <Download className="w-4 h-4" />
              <span>Download .docx</span>
            </button>
          )}
        </div>
      </div>

      {/* Default Format / Template Upload Control */}
      <TemplateManagerPanel
        templates={templates}
        onUploadTemplate={onUploadTemplate}
        onRemoveTemplate={onRemoveTemplate}
        singleDocType={engine.id}
      />

      {activeTemplate ? (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h3 className="font-display text-sm font-bold text-slate-900">
                Populated Template Content Preview — {activeTemplate.fileName}
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Preview of populated text &amp; tables. Exported .docx strictly preserves original Word page margins, orientation, fonts, borders, headers, footers, and page breaks.
              </p>
            </div>
            {lastAppliedAt && !hasPendingChanges && (
              <span className="inline-flex items-center gap-1 text-xs font-medium text-emerald-700">
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>Synced at {lastAppliedAt}</span>
              </span>
            )}
          </div>

          <div className="p-6 sm:p-8 bg-slate-100/70">
            <div className="ic-memo-sheet max-w-[794px] mx-auto bg-white border border-slate-300 shadow-sm px-10 py-8 space-y-3 text-black">
              {previewBlocks.length === 0 ? (
                <p className="text-xs text-slate-500 font-sans">
                  Template loaded cleanly. Click &ldquo;Download .docx&rdquo; to generate the formatted output document.
                </p>
              ) : (
                previewBlocks.map((block, idx) =>
                  block.type === 'paragraph' ? (
                    <p key={idx} className="whitespace-pre-line text-[10pt] leading-snug">
                      {block.text}
                    </p>
                  ) : (
                    <table key={idx} className="w-full text-[9.5pt] border border-dotted border-black">
                      <tbody>
                        {(block.rows || []).map((row, rIdx) => (
                          <tr key={rIdx}>
                            {row.map((cellVal, cIdx) => (
                              <td
                                key={cIdx}
                                className="border border-dotted border-black px-2 py-1 whitespace-pre-line"
                              >
                                {cellVal}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )
                )
              )}
            </div>
          </div>
        </div>
      ) : (
        /* Live Preview Placeholder Sheet (Unchanged existing fallback when no template uploaded) */
        <div className="bg-white border border-slate-200 rounded-lg p-12 text-center space-y-5">
          <div className="w-12 h-12 mx-auto rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-600">
            <Clock className="w-6 h-6" />
          </div>

          <div className="space-y-2 max-w-md mx-auto">
            <h3 className="font-display text-xl font-bold text-slate-900">Rules coming soon</h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              The mapping rules, template layout, and .docx generator for{' '}
              <span className="font-semibold text-slate-900">{engine.name}</span> will be wired into
              this modular tab as soon as the specification is supplied — or upload a default{' '}
              <span className="font-semibold text-slate-900">.docx</span> template above to generate
              formatted outputs right away.
            </p>
          </div>

          <div className="max-w-lg mx-auto pt-4 border-t border-slate-200 grid grid-cols-3 gap-4 text-left text-xs">
            <div>
              <p className="text-slate-500">Account No.</p>
              <p className="font-mono-tabular font-medium text-slate-900 mt-0.5">
                {accountNo || 'Pending upload'}
              </p>
            </div>
            <div>
              <p className="text-slate-500">Branch</p>
              <p className="font-medium text-slate-900 mt-0.5">{branchName || 'Pending upload'}</p>
            </div>
            <div>
              <p className="text-slate-500">Proposal Date</p>
              <p className="font-mono-tabular font-medium text-slate-900 mt-0.5">
                {proposalDate || 'Pending input'}
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
