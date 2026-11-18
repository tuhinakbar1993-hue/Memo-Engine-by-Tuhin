import React from 'react';
import { Clock, Download, FileText } from 'lucide-react';
import type { MemoEngineDefinition } from '../types/memo';

interface ComingSoonMemoTabProps {
  engine: MemoEngineDefinition;
  accountNo?: string;
  branchName?: string;
  proposalDate?: string;
}

export const ComingSoonMemoTab: React.FC<ComingSoonMemoTabProps> = ({
  engine,
  accountNo,
  branchName,
  proposalDate,
}) => {
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

        <button
          type="button"
          disabled
          title="Rules coming soon"
          className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-400 bg-slate-100 border border-slate-200 rounded-lg cursor-not-allowed whitespace-nowrap shrink-0"
        >
          <Download className="w-4 h-4" />
          <span>Download .docx</span>
        </button>
      </div>

      {/* Live Preview Placeholder Sheet */}
      <div className="bg-white border border-slate-200 rounded-lg p-12 text-center space-y-5">
        <div className="w-12 h-12 mx-auto rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-600">
          <Clock className="w-6 h-6" />
        </div>

        <div className="space-y-2 max-w-md mx-auto">
          <h3 className="font-display text-xl font-bold text-slate-900">Rules coming soon</h3>
          <p className="text-sm text-slate-600 leading-relaxed">
            The mapping rules, template layout, and .docx generator for{' '}
            <span className="font-semibold text-slate-900">{engine.name}</span> will be wired into
            this modular tab as soon as the specification is supplied.
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
    </div>
  );
};
