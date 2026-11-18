import React, { useState } from 'react';
import { Download, Loader2, Sparkles } from 'lucide-react';
import type { AiSectionsData, ResolvedICMemo } from '../types/memo';
import { ValidationPanel } from './ValidationPanel';

interface ICMemoPreviewProps {
  memo: ResolvedICMemo;
  aiSections: AiSectionsData;
  overrides: Record<string, string>;
  onOverrideChange: (code: string, value: string) => void;
  onGenerateAiSection: (section: 'justification' | 'branchVisit' | 'samdObservation') => Promise<void>;
  onUpdateAiSection: (section: 'justification' | 'branchVisit' | 'samdObservation', value: string | null) => void;
  generatingSection: 'justification' | 'branchVisit' | 'samdObservation' | null;
  aiError: string | null;
  onDownloadDocx: () => Promise<void>;
  downloading: boolean;
}

const Token: React.FC<{ ph: Record<string, string>; code: string; bold?: boolean }> = ({
  ph,
  code,
  bold = true,
}) => {
  const val = ph[code];
  const isMissing = !val || val.trim() === '';
  if (isMissing) {
    return (
      <span className="bg-amber-100 text-amber-900 px-1 rounded font-mono text-[9pt] font-semibold">
        {code}
      </span>
    );
  }
  return <span className={bold ? 'font-bold' : ''}>{val}</span>;
};

export const ICMemoPreview: React.FC<ICMemoPreviewProps> = ({
  memo,
  aiSections,
  overrides,
  onOverrideChange,
  onGenerateAiSection,
  onUpdateAiSection,
  generatingSection,
  aiError,
  onDownloadDocx,
  downloading,
}) => {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;

  const [editingJustification, setEditingJustification] = useState(false);
  const [editingBranchVisit, setEditingBranchVisit] = useState(false);
  const [editingSamd, setEditingSamd] = useState(false);

  const defaultCibStatuses = [
    'STD',
    'SMA',
    'SS',
    'DF',
    'BL',
    'Willful Default (WD)',
    'Write-Off',
    'Non-Funded',
    'Stay Order',
    'TOTAL',
  ];

  const sec3Rows =
    bp?.section3Rows && bp.section3Rows.length > 0
      ? bp.section3Rows
      : [
          { sl: 'a.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'b.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'c.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        ];

  const sec4Rows =
    bp?.section4Rows && bp.section4Rows.length > 0
      ? bp.section4Rows
      : defaultCibStatuses.map((s) => ({
          status: s,
          unitBorrower: '',
          unitGuarantor: '',
          otherBorrower: '',
          otherGuarantor: '',
          remarks: '',
        }));

  const sec5Rows =
    bp?.section5Rows && bp.section5Rows.length > 0
      ? bp.section5Rows
      : [{ sl: '', typeOfCase: '', caseNo: '', filingDate: '', amount: '', previousDate: '', nextDate: '', presentStatus: '' }];

  const defaultSec6 = [
    { sl: 'a.', authority: 'IC & CD', lastInspectionDate: '', objectionStatus: '', objections: '', branchReply: '' },
    { sl: 'b.', authority: 'DBI, Bangladesh Bank', lastInspectionDate: '', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'c.', authority: 'DOS, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'd.', authority: 'BFIU, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'e', authority: 'BRD, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
  ];
  const sec6Rows = bp?.section6Rows && bp.section6Rows.length > 0 ? bp.section6Rows : defaultSec6;

  const sec7Primary =
    bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0
      ? bp.section7PrimaryRows
      : [
          { particulars: 'i) Hypothecation of Stock as on ………………..', value: '' },
          { particulars: 'ii) Hypothecation of Machinery', value: '' },
          { particulars: 'iii) Assignment of Bills against Work Order', value: '' },
          { particulars: 'iv) Margin held against existing LC & BG liability as on……..', value: '' },
        ];

  const sec7Schedules =
    bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0
      ? bp.section7CollateralSchedules
      : [
          'Schedule-01: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-02: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-03: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-04: Full Schedule with owner name as per Mortgage Deed.',
        ];

  const sec7ValRows =
    bp?.section7ValuationRows && bp.section7ValuationRows.length > 0
      ? bp.section7ValuationRows
      : ['Schedule-01', 'Schedule-02', 'Schedule-03', 'Schedule-04', 'Building', 'Other', 'Grand Total'].map((s) => ({
          schedule: s,
          origBranchMv: '',
          origBranchFsv: '',
          origSurvMv: '',
          origSurvFsv: '',
          presBranchMv: '',
          presBranchFsv: '',
          presSurvMv: '',
          presSurvFsv: '',
          presAvMv: '',
          presAvFsv: '',
        }));

  const sec9Rows =
    bp?.section9Rows && bp.section9Rows.length > 0
      ? bp.section9Rows
      : ['a.', 'b.', 'c.', 'd.', 'e.', 'f.', 'g.'].map((sl) => ({
          sl,
          referenceNo: '',
          date: '',
          limit: '',
          remarks: '',
        }));

  const sp = bp?.section10Particulars;
  const sec10Items: Array<[string, string, string]> = [
    ['a.', 'Name of the Client', sp?.nameOfClient || ''],
    ['b.', 'Unique Customer ID', sp?.uniqueCustomerId || ''],
    ['c.', 'AWCD/MSD Account with opening date', sp?.awcdMsdAccount || ''],
    ['d.', 'Business Type/Nature', sp?.businessTypeNature || ''],
    ['e.', 'Business Management Type', sp?.businessManagementType || ''],
    ['f.', 'Client Type', sp?.clientType || ''],
    ['g.', 'Key Person with Phone No.', sp?.keyPersonPhone || ''],
    ['', 'Address', sp?.address || ''],
    ['h.', 'Office/Business Address', sp?.officeBusinessAddress || ''],
    ['', 'Factory Address', sp?.factoryAddress || ''],
    ['i.', 'Purpose of Investment', sp?.purposeOfInvestment || ''],
    ['j.', 'Limit of Investment', sp?.limitOfInvestment || ''],
    ['k.', 'Mode of Investment', sp?.modeOfInvestment || ''],
    ['l.', 'Last Visit Date', sp?.lastVisitDate || ''],
    ['m.', 'Stock Position', sp?.stockPosition || ''],
    ['n.', 'Fixed Asset', sp?.fixedAsset || ''],
    ['o.', 'Business Net Worth', sp?.businessNetWorth || ''],
    ['p.', 'Insurance Coverage', sp?.insuranceCoverage || ''],
    ['q.', 'Trade License & Validity', sp?.tradeLicenseValidity || ''],
    ['r.', 'Other Regulatory Documents', sp?.otherRegulatoryDocuments || ''],
    ['s.', 'Date of Original Investment', sp?.dateOfOriginalInvestment || ''],
    ['t.', '1st Disbursement Manager', sp?.firstDisbursementManager || ''],
    ['u.', 'First Date of Classification with Status', sp?.firstDateOfClassification || ''],
  ];

  const dirRows =
    bp?.section10Directors && bp.section10Directors.length > 0
      ? bp.section10Directors
      : [{ name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' }];

  const bankDirRows =
    bp?.section10BankDirectors && bp.section10BankDirectors.length > 0
      ? bp.section10BankDirectors
      : [
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
        ];

  const finRows =
    bp?.section11Rows && bp.section11Rows.length > 0
      ? bp.section11Rows
      : [{ year: '', grossSales: bp?.section11FinancialsNote || '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' }];

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      {/* Top Preview Controls */}
      <div className="bg-white border border-slate-200 rounded-lg p-4 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="font-display text-lg font-bold text-slate-900">IC Memo Live Preview</h2>
          <p className="text-xs text-slate-500">
            Matches the official Investment Committee .docx layout. Use inline Generate buttons for AI-assisted sections before downloading.
          </p>
        </div>

        <button
          type="button"
          onClick={onDownloadDocx}
          disabled={downloading}
          className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
        >
          {downloading ? (
            <Loader2 className="w-4 h-4 animate-spin" />
          ) : (
            <Download className="w-4 h-4" />
          )}
          <span>{downloading ? 'Building .docx...' : 'Download .docx'}</span>
        </button>
      </div>

      {aiError && (
        <div className="bg-red-50 border border-red-200 text-red-800 rounded-lg px-4 py-3 text-sm">
          {aiError}
        </div>
      )}

      {/* Validation Panel */}
      <ValidationPanel
        missingFields={memo.missingFields}
        overrides={overrides}
        onOverrideChange={onOverrideChange}
      />

      {/* Official Document Paper Sheet */}
      <div className="bg-slate-200/70 p-4 sm:p-8 rounded-lg overflow-x-auto">
        <div className="ic-memo-sheet max-w-[820px] mx-auto bg-white shadow-sm border border-slate-300 px-10 py-8 space-y-4">
          {/* Header: Bismillah + Logo + Bank Title + Exit with Waiver Box */}
          <div className="flex flex-col items-center">
            <img
              src="/assets/bismillah.jpg"
              alt="Bismillahir Rahmanir Rahim"
              referrerPolicy="no-referrer"
              className="h-6 object-contain mb-1"
            />
          </div>

          <div className="flex items-start justify-between gap-4">
            <img
              src="/assets/sib_logo.jpg"
              alt="Sammilito Islami Bank PLC"
              referrerPolicy="no-referrer"
              className="w-16 h-auto object-contain shrink-0"
            />
            <div className="text-center flex-1 leading-tight">
              <p className="text-[13pt] font-bold underline">SAMMILITO ISLAMI BANK PLC.</p>
              <p className="text-[11pt] font-bold underline">(State-owned Shariah-based Bank)</p>
              <p className="text-[10.5pt] font-bold underline">UNION BANK UNIT</p>
              <p className="text-[9.5pt] underline">
                Special Asset Management Division Head Office, Dhaka
              </p>
            </div>
            <div className="border border-black px-3 py-1.5 text-[10.5pt] font-bold whitespace-nowrap shrink-0">
              Exit with Waiver
            </div>
          </div>

          {/* Metadata Table */}
          <table className="w-[52%] text-[9.5pt] border border-black">
            <tbody>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold w-[28%]">Dated</td>
                <td className="border border-black px-2 py-0.5 font-bold text-center w-[8%]">:</td>
                <td className="border border-black px-2 py-0.5">
                  <Token ph={ph} code="#1#" />
                </td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">Memo no</td>
                <td className="border border-black px-2 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">To</td>
                <td className="border border-black px-2 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5 font-bold">Investment Committee</td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">From</td>
                <td className="border border-black px-2 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5 font-bold">SAMD</td>
              </tr>
            </tbody>
          </table>

          {/* A/C & Branch + Subject */}
          <table className="w-full text-[9.5pt] border border-dotted border-black">
            <tbody>
              <tr>
                <td className="border border-dotted border-black px-2 py-1 font-bold w-[60%]">
                  A/C: <Token ph={ph} code="#2#" /> (CIS: <Token ph={ph} code="#69#" />)
                </td>
                <td className="border border-dotted border-black px-2 py-1 font-bold w-[40%]">
                  Branch: <Token ph={ph} code="#3#" />
                </td>
              </tr>
              <tr>
                <td colSpan={2} className="border border-dotted border-black px-2 py-1.5 font-bold">
                  Subject: The customer proposed for full and final settlement of their liability by
                  allowing waiver of Tk.<Token ph={ph} code="#4#" /> along with profit/compensation to
                  be accrued up to the date of full and final settlement as follows:
                </td>
              </tr>
            </tbody>
          </table>

          <p className="text-right text-[9pt] font-bold">
            [All Figures in this Memo are in Million BDT Unless Stated Otherwise]
          </p>

          {/* Main Settlement Table */}
          <table className="w-full text-[9pt] border border-dotted border-black">
            <thead>
              <tr>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-1 text-center w-[12%]">
                  Mode
                </th>
                <th colSpan={2} className="border border-dotted border-black px-1.5 py-1 text-center w-[14%]">
                  Limit
                </th>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-1 text-center w-[21%]">
                  Particulars
                </th>
                <th className="border border-dotted border-black px-1.5 py-1 text-center w-[12%]">
                  O/S as on <Token ph={ph} code="#68#" />
                </th>
                <th className="border border-dotted border-black px-1.5 py-1 text-center w-[12%]">
                  Proposed Waiver
                </th>
                <th className="border border-dotted border-black px-1.5 py-1 text-center w-[12%]">
                  Settlement Amount
                </th>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-1 text-center w-[17%]">
                  Remarks
                </th>
              </tr>
              <tr>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">Original</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">RSD</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">01</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">02</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">03 (1-2)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-1 font-bold">
                  <Token ph={ph} code="#5#" />
                </td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-1 text-center font-bold">
                  <Token ph={ph} code="#6#" />
                </td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-1 text-center">
                  {bp?.limitRSD || '-'}
                </td>
                <td className="border border-dotted border-black px-1.5 py-1">Principal</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">
                  <Token ph={ph} code="#7#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-1 font-bold leading-snug">
                  <div>1)Settlement Deadline:</div>
                  <div className="mb-2">
                    <Token ph={ph} code="#28#" />
                  </div>
                  <div>2) Break up</div>
                  <div>
                    Principal: <Token ph={ph} code="#7#" />
                  </div>
                  <div>
                    Charges: <Token ph={ph} code="#8#" />
                  </div>
                  <div>
                    COF: <Token ph={ph} code="#31#" />
                  </div>
                  <div className="mb-2">
                    Total: <Token ph={ph} code="#32#" />
                  </div>
                  <div>3) Security Value:</div>
                  <div>
                    MV: <Token ph={ph} code="#33#" />
                  </div>
                  <div>
                    FSV: <Token ph={ph} code="#34#" />
                  </div>
                </td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">
                  Charges (Legal &amp; Others)
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">
                  <Token ph={ph} code="#8#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">Gestation Rent</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">Charged Profit</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">
                  <Token ph={ph} code="#10#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">Compensation</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">
                  <Token ph={ph} code="#11#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">
                  Uncharged Compensation
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">
                  <Token ph={ph} code="#12#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-1">
                  Uncharged Profit (for T/L)
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center">-</td>
              </tr>
              <tr>
                <td colSpan={4} className="border border-dotted border-black px-2 py-1 font-bold text-right">
                  Grand-Total
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center font-bold">
                  <Token ph={ph} code="#25#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center font-bold">
                  <Token ph={ph} code="#4#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1 text-center font-bold">
                  <Token ph={ph} code="#27#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-1"></td>
              </tr>
            </tbody>
          </table>

          {/* Note */}
          <div className="border border-dotted border-black px-2 py-1 text-[9pt] font-bold">
            Note: Information Regarding Down payment (deposited Tk.<Token ph={ph} code="#35#" />{' '}
            Million in A/C: <Token ph={ph} code="#36#" /> on <Token ph={ph} code="#68#" />.)
          </div>

          {/* Key Points */}
          <div className="space-y-1">
            <p className="font-bold underline text-[9.5pt]">Key Points:</p>
            <p className="text-[9pt]">
              This Memo is placed in conformity with{' '}
              <span className="font-bold">BRPD 1 Circular Letter No. 23 dated 29.06.2026</span> under
              the following contexts:
            </p>
            <table className="w-full text-[9pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th colSpan={3} className="border border-dotted border-black px-2 py-1 text-left w-[58%]">
                    Particulars
                  </th>
                  <th colSpan={3} className="border border-dotted border-black px-2 py-1 text-left w-[42%]">
                    Waiver recommended by:
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[5%]">a.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[40%]">
                    Total [Profit+ Compensation+ Uncharged Comp.&amp; Profit]
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[13%]">
                    : <Token ph={ph} code="#42#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[5%]">d.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[22%]">Client</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[15%]">
                    : <Token ph={ph} code="#45#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Cost of Fund</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#31#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Branch</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#45#" />
                  </td>
                </tr>
                <tr>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">c.</td>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">
                    Maximum Waiver Considering COF [a-b]
                  </td>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#44#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">SAMD, HO</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">g.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">HOIC</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">h.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">MEC</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">i.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">IC &amp; CD</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Justification of The Proposed Facility (AI Generate #51#) */}
          <div className="space-y-1 text-[9.5pt]">
            <div className="flex items-center justify-between gap-3">
              <p className="font-bold underline">Justification of The Proposed Facility:</p>
              <div className="flex items-center gap-2 font-sans">
                {aiSections.justification !== null && (
                  <button
                    type="button"
                    onClick={() => setEditingJustification((v) => !v)}
                    className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded border border-slate-300 cursor-pointer"
                  >
                    {editingJustification ? 'Done Editing' : 'Edit Text'}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => onGenerateAiSection('justification')}
                  disabled={generatingSection !== null}
                  className="inline-flex items-center gap-1.5 px-3 py-1 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded cursor-pointer disabled:opacity-60"
                >
                  {generatingSection === 'justification' ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Sparkles className="w-3.5 h-3.5" />
                  )}
                  <span>{aiSections.justification !== null ? 'Regenerate' : 'Generate'}</span>
                </button>
              </div>
            </div>

            <div className="pl-4 space-y-1">
              {editingJustification ? (
                <textarea
                  rows={4}
                  value={aiSections.justification ?? ''}
                  onChange={(e) => onUpdateAiSection('justification', e.target.value)}
                  className="w-full font-memo text-[9.5pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
                />
              ) : (
                <p className="whitespace-pre-line">
                  a. <Token ph={ph} code="#51#" bold={false} />
                </p>
              )}
              <p>b. Investment was not created against any fraud and forgery.</p>
              <p>c. The client is not a willful defaulter.</p>
            </div>
          </div>

          <p className="font-bold text-[9.5pt]">
            Previous Board Decision(s) Regarding Rescheduling/Waiver/Written-Off: Nil
          </p>

          {/* Section 1 */}
          <div className="border border-dotted border-black px-2 py-1 font-bold text-[9.5pt]">
            1. Nature of Business: <Token ph={ph} code="#54#" />
          </div>

          {/* Section 2 */}
          <div className="space-y-1">
            <p className="font-bold text-[9.5pt]">
              2. Liability Position of the Clients and Its Sister Concerns with Our Bank as on{' '}
              <Token ph={ph} code="#68#" />:
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">Name of the Concern</th>
                  <th className="border border-dotted border-black px-1 py-1">Mode of Investment</th>
                  <th className="border border-dotted border-black px-1 py-1">Limit</th>
                  <th className="border border-dotted border-black px-1 py-1">Outstanding</th>
                  <th className="border border-dotted border-black px-1 py-1">Validity</th>
                  <th className="border border-dotted border-black px-1 py-1">Overdue</th>
                  <th className="border border-dotted border-black px-1 py-1">
                    Recovery in existing deals till date [Mode wise]
                  </th>
                  <th className="border border-dotted border-black px-1 py-1">Remarks</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td rowSpan={2} className="border border-dotted border-black px-1.5 py-1 font-bold">
                    <Token ph={ph} code="#2#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 font-bold">
                    <div>
                      <Token ph={ph} code="#55#" />
                    </div>
                    <div>Funded</div>
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#56#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#25#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#58#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#59#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#60#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1">
                    {bp?.sec2Remarks || ''}
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-1 font-bold">Non Funded</td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                </tr>
                <tr>
                  <td colSpan={2} className="border border-dotted border-black px-1.5 py-1 font-bold text-right">
                    Total
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#56#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#25#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#58#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#59#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1 text-center">
                    <Token ph={ph} code="#60#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-1"></td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Section 3 */}
          <div className="space-y-1">
            <p className="font-bold underline text-[9.5pt]">
              3. Liability of the Client &amp; Its Sister/Allied Concerns with Other Units of SIB:
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">SL</th>
                  <th className="border border-dotted border-black px-1 py-1">Name of Concern</th>
                  <th className="border border-dotted border-black px-1 py-1">Name of the Banks/NBFIs</th>
                  <th className="border border-dotted border-black px-1 py-1">Nature of Facility</th>
                  <th className="border border-dotted border-black px-1 py-1">Limit</th>
                  <th className="border border-dotted border-black px-1 py-1">Outstanding with profit/Interest</th>
                  <th className="border border-dotted border-black px-1 py-1">Overdue/EOL</th>
                  <th className="border border-dotted border-black px-1 py-1">Expiry Date</th>
                  <th className="border border-dotted border-black px-1 py-1">CL Status</th>
                </tr>
              </thead>
              <tbody>
                {sec3Rows.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.sl}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.nameOfConcern}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.bankOrNbfi}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.natureOfFacility}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.limit}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.outstanding}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.overdueEol}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.expiryDate}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.clStatus}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 4 */}
          <div className="space-y-1">
            <p className="font-bold text-[9.5pt]">
              <span className="underline">4. CIB Status:</span> Report dated{' '}
              <Token ph={ph} code="#62#" /> shows:
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th rowSpan={2} className="border border-dotted border-black px-1 py-1">STATUS</th>
                  <th colSpan={2} className="border border-dotted border-black px-1 py-1 text-center">
                    With respective unit
                  </th>
                  <th colSpan={2} className="border border-dotted border-black px-1 py-1 text-center">
                    With Other Bank
                  </th>
                  <th rowSpan={2} className="border border-dotted border-black px-1 py-1 text-center">
                    Remarks (Pls. give concern wise Breakup)
                  </th>
                </tr>
                <tr>
                  <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
                  <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
                  <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
                  <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
                </tr>
              </thead>
              <tbody>
                {sec4Rows.map((r, i) => (
                  <tr key={i} className={/TOTAL/i.test(r.status) ? 'font-bold' : ''}>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.status}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitBorrower}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitGuarantor}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherBorrower}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherGuarantor}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 5 */}
          <div className="space-y-1">
            <p className="font-bold text-[9.5pt]">
              <span className="underline">5. Legal Action Status as on Memo Date:</span>{' '}
              {bp?.section5StatusNote || ''}
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">Sl</th>
                  <th className="border border-dotted border-black px-1 py-1">Type of Case</th>
                  <th className="border border-dotted border-black px-1 py-1">Case No.</th>
                  <th className="border border-dotted border-black px-1 py-1">Filling date</th>
                  <th className="border border-dotted border-black px-1 py-1">Amount</th>
                  <th className="border border-dotted border-black px-1 py-1">Previous date</th>
                  <th className="border border-dotted border-black px-1 py-1">Next date</th>
                  <th className="border border-dotted border-black px-1 py-1">Present status</th>
                </tr>
              </thead>
              <tbody>
                {sec5Rows.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5 min-h-4">{r.sl}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.typeOfCase}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.caseNo}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.filingDate}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.amount}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.previousDate}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.nextDate}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presentStatus}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 6 */}
          <div className="space-y-1">
            <p className="font-bold text-[9.5pt]">
              <span className="underline">
                6. Audit, Compliance and Regulatory Authority Related Status:
              </span>{' '}
              {bp?.section6StatusNote || ''}
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">SL</th>
                  <th className="border border-dotted border-black px-1 py-1">Name of the Authority</th>
                  <th className="border border-dotted border-black px-1 py-1">Last Inspection / Report Date</th>
                  <th className="border border-dotted border-black px-1 py-1">Objection Status (Yes/No)</th>
                  <th className="border border-dotted border-black px-1 py-1">Objections (If yes)</th>
                  <th className="border border-dotted border-black px-1 py-1">Branch Reply/ Current Position</th>
                </tr>
              </thead>
              <tbody>
                {sec6Rows.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.sl}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.authority}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.lastInspectionDate}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.objectionStatus}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.objections}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.branchReply}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 7 */}
          <div className="space-y-1.5 text-[9pt]">
            <p className="font-bold underline text-[9.5pt]">7. Security:</p>
            <p className="font-bold pl-2">
              <span className="underline">a. Primary:</span> {bp?.section7PrimaryNote || ''}
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-2 py-1 text-left w-[75%]">Particulars</th>
                  <th className="border border-dotted border-black px-2 py-1 text-center w-[25%]">Value</th>
                </tr>
              </thead>
              <tbody>
                {sec7Primary.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-2 py-0.5">{r.particulars}</td>
                    <td className="border border-dotted border-black px-2 py-0.5 text-center">{r.value}</td>
                  </tr>
                ))}
                <tr className="font-bold">
                  <td className="border border-dotted border-black px-2 py-0.5 text-center">
                    Total value of primary security
                  </td>
                  <td className="border border-dotted border-black px-2 py-0.5 text-center">
                    {bp?.section7PrimaryTotal || ''}
                  </td>
                </tr>
              </tbody>
            </table>

            <p className="font-bold pl-2 pt-1">
              <span className="underline">b. Collateral:</span> {bp?.section7CollateralNote || ''}
            </p>
            <div className="pl-2 space-y-0.5 text-[8.5pt]">
              {sec7Schedules.map((s, i) => (
                <p key={i}>{s}</p>
              ))}
            </div>

            <table className="w-full text-[7.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th rowSpan={3} className="border border-dotted border-black px-1 py-1">Schedule</th>
                  <th colSpan={10} className="border border-dotted border-black px-1 py-1 text-center">
                    Value of the mortgaged properties as assessed
                  </th>
                </tr>
                <tr>
                  <th colSpan={4} className="border border-dotted border-black px-1 py-0.5 text-center">
                    At the time of allowing investment
                  </th>
                  <th colSpan={6} className="border border-dotted border-black px-1 py-0.5 text-center">
                    At present
                  </th>
                </tr>
                <tr>
                  <th className="border border-dotted border-black px-1 py-0.5">By Br MV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">By Br FSV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">Surv MV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">Surv FSV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">By Br MV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">By Br FSV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">Surv MV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">Surv FSV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">AV Cell MV</th>
                  <th className="border border-dotted border-black px-1 py-0.5">AV Cell FSV</th>
                </tr>
              </thead>
              <tbody>
                {sec7ValRows.map((r, i) => (
                  <tr key={i} className={/Grand/i.test(r.schedule) ? 'font-bold' : ''}>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.schedule}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.origBranchMv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.origBranchFsv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.origSurvMv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.origSurvFsv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presBranchMv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presBranchFsv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presSurvMv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presSurvFsv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presAvMv}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.presAvFsv}</td>
                  </tr>
                ))}
                <tr>
                  <td colSpan={11} className="border border-dotted border-black px-2 py-1 text-center font-bold">
                    Outstanding to Lowest Security value Coverage (%): {bp?.section7CoveragePct || '……………%'}
                  </td>
                </tr>
              </tbody>
            </table>

            <div className="pl-2 font-bold text-[8.5pt]">
              <p>c. Others:</p>
              <p className="pl-4">a. Usual Charge documents.</p>
              <p className="pl-4">b. Acceptance of Branch Sanction by the Client.</p>
            </div>
          </div>

          {/* Section 8 */}
          <p className="text-[9.5pt]">
            <span className="font-bold underline">8. Source of Repayment:</span>{' '}
            {memo.sourceOfRepayment || (
              <span className="text-slate-400 italic">[Enter in Manual Inputs]</span>
            )}
          </p>

          {/* Section 9 */}
          <div className="space-y-1">
            <p className="font-bold underline text-[9.5pt]">9. Sequence of Head Office Sanction(S):</p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1.5 py-1 w-[8%]">SL</th>
                  <th className="border border-dotted border-black px-1.5 py-1 w-[36%]">HO Sanction Reference No.</th>
                  <th className="border border-dotted border-black px-1.5 py-1 w-[16%]">Date</th>
                  <th className="border border-dotted border-black px-1.5 py-1 w-[14%]">Limit</th>
                  <th className="border border-dotted border-black px-1.5 py-1 w-[26%]">Remarks</th>
                </tr>
              </thead>
              <tbody>
                {sec9Rows.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.sl}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.referenceNo}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.date}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.limit}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 10 */}
          <div className="space-y-2">
            <p className="font-bold underline text-[9.5pt]">10. Brief Particulars of The Client:</p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <tbody>
                {sec10Items.map(([idx, label, val], i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1.5 py-0.5 w-[6%]">{idx}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 w-[36%]">{label}</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 text-center w-[4%]">:</td>
                    <td className="border border-dotted border-black px-1.5 py-0.5 w-[54%]">{val}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <p className="font-bold underline text-[9pt]">
              • Particular of the Proprietor/Partner/Directors:
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">Name of the Directors</th>
                  <th className="border border-dotted border-black px-1 py-1">Age</th>
                  <th className="border border-dotted border-black px-1 py-1">Education</th>
                  <th className="border border-dotted border-black px-1 py-1">Designation</th>
                  <th className="border border-dotted border-black px-1 py-1">Share (%)</th>
                  <th className="border border-dotted border-black px-1 py-1">Experience</th>
                  <th className="border border-dotted border-black px-1 py-1">Net worth</th>
                </tr>
              </thead>
              <tbody>
                {dirRows.map((d, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.age}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.education}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.sharePct}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.experience}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.netWorth}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <p className="font-bold underline text-[9pt]">
              • Remarks: Director of any Bank/ FI/ Insurance company:
            </p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">Name of the Director</th>
                  <th className="border border-dotted border-black px-1 py-1">Name of the bank/NBFIs</th>
                  <th className="border border-dotted border-black px-1 py-1">Designation/Position</th>
                  <th className="border border-dotted border-black px-1 py-1">Percentage of Shares</th>
                  <th className="border border-dotted border-black px-1 py-1">Number of Shares</th>
                </tr>
              </thead>
              <tbody>
                {bankDirRows.map((d, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.bankNbfi}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.sharePct}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{d.numShares}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 11 */}
          <div className="space-y-1">
            <p className="font-bold underline text-[9.5pt]">11. Financials of The Client:</p>
            <table className="w-full text-[8.5pt] border border-dotted border-black">
              <thead>
                <tr>
                  <th className="border border-dotted border-black px-1 py-1">Year</th>
                  <th className="border border-dotted border-black px-1 py-1">Gross Sales</th>
                  <th className="border border-dotted border-black px-1 py-1">Gross Profit</th>
                  <th className="border border-dotted border-black px-1 py-1">Net Profit</th>
                  <th className="border border-dotted border-black px-1 py-1">Total Asset</th>
                  <th className="border border-dotted border-black px-1 py-1">Total liabilities</th>
                  <th className="border border-dotted border-black px-1 py-1">Net Worth</th>
                </tr>
              </thead>
              <tbody>
                {finRows.map((r, i) => (
                  <tr key={i}>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.year}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.grossSales}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.grossProfit}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.netProfit}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.totalAsset}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.totalLiabilities}</td>
                    <td className="border border-dotted border-black px-1 py-0.5">{r.netWorth}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Section 12 (AI Generate: third person) */}
          <div className="space-y-1 text-[9.5pt]">
            <div className="flex items-center justify-between gap-3">
              <p className="font-bold underline">12. Branch Visit Report:</p>
              <div className="flex items-center gap-2 font-sans">
                {aiSections.branchVisit !== null && (
                  <button
                    type="button"
                    onClick={() => setEditingBranchVisit((v) => !v)}
                    className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded border border-slate-300 cursor-pointer"
                  >
                    {editingBranchVisit ? 'Done Editing' : 'Edit Text'}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => onGenerateAiSection('branchVisit')}
                  disabled={generatingSection !== null}
                  className="inline-flex items-center gap-1.5 px-3 py-1 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded cursor-pointer disabled:opacity-60"
                >
                  {generatingSection === 'branchVisit' ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Sparkles className="w-3.5 h-3.5" />
                  )}
                  <span>{aiSections.branchVisit !== null ? 'Regenerate' : 'Generate'}</span>
                </button>
              </div>
            </div>
            {editingBranchVisit ? (
              <textarea
                rows={3}
                value={aiSections.branchVisit ?? ''}
                onChange={(e) => onUpdateAiSection('branchVisit', e.target.value)}
                className="w-full font-memo text-[9.5pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.branchVisitReport ? (
              <p className="whitespace-pre-line">{memo.branchVisitReport}</p>
            ) : null}
          </div>

          {/* Section 13 */}
          <p className="text-[9.5pt]">
            <span className="font-bold underline">13. Recommendation of The Branch:</span> The Branch
            Recommended the following vide Proposal Ref: <Token ph={ph} code="#64#" /> dated{' '}
            <Token ph={ph} code="#65#" />
          </p>

          {/* Section 14 */}
          <p className="text-[9.5pt] font-bold">
            <span className="underline">
              14. Recommendation of Internal Control &amp; Compliance Division:
            </span>{' '}
            N/A
          </p>

          {/* Section 15 (AI Generate: 3 observation/recommendation lines) */}
          <div className="space-y-1 text-[9.5pt]">
            <div className="flex items-center justify-between gap-3">
              <p className="font-bold underline">15. Observation and Recommendation of SAMD:</p>
              <div className="flex items-center gap-2 font-sans">
                {aiSections.samdObservation !== null && (
                  <button
                    type="button"
                    onClick={() => setEditingSamd((v) => !v)}
                    className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded border border-slate-300 cursor-pointer"
                  >
                    {editingSamd ? 'Done Editing' : 'Edit Text'}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => onGenerateAiSection('samdObservation')}
                  disabled={generatingSection !== null}
                  className="inline-flex items-center gap-1.5 px-3 py-1 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded cursor-pointer disabled:opacity-60"
                >
                  {generatingSection === 'samdObservation' ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Sparkles className="w-3.5 h-3.5" />
                  )}
                  <span>{aiSections.samdObservation !== null ? 'Regenerate' : 'Generate'}</span>
                </button>
              </div>
            </div>
            {editingSamd ? (
              <textarea
                rows={4}
                value={aiSections.samdObservation ?? ''}
                onChange={(e) => onUpdateAiSection('samdObservation', e.target.value)}
                className="w-full font-memo text-[9.5pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.samdObservation ? (
              <div className="pl-4 whitespace-pre-line">{memo.samdObservation}</div>
            ) : null}
          </div>

          {/* Section 16 */}
          <div className="space-y-1">
            <p className="font-bold underline text-[9.5pt]">16. Particulars of the facility:</p>
            <table className="w-full text-[9pt] border border-dotted border-black">
              <tbody>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[5%]">a.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 w-[24%]">Facility</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    EXIT with Waiver under BRPD 1 Circular Letter No. 23 dated 29.06.2026
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Outstanding</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#25#" /> million as on <Token ph={ph} code="#68#" />.
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">c.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Down Payment</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Down-payment of Tk. <Token ph={ph} code="#35#" /> Million is kept in{' '}
                    <Token ph={ph} code="#36#" /> A/C <Token ph={ph} code="#36#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">d.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Waiver</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#4#" /> million as on <Token ph={ph} code="#68#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Settlement/Exit amount</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#27#" /> Million as on <Token ph={ph} code="#68#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Settlement period</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    <Token ph={ph} code="#28#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">g.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Profit/Rent rate</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    Without charging further profit.
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">h.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Repayment Method</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Lump sum.</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Section 17 */}
          <div className="space-y-1 text-[9pt]">
            <p className="font-bold underline text-[9.5pt]">17. Special Term(s) &amp; Condition(s):</p>
            <p className="pl-4">
              If the proposal approved, the branch is advised to comply the following special terms
              and conditions:
            </p>
            <ol className="pl-8 space-y-1 list-none">
              <li>
                i. The waiver of profit shall be adjusted from FAD, Head Office from the provision
                amount kept (if any) and the rest amount shall be adjusted by branch gradually after
                adjustment of the settlement amount within the approved time frame.
              </li>
              <li>
                ii. The final settlement amount will be determined at the date of execution of
                settlement considering up to date liability position.
              </li>
              <li>
                iii. The Branch must ensure adjustment of all other direct and indirect liabilities
                of the client before allowing the final redemption of the mortgaged collateral
                security(s).
              </li>
              <li>
                iv. The Branch must ensure adjustment of all non-funded liabilities (L/C, Bank
                Guarantee, Credit Card etc.) before allowing final redemption of the mortgaged
                collateral security(s). The Branch must also ensure that any encashable security
                (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client)
                including collateral is properly mentioned in the proposal. If any encashable
                security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the
                client) is found not mentioned in the proposal and later encashed in favor of the
                client, the Manager will be held responsible.
              </li>
              <li>
                v. If the client fails to repay or fails to adjust the liability within the given
                time or fails to pay installments regularly, all facilities shall be cancelled and
                legal action shall be initiated by the Branch.
              </li>
              <li>
                vi. Branch should be utmost sincere to protect bank interest in the course of
                execution partial redemption.
              </li>
              <li>
                vii. Branch is advised to follow the guidelines of Bangladesh Bank BRPD circular No.
                06/22, Circular Letter No.18/22, 46/22 and subsequent up-dates as well as all
                instruction circular(s) of SIBL.
              </li>
              <li className="font-bold">
                viii. The Profit rate of the stipulated investment shall not be changed or reduced to
                0.00% until full and final adjustment of the Investment, as approved by Head Office.
              </li>
              <li>
                ix. All the legal expenses and other charges (if any) already incurred and to be
                incurred in connection with the withdrawal of case/suits and others issues shall be
                borne by the client.
              </li>
            </ol>
          </div>

          {/* Closing Box & Signatories */}
          <div className="pt-4 space-y-6 text-[9.5pt]">
            <div className="border border-black px-3 py-1 text-center font-bold">
              The proposal is hereby placed before the Investment Committee for kind recommendation.
            </div>

            <div className="font-bold leading-snug">
              <p>Union Bank PLC Unit</p>
              <p>Special Asset Management Division</p>
            </div>

            <p className="text-center font-bold">Submitted for kind decision.</p>

            <div className="grid grid-cols-2 gap-8 pt-6">
              <div>
                <p className="font-bold">(Shahidur Rahman khan)</p>
                <p>Deputy Managing Director &amp; CRO</p>
                <p>Sammilito Islami Bank PLC.</p>
              </div>
              <div>
                <p className="font-bold">(Ahmed Zafrul Hasan)</p>
                <p>Additional Managing Director &amp; COO</p>
                <p>Sammilito Islami Bank PLC.</p>
              </div>
            </div>

            <div className="pt-6 text-center leading-snug">
              <p className="font-bold">(Abedur Rahman Sikder)</p>
              <p>Managing Director</p>
              <p>Sammilito Islami Bank PLC.</p>
              <p>(Transferee of the assets and liabilities of</p>
              <p>EXIM Bank, SIBL, FSIB, Global Islami</p>
              <p>Bank and Union Bank)</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
