import type {
  AiSectionsData,
  ExtractedDocumentsState,
  ManualInputsData,
  MissingFieldItem,
  ResolvedICMemo,
} from '../types/memo';
import { bdtToMillionStr, formatMillionNumber, parseIndianNumber } from '../utils/numberFormat';

export interface PlaceholderSpec {
  code: string;
  label: string;
  source: string;
}

export const IC_PLACEHOLDER_SPECS: PlaceholderSpec[] = [
  { code: '#1#', label: 'Proposal Date', source: 'Manual Inputs' },
  { code: '#2#', label: 'Account Number / Name of Concern', source: 'Account Statement (Account number)' },
  { code: '#3#', label: 'Branch Name', source: 'Account Statement (2nd line)' },
  { code: '#4#', label: 'Waiver Amount (Million BDT)', source: 'Manual Inputs' },
  { code: '#5#', label: 'Mode / Product Name', source: 'Party Performance (Product Name)' },
  { code: '#6#', label: 'Original Limit (Million BDT)', source: 'Branch Proposal (Main Table)' },
  { code: '#7#', label: 'Principal Outstanding (Million BDT)', source: 'Outstanding (Grand Total -> Princ Out)' },
  { code: '#8#', label: 'Charges (Million BDT)', source: 'Outstanding (Grand Total -> Charge)' },
  { code: '#10#', label: 'Charged Profit (Million BDT)', source: 'Outstanding (Grand Total -> Profit Out)' },
  { code: '#11#', label: 'Compensation (Million BDT)', source: 'Outstanding (Grand Total -> Compensation)' },
  { code: '#12#', label: 'Unapplied / Uncharged Comp. (Million BDT)', source: 'Outstanding (Grand Total -> Unapplied Pro)' },
  { code: '#25#', label: 'NET Outstanding (Million BDT)', source: 'Outstanding (Grand Total -> NET Out)' },
  { code: '#27#', label: 'Settlement Amount (Million BDT)', source: 'Manual Inputs' },
  { code: '#28#', label: 'Settlement Period / Deadline', source: 'Manual Inputs' },
  { code: '#31#', label: 'Cost of Fund Surplus/Deficit (Million BDT)', source: 'Cost of Fund (Excel)' },
  { code: '#32#', label: 'Breakup Total (#7# + #8# + #31#)', source: 'Calculated' },
  { code: '#33#', label: 'Security Value MV', source: 'Branch Proposal (Remarks -> MV)' },
  { code: '#34#', label: 'Security Value FSV', source: 'Branch Proposal (Remarks -> FSV)' },
  { code: '#35#', label: 'Down Payment / Last Balance (Million BDT)', source: 'Account Statement (Last Balance)' },
  { code: '#36#', label: 'Down Payment A/C No.', source: 'Account Statement (A/C No.)' },
  { code: '#42#', label: 'Total Profit+Comp+Unapplied (#10#+#8#+#11#+#12#)', source: 'Calculated' },
  { code: '#44#', label: 'Max Waiver Considering COF (#42# - #31#)', source: 'Calculated' },
  { code: '#45#', label: 'Waiver Recommended by Client & Branch', source: 'Manual Inputs' },
  { code: '#47#', label: 'Waiver Recommended by H/O', source: 'Manual Inputs' },
  { code: '#51#', label: 'Justification of Proposed Facility (3 lines)', source: 'AI Generate / Manual' },
  { code: '#54#', label: 'Nature of Business', source: 'Branch Proposal (Section 1)' },
  { code: '#55#', label: 'Sec. 2 Mode of Investment', source: 'Party Performance (Product Name)' },
  { code: '#56#', label: 'Sec. 2 Limit (Million BDT)', source: 'Branch Proposal (Section 2)' },
  { code: '#58#', label: 'Sec. 2 Validity', source: 'Branch Proposal (Section 2)' },
  { code: '#59#', label: 'Overdue Amount (Million BDT)', source: 'Overdue (Grand Total)' },
  { code: '#60#', label: 'Recovery Till Date (Million BDT)', source: 'Party Performance (Customer Total)' },
  { code: '#62#', label: 'CIB Report Date', source: 'CIB Report (Date of Inquiry)' },
  { code: '#64#', label: 'Branch Proposal Reference No.', source: 'Branch Proposal (Ref. No.)' },
  { code: '#65#', label: 'Branch Proposal Date', source: 'Branch Proposal (Proposal Date)' },
  { code: '#68#', label: 'O/S As On / Print Date', source: 'Party Performance (Print Date)' },
  { code: '#69#', label: 'Customer ID / CIS', source: 'Account Statement (Customer ID)' },
];

export function resolveICMemo(
  docs: ExtractedDocumentsState,
  manual: ManualInputsData,
  ai: AiSectionsData,
  overrides: Record<string, string>
): ResolvedICMemo {
  const bp = docs.branchProposal;
  const out = docs.outstanding;
  const pp = docs.partyPerformance;
  const ovd = docs.overdue;
  const stmt = docs.accountStatement;
  const cib = docs.cibReport;
  const cof = docs.costOfFund;

  const base: Record<string, string> = {
    // 5.1 Manual Inputs
    '#1#': manual.proposalDate.trim(),
    '#4#': manual.waiverAmount.trim(),
    '#27#': manual.settlementAmount.trim(),
    '#28#': manual.settlementPeriod.trim(),
    '#45#': manual.waiverRecClientBranch.trim(),
    '#47#': manual.waiverRecHO.trim(),

    // 5.2 Account Statement
    '#2#': stmt?.accountNumber?.trim() || '',
    '#36#': stmt?.acNo?.trim() || '',
    '#3#': stmt?.branchName?.trim() || '',
    '#69#': stmt?.customerId?.trim() || '',
    '#35#': stmt ? bdtToMillionStr(stmt.lastBalanceBDT) : '',

    // 5.3 Outstanding (Grand Total, BDT -> Million BDT)
    '#7#': out ? bdtToMillionStr(out.princOutBDT) : '',
    '#8#': out ? bdtToMillionStr(out.chargeBDT) : '',
    '#10#': out ? bdtToMillionStr(out.profitOutBDT) : '',
    '#11#': out ? bdtToMillionStr(out.compensationBDT) : '',
    '#12#': out ? bdtToMillionStr(out.unappliedProBDT) : '',
    '#25#': out ? bdtToMillionStr(out.netOutBDT) : '',

    // 5.4 Party Performance
    '#5#': pp?.productName?.trim() || '',
    '#55#': pp?.productName?.trim() || '',
    '#60#': pp ? bdtToMillionStr(pp.recoveryAmtBDT) : '',
    '#68#': pp?.printDate?.trim() || '',

    // 5.5 Overdue
    '#59#': ovd ? bdtToMillionStr(ovd.overdueAmountBDT) : '',

    // 5.6 CIB Report
    '#62#': cib?.cibDate?.trim() || '',

    // 5.7 Cost of Fund
    '#31#': cof ? bdtToMillionStr(cof.surplusDeficitBDT) : '',

    // 5.8 Branch Proposal (already in Million BDT — copy as written)
    '#6#': bp?.limitOriginal?.trim() || '',
    '#33#': bp?.securityMV?.trim() || '',
    '#34#': bp?.securityFSV?.trim() || '',
    '#54#': bp?.natureOfBusiness?.trim() || '',
    '#56#': bp?.sec2Limit?.trim() || '',
    '#58#': bp?.sec2Validity?.trim() || '',
    '#64#': bp?.referenceNo?.trim() || '',
    '#65#': bp?.proposalDate?.trim() || '',

    // 6. AI Generate (#51#)
    '#51#': ai.justification !== null ? ai.justification.trim() : '',
  };

  // Apply non-calculated manual overrides first so calculated fields react to user edits
  const resolved: Record<string, string> = { ...base };
  for (const [key, val] of Object.entries(overrides)) {
    if (key !== '#32#' && key !== '#42#' && key !== '#44#' && val !== undefined) {
      resolved[key] = val.trim();
    }
  }

  // 5.9 Calculated fields
  const n7 = parseIndianNumber(resolved['#7#']);
  const n8 = parseIndianNumber(resolved['#8#']);
  const n10 = parseIndianNumber(resolved['#10#']);
  const n11 = parseIndianNumber(resolved['#11#']);
  const n12 = parseIndianNumber(resolved['#12#']);
  const n31 = parseIndianNumber(resolved['#31#']);

  resolved['#32#'] =
    n7 !== null && n8 !== null && n31 !== null ? formatMillionNumber(n7 + n8 + n31) : '';

  resolved['#42#'] =
    n10 !== null && n8 !== null && n11 !== null && n12 !== null
      ? formatMillionNumber(n10 + n8 + n11 + n12)
      : '';

  const n42 = parseIndianNumber(resolved['#42#']);
  resolved['#44#'] =
    n42 !== null && n31 !== null ? formatMillionNumber(n42 - n31) : '';

  // Allow direct override on calculated fields if explicitly edited
  for (const calcKey of ['#32#', '#42#', '#44#']) {
    if (overrides[calcKey] !== undefined && overrides[calcKey].trim() !== '') {
      resolved[calcKey] = overrides[calcKey].trim();
    }
  }

  // Collect only unfilled placeholders for Validation Panel
  const missingFields: MissingFieldItem[] = [];
  for (const spec of IC_PLACEHOLDER_SPECS) {
    const val = resolved[spec.code];
    if (!val || val.trim() === '') {
      missingFields.push({
        code: spec.code,
        label: spec.label,
        source: spec.source,
        reason:
          spec.code === '#51#'
            ? 'Not generated yet (click "Generate" in IC Memo or enter manually).'
            : `Missing or unreadable from ${spec.source}.`,
      });
    }
  }

  return {
    placeholders: resolved,
    missingFields,
    sourceOfRepayment: manual.sourceOfRepayment.trim(),
    justificationLines: resolved['#51#'] || '',
    branchVisitReport: ai.branchVisit !== null ? ai.branchVisit.trim() : '',
    samdObservation: ai.samdObservation !== null ? ai.samdObservation.trim() : '',
    branchProposal: bp,
  };
}
