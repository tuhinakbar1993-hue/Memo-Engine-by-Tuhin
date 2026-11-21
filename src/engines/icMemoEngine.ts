import { createEmptyBranchProposal } from '../parsers/branchProposalParser';
import type {
  AiSectionsData,
  BranchProposalData,
  ExtractedDocumentsState,
  ManualInputsData,
  MissingFieldItem,
  ResolvedICMemo,
  ResolvedStageMemo,
  ScrutinySectionKey,
  StageAiSectionsData,
  StageMemoDocType,
  StageMemoInputs,
  StageScrutinySectionKey,
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
  overrides: Record<string, string>,
  manualEdited: Partial<Record<ScrutinySectionKey, boolean>> = {}
): ResolvedICMemo {
  const bp = docs.branchProposal;
  const out = docs.outstanding;
  const pp = docs.partyPerformance;
  const ovd = docs.overdue;
  const stmt = docs.accountStatement;
  const cib = docs.cibReport;
  const cof = docs.costOfFund;

  const hasStructuredAi = Boolean(
    ai.sec2LiabilityOurBank ||
      ai.sec3LiabilityOtherUnits ||
      ai.sec4CibStatus ||
      ai.sec5LegalAction ||
      ai.sec6AuditCompliance ||
      ai.sec7Security ||
      ai.sec9HoSanctions ||
      ai.sec10ClientParticulars ||
      ai.sec11Financials
  );

  let mergedBp: BranchProposalData | null = bp
    ? {
        ...bp,
        section10Particulars: { ...bp.section10Particulars },
      }
    : hasStructuredAi
      ? createEmptyBranchProposal()
      : null;

  if (mergedBp) {
    if (ai.sec2LiabilityOurBank) {
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.limit.trim()) {
        mergedBp.sec2Limit = ai.sec2LiabilityOurBank.limit.trim();
      }
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.validity.trim()) {
        mergedBp.sec2Validity = ai.sec2LiabilityOurBank.validity.trim();
      }
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.remarks.trim()) {
        mergedBp.sec2Remarks = ai.sec2LiabilityOurBank.remarks.trim();
      }
    }
    if (ai.sec3LiabilityOtherUnits) {
      mergedBp.section3Rows = ai.sec3LiabilityOtherUnits.rows;
    }
    if (ai.sec4CibStatus) {
      mergedBp.section4Rows = ai.sec4CibStatus.rows;
      if (manualEdited.sec4CibStatus || ai.sec4CibStatus.cibDate.trim()) {
        mergedBp.section4CibDate = ai.sec4CibStatus.cibDate.trim();
      }
    }
    if (ai.sec5LegalAction) {
      mergedBp.section5StatusNote = ai.sec5LegalAction.statusNote;
      mergedBp.section5Rows = ai.sec5LegalAction.rows;
    }
    if (ai.sec6AuditCompliance) {
      mergedBp.section6StatusNote = ai.sec6AuditCompliance.statusNote;
      mergedBp.section6Rows = ai.sec6AuditCompliance.rows;
    }
    if (ai.sec7Security) {
      mergedBp.section7PrimaryNote = ai.sec7Security.primaryNote;
      mergedBp.section7PrimaryRows = ai.sec7Security.primaryRows;
      mergedBp.section7PrimaryTotal = ai.sec7Security.primaryTotal;
      mergedBp.section7CollateralNote = ai.sec7Security.collateralNote;
      mergedBp.section7CollateralSchedules = ai.sec7Security.collateralSchedules;
      mergedBp.section7ValuationRows = ai.sec7Security.valuationRows;
      mergedBp.section7CoveragePct = ai.sec7Security.coveragePct;
      mergedBp.section7Others = ai.sec7Security.othersNote;
    }
    if (ai.sec9HoSanctions) {
      mergedBp.section9Rows = ai.sec9HoSanctions.rows;
    }
    if (ai.sec10ClientParticulars) {
      mergedBp.section10Particulars = { ...ai.sec10ClientParticulars.particulars };
      mergedBp.section10Directors = ai.sec10ClientParticulars.directors;
      mergedBp.section10BankDirectors = ai.sec10ClientParticulars.bankDirectors;
    }
    if (ai.sec11Financials) {
      mergedBp.section11FinancialsNote = ai.sec11Financials.financialsNote;
      mergedBp.section11Rows = ai.sec11Financials.rows;
    }
  }

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

    // 5.6 CIB Report (authoritative CIB Report first, fallback to Branch Proposal sec 4 or manual Sec 4 edit)
    '#62#':
      manualEdited.sec4CibStatus && ai.sec4CibStatus
        ? ai.sec4CibStatus.cibDate.trim() || cib?.cibDate?.trim() || ''
        : cib?.cibDate?.trim() || ai.sec4CibStatus?.cibDate?.trim() || bp?.section4CibDate?.trim() || '',

    // 5.7 Cost of Fund
    '#31#': cof ? bdtToMillionStr(cof.surplusDeficitBDT) : '',

    // 5.8 Branch Proposal (already in Million BDT — copy as written)
    '#6#': bp?.limitOriginal?.trim() || '',
    '#33#': bp?.securityMV?.trim() || '',
    '#34#': bp?.securityFSV?.trim() || '',
    '#54#':
      ai.sec1NatureOfBusiness !== null
        ? ai.sec1NatureOfBusiness.trim()
        : bp?.natureOfBusiness?.trim() || '',
    '#56#': mergedBp?.sec2Limit?.trim() || '',
    '#58#': mergedBp?.sec2Validity?.trim() || '',
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
            ? 'Not generated yet (click "Scrutinize and Fill Up Through AI" in IC Memo or enter manually).'
            : `Missing or unreadable from ${spec.source}.`,
      });
    }
  }

  const sectionSummaries: Partial<Record<ScrutinySectionKey, string>> = {};
  if (ai.sec2LiabilityOurBank?.summary?.trim()) {
    sectionSummaries.sec2LiabilityOurBank = ai.sec2LiabilityOurBank.summary.trim();
  }
  if (ai.sec3LiabilityOtherUnits?.summary?.trim()) {
    sectionSummaries.sec3LiabilityOtherUnits = ai.sec3LiabilityOtherUnits.summary.trim();
  }
  if (ai.sec4CibStatus?.summary?.trim()) {
    sectionSummaries.sec4CibStatus = ai.sec4CibStatus.summary.trim();
  }
  if (ai.sec5LegalAction?.summary?.trim()) {
    sectionSummaries.sec5LegalAction = ai.sec5LegalAction.summary.trim();
  }
  if (ai.sec6AuditCompliance?.summary?.trim()) {
    sectionSummaries.sec6AuditCompliance = ai.sec6AuditCompliance.summary.trim();
  }
  if (ai.sec7Security?.summary?.trim()) {
    sectionSummaries.sec7Security = ai.sec7Security.summary.trim();
  }
  if (ai.sec9HoSanctions?.summary?.trim()) {
    sectionSummaries.sec9HoSanctions = ai.sec9HoSanctions.summary.trim();
  }
  if (ai.sec10ClientParticulars?.summary?.trim()) {
    sectionSummaries.sec10ClientParticulars = ai.sec10ClientParticulars.summary.trim();
  }
  if (ai.sec11Financials?.summary?.trim()) {
    sectionSummaries.sec11Financials = ai.sec11Financials.summary.trim();
  }

  return {
    placeholders: resolved,
    missingFields,
    sourceOfRepayment:
      manual.sourceOfRepayment.trim() || (ai.sec8SourceOfRepayment ?? '').trim(),
    justificationLines: resolved['#51#'] || '',
    clientPrayer: ai.clientPrayer !== null ? ai.clientPrayer.trim() : '',
    branchVisitReport: ai.branchVisit !== null ? ai.branchVisit.trim() : '',
    samdObservation: ai.samdObservation !== null ? ai.samdObservation.trim() : '',
    sec13BranchRecommendation:
      ai.sec13BranchRecommendation !== null ? ai.sec13BranchRecommendation.trim() : '',
    sec14IccdRecommendation:
      ai.sec14IccdRecommendation !== null ? ai.sec14IccdRecommendation.trim() : '',
    sectionSummaries,
    sec2NonFundedNote: ai.sec2LiabilityOurBank?.nonFundedNote?.trim() || '',
    sec7OthersNote: ai.sec7Security?.othersNote?.trim() || '',
    branchProposal: mergedBp,
  };
}

export const INITIAL_STAGE_INPUTS: StageMemoInputs = {
  icMemoNumber: '',
  icMemoDate: '',
  mecMemoNumber: '',
  mecMemoDate: '',
  boardMemoNumber: '',
  boardMemoDate: '',
  sanctionRefNumber: '',
  sanctionDate: '',
};

export const INITIAL_STAGE_AI_SECTIONS: StageAiSectionsData = {
  mecIcRecommendation: null,
  mecExecutiveSummary: null,
  mecSamdRecommendation: null,
  boardMecRecommendation: null,
  boardProposalRationale: null,
  boardResolutionTerms: null,
  sanctionBoardAuthority: null,
  sanctionCrossVerifiedSummary: null,
  sanctionSecurityAndCompliance: null,
  sanctionSpecialConditions: null,
};

export const STAGE_SECTIONS_BY_DOC: Record<StageMemoDocType, StageScrutinySectionKey[]> = {
  mec: ['mecIcRecommendation', 'mecExecutiveSummary', 'mecSamdRecommendation'],
  board: ['boardMecRecommendation', 'boardProposalRationale', 'boardResolutionTerms'],
  sanction: [
    'sanctionBoardAuthority',
    'sanctionCrossVerifiedSummary',
    'sanctionSecurityAndCompliance',
    'sanctionSpecialConditions',
  ],
};

export const STAGE_SECTION_LABELS: Record<StageScrutinySectionKey, string> = {
  mecIcRecommendation: 'Recommendation of the Investment Committee (IC)',
  mecExecutiveSummary: 'Executive Appraisal & Waiver Rationale for MEC',
  mecSamdRecommendation: 'Observation and Recommendation of SAMD for MEC',
  boardMecRecommendation: 'Management Executive Committee (MEC) Recommendation',
  boardProposalRationale: 'Board Memorandum Summary & Financial Justification',
  boardResolutionTerms: 'Proposed Board Resolution & Approval Terms',
  sanctionBoardAuthority: 'Board Approval Reference & Sanction Basis',
  sanctionCrossVerifiedSummary: 'AI Cross-Verified Facility & Liability Confirmation',
  sanctionSecurityAndCompliance: 'Collateral Security, Legal & Regulatory Clearance Schedule',
  sanctionSpecialConditions: 'Operative Terms, Conditions & Redemption Covenants',
};

export function formatChronologicOrderWithYear(
  rawValue: string,
  fallbackPlaceholder: string
): string {
  const trimmed = (rawValue || '').trim();
  if (!trimmed) return `${fallbackPlaceholder}/26`;
  if (/\/26$/i.test(trimmed)) return trimmed;
  return `${trimmed}/26`;
}

export function buildMecIcRecommendationText(
  icMemoNumber: string,
  icMemoDate: string
): string {
  const orderWithYear = formatChronologicOrderWithYear(
    icMemoNumber,
    '[Chronologic Order of the IC Memo]'
  );
  const dateStr = (icMemoDate || '').trim() || '[IC Memo Date]';
  return `This proposal was placed before the Investment Committee (IC) in its ${orderWithYear} meeting held on ${dateStr}. The Committee recommended the proposal as presented.`;
}

export function buildBoardMecRecommendationText(
  mecMemoNumber: string,
  mecMemoDate: string
): string {
  const orderWithYear = formatChronologicOrderWithYear(
    mecMemoNumber,
    '[Chronologic Order of the MEC Memo]'
  );
  const dateStr = (mecMemoDate || '').trim() || '[MEC Memo Date]';
  return `This proposal was placed before the Management Executive Committee (MEC) in its ${orderWithYear} meeting held on ${dateStr}. The Committee recommended the proposal as presented.`;
}

function buildDefaultSanctionBoardAuthority(
  icMemo: ResolvedICMemo,
  stageInputs: StageMemoInputs,
  stageAi: StageAiSectionsData
): string {
  const ph = icMemo.placeholders;
  const hasCore = Boolean(
    ph['#2#'] ||
      ph['#3#'] ||
      ph['#4#'] ||
      ph['#27#'] ||
      stageInputs.boardMemoNumber ||
      stageInputs.boardMemoDate ||
      stageAi.boardResolutionTerms ||
      stageAi.boardProposalRationale
  );
  if (!hasCore) return '';

  const boardOrder = formatChronologicOrderWithYear(
    stageInputs.boardMemoNumber,
    '[Chronologic Order of the Board Memo]'
  );
  const boardDate = stageInputs.boardMemoDate || '[Board Memo Date]';
  const mecOrder = stageInputs.mecMemoNumber
    ? formatChronologicOrderWithYear(stageInputs.mecMemoNumber, '')
    : '';
  const icOrder = stageInputs.icMemoNumber
    ? formatChronologicOrderWithYear(stageInputs.icMemoNumber, '')
    : '';

  const chainParts: string[] = [];
  if (mecOrder || stageInputs.mecMemoDate) {
    chainParts.push(`MEC Memo ${mecOrder || '—'} dated ${stageInputs.mecMemoDate || '—'}`);
  }
  if (icOrder || stageInputs.icMemoDate) {
    chainParts.push(`IC Memo ${icOrder || '—'} dated ${stageInputs.icMemoDate || '—'}`);
  }
  const chainText =
    chainParts.length > 0 ? ` following prior recommendations vide ${chainParts.join(' and ')}` : '';

  const acct = ph['#2#'] || '[Account No.]';
  const branch = ph['#3#'] || '[Branch Name]';
  const clientName = icMemo.branchProposal?.section10Particulars?.nameOfClient?.trim();
  const acctLabel = clientName && clientName !== acct ? `${clientName} (${acct})` : acct;

  return (
    `Pursuant to the approval accorded by the Board of Directors of Sammilito Islami Bank PLC vide Board Memo ${boardOrder} dated ${boardDate}${chainText}, Head Office hereby conveys formal sanction for the full and final Exit with Waiver facility in respect of A/C ${acctLabel} under ${branch} Branch under Bangladesh Bank BRPD Circular Letter No. 23 dated 29.06.2026.` +
    (stageAi.boardResolutionTerms?.trim()
      ? `\nApproved Board Resolution Summary: ${stageAi.boardResolutionTerms.trim()}`
      : '')
  );
}

function buildDefaultSanctionCrossVerifiedSummary(
  icMemo: ResolvedICMemo,
  stageInputs: StageMemoInputs
): string {
  const ph = icMemo.placeholders;
  const hasCore = Boolean(
    ph['#2#'] || ph['#25#'] || ph['#27#'] || ph['#7#'] || icMemo.branchProposal
  );
  if (!hasCore) return '';

  const lines: string[] = [];
  lines.push(
    `1. Facility & Settlement Parameters: Mode of Investment: ${ph['#5#'] || ph['#55#'] || '[Pending]'}; Present Limit: Tk. ${ph['#56#'] || ph['#6#'] || '[Pending]'} Million (Validity: ${ph['#58#'] || '[Pending]'}); NET Outstanding: Tk. ${ph['#25#'] || '[Pending]'} Million as on ${ph['#68#'] || '[Pending]'} (Principal: Tk. ${ph['#7#'] || '—'} M, Charges: Tk. ${ph['#8#'] || '—'} M, Charged Profit: Tk. ${ph['#10#'] || '—'} M, Compensation: Tk. ${ph['#11#'] || '—'} M, Unapplied Comp.: Tk. ${ph['#12#'] || '—'} M).`
  );
  lines.push(
    `2. Sanctioned Waiver & Recovery Position: Approved Settlement/Exit Amount of Tk. ${ph['#27#'] || '[Pending]'} Million (within ${ph['#28#'] || '[Settlement Period pending]'}) against Waiver of Tk. ${ph['#4#'] || '[Pending]'} Million (Max Waiver Considering Cost of Fund: Tk. ${ph['#44#'] || '—'} M; Cost of Fund Surplus/Deficit: Tk. ${ph['#31#'] || '—'} M; Breakup Total: Tk. ${ph['#32#'] || '—'} M), backed by deposited Down Payment of Tk. ${ph['#35#'] || '[Pending]'} Million in A/C ${ph['#36#'] || '[Pending]'}.`
  );
  const bp = icMemo.branchProposal;
  if (icMemo.sourceOfRepayment || bp?.referenceNo || stageInputs.boardMemoNumber) {
    lines.push(
      `3. Source of Repayment & Verification Chain: Recovery Till Date: Tk. ${ph['#60#'] || '—'} Million; Overdue: Tk. ${ph['#59#'] || '—'} Million; Repayment Source: ${icMemo.sourceOfRepayment || 'As stated in approved Board proposal'}. Verified against Branch Proposal Ref. ${ph['#64#'] || '—'} dated ${ph['#65#'] || '—'}.`
    );
  }
  return lines.join('\n');
}

function buildDefaultSanctionSecurityAndCompliance(icMemo: ResolvedICMemo): string {
  const ph = icMemo.placeholders;
  const bp = icMemo.branchProposal;
  if (!ph['#2#'] && !bp) return '';

  const cibDate = ph['#62#'] || bp?.section4CibDate || '[CIB Date pending]';
  const legalStatus =
    bp?.section5StatusNote?.trim() ||
    (bp?.section5Rows && bp.section5Rows.some((r) => r.caseNo || r.typeOfCase)
      ? `${bp.section5Rows.filter((r) => r.caseNo || r.typeOfCase).length} legal matter(s) recorded for settlement-linked withdrawal`
      : 'Nil / None reported');
  const auditStatus =
    bp?.section6StatusNote?.trim() ||
    (bp?.section6Rows && bp.section6Rows.some((r) => r.objectionStatus)
      ? 'Reviewed per IC & CD / Bangladesh Bank regulatory records'
      : 'Reviewed per Board Memo compliance record');
  const collateralNote = bp?.section7CollateralNote?.trim();
  const primaryNote = bp?.section7PrimaryNote?.trim();

  const parts: string[] = [
    `Security Valuation: Market Value (MV): ${ph['#33#'] || 'As per Board Memo'}; Forced Sale Value (FSV): ${ph['#34#'] || 'As per Board Memo'}${bp?.section7CoveragePct ? `; Coverage: ${bp.section7CoveragePct}` : ''}.`,
  ];
  if (primaryNote) parts.push(`Primary Security: ${primaryNote}`);
  if (collateralNote) parts.push(`Collateral Particulars: ${collateralNote}`);
  parts.push(`CIB Status (as on ${cibDate}): Verified prior to sanction issuance.`);
  parts.push(`Legal & Audit Compliance: Legal Status — ${legalStatus}; Audit & Regulatory Position — ${auditStatus}.`);
  return parts.join('\n');
}

function buildDefaultSanctionSpecialConditions(
  icMemo: ResolvedICMemo,
  stageInputs: StageMemoInputs,
  stageAi: StageAiSectionsData
): string {
  const ph = icMemo.placeholders;
  const hasCore = Boolean(
    ph['#2#'] ||
      ph['#27#'] ||
      ph['#28#'] ||
      stageInputs.boardMemoNumber ||
      stageAi.boardResolutionTerms
  );
  if (!hasCore) return '';

  const period = ph['#28#'] || '[Settlement Period]';
  const settlementAmt = ph['#27#'] || '[Settlement Amount]';
  const waiverAmt = ph['#4#'] || '[Waiver Amount]';
  const downPay = ph['#35#'] || '[Down Payment]';
  const downAc = ph['#36#'] || '[A/C No.]';

  const lines = [
    `i. The Branch shall ensure full and final realization of the approved Settlement/Exit Amount of Tk. ${settlementAmt} Million on or before ${period} strictly in lump sum, along with adjustment of the deposited Down Payment of Tk. ${downPay} Million held in A/C ${downAc}.`,
    `ii. The approved Waiver of Tk. ${waiverAmt} Million (together with unapplied/accrued profit up to execution) shall be adjusted from FAD, Head Office against provision held (if any), and the balance gradually by the Branch only after full realization of the settlement amount within ${period}.`,
    `iii. Final payoff balance shall be computed on the exact date of execution taking into account the up-to-date liability position.`,
    `iv. Prior to allowing redemption or release of any mortgaged collateral or encashable security (MTDR, margin, scheme deposit, etc.), the Branch shall ensure complete adjustment of all direct, indirect, and non-funded liabilities (L/C, Bank Guarantee, Credit Card, etc.) of the borrower and sister concerns.`,
    `v. The Profit/Rent rate shall NOT be altered or reduced to 0.00% until full and final adjustment of the investment as approved by Head Office.`,
    `vi. Upon any failure to adjust within ${period}, this sanction shall automatically stand cancelled and the Branch shall immediately pursue/resume legal recovery proceedings, with all legal expenses borne by the client.`,
  ];
  return lines.join('\n');
}

export function resolveStageMemo(
  docType: StageMemoDocType,
  icMemo: ResolvedICMemo,
  stageInputs: StageMemoInputs,
  stageAi: StageAiSectionsData
): ResolvedStageMemo {
  const cleanInputs: StageMemoInputs = {
    icMemoNumber: stageInputs.icMemoNumber.trim(),
    icMemoDate: stageInputs.icMemoDate.trim(),
    mecMemoNumber: stageInputs.mecMemoNumber.trim(),
    mecMemoDate: stageInputs.mecMemoDate.trim(),
    boardMemoNumber: stageInputs.boardMemoNumber.trim(),
    boardMemoDate: stageInputs.boardMemoDate.trim(),
    sanctionRefNumber: stageInputs.sanctionRefNumber.trim(),
    sanctionDate: stageInputs.sanctionDate.trim(),
  };

  const hasMecRefOrCase = Boolean(
    cleanInputs.icMemoNumber ||
      cleanInputs.icMemoDate ||
      icMemo.placeholders['#2#'] ||
      icMemo.branchProposal
  );
  const hasBoardRefOrCase = Boolean(
    cleanInputs.mecMemoNumber ||
      cleanInputs.mecMemoDate ||
      hasMecRefOrCase
  );

  const mecIcText =
    (stageAi.mecIcRecommendation ?? '').trim() ||
    (hasMecRefOrCase
      ? buildMecIcRecommendationText(cleanInputs.icMemoNumber, cleanInputs.icMemoDate)
      : '');

  const boardMecText =
    (stageAi.boardMecRecommendation ?? '').trim() ||
    (hasBoardRefOrCase
      ? buildBoardMecRecommendationText(cleanInputs.mecMemoNumber, cleanInputs.mecMemoDate)
      : '');

  const stageSections: Record<StageScrutinySectionKey, string> = {
    mecIcRecommendation: mecIcText,
    mecExecutiveSummary: (stageAi.mecExecutiveSummary ?? '').trim(),
    mecSamdRecommendation: (stageAi.mecSamdRecommendation ?? '').trim(),
    boardMecRecommendation: boardMecText,
    boardProposalRationale: (stageAi.boardProposalRationale ?? '').trim(),
    boardResolutionTerms: (stageAi.boardResolutionTerms ?? '').trim(),
    sanctionBoardAuthority:
      (stageAi.sanctionBoardAuthority ?? '').trim() ||
      buildDefaultSanctionBoardAuthority(icMemo, cleanInputs, stageAi),
    sanctionCrossVerifiedSummary:
      (stageAi.sanctionCrossVerifiedSummary ?? '').trim() ||
      buildDefaultSanctionCrossVerifiedSummary(icMemo, cleanInputs),
    sanctionSecurityAndCompliance:
      (stageAi.sanctionSecurityAndCompliance ?? '').trim() ||
      buildDefaultSanctionSecurityAndCompliance(icMemo),
    sanctionSpecialConditions:
      (stageAi.sanctionSpecialConditions ?? '').trim() ||
      buildDefaultSanctionSpecialConditions(icMemo, cleanInputs, stageAi),
  };

  return {
    docType,
    icMemo,
    stageInputs: cleanInputs,
    stageSections,
  };
}

