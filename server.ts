import 'dotenv/config';
import { createHash } from 'crypto';
import express from 'express';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, ThinkingLevel, Type, type Schema } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;

function getAiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not configured in environment variables.');
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

type ScrutinySectionKey =
  | 'justification'
  | 'sec1NatureOfBusiness'
  | 'sec2LiabilityOurBank'
  | 'sec3LiabilityOtherUnits'
  | 'sec4CibStatus'
  | 'sec5LegalAction'
  | 'sec6AuditCompliance'
  | 'sec7Security'
  | 'sec8SourceOfRepayment'
  | 'sec9HoSanctions'
  | 'sec10ClientParticulars'
  | 'sec11Financials'
  | 'clientPrayer'
  | 'branchVisit'
  | 'sec13BranchRecommendation'
  | 'sec14IccdRecommendation'
  | 'samdObservation';

type StageMemoDocType = 'mec' | 'board' | 'sanction';

type StageScrutinySectionKey =
  | 'mecIcRecommendation'
  | 'mecExecutiveSummary'
  | 'mecSamdRecommendation'
  | 'boardMecRecommendation'
  | 'boardProposalRationale'
  | 'boardResolutionTerms'
  | 'sanctionBoardAuthority'
  | 'sanctionCrossVerifiedSummary'
  | 'sanctionSpecialConditions';

function strField(description: string): Schema {
  return { type: Type.STRING, description };
}

function getScrutinyConfig(section: ScrutinySectionKey): {
  prompt: string;
  schema: Schema;
} {
  switch (section) {
    case 'justification':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal (especially "Justification of The Proposed Facility", subject/terms, business/distress background, recovery/down payment details, and branch recommendations) together with verified memo figures below.
Write clean, formal banking text strictly for IC Memo item "Justification of The Proposed Facility:" ("a.").
Format as exactly 3 numbered justification points separated by newlines ("1. ...", "2. ...", "3. ..."), covering:
1. Client's current business/financial position and genuine distress reflected in the Branch Proposal.
2. Recovery and settlement rationale (down payment, principal/Cost of Fund recovery, and protection of bank interest) grounded strictly in the Branch Proposal and provided figures.
3. Immediate reduction/regularization of classified NPL exposure and avoidance of prolonged litigation.
CRITICAL CLEAN-TEXT RULES: Output ONLY formal banking justification points in "text". Do NOT include conversational openings, AI notes, or explanations ("Based on the uploaded document", "Here is...", etc.), and do NOT use markdown asterisks.
If the Branch Proposal lacks information needed for this section, include ONLY what is factually supported (or "" if none) in "text" and explain what is unavailable strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean official banking justification points (1., 2., 3.) with no AI notes or disclaimers, or "" if source info is completely absent.'),
            insufficientReason: strField('Empty string "" if sufficiently supported; otherwise brief software UI note stating what info was unavailable in Branch Proposal.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec1NatureOfBusiness':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal (Section "1. Nature of Business", Section 10 "Business Type/Nature", and business description) to populate IC Memo Section "1. Nature of Business:".
Return ONLY the concise official Nature of Business phrase in "text" (e.g., "Agriculture (Fisheries/Dairy)" or exact activity wording from the Branch Proposal).
Do NOT include conversational remarks, "Based on...", AI notes, or markdown asterisks in "text".
If the Branch Proposal does not state the nature of business, set "text" to "" and explain strictly in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Exact concise Nature of Business label/phrase for the memo, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if found; otherwise concise UI notice explaining missing information.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec2LiabilityOurBank':
      return {
        prompt: `Scrutinize Section "2. Liability Position of the Clients and Its Sister Concerns with Our Bank" in the uploaded Branch Proposal.
Extract strictly what appears in the Branch Proposal:
- limit: Sanctioned limit in Million BDT (e.g. "0.50" or "0.50 Million"; "" if absent).
- validity: Validity / expiry date and status (e.g. "01.09.2025 (Exp.)"; "" if absent).
- remarks: Exact remarks in Section 2 table ("" if blank).
- nonFundedNote: Exact Non-Funded or sister concern note in Section 2 table ("" if blank/Nil).
- summary: For software UI audit log only — if key fields are blank/insufficient in the Branch Proposal, state concisely why; otherwise "". Never fabricate figures.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            limit: strField('Section 2 Limit from Branch Proposal, or "" if absent.'),
            validity: strField('Section 2 Validity/Expiry from Branch Proposal, or "" if absent.'),
            remarks: strField('Section 2 Remarks from Branch Proposal, or "" if blank.'),
            nonFundedNote: strField('Non-Funded or sister concern remarks from Section 2, or "" if none.'),
            summary: strField('Software UI insufficient-data note only if info is missing/blank in Branch Proposal; else "".'),
          },
          required: ['limit', 'validity', 'remarks', 'nonFundedNote', 'summary'],
        },
      };

    case 'sec3LiabilityOtherUnits':
      return {
        prompt: `Scrutinize Section "3. Liability of the Client & Its Sister/Allied Concerns with Other Units of SIB" (or other banks/NBFIs) in the uploaded Branch Proposal.
Extract every row present in Section 3 of the Branch Proposal into "rows" (with exact columns: sl, nameOfConcern, bankOrNbfi, natureOfFacility, limit, outstanding, overdueEol, expiryDate, clStatus).
- If the Branch Proposal table has blank/empty rows or states "Nil" / "N/A", reflect that accurately (do NOT fabricate bank names, limits, or facilities; keep empty cells as "" or "Nil" only where the Branch Proposal actually states Nil).
- In "summary", write a brief formal statement explaining what the Branch Proposal shows for Section 3 (e.g., whether any sister/allied or other-unit/other-bank liabilities exist, or explicitly stating that the Branch Proposal records Nil / leaves the table blank / provides insufficient information).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL (e.g. a., b., c., Sub-Total, Grand Total)'),
                  nameOfConcern: strField('Name of Concern'),
                  bankOrNbfi: strField('Name of the Banks/NBFIs'),
                  natureOfFacility: strField('Nature of Facility'),
                  limit: strField('Limit'),
                  outstanding: strField('Outstanding with profit/Interest'),
                  overdueEol: strField('Overdue/EOL'),
                  expiryDate: strField('Expiry Date'),
                  clStatus: strField('CL Status'),
                },
                required: [
                  'sl',
                  'nameOfConcern',
                  'bankOrNbfi',
                  'natureOfFacility',
                  'limit',
                  'outstanding',
                  'overdueEol',
                  'expiryDate',
                  'clStatus',
                ],
              },
            },
            summary: strField('Formal banking summary of Section 3 findings or note on Nil/insufficient info.'),
          },
          required: ['rows', 'summary'],
        },
      };

    case 'sec4CibStatus':
      return {
        prompt: `Scrutinize Section "4. CIB Status" in the uploaded Branch Proposal thoroughly.
Extract:
- cibDate: The CIB report date mentioned in Section 4 of the Branch Proposal (e.g. "17.11.2026"), or "" if absent.
- rows: Extract the exact values from the Section 4 CIB Status table for each status row (STD, SMA, SS, DF, BL, Willful Default (WD), Write-Off, Non-Funded, Stay Order, TOTAL) across columns: status, unitBorrower, unitGuarantor, otherBorrower, otherGuarantor, remarks.
- Fix obvious scanner/OCR typos ONLY when clearly verifiable from the visual table image (e.g. "Nit" -> "Nil", "10/-" if actually "0/-" or "10/-" per image — read the visual page image carefully!). Leave genuinely blank cells as "". Do NOT invent figures.
- summary: Write 1-2 formal banking sentences summarizing the CIB status as recorded in the Branch Proposal (or clearly state if CIB information is unavailable/insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            cibDate: strField('CIB report date from Branch Proposal Section 4, or empty string.'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  status: strField('Status label (STD, SMA, SS, DF, BL, Willful Default (WD), Write-Off, Non-Funded, Stay Order, TOTAL)'),
                  unitBorrower: strField('With respective unit - Borrower'),
                  unitGuarantor: strField('With respective unit - Guarantor'),
                  otherBorrower: strField('With Other Bank - Borrower'),
                  otherGuarantor: strField('With Other Bank - Guarantor'),
                  remarks: strField('Remarks'),
                },
                required: ['status', 'unitBorrower', 'unitGuarantor', 'otherBorrower', 'otherGuarantor', 'remarks'],
              },
            },
            summary: strField('Concise formal summary of CIB status from the Branch Proposal.'),
          },
          required: ['cibDate', 'rows', 'summary'],
        },
      };

    case 'sec5LegalAction':
      return {
        prompt: `Scrutinize Section "5. Legal Action Status as on Memo Date" in the uploaded Branch Proposal.
Extract:
- statusNote: Any header status note such as "Nil" or "N/A" next to the heading in the Branch Proposal (or "" if none).
- rows: Any legal case rows listed in Section 5 of the Branch Proposal (sl, typeOfCase, caseNo, filingDate, amount, previousDate, nextDate, presentStatus). If no cases are filed / table is empty, return an empty array []. Do NOT invent case numbers or dates.
- summary: A brief formal banking statement describing the legal action status strictly per the Branch Proposal (or clearly stating if legal action details are marked Nil/N/A or insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            statusNote: strField('Status note next to Section 5 heading (e.g. Nil, N/A, or empty string).'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('Sl'),
                  typeOfCase: strField('Type of Case'),
                  caseNo: strField('Case No.'),
                  filingDate: strField('Filing date'),
                  amount: strField('Amount'),
                  previousDate: strField('Previous date'),
                  nextDate: strField('Next date'),
                  presentStatus: strField('Present status'),
                },
                required: ['sl', 'typeOfCase', 'caseNo', 'filingDate', 'amount', 'previousDate', 'nextDate', 'presentStatus'],
              },
            },
            summary: strField('Formal banking summary of Legal Action Status from the Branch Proposal.'),
          },
          required: ['statusNote', 'rows', 'summary'],
        },
      };

    case 'sec6AuditCompliance':
      return {
        prompt: `Scrutinize Section "6. Audit, Compliance and Regulatory Authority Related Status" in the uploaded Branch Proposal.
Extract:
- statusNote: Any note next to the Section 6 heading (e.g., "Nil", "N/A", or "").
- rows: Extract the exact inspection dates, objection status (Yes/No), objections, and branch reply/current position for each authority row in Section 6 (a. IC & CD, b. DBI, Bangladesh Bank, c. DOS, Bangladesh Bank, d. BFIU, Bangladesh Bank, e. BRD, Bangladesh Bank, f. ACC(Anti-Corruption Commission)).
- Preserve exact values visible in the Branch Proposal image/text. Do NOT fabricate inspection dates or replies.
- summary: A concise formal banking summary of the audit, compliance, and regulatory inspection position in the Branch Proposal (or clearly note if insufficient/unavailable).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            statusNote: strField('Status note next to Section 6 heading (e.g., N/A, Nil, or empty string).'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL'),
                  authority: strField('Name of the Authority'),
                  lastInspectionDate: strField('Last Inspection / Report Date'),
                  objectionStatus: strField('Objection Status (Yes/No)'),
                  objections: strField('Objections (If yes)'),
                  branchReply: strField('Branch Reply/ Current Position'),
                },
                required: ['sl', 'authority', 'lastInspectionDate', 'objectionStatus', 'objections', 'branchReply'],
              },
            },
            summary: strField('Formal summary of Audit, Compliance & Regulatory Status from Branch Proposal.'),
          },
          required: ['statusNote', 'rows', 'summary'],
        },
      };

    case 'sec7Security':
      return {
        prompt: `Scrutinize Section "7. Security" (including a. Primary, b. Collateral, schedule descriptions, valuation table, coverage %, and c. Others, plus MV/FSV in Remarks if noted) in the uploaded Branch Proposal.
Extract strictly what is recorded in the Branch Proposal:
- primaryNote: Note next to "a. Primary:" (e.g. "Nil", "N/A", or "").
- primaryRows: Rows in the Primary Security table (particulars, value).
- primaryTotal: Total value of primary security (or "" if blank/Nil).
- collateralNote: Note next to "b. Collateral:" (e.g. "N/A", "Nil", or "").
- collateralSchedules: List of collateral schedule descriptions if filled in the Branch Proposal (if template placeholders are blank/unfilled, return standard schedule lines only as shown or empty array if marked N/A).
- valuationRows: Rows in the Collateral Valuation table (schedule, origBranchMv, origBranchFsv, origSurvMv, origSurvFsv, presBranchMv, presBranchFsv, presSurvMv, presSurvFsv, presAvMv, presAvFsv). Leave unfilled cells as "".
- coveragePct: "Outstanding to Lowest Security value Coverage (%)" value if stated in the Branch Proposal (or "" if blank).
- othersNote: Any additional text under "c. Others:" if stated beyond standard charge documents.
- summary: A concise formal banking assessment of primary & collateral security strictly grounded in the Branch Proposal (explicitly stating if primary/collateral security is marked Nil/N/A or lacks valuation figures in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            primaryNote: strField('Note next to a. Primary (e.g. Nil, N/A, or empty).'),
            primaryRows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  particulars: strField('Particulars'),
                  value: strField('Value'),
                },
                required: ['particulars', 'value'],
              },
            },
            primaryTotal: strField('Total value of primary security'),
            collateralNote: strField('Note next to b. Collateral (e.g. N/A, Nil, or empty).'),
            collateralSchedules: { type: Type.ARRAY, items: { type: Type.STRING } },
            valuationRows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  schedule: strField('Schedule'),
                  origBranchMv: strField('Original Branch MV'),
                  origBranchFsv: strField('Original Branch FSV'),
                  origSurvMv: strField('Original Surveyor MV'),
                  origSurvFsv: strField('Original Surveyor FSV'),
                  presBranchMv: strField('Present Branch MV'),
                  presBranchFsv: strField('Present Branch FSV'),
                  presSurvMv: strField('Present Surveyor MV'),
                  presSurvFsv: strField('Present Surveyor FSV'),
                  presAvMv: strField('Present AV Cell MV'),
                  presAvFsv: strField('Present AV Cell FSV'),
                },
                required: [
                  'schedule',
                  'origBranchMv',
                  'origBranchFsv',
                  'origSurvMv',
                  'origSurvFsv',
                  'presBranchMv',
                  'presBranchFsv',
                  'presSurvMv',
                  'presSurvFsv',
                  'presAvMv',
                  'presAvFsv',
                ],
              },
            },
            coveragePct: strField('Coverage percentage if stated, else empty string.'),
            othersNote: strField('Additional note under c. Others if any, else empty string.'),
            summary: strField('Concise formal banking summary of Section 7 Security in the Branch Proposal.'),
          },
          required: [
            'primaryNote',
            'primaryRows',
            'primaryTotal',
            'collateralNote',
            'collateralSchedules',
            'valuationRows',
            'coveragePct',
            'othersNote',
            'summary',
          ],
        },
      };

    case 'sec8SourceOfRepayment':
      return {
        prompt: `Scrutinize Section "8. Source of Repayment" (and repayment/settlement funding details across the Branch Proposal).
Write clean, official banking prose strictly stating the source of repayment for IC Memo Section "8. Source of Repayment:" based solely on the Branch Proposal.
Do NOT include conversational openings, "Based on the uploaded document", AI notes, or markdown asterisks in "text".
If the Branch Proposal contains insufficient or no information on the source of repayment, set "text" to "" and state what is missing in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean formal Source of Repayment prose for the memo body, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note explaining missing information in Branch Proposal.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec9HoSanctions':
      return {
        prompt: `Scrutinize Section "9. Sequence of Head Office Sanction(S)" in the uploaded Branch Proposal.
Extract every sanction row visible in Section 9 (columns: sl, referenceNo, date, limit, remarks). Read the visual page image carefully so reference numbers, dates, and limits are accurate.
If rows are blank or no Head Office sanctions are listed, return only populated rows (or empty array if none) and clearly indicate in "summary" what the Branch Proposal shows (or that Head Office sanction details are unavailable/insufficient in the Branch Proposal). Do NOT invent reference numbers or dates.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL (a., b., c., etc.)'),
                  referenceNo: strField('HO Sanction Reference No.'),
                  date: strField('Date'),
                  limit: strField('Limit'),
                  remarks: strField('Remarks'),
                },
                required: ['sl', 'referenceNo', 'date', 'limit', 'remarks'],
              },
            },
            summary: strField('Concise formal summary of Head Office Sanction sequence or unavailable/insufficient note.'),
          },
          required: ['rows', 'summary'],
        },
      };

    case 'sec10ClientParticulars':
      return {
        prompt: `Scrutinize Section "10. Brief Particulars of The Client" (items a. through u., plus "Particular of the Proprietor/Partner/Directors" and "Remarks: Director of any Bank/ FI/ Insurance company") in the uploaded Branch Proposal.
Carefully read the visual page images and text to extract EVERY available item accurately without OCR errors:
- particulars: nameOfClient (a), uniqueCustomerId (b), awcdMsdAccount (c), businessTypeNature (d), businessManagementType (e), clientType (f), keyPersonPhone (g), address (g), officeBusinessAddress (h), factoryAddress (h), purposeOfInvestment (i), limitOfInvestment (j), modeOfInvestment (k), lastVisitDate (l), stockPosition (m), fixedAsset (n), businessNetWorth (o), insuranceCoverage (p), tradeLicenseValidity (q), otherRegulatoryDocuments (r), dateOfOriginalInvestment (s), firstDisbursementManager (t), firstDateOfClassification (u). Leave genuinely missing/blank items as "".
- directors: rows from "Particular of the Proprietor/Partner/Directors" table (empty [] if blank/N/A).
- bankDirectors: rows from "Remarks: Director of any Bank/ FI/ Insurance company" table (empty [] if blank/N/A).
- summary: Concise formal note summarizing key client particulars and explicitly noting if director/bank-director tables or specific particulars are marked N/A or unavailable in the Branch Proposal.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            particulars: {
              type: Type.OBJECT,
              properties: {
                nameOfClient: strField('a. Name of the Client'),
                uniqueCustomerId: strField('b. Unique Customer ID'),
                awcdMsdAccount: strField('c. AWCD/MSD Account with opening date'),
                businessTypeNature: strField('d. Business Type/Nature'),
                businessManagementType: strField('e. Business Management Type'),
                clientType: strField('f. Client Type'),
                keyPersonPhone: strField('g. Key Person with Phone No.'),
                address: strField('g. Address'),
                officeBusinessAddress: strField('h. Office/Business Address'),
                factoryAddress: strField('h. Factory Address'),
                purposeOfInvestment: strField('i. Purpose of Investment'),
                limitOfInvestment: strField('j. Limit of Investment'),
                modeOfInvestment: strField('k. Mode of Investment'),
                lastVisitDate: strField('l. Last Visit Date'),
                stockPosition: strField('m. Stock Position'),
                fixedAsset: strField('n. Fixed Asset'),
                businessNetWorth: strField('o. Business Net Worth'),
                insuranceCoverage: strField('p. Insurance Coverage'),
                tradeLicenseValidity: strField('q. Trade License & Validity'),
                otherRegulatoryDocuments: strField('r. Other Regulatory Documents'),
                dateOfOriginalInvestment: strField('s. Date of Original Investment'),
                firstDisbursementManager: strField('t. 1st Disbursement Manager'),
                firstDateOfClassification: strField('u. First Date of Classification with Status'),
              },
              required: [
                'nameOfClient',
                'uniqueCustomerId',
                'awcdMsdAccount',
                'businessTypeNature',
                'businessManagementType',
                'clientType',
                'keyPersonPhone',
                'address',
                'officeBusinessAddress',
                'factoryAddress',
                'purposeOfInvestment',
                'limitOfInvestment',
                'modeOfInvestment',
                'lastVisitDate',
                'stockPosition',
                'fixedAsset',
                'businessNetWorth',
                'insuranceCoverage',
                'tradeLicenseValidity',
                'otherRegulatoryDocuments',
                'dateOfOriginalInvestment',
                'firstDisbursementManager',
                'firstDateOfClassification',
              ],
            },
            directors: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  name: strField('Name of the Directors'),
                  age: strField('Age'),
                  education: strField('Education'),
                  designation: strField('Designation'),
                  sharePct: strField('Share (%)'),
                  experience: strField('Experience'),
                  netWorth: strField('Net worth'),
                },
                required: ['name', 'age', 'education', 'designation', 'sharePct', 'experience', 'netWorth'],
              },
            },
            bankDirectors: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  name: strField('Name of the Director'),
                  bankNbfi: strField('Name of the bank/NBFIs'),
                  designation: strField('Designation/Position'),
                  sharePct: strField('Percentage of Shares'),
                  numShares: strField('Number of Shares'),
                },
                required: ['name', 'bankNbfi', 'designation', 'sharePct', 'numShares'],
              },
            },
            summary: strField('Concise formal summary of client particulars and any unavailable/N/A items.'),
          },
          required: ['particulars', 'directors', 'bankDirectors', 'summary'],
        },
      };

    case 'sec11Financials':
      return {
        prompt: `Scrutinize Section "11. Financials of The Client" in the uploaded Branch Proposal.
Extract:
- financialsNote: Any note such as "Nil" or "N/A" if the client has no business or financials are marked Nil in the Branch Proposal (or "" if actual figures exist).
- rows: Financial rows (year, grossSales, grossProfit, netProfit, totalAsset, totalLiabilities, netWorth) if figures appear in the Branch Proposal (empty [] if marked Nil/N/A/blank). Do NOT fabricate financial figures.
- summary: A brief formal banking statement summarizing the financial position reported in the Branch Proposal (or explicitly noting that financials are marked Nil / unavailable / insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            financialsNote: strField('Note such as Nil or N/A if stated, else empty string.'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  year: strField('Year'),
                  grossSales: strField('Gross Sales'),
                  grossProfit: strField('Gross Profit'),
                  netProfit: strField('Net Profit'),
                  totalAsset: strField('Total Asset'),
                  totalLiabilities: strField('Total liabilities'),
                  netWorth: strField('Net Worth'),
                },
                required: ['year', 'grossSales', 'grossProfit', 'netProfit', 'totalAsset', 'totalLiabilities', 'netWorth'],
              },
            },
            summary: strField('Formal banking summary of Client Financials or Nil/insufficient notice.'),
          },
          required: ['financialsNote', 'rows', 'summary'],
        },
      };

    case 'clientPrayer':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal thoroughly for any details of the customer's written application / prayer ("Contents of the Client's Prayer", client's request letter/application, subject clause, requested waiver amount, offered settlement amount & down payment, proposed repayment period/timeline, stated business distress/reasons for default, and commitment for full and final exit).
Write clean, concise, official banking prose for IC Memo Section "12. Contents of the Client's Prayer:" strictly grounded in the Branch Proposal and verified memo figures below.
CRITICAL CLEAN-TEXT RULES: Output ONLY formal third-person banking memorandum prose describing the client's prayer/request in "text" (1-2 concise formal paragraphs). Do NOT include "Based on the uploaded document", conversational framing, AI notes, or markdown asterisks.
If the Branch Proposal contains insufficient information about the client's application/prayer, keep ONLY verified facts (or "" if none) in "text" and explain what is missing strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField("Clean official banking prose summarizing Contents of the Client's Prayer, or '' if unavailable."),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note listing missing client prayer details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'branchVisit':
      return {
        prompt: `Scrutinize all relevant sections of the uploaded Branch Proposal (Section 12 "Branch Visit Report", Section 10 Client Particulars, office/factory addresses, last visit date, stock position, fixed assets, business net worth, Section 7 security details, and branch recommendations) to write Section "12. Branch Visit Report:" for the Head Office IC Memo.
CRITICAL REQUIREMENT: Write strictly in the THIRD PERSON regarding the branch (e.g., "The officials of [Branch Name] Branch...", "Branch management conducted a physical inspection...", never use "we", "us", "our", or "our branch").
Write 1 to 2 clean, formal banking paragraphs containing ONLY verified officials' names/designations, visit date, client/concern name, premises inspected, physical business/stock findings, and inspection outcome from the Branch Proposal.
Do NOT include "Based on the uploaded Branch Proposal", AI notes, disclaimers, or markdown asterisks inside "text".
If visit details are absent/insufficient in the Branch Proposal, keep ONLY verified facts (or "" if none) in "text" and record what is missing strictly in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean third-person formal Branch Visit Report prose with no AI notes or disclaimers, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if sufficient visit facts were found; otherwise brief UI notice of missing visit details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec13BranchRecommendation':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal for the Branch's specific recommendation (Section 15 "Recommendation of Branch Investment Committee", Section 16 "Branch Recommendation", Proposal Reference No., Proposal Date, settlement amount, waiver amount, settlement period, and conditions).
Write clean formal banking prose for IC Memo Section "13. Recommendation of The Branch:" stating the exact Branch / BIC recommendation.
Do NOT include "Based on the uploaded document", conversational framing, AI commentary, or markdown asterisks inside "text".
If recommendation details are blank or insufficient in the Branch Proposal, set "text" to verified portions only (or "" if none) and record the gap strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean formal Branch Recommendation prose for the memo body, or "" if absent.'),
            insufficientReason: strField('Empty string "" if sufficient; otherwise brief UI notice of missing recommendation details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec14IccdRecommendation':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal for any recommendation or comment from the Internal Control & Compliance Division (ICCD) (Section 14 or ICCD remarks).
Write clean formal content for IC Memo Section "14. Recommendation of Internal Control & Compliance Division:":
- If the Branch Proposal contains an ICCD recommendation, summarize it directly in clean formal banking prose without meta-commentary.
- If the Branch Proposal states "N/A", "Nil", or has no ICCD recommendation, return strictly "N/A" in "text" (no conversational notes or parenthetical AI explanations!).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean ICCD Recommendation text or exact "N/A".'),
            insufficientReason: strField('Empty string "" unless Branch Proposal appears unreadable/missing.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'samdObservation':
      return {
        prompt: `Thoroughly scrutinize the uploaded Branch Proposal together with the verified IC Memo figures below, and write Section "15. Observation and Recommendation of SAMD:" as Head Office SAMD.
Write exactly five (5) substantive, formal banking points separated by newlines ("1. ...", "2. ...", "3. ...", "4. ...", "5. ...") addressing:
1. Verification of total outstanding liability breakup, classification/overdue/CIB status, and confirmation of absence of fraud, forgery, or willful default per records.
2. Assessment of primary & collateral security position/coverage, legal action status, and audit/regulatory compliance position.
3. Cost of Fund & waiver adequacy analysis (down payment, recovery vs. principal/charges/Cost of Fund, maximum permissible waiver considering COF, and conformity with BRPD Circular Letter No. 23 dated 29.06.2026).
4. Evaluation of the client's business/financial condition, Branch Visit findings, Branch Recommendation, and source of repayment.
5. Formal SAMD recommendation placing the Exit with Waiver proposal before the Investment Committee for exit within the stipulated settlement period subject to compliance with special terms and conditions.
CRITICAL CLEAN-TEXT RULES: Output ONLY crisp official SAMD memorandum points in "text". Do NOT add conversational remarks ("Here is...", "Based on the uploaded document..."), disclaimers, meta-commentary, or markdown asterisks. Ground every figure strictly in the Branch Proposal and context; record any missing source inputs in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean official 5 numbered SAMD observation & recommendation points separated by newlines.'),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note listing missing source inputs.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };
  }
}

const VALID_SCRUTINY_SECTIONS = new Set<ScrutinySectionKey>([
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
]);

interface IncomingBranchProposalSource {
  sourceId?: string;
  fileName?: string;
  rawText?: string;
  pageImages?: Array<{ mimeType: string; data: string }>;
}

interface CachedBranchProposalSource {
  sourceId: string;
  fileName: string;
  rawText: string;
  imageParts: Array<{ inlineData: { mimeType: string; data: string } }>;
}

const MAX_CACHED_SOURCES = 20;
const branchProposalCache = new Map<string, CachedBranchProposalSource>();

function buildCachedSource(
  sourceId: string,
  incoming: IncomingBranchProposalSource
): CachedBranchProposalSource | null {
  const rawText = typeof incoming.rawText === 'string' ? incoming.rawText.slice(0, 45000) : '';
  const pageImages = Array.isArray(incoming.pageImages) ? incoming.pageImages.slice(0, 12) : [];
  const imageParts: Array<{ inlineData: { mimeType: string; data: string } }> = [];
  for (const img of pageImages) {
    if (img?.data && img?.mimeType) {
      imageParts.push({
        inlineData: {
          mimeType: img.mimeType,
          data: img.data,
        },
      });
    }
  }

  if (!rawText.trim() && imageParts.length === 0) {
    return null;
  }

  return {
    sourceId,
    fileName: incoming.fileName || 'Branch Proposal',
    rawText,
    imageParts,
  };
}

function storeCachedSource(sourceId: string, entry: CachedBranchProposalSource): void {
  if (branchProposalCache.has(sourceId)) {
    branchProposalCache.delete(sourceId);
  }
  branchProposalCache.set(sourceId, entry);
  while (branchProposalCache.size > MAX_CACHED_SOURCES) {
    const oldestKey = branchProposalCache.keys().next().value;
    if (oldestKey === undefined) break;
    branchProposalCache.delete(oldestKey);
  }
}

function resolveBranchProposalSource(
  sourceId?: string,
  incoming?: IncomingBranchProposalSource | null
):
  | { ok: true; source: CachedBranchProposalSource }
  | { ok: false; status: number; error: string; code?: string } {
  const effectiveId = (incoming?.sourceId || sourceId || '').trim();

  if (incoming) {
    const built = buildCachedSource(effectiveId || `bp_${Date.now()}`, incoming);
    if (!built) {
      return {
        ok: false,
        status: 400,
        error:
          'Branch Proposal is not uploaded or readable. Please upload the Branch Proposal first before running AI scrutiny.',
      };
    }
    if (effectiveId) {
      storeCachedSource(effectiveId, built);
    }
    return { ok: true, source: built };
  }

  if (effectiveId) {
    const hit = branchProposalCache.get(effectiveId);
    if (hit) {
      return { ok: true, source: hit };
    }
    return {
      ok: false,
      status: 404,
      code: 'SOURCE_CACHE_MISS',
      error: 'Cached Branch Proposal source not found on server; resending full source.',
    };
  }

  return {
    ok: false,
    status: 400,
    error:
      'Branch Proposal is not uploaded or readable. Please upload the Branch Proposal first before running AI scrutiny.',
  };
}

function buildCompactSectionContext(
  section: ScrutinySectionKey,
  context?: Record<string, unknown>
): string {
  if (!context) return '{}';

  // Sections that only extract directly from Branch Proposal don't need redundant branchProposal JSON or all narrative drafts
  if (
    section === 'justification' ||
    section === 'sec2LiabilityOurBank' ||
    section === 'samdObservation'
  ) {
    const { branchProposal: _omitBp, ...rest } = context;
    return JSON.stringify(rest);
  }

  // Extraction sections get verified numeric placeholders + summaries without bloated narrative text or redundant branchProposal dump
  return JSON.stringify({
    placeholders: context.placeholders ?? {},
    extractedSummaries: context.extractedSummaries ?? {},
  });
}

function sanitizeAiMemoText(raw: string): string {
  let cleaned = raw
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/^\s*\*\s+/gm, '')
    .trim();

  // Remove unwanted conversational AI preambles or processing disclaimers
  const unwantedLinePrefix =
    /^(?:(?:here\s+is|below\s+is|following\s+is)\s+(?:the\s+)?(?:generated|extracted|formal|requested)[^:\n]*:?\s*|ai[-\s]*generated\s*(?:content|text|section|summary)\s*:?\s*|based\s+on\s+(?:a\s+thorough\s+)?scrutiny\s+of\s+the\s+uploaded\s+branch\s+proposal[,:\s]*)/i;
  cleaned = cleaned.replace(unwantedLinePrefix, '').trim();

  const metaDisclaimerLine =
    /^(?:\[?\s*(?:ai[-\s]*note|processing\s+note|disclaimer|internal\s+note|note\s*:\s*(?:based\s+on\s+the\s+uploaded|information\s+(?:regarding|is)\s+(?:unavailable|insufficient|not\s+provided)))[^\n]*)$/gim;
  cleaned = cleaned.replace(metaDisclaimerLine, '').replace(/\n{3,}/g, '\n\n').trim();

  // If the entire text is literally a "no info / unavailable" meta-message, return "" so clean memo body isn't polluted
  if (
    /^(?:information\s+(?:regarding|on|for)[^.]{0,80}\s+is\s+)?(?:unavailable|insufficient|not\s+(?:specified|provided|mentioned|available))(?:\s+in\s+the\s+(?:uploaded\s+)?branch\s+proposal)?\.?$/i.test(
      cleaned
    )
  ) {
    return '';
  }

  return cleaned;
}

const NARRATIVE_SECTION_KEYS = new Set<ScrutinySectionKey>([
  'justification',
  'sec1NatureOfBusiness',
  'sec8SourceOfRepayment',
  'clientPrayer',
  'branchVisit',
  'sec13BranchRecommendation',
  'sec14IccdRecommendation',
  'samdObservation',
]);

function evaluateSectionResult(
  section: ScrutinySectionKey,
  parsed: Record<string, unknown>
): { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string } {
  const insufficientPattern =
    /\b(unavailable|insufficient|not\s+(?:specified|provided|mentioned|available|stated)|lacks?\s+sufficient|left\s+blank)\b/i;

  if (NARRATIVE_SECTION_KEYS.has(section)) {
    const rawText = String(parsed.text ?? '').trim();
    const cleanText = sanitizeAiMemoText(rawText);
    const modelReason = String(parsed.insufficientReason ?? '').trim();

    let reason = modelReason;
    if (!cleanText && !reason) {
      reason = insufficientPattern.test(rawText)
        ? rawText
        : 'Necessary details were not found in the uploaded Branch Proposal.';
    }

    const insufficient = Boolean(reason) || cleanText === '';
    return {
      result: { text: cleanText, insufficientReason: reason },
      insufficient,
      insufficientReason: reason,
    };
  }

  const summary = String(parsed.summary ?? '').trim();
  let reason = insufficientPattern.test(summary) ? summary : '';

  if (section === 'sec2LiabilityOurBank') {
    const hasAnyField = Boolean(
      String(parsed.limit ?? '').trim() ||
        String(parsed.validity ?? '').trim() ||
        String(parsed.remarks ?? '').trim() ||
        String(parsed.nonFundedNote ?? '').trim()
    );
    if (!hasAnyField && !reason) {
      reason = summary || 'Section 2 liability details were blank or unavailable in the Branch Proposal.';
    }
  } else if (section === 'sec10ClientParticulars') {
    const particulars = (parsed.particulars ?? {}) as Record<string, unknown>;
    const filledCount = Object.values(particulars).filter((v) => String(v ?? '').trim() !== '').length;
    if (filledCount === 0 && !reason) {
      reason = summary || 'Client particulars fields were blank or unavailable in the Branch Proposal.';
    }
  }

  return {
    result: parsed,
    insufficient: Boolean(reason),
    insufficientReason: reason,
  };
}

const FREE_TIER_MODELS = ['gemini-4-hydrogen', 'gemini-3.1-flash-lite'] as const;
type SupportedScrutinyModel = (typeof FREE_TIER_MODELS)[number];

const modelCooldownUntil = new Map<string, number>();

function isQuotaOrRateError(msg: string): boolean {
  return /\b429\b|RESOURCE_EXHAUSTED|quota|rate\s*limit/i.test(msg);
}

function isDailyQuotaError(msg: string): boolean {
  return /GenerateRequestsPerDay|per\s*day|daily|free_tier/i.test(msg);
}

function isModelNotFoundError(msg: string): boolean {
  return /\b404\b|NOT_FOUND|is not found|not supported for generateContent/i.test(msg);
}

function formatGeminiError(err: unknown, fallback = 'Failed to scrutinize section.'): string {
  const raw = err instanceof Error ? err.message : String(err ?? '');
  if (!raw) return fallback;
  try {
    const parsed = JSON.parse(raw) as { error?: { code?: number; status?: string; message?: string } };
    if (parsed?.error) {
      const code = parsed.error.code;
      const status = parsed.error.status ?? '';
      if (code === 429 || status === 'RESOURCE_EXHAUSTED') {
        return isDailyQuotaError(raw)
          ? 'Gemini API daily request quota reached across available free-tier models — click Retry once quota resets or select a billing-enabled API key in Settings > Secrets.'
          : 'Gemini API rate limit reached — click Retry in a moment to generate this section.';
      }
      if (parsed.error.message) {
        return parsed.error.message.split('\n')[0];
      }
    }
  } catch {
    // Plain text error message
  }
  if (isQuotaOrRateError(raw)) {
    return isDailyQuotaError(raw)
      ? 'Gemini API daily request quota reached across available free-tier models — click Retry once quota resets or select a billing-enabled API key in Settings > Secrets.'
      : 'Gemini API rate limit reached — click Retry in a moment to generate this section.';
  }
  return raw.split('\n')[0] || fallback;
}

function buildModelConfig(
  model: SupportedScrutinyModel,
  baseConfig: Record<string, unknown>
): Record<string, unknown> {
  if (model === 'gemini-4-hydrogen') {
    return {
      ...baseConfig,
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
    };
  }
  return baseConfig;
}

async function executeWithModelFallback<T>(
  runner: (model: SupportedScrutinyModel) => Promise<T>,
  shouldStop?: () => boolean
): Promise<T> {
  const now = Date.now();
  const active = FREE_TIER_MODELS.filter((m) => (modelCooldownUntil.get(m) ?? 0) <= now);
  const candidates: SupportedScrutinyModel[] =
    active.length > 0 ? active : [...FREE_TIER_MODELS];

  let lastErr: unknown;
  for (let idx = 0; idx < candidates.length; idx++) {
    if (shouldStop?.()) {
      throw new Error('Request cancelled.');
    }
    const model = candidates[idx];
    try {
      const result = await runner(model);
      modelCooldownUntil.delete(model);
      return result;
    } catch (err) {
      lastErr = err;
      const msg = err instanceof Error ? err.message : String(err);

      if (isDailyQuotaError(msg) || isModelNotFoundError(msg)) {
        modelCooldownUntil.set(model, Date.now() + 10 * 60_000);
        continue;
      }

      if (isQuotaOrRateError(msg)) {
        modelCooldownUntil.set(model, Date.now() + 20_000);
        continue;
      }

      const isTransientNet = /\b(500|502|503|504)\b|timeout|network|econnreset/i.test(msg);
      if (isTransientNet && !shouldStop?.()) {
        await new Promise((resolve) => setTimeout(resolve, 400));
        try {
          return await runner(model);
        } catch (retryErr) {
          lastErr = retryErr;
          continue;
        }
      }

      throw err;
    }
  }

  throw lastErr;
}

const MAX_CACHED_RESULTS = 200;
const sectionResultCache = new Map<
  string,
  { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }
>();

function hashString(input: string): string {
  return createHash('sha256').update(input).digest('hex').slice(0, 16);
}

function getSectionCacheKey(
  sourceId: string,
  section: ScrutinySectionKey,
  compactContext: string
): string {
  return `${sourceId}::${section}::${hashString(compactContext)}`;
}

function storeSectionCache(
  key: string,
  value: { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }
): void {
  if (sectionResultCache.has(key)) {
    sectionResultCache.delete(key);
  }
  sectionResultCache.set(key, value);
  while (sectionResultCache.size > MAX_CACHED_RESULTS) {
    const oldest = sectionResultCache.keys().next().value;
    if (oldest === undefined) break;
    sectionResultCache.delete(oldest);
  }
}

const SCRUTINY_SYSTEM_INSTRUCTION =
  'You are a Senior Credit Analyst at the Special Asset Management Division (SAMD), Head Office, Sammilito Islami Bank PLC (Union Bank PLC Unit), Dhaka. ' +
  'Your primary source document is the uploaded Branch Proposal (both its page images and extracted text). ' +
  'Scrutinize the Branch Proposal thoroughly and produce content strictly tailored to the requested IC Memo section. ' +
  'CRITICAL INTEGRITY RULE: Never invent, assume, or fabricate any facts, figures, dates, case numbers, valuations, names, observations, or recommendations not supported by the Branch Proposal (and verified memo figures provided in context). ' +
  'OUTPUT PURITY RULE: Output ONLY the official formal banking memorandum text/fields matching the section heading. NEVER add conversational remarks ("Here is...", "Based on the uploaded document..."), disclaimers, AI processing notes, scrutiny commentary, or markdown asterisks inside memo text. If required information is unavailable in the Branch Proposal, leave missing fields blank ("") and report what is missing strictly inside "insufficientReason" (or "summary" for structured tables).';

async function scrutinizeSingleSection(
  ai: GoogleGenAI,
  section: ScrutinySectionKey,
  source: CachedBranchProposalSource,
  context?: Record<string, unknown>,
  shouldStop?: () => boolean,
  bypassCache = false
): Promise<{ result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }> {
  const { prompt, schema } = getScrutinyConfig(section);
  const compactContext = buildCompactSectionContext(section, context);
  const cacheKey = getSectionCacheKey(source.sourceId, section, compactContext);

  if (!bypassCache) {
    const cached = sectionResultCache.get(cacheKey);
    if (cached) {
      return cached;
    }
  }

  const parts: Array<
    | { inlineData: { mimeType: string; data: string } }
    | { text: string }
  > = [
    ...source.imageParts,
    {
      text: `${prompt}

Uploaded Branch Proposal File: ${source.fileName}

Extracted Branch Proposal Text (may contain minor scanner OCR noise — cross-verify against attached page images when present):
${source.rawText}

Additional Verified IC Memo Context (from other uploaded documents & manual inputs):
${compactContext}`,
    },
  ];

  const response = await executeWithModelFallback(
    (model) =>
      ai.models.generateContent({
        model,
        contents: { parts },
        config: buildModelConfig(model, {
          systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
          responseMimeType: 'application/json',
          responseSchema: schema,
        }),
      }),
    shouldStop
  );

  const rawJson = (response.text ?? '{}').trim();
  const parsed = JSON.parse(rawJson) as Record<string, unknown>;
  const evaluated = evaluateSectionResult(section, parsed);
  storeSectionCache(cacheKey, evaluated);
  return evaluated;
}

/**
 * Incrementally extracts completed top-level `"key": { ... }` object entries from a streaming JSON root object.
 */
function extractCompletedTopLevelObjects(
  jsonBuffer: string,
  emittedKeys: Set<string>
): Array<{ key: string; value: Record<string, unknown> }> {
  const out: Array<{ key: string; value: Record<string, unknown> }> = [];
  const len = jsonBuffer.length;
  let i = 0;

  while (i < len && /\s/.test(jsonBuffer[i])) i++;
  if (i >= len || jsonBuffer[i] !== '{') return out;
  i++;

  while (i < len) {
    while (i < len && (/\s/.test(jsonBuffer[i]) || jsonBuffer[i] === ',')) i++;
    if (i >= len || jsonBuffer[i] === '}') break;

    if (jsonBuffer[i] !== '"') break;
    const keyStart = i;
    i++;
    let escaped = false;
    let keyClosed = false;
    while (i < len) {
      const ch = jsonBuffer[i];
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === '"') {
        keyClosed = true;
        i++;
        break;
      }
      i++;
    }
    if (!keyClosed) break;

    let key: string;
    try {
      key = JSON.parse(jsonBuffer.slice(keyStart, i)) as string;
    } catch {
      break;
    }

    while (i < len && /\s/.test(jsonBuffer[i])) i++;
    if (i >= len || jsonBuffer[i] !== ':') break;
    i++;
    while (i < len && /\s/.test(jsonBuffer[i])) i++;
    if (i >= len) break;

    if (jsonBuffer[i] !== '{') {
      break;
    }

    const valStart = i;
    let depth = 0;
    let inStr = false;
    let strEsc = false;
    let valEnd = -1;

    while (i < len) {
      const ch = jsonBuffer[i];
      if (inStr) {
        if (strEsc) {
          strEsc = false;
        } else if (ch === '\\') {
          strEsc = true;
        } else if (ch === '"') {
          inStr = false;
        }
      } else {
        if (ch === '"') {
          inStr = true;
        } else if (ch === '{') {
          depth++;
        } else if (ch === '}') {
          depth--;
          if (depth === 0) {
            valEnd = i;
            i++;
            break;
          }
        }
      }
      i++;
    }

    if (valEnd === -1) {
      break;
    }

    if (!emittedKeys.has(key)) {
      try {
        const parsedObj = JSON.parse(jsonBuffer.slice(valStart, valEnd + 1)) as Record<
          string,
          unknown
        >;
        emittedKeys.add(key);
        out.push({ key, value: parsedObj });
      } catch {
        // Defer to full JSON parse at stream end
      }
    }
  }

  return out;
}

function buildConsolidatedBatchSpec(sections: ScrutinySectionKey[]): {
  prompt: string;
  schema: Schema;
} {
  const properties: Record<string, Schema> = {};
  const promptBlocks: string[] = [];

  for (const sec of sections) {
    const cfg = getScrutinyConfig(sec);
    properties[sec] = cfg.schema;
    promptBlocks.push(`### SECTION KEY "${sec}":\n${cfg.prompt}`);
  }

  const prompt = `Scrutinize the uploaded Branch Proposal thoroughly once and populate ALL ${sections.length} requested IC Memo section keys below in a single structured JSON object.
Follow EVERY section's exact instructions and integrity rules:
- Never fabricate, assume, or invent any facts, dates, figures, or observations.
- Output ONLY clean formal banking text/data for each section (no conversational preambles, AI notes, or markdown asterisks).
- If information for a specific section is absent or insufficient in the Branch Proposal, keep missing fields blank ("") and state why strictly inside that section's "insufficientReason" or "summary" field.

${promptBlocks.join('\n\n')}`;

  return {
    prompt,
    schema: {
      type: Type.OBJECT,
      properties,
      required: sections,
    },
  };
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));

  // Pre-warm server-side Branch Proposal source cache so subsequent section/batch calls skip re-uploading base64 images
  app.post('/api/gemini/cache-source', (req, res) => {
    try {
      const { sourceId, branchProposalSource } = req.body as {
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
      };
      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }
      res.json({ cached: true, sourceId: resolved.source.sourceId });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to cache source.';
      res.status(500).json({ error: message });
    }
  });

  // Multimodal Branch Proposal Scrutiny for any individual IC Memo section (Justification .. Section 15)
  app.post('/api/gemini/scrutinize-section', async (req, res) => {
    try {
      const { section, sourceId, branchProposalSource, context, bypassCache } = req.body as {
        section: ScrutinySectionKey;
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
        bypassCache?: boolean;
      };

      if (!section || !VALID_SCRUTINY_SECTIONS.has(section)) {
        res.status(400).json({ error: 'Invalid or missing section key.' });
        return;
      }

      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }

      const ai = getAiClient();
      const { result, insufficient, insufficientReason } = await scrutinizeSingleSection(
        ai,
        section,
        resolved.source,
        context,
        undefined,
        Boolean(bypassCache)
      );
      res.json({ result, insufficient, insufficientReason, sourceId: resolved.source.sourceId });
    } catch (error) {
      console.error('Error scrutinizing IC Memo section:', error);
      const message = formatGeminiError(error, 'Failed to scrutinize section.');
      res.status(500).json({ error: message });
    }
  });

  // Consolidated Streaming ("Generate and Fill Up All Together") via SSE:
  // Executes all independent sections in 1 consolidated streaming Gemini call (emitting each section incrementally as its JSON property finishes)
  // followed by 1 SAMD synthesis call — reducing quota consumption by 8x-15x while preserving live progressive UI updates.
  app.post('/api/gemini/scrutinize-all', async (req, res) => {
    try {
      const { sections, sourceId, branchProposalSource, context } = req.body as {
        sections?: ScrutinySectionKey[];
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      const targetSections: ScrutinySectionKey[] = (
        Array.isArray(sections) && sections.length > 0
          ? sections
          : Array.from(VALID_SCRUTINY_SECTIONS)
      ).filter((s): s is ScrutinySectionKey => VALID_SCRUTINY_SECTIONS.has(s));

      if (targetSections.length === 0) {
        res.status(400).json({ error: 'No valid sections specified for batch scrutiny.' });
        return;
      }

      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }

      const ai = getAiClient();

      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('X-Accel-Buffering', 'no');
      res.flushHeaders?.();

      let clientClosed = false;
      res.on('close', () => {
        if (!res.writableEnded) {
          clientClosed = true;
        }
      });

      const writeEvent = (payload: Record<string, unknown>) => {
        if (clientClosed || res.writableEnded) return;
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
      };

      const includesSamd = targetSections.includes('samdObservation');
      const independentQueue = targetSections.filter((s) => s !== 'samdObservation');
      const samdContext: Record<string, unknown> = { ...(context ?? {}) };

      const feedSamdContext = (section: ScrutinySectionKey, result: Record<string, unknown>) => {
        if (section === 'justification' && result.text) {
          samdContext.justificationLines = String(result.text);
        } else if (section === 'sec8SourceOfRepayment' && result.text) {
          samdContext.sourceOfRepayment = String(result.text);
        } else if (section === 'clientPrayer' && result.text) {
          samdContext.clientPrayer = String(result.text);
        } else if (section === 'branchVisit' && result.text) {
          samdContext.branchVisitReport = String(result.text);
        } else if (section === 'sec13BranchRecommendation' && result.text) {
          samdContext.sec13BranchRecommendation = String(result.text);
        } else if (section === 'sec14IccdRecommendation' && result.text) {
          samdContext.sec14IccdRecommendation = String(result.text);
        }
      };

      const emitEvaluatedSection = (
        section: ScrutinySectionKey,
        evaluated: {
          result: Record<string, unknown>;
          insufficient: boolean;
          insufficientReason: string;
        }
      ) => {
        feedSamdContext(section, evaluated.result);
        writeEvent({
          type: 'section',
          section,
          result: evaluated.result,
          insufficient: evaluated.insufficient,
          insufficientReason: evaluated.insufficientReason,
        });
      };

      // Step 1: Serve any already-cached unchanged sections immediately
      const uncachedQueue: ScrutinySectionKey[] = [];
      for (const sec of independentQueue) {
        const compactCtx = buildCompactSectionContext(sec, context);
        const cacheKey = getSectionCacheKey(resolved.source.sourceId, sec, compactCtx);
        const cached = sectionResultCache.get(cacheKey);
        if (cached) {
          emitEvaluatedSection(sec, cached);
        } else {
          uncachedQueue.push(sec);
        }
      }

      // Step 2: Scrutinize uncached independent sections in 1 consolidated streaming request
      if (uncachedQueue.length === 1 && !clientClosed) {
        const sec = uncachedQueue[0];
        try {
          const evaluated = await scrutinizeSingleSection(
            ai,
            sec,
            resolved.source,
            context,
            () => clientClosed
          );
          emitEvaluatedSection(sec, evaluated);
        } catch (err) {
          if (!clientClosed) {
            writeEvent({
              type: 'error',
              section: sec,
              error: formatGeminiError(err, 'Failed to scrutinize section.'),
            });
          }
        }
      } else if (uncachedQueue.length > 1 && !clientClosed) {
        const emittedKeys = new Set<string>();
        let batchErrorMessage: string | null = null;
        let hitHardQuota = false;

        try {
          const { prompt, schema } = buildConsolidatedBatchSpec(uncachedQueue);
          const batchContext = JSON.stringify({
            placeholders: context?.placeholders ?? {},
            extractedSummaries: context?.extractedSummaries ?? {},
          });

          const parts: Array<
            | { inlineData: { mimeType: string; data: string } }
            | { text: string }
          > = [
            ...resolved.source.imageParts,
            {
              text: `${prompt}

Uploaded Branch Proposal File: ${resolved.source.fileName}

Extracted Branch Proposal Text (may contain minor scanner OCR noise — cross-verify against attached page images when present):
${resolved.source.rawText}

Additional Verified IC Memo Context (from other uploaded documents & manual inputs):
${batchContext}`,
            },
          ];

          const fullBuffer = await executeWithModelFallback(async (model) => {
            const stream = await ai.models.generateContentStream({
              model,
              contents: { parts },
              config: buildModelConfig(model, {
                systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
                responseMimeType: 'application/json',
                responseSchema: schema,
              }),
            });

            let acc = '';
            for await (const chunk of stream) {
              if (clientClosed) break;
              const delta = chunk.text ?? '';
              if (!delta) continue;
              acc += delta;

              const completed = extractCompletedTopLevelObjects(acc, emittedKeys);
              for (const item of completed) {
                const sec = item.key as ScrutinySectionKey;
                if (!uncachedQueue.includes(sec)) continue;
                const evaluated = evaluateSectionResult(sec, item.value);
                const cacheKey = getSectionCacheKey(
                  resolved.source.sourceId,
                  sec,
                  buildCompactSectionContext(sec, context)
                );
                storeSectionCache(cacheKey, evaluated);
                emitEvaluatedSection(sec, evaluated);
              }
            }
            return acc;
          }, () => clientClosed);

          if (!clientClosed && fullBuffer.trim()) {
            try {
              const finalParsed = JSON.parse(fullBuffer) as Record<
                string,
                Record<string, unknown>
              >;
              for (const sec of uncachedQueue) {
                if (emittedKeys.has(sec)) continue;
                const rawObj = finalParsed[sec];
                if (rawObj && typeof rawObj === 'object') {
                  emittedKeys.add(sec);
                  const evaluated = evaluateSectionResult(sec, rawObj);
                  const cacheKey = getSectionCacheKey(
                    resolved.source.sourceId,
                    sec,
                    buildCompactSectionContext(sec, context)
                  );
                  storeSectionCache(cacheKey, evaluated);
                  emitEvaluatedSection(sec, evaluated);
                }
              }
            } catch {
              // Handled below by missing-section recovery
            }
          }
        } catch (err) {
          const rawMsg = err instanceof Error ? err.message : String(err);
          hitHardQuota = isQuotaOrRateError(rawMsg);
          batchErrorMessage = formatGeminiError(err, 'Failed to scrutinize section.');
        }

        const remaining = uncachedQueue.filter((sec) => !emittedKeys.has(sec));
        if (remaining.length > 0 && !clientClosed) {
          if (hitHardQuota && batchErrorMessage) {
            for (const sec of remaining) {
              writeEvent({
                type: 'error',
                section: sec,
                error: batchErrorMessage,
              });
            }
          } else {
            for (const sec of remaining) {
              if (clientClosed) break;
              try {
                const evaluated = await scrutinizeSingleSection(
                  ai,
                  sec,
                  resolved.source,
                  context,
                  () => clientClosed
                );
                emitEvaluatedSection(sec, evaluated);
              } catch (secErr) {
                if (!clientClosed) {
                  writeEvent({
                    type: 'error',
                    section: sec,
                    error: formatGeminiError(
                      secErr,
                      batchErrorMessage || 'Failed to scrutinize section.'
                    ),
                  });
                }
              }
            }
          }
        }
      }

      // Step 3: SAMD Observation synthesis (uses freshly populated sections in samdContext)
      if (includesSamd && !clientClosed) {
        const samdCompactCtx = buildCompactSectionContext('samdObservation', samdContext);
        const samdCacheKey = getSectionCacheKey(
          resolved.source.sourceId,
          'samdObservation',
          samdCompactCtx
        );
        const cachedSamd = sectionResultCache.get(samdCacheKey);
        if (cachedSamd) {
          emitEvaluatedSection('samdObservation', cachedSamd);
        } else {
          try {
            const evaluated = await scrutinizeSingleSection(
              ai,
              'samdObservation',
              resolved.source,
              samdContext,
              () => clientClosed
            );
            emitEvaluatedSection('samdObservation', evaluated);
          } catch (err) {
            if (!clientClosed) {
              writeEvent({
                type: 'error',
                section: 'samdObservation',
                error: formatGeminiError(err, 'Failed to scrutinize SAMD section.'),
              });
            }
          }
        }
      }

      if (!clientClosed) {
        writeEvent({ type: 'done', sourceId: resolved.source.sourceId });
        res.end();
      }
    } catch (error) {
      console.error('Error during batch scrutiny:', error);
      const message = formatGeminiError(error, 'Batch scrutiny failed.');
      if (!res.headersSent) {
        res.status(500).json({ error: message });
      } else if (!res.writableEnded) {
        res.write(
          `data: ${JSON.stringify({
            type: 'fatal',
            error: message,
          })}\n\n`
        );
        res.end();
      }
    }
  });

  const VALID_STAGE_DOCS = new Set<StageMemoDocType>(['mec', 'board', 'sanction']);
  const STAGE_SECTIONS_MAP: Record<StageMemoDocType, StageScrutinySectionKey[]> = {
    mec: ['mecIcRecommendation', 'mecExecutiveSummary', 'mecSamdRecommendation'],
    board: ['boardMecRecommendation', 'boardProposalRationale', 'boardResolutionTerms'],
    sanction: ['sanctionBoardAuthority', 'sanctionCrossVerifiedSummary', 'sanctionSpecialConditions'],
  };

  function getStageSectionConfig(
    section: StageScrutinySectionKey,
    context?: Record<string, unknown>
  ): { prompt: string; schema: Schema } {
    const stageInputs = (context?.stageInputs ?? {}) as Record<string, string>;
    const icNo = (stageInputs.icMemoNumber || '').trim() || '[Chronologic Order of the IC Memo]';
    const icDt = (stageInputs.icMemoDate || '').trim() || '[IC Memo Date]';
    const mecNo = (stageInputs.mecMemoNumber || '').trim() || '[Chronologic Order of the MEC Memo]';
    const mecDt = (stageInputs.mecMemoDate || '').trim() || '[MEC Memo Date]';
    const boardNo = (stageInputs.boardMemoNumber || '').trim() || '[Chronologic Order of the Board Memo]';
    const boardDt = (stageInputs.boardMemoDate || '').trim() || '[Board Memo Date]';

    const icCentralRule = `PRIMARY SOURCE INSTRUCTION: Treat the finalized IC Memo data in the context ("icMemoFinalized", "placeholders", "icMemoAiSections", and carried sections) as the primary central data reference. Preserve 100% consistency with every client detail, facility mode, financial figure, security value, justification, observation, and recommendation established in the IC Memo — never alter, contradict, omit, or invent inconsistent figures. Adapt only the framing/terminology for this approval stage.`;

    const schema: Schema = {
      type: Type.OBJECT,
      properties: {
        text: strField('Clean official formal banking text for this memo section, or "" if source inputs are absent.'),
        insufficientReason: strField('Empty string "" if supported; otherwise concise UI note listing missing fields/inputs.'),
      },
      required: ['text', 'insufficientReason'],
    };

    switch (section) {
      case 'mecIcRecommendation':
        return {
          prompt: `${icCentralRule}
Write Section "Recommendation of the Investment Committee (IC)" for the Management Executive Committee (MEC) Memorandum.
You MUST explicitly incorporate the Chronologic Order of the IC Memo "${icNo}" and IC Memo Date "${icDt}" into the formal recommendation wording (e.g., referencing the Investment Committee's recommendation vide IC Memo (Chronologic Order: ${icNo}) dated ${icDt}), accurately stating the recommended Exit with Waiver facility under BRPD Circular Letter No. 23 dated 29.06.2026, NET Outstanding (#25#), Proposed Waiver (#4#), Settlement Amount (#27#), Down Payment (#35#), Settlement Period (#28#), Cost of Fund compliance (#31#, #44#), and placement before the MEC for kind consideration.
Output ONLY clean official banking text with no AI preambles or markdown asterisks.`,
          schema,
        };
      case 'mecExecutiveSummary':
        return {
          prompt: `${icCentralRule}
Scrutinize the finalized IC Memo (including Contents of the Client's Prayer, Justification, Liability Positions, CIB Status, Legal/Audit Status, Security Valuation MV/FSV, Source of Repayment, Branch Visit Report, Branch Recommendation, and SAMD Observation) alongside the Branch Proposal and write Section "Executive Appraisal & Waiver Rationale for MEC".
Write 2 concise, structured formal banking paragraphs synthesizing the borrower's distress, security coverage, Cost of Fund surplus/deficit vs maximum permissible waiver, and strategic recovery benefit for the Management Executive Committee strictly preserving all IC Memo facts and figures. Clean official prose only, no asterisks.`,
          schema,
        };
      case 'mecSamdRecommendation':
        return {
          prompt: `${icCentralRule}
Write Section "Observation and Recommendation of SAMD for MEC" placing this Exit with Waiver proposal before the Management Executive Committee after Investment Committee scrutiny (explicitly referencing Chronologic Order of the IC Memo "${icNo}" dated "${icDt}" if provided).
Write 3 to 4 crisp numbered formal points ("1. ...", "2. ...", "3. ...") summarizing verified liability & COF recovery protection carried forward from the IC Memo, compliance with special terms & conditions, and formal recommendation for MEC approval/onward placement before the Board. Clean official prose only.`,
          schema,
        };
      case 'boardMecRecommendation':
        return {
          prompt: `${icCentralRule}
Write Section "Recommendation of the Management Executive Committee (MEC)" for the Board of Directors Memorandum.
You MUST explicitly incorporate the Chronologic Order of the MEC Memo "${mecNo}" and MEC Memo Date "${mecDt}" (while preserving reference to Chronologic Order of the IC Memo "${icNo}" dated "${icDt}" where available), accurately summarizing the MEC's recommendation for Exit with Waiver under BRPD Circular Letter No. 23 dated 29.06.2026 covering Account (#2#), Branch (#3#), NET Outstanding (#25#), Proposed Waiver (#4#), Settlement Amount (#27#), Down Payment (#35#), and Settlement Period (#28#).
Clean formal banking prose only; no markdown asterisks or AI notes.`,
          schema,
        };
      case 'boardProposalRationale':
        return {
          prompt: `${icCentralRule}
Using the finalized IC Memo and carried-forward MEC Memo data (alongside the Branch Proposal), write Section "Board Memorandum Summary & Financial Justification" for submission to the Board of Directors.
Cover: (1) account classification, CIB/legal/regulatory verification established in the IC Memo, (2) financial recovery & Cost of Fund protection analysis (Principal #7#, Charges #8#, COF #31#, Breakup Total #32# vs Settlement Amount #27#, Max Waiver #44# vs Proposed Waiver #4#), and (3) why full and final settlement with waiver maximizes timely cash recovery over prolonged recovery proceedings. Clean formal banking prose only.`,
          schema,
        };
      case 'boardResolutionTerms':
        return {
          prompt: `${icCentralRule}
Write Section "Proposed Board Resolution & Approval Terms" for the Board Memorandum, framing a formal resolution for approval by the Board of Directors of Sammilito Islami Bank PLC for the proposed Exit with Waiver (Account #2#, Branch #3#, Settlement Amount #27# Million, Waiver #4# Million, Deadline #28#; referencing MEC Chronologic Order ${mecNo} dated ${mecDt}) subject to compliance with Special Terms & Conditions and Bangladesh Bank BRPD Circular Letter No. 23 dated 29.06.2026. Clean formal banking prose only.`,
          schema,
        };
      case 'sanctionBoardAuthority':
        return {
          prompt: `${icCentralRule}
Write Section "Board Approval Reference & Sanction Basis" for the Head Office Sanction Letter.
You MUST explicitly incorporate the Chronologic Order of the Board Memo "${boardNo}" and Board Memo Date "${boardDt}" (and preserve chain references to Chronologic Order of the MEC Memo "${mecNo}" dated "${mecDt}" and Chronologic Order of the IC Memo "${icNo}" dated "${icDt}" where available), formally conveying Head Office sanction for the approved Exit with Waiver facility for A/C ${'#2#'} of ${'#3#'} Branch under BRPD Circular Letter No. 23 dated 29.06.2026. Clean official Head Office Sanction Letter prose only.`,
          schema,
        };
      case 'sanctionCrossVerifiedSummary':
        return {
          prompt: `${icCentralRule}
Perform an AI cross-verification across ALL uploaded source documents (Outstanding, Party Performance, Overdue, Account Statement, CIB Report, Cost of Fund, Branch Proposal) and the chained IC -> MEC -> Board memo data to write Section "AI Cross-Verified Facility & Liability Confirmation" for the Sanction Letter.
Confirm exact verified figures in Million BDT (NET Outstanding #25#, Principal #7#, Charges #8#, Down Payment #35# in A/C #36#, Cost of Fund #31#, Breakup Total #32#, Proposed Waiver #4#, Settlement/Exit Amount #27#, Settlement Period #28#, Security MV #33# & FSV #34#) and confirm full consistency with the IC Memo, Branch Proposal, and Board sanction parameters. Clean official prose only; flag any missing verification input strictly inside "insufficientReason".`,
          schema,
        };
      case 'sanctionSpecialConditions':
        return {
          prompt: `${icCentralRule}
Write Section "Operative Terms, Conditions & Redemption Covenants" tailored for this Head Office Sanction Letter based on the verified IC Memo particulars, security position, legal/CIB status, and Board Memo (Chronologic Order: ${boardNo} dated ${boardDt}).
Write crisp numbered operative covenants ("i. ...", "ii. ...", etc.) covering adjustment of settlement within ${'#28#'}, determination of final up-to-date settlement balance at execution date, clearance of all direct/indirect/non-funded (L/C, BG, Credit Card) liabilities prior to collateral redemption, non-reduction of profit rate to 0.00% until full adjustment, legal/suit withdrawal cost recovery, and automatic cancellation on default. Clean official banking prose only.`,
          schema,
        };
    }
  }

  function resolveOptionalStageSource(
    sourceId?: string,
    incoming?: IncomingBranchProposalSource | null
  ): CachedBranchProposalSource | null {
    if (incoming) {
      const id = (incoming.sourceId || sourceId || `stage_${Date.now()}`).trim();
      const built = buildCachedSource(id, incoming);
      if (built && (incoming.sourceId || sourceId)) {
        storeCachedSource(id, built);
      }
      return built;
    }
    if (sourceId && branchProposalCache.has(sourceId.trim())) {
      return branchProposalCache.get(sourceId.trim()) ?? null;
    }
    return null;
  }

  async function scrutinizeSingleStageSection(
    ai: GoogleGenAI,
    section: StageScrutinySectionKey,
    source: CachedBranchProposalSource | null,
    context?: Record<string, unknown>,
    shouldStop?: () => boolean
  ): Promise<{ result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }> {
    const { prompt, schema } = getStageSectionConfig(section, context);
    const contextStr = JSON.stringify(context ?? {});
    const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${section}::${hashString(contextStr)}`;
    const cached = sectionResultCache.get(cacheKey);
    if (cached) return cached;

    const textPayload = `${prompt}

${source ? `Uploaded Branch Proposal File: ${source.fileName}\nExtracted Text:\n${source.rawText}\n` : ''}
Chained Memo & Verified Document Context (IC Memo -> MEC Memo -> Board Memo -> Sanction + Stage Inputs):
${contextStr}`;

    const parts: Array<
      | { inlineData: { mimeType: string; data: string } }
      | { text: string }
    > = source ? [...source.imageParts, { text: textPayload }] : [{ text: textPayload }];

    const response = await executeWithModelFallback(
      (model) =>
        ai.models.generateContent({
          model,
          contents: { parts },
          config: buildModelConfig(model, {
            systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
            responseMimeType: 'application/json',
            responseSchema: schema,
          }),
        }),
      shouldStop
    );

    const parsed = JSON.parse((response.text ?? '{}').trim()) as Record<string, unknown>;
    const rawText = String(parsed.text ?? '').trim();
    const cleanText = sanitizeAiMemoText(rawText);
    const modelReason = String(parsed.insufficientReason ?? '').trim();
    const insufficient = Boolean(modelReason) || cleanText === '';
    const evaluated = {
      result: { text: cleanText, insufficientReason: modelReason },
      insufficient,
      insufficientReason:
        modelReason || (!cleanText ? 'Required chained memo context or source data is unavailable.' : ''),
    };
    storeSectionCache(cacheKey, evaluated);
    return evaluated;
  }

  app.post('/api/gemini/scrutinize-stage-section', async (req, res) => {
    try {
      const { docType, section, sourceId, branchProposalSource, context } = req.body as {
        docType: StageMemoDocType;
        section: StageScrutinySectionKey;
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      if (!docType || !VALID_STAGE_DOCS.has(docType) || !STAGE_SECTIONS_MAP[docType].includes(section)) {
        res.status(400).json({ error: 'Invalid memo stage or section key.' });
        return;
      }

      const source = resolveOptionalStageSource(sourceId, branchProposalSource);
      const ai = getAiClient();
      const evaluated = await scrutinizeSingleStageSection(ai, section, source, context);
      res.json({
        result: evaluated.result,
        insufficient: evaluated.insufficient,
        insufficientReason: evaluated.insufficientReason,
        sourceId: source?.sourceId,
      });
    } catch (error) {
      console.error('Error scrutinizing stage memo section:', error);
      res.status(500).json({ error: formatGeminiError(error, 'Failed to generate stage section.') });
    }
  });

  app.post('/api/gemini/scrutinize-stage-all', async (req, res) => {
    try {
      const { docType, sections, sourceId, branchProposalSource, context } = req.body as {
        docType: StageMemoDocType;
        sections?: StageScrutinySectionKey[];
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      if (!docType || !VALID_STAGE_DOCS.has(docType)) {
        res.status(400).json({ error: 'Invalid stage memo type.' });
        return;
      }

      const allowed = STAGE_SECTIONS_MAP[docType];
      const targetSections = (
        Array.isArray(sections) && sections.length > 0 ? sections : allowed
      ).filter((s): s is StageScrutinySectionKey => allowed.includes(s));

      if (targetSections.length === 0) {
        res.status(400).json({ error: 'No valid stage sections requested.' });
        return;
      }

      const source = resolveOptionalStageSource(sourceId, branchProposalSource);
      const ai = getAiClient();

      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('X-Accel-Buffering', 'no');
      res.flushHeaders?.();

      let clientClosed = false;
      res.on('close', () => {
        if (!res.writableEnded) clientClosed = true;
      });

      const writeEvent = (payload: Record<string, unknown>) => {
        if (clientClosed || res.writableEnded) return;
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
      };

      const contextStr = JSON.stringify(context ?? {});
      const uncached: StageScrutinySectionKey[] = [];
      for (const sec of targetSections) {
        const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${sec}::${hashString(contextStr)}`;
        const cached = sectionResultCache.get(cacheKey);
        if (cached) {
          writeEvent({
            type: 'section',
            section: sec,
            result: cached.result,
            insufficient: cached.insufficient,
            insufficientReason: cached.insufficientReason,
          });
        } else {
          uncached.push(sec);
        }
      }

      if (uncached.length === 1 && !clientClosed) {
        const sec = uncached[0];
        try {
          const ev = await scrutinizeSingleStageSection(ai, sec, source, context, () => clientClosed);
          writeEvent({
            type: 'section',
            section: sec,
            result: ev.result,
            insufficient: ev.insufficient,
            insufficientReason: ev.insufficientReason,
          });
        } catch (err) {
          if (!clientClosed) {
            writeEvent({
              type: 'error',
              section: sec,
              error: formatGeminiError(err, 'Failed to generate stage section.'),
            });
          }
        }
      } else if (uncached.length > 1 && !clientClosed) {
        const properties: Record<string, Schema> = {};
        const promptBlocks: string[] = [];
        for (const sec of uncached) {
          const cfg = getStageSectionConfig(sec, context);
          properties[sec] = cfg.schema;
          promptBlocks.push(`### STAGE SECTION KEY "${sec}":\n${cfg.prompt}`);
        }
        const batchPrompt = `Generate ALL ${uncached.length} requested ${docType.toUpperCase()} document sections in one clean JSON object. Follow every integrity and formatting rule.\n\n${promptBlocks.join('\n\n')}`;
        const textPayload = `${batchPrompt}\n\n${source ? `Uploaded Branch Proposal File: ${source.fileName}\nExtracted Text:\n${source.rawText}\n` : ''}\nChained Memo & Verified Document Context:\n${contextStr}`;
        const parts: Array<
          | { inlineData: { mimeType: string; data: string } }
          | { text: string }
        > = source ? [...source.imageParts, { text: textPayload }] : [{ text: textPayload }];

        const emittedKeys = new Set<string>();
        let batchErr: string | null = null;
        try {
          const fullBuffer = await executeWithModelFallback(async (model) => {
            const stream = await ai.models.generateContentStream({
              model,
              contents: { parts },
              config: buildModelConfig(model, {
                systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
                responseMimeType: 'application/json',
                responseSchema: {
                  type: Type.OBJECT,
                  properties,
                  required: uncached,
                },
              }),
            });
            let acc = '';
            for await (const chunk of stream) {
              if (clientClosed) break;
              const delta = chunk.text ?? '';
              if (!delta) continue;
              acc += delta;
              const completed = extractCompletedTopLevelObjects(acc, emittedKeys);
              for (const item of completed) {
                const sec = item.key as StageScrutinySectionKey;
                if (!uncached.includes(sec)) continue;
                const cleanText = sanitizeAiMemoText(String(item.value.text ?? '').trim());
                const reason = String(item.value.insufficientReason ?? '').trim();
                const ev = {
                  result: { text: cleanText, insufficientReason: reason },
                  insufficient: Boolean(reason) || cleanText === '',
                  insufficientReason: reason,
                };
                const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${sec}::${hashString(contextStr)}`;
                storeSectionCache(cacheKey, ev);
                writeEvent({
                  type: 'section',
                  section: sec,
                  result: ev.result,
                  insufficient: ev.insufficient,
                  insufficientReason: ev.insufficientReason,
                });
              }
            }
            return acc;
          }, () => clientClosed);

          if (!clientClosed && fullBuffer.trim()) {
            try {
              const finalParsed = JSON.parse(fullBuffer) as Record<string, Record<string, unknown>>;
              for (const sec of uncached) {
                if (emittedKeys.has(sec)) continue;
                const obj = finalParsed[sec];
                if (obj && typeof obj === 'object') {
                  emittedKeys.add(sec);
                  const cleanText = sanitizeAiMemoText(String(obj.text ?? '').trim());
                  const reason = String(obj.insufficientReason ?? '').trim();
                  const ev = {
                    result: { text: cleanText, insufficientReason: reason },
                    insufficient: Boolean(reason) || cleanText === '',
                    insufficientReason: reason,
                  };
                  const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${sec}::${hashString(contextStr)}`;
                  storeSectionCache(cacheKey, ev);
                  writeEvent({
                    type: 'section',
                    section: sec,
                    result: ev.result,
                    insufficient: ev.insufficient,
                    insufficientReason: ev.insufficientReason,
                  });
                }
              }
            } catch {
              // Handled below
            }
          }
        } catch (err) {
          batchErr = formatGeminiError(err, 'Failed to generate stage memo sections.');
        }

        const remaining = uncached.filter((s) => !emittedKeys.has(s));
        for (const sec of remaining) {
          if (clientClosed) break;
          writeEvent({
            type: 'error',
            section: sec,
            error: batchErr || 'Section did not finish generating — click Retry.',
          });
        }
      }

      if (!clientClosed) {
        writeEvent({ type: 'done', sourceId: source?.sourceId });
        res.end();
      }
    } catch (error) {
      const message = formatGeminiError(error, 'Stage batch generation failed.');
      if (!res.headersSent) {
        res.status(500).json({ error: message });
      } else if (!res.writableEnded) {
        res.write(`data: ${JSON.stringify({ type: 'fatal', error: message })}\n\n`);
        res.end();
      }
    }
  });

  // Legacy AI "Generate" sections endpoint (preserved for backwards compatibility)
  app.post('/api/gemini/generate-section', async (req, res) => {
    try {
      const { section, context } = req.body as {
        section: 'justification' | 'branchVisit' | 'samdObservation';
        context: Record<string, unknown>;
      };

      if (!section || !context) {
        res.status(400).json({ error: 'Missing section type or memo context.' });
        return;
      }

      const ai = getAiClient();

      const systemInstruction =
        'You are a Senior Credit Analyst at the Special Asset Management Division (SAMD), Head Office, Sammilito Islami Bank PLC (Union Bank PLC Unit), Dhaka. Write formal, concise, objective banking memorandum prose strictly grounded in the provided document data and figures in Million BDT. Do not invent figures or facts not present in the context.';

      let userPrompt = '';

      if (section === 'justification') {
        userPrompt = `Scrutinize the IC Memo data (placeholders, settlement/waiver figures, Cost of Fund, security, classification, source of repayment) and all relevant details of the Branch Proposal below, and write exactly 3 key and relevant justification points for placeholder #51# under "Justification of The Proposed Facility".
Format as exactly 3 numbered justification points separated by newlines ("1. ...", "2. ...", "3. ..."), covering: (1) client's current business/financial position and genuine distress without willful default, (2) recovery & settlement math (down payment / principal / cost of fund coverage & protection of bank interest), and (3) immediate reduction of classified/NPL exposure and avoidance of prolonged litigation.
Keep each point crisp, factual, and substantive (1-2 sentences each). Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'branchVisit') {
        userPrompt = `Scrutinize all relevant sections of the Branch Proposal (including Section 12 Branch Visit Report, Section 10 Client Particulars, office/factory addresses, last visit date, stock position, fixed assets, business net worth, Section 7 security details, Section 11 financials, and Section 13/15/16 branch recommendations) and the IC Memo client particulars below, and write a comprehensive Section 12 "Branch Visit Report" for the Head Office IC Memo.
CRITICAL REQUIREMENT: Write from the perspective of the branch while strictly treating the branch in the THIRD PERSON (e.g., "The officials of [Branch Name] Branch...", "Branch management conducted an inspection...", "The Branch observed that...", never use "we", "us", "our", or "our branch").
Include the visiting officials' names/designations if available, visit date, client/concern name, premises/factory/collateral inspected, physical stock/asset & operational business findings, client interaction, and branch verification outcome accurately. Write 1 to 2 comprehensive, formal banking paragraphs. Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'samdObservation') {
        userPrompt = `Thoroughly scrutinize the IC Memo figures and sections (Principal, Charges, Charged Profit, Compensation, Uncharged Compensation, NET Outstanding, Cost of Fund Surplus/Deficit, Maximum Waiver Considering COF #44#, Down Payment, Proposed Waiver, Settlement/Exit Amount, Settlement Period, CIB & Classification status, Legal Action status, Audit/Regulatory status, Security MV/FSV, Source of Repayment, Branch Visit Report) and the relevant details of the Branch Proposal below, and write Section 15 "Observation and Recommendation of SAMD".
Write exactly five (5) key and substantive justification points separated by newlines ("1. ...", "2. ...", "3. ...", "4. ...", "5. ...") reflecting SAMD's observations and recommendations, addressing:
1. Verification of total outstanding liability breakup, classification/overdue/CIB status, and confirmation of absence of fraud, forgery, or willful default.
2. Assessment of primary & collateral security valuation (MV/FSV coverage), legal action status, and audit/regulatory compliance position.
3. Cost of Fund & waiver adequacy analysis (confirming down payment, recovery covering principal/charges/Cost of Fund, maximum permissible waiver considering COF, and conformity with BRPD Circular Letter No. 23 dated 29.06.2026).
4. Evaluation of client's genuine business/financial condition, branch visit & branch recommendation findings, and viability of the proposed source of repayment within the settlement deadline.
5. Formal SAMD recommendation to place the Exit with Waiver proposal before the Investment Committee for exit within the stipulated settlement period subject to compliance with special terms and conditions.
Each point should be 1-2 substantive, formal banking sentences. Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else {
        res.status(400).json({ error: 'Unknown section identifier.' });
        return;
      }

      const response = await executeWithModelFallback((model) =>
        ai.models.generateContent({
          model,
          contents: userPrompt,
          config: buildModelConfig(model, { systemInstruction }),
        })
      );

      const text = (response.text ?? '').trim();
      res.json({ text });
    } catch (error) {
      console.error('Error generating IC Memo section:', error);
      const message = formatGeminiError(error, 'Failed to generate section.');
      res.status(500).json({ error: message });
    }
  });

  const httpServer = http.createServer(app);

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        allowedHosts: true,
        ws: {
          server: httpServer,
        },
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`Memo Engine by Tuhin running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
