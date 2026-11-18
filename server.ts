import 'dotenv/config';
import express from 'express';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, ThinkingLevel, Type } from '@google/genai';

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

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));

  // Multimodal extraction for scanned Branch Proposal PDF pages
  app.post('/api/gemini/extract-branch-proposal', async (req, res) => {
    try {
      const { pages } = req.body as { pages?: Array<{ mimeType: string; data: string }> };
      if (!pages || !Array.isArray(pages) || pages.length === 0) {
        res.status(400).json({ error: 'No rendered PDF pages provided for extraction.' });
        return;
      }

      const ai = getAiClient();

      const imageParts = pages.map((page) => ({
        inlineData: {
          mimeType: page.mimeType || 'image/jpeg',
          data: page.data,
        },
      }));

      const promptText = `You are an expert Bangladeshi banking credit analyst reading a scanned, noisy "Exit with Waiver" Branch Proposal from Sammilito Islami Bank PLC (Union Bank PLC Unit).
Carefully inspect all pages (including stamps, handwritten corrections, and tables) and extract the exact text and figures.

CRITICAL RULES:
1. All figures in the Branch Proposal are ALREADY in Million BDT. Copy figures as written (do NOT divide or convert units).
2. Read handwritten reference numbers and dates carefully (for example: Ref no like UB/PATIYA/INV/2026/487, Proposal Date like 18/11/2026).
3. For tables in sections 3, 4, 5, 6, 7, 9, 10, 11, extract every non-empty row accurately. If a section is marked "N/A" or "Nil", capture that status and still extract any filled table rows.
4. For Security Value in Remarks of the main table (item 3), extract MV (e.g. "N/A" or amount) and FSV (if blank, return "").
5. Return strictly valid JSON matching the schema.`;

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: {
          parts: [...imageParts, { text: promptText }],
        },
        config: {
          thinkingConfig: { thinkingLevel: ThinkingLevel.MEDIUM },
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              referenceNo: { type: Type.STRING, description: 'Proposal reference number on page 1 top-left, e.g. UB/PATIYA/INV/2026/487 (#64#)' },
              proposalDate: { type: Type.STRING, description: 'Proposal date on page 1 top-left, e.g. 18/11/2026 (#65#)' },
              branchName: { type: Type.STRING, description: 'Branch name shown in header or A/C bar' },
              accountNo: { type: Type.STRING, description: 'Investment A/C number' },
              subjectText: { type: Type.STRING, description: 'Full subject paragraph text' },
              limitOriginal: { type: Type.STRING, description: 'Original Limit from the main settlement table (#6#), copied as written' },
              limitRSD: { type: Type.STRING, description: 'RSD limit if present, else empty or -' },
              securityMV: { type: Type.STRING, description: 'Security Value MV from Remarks item 3 (#33#), e.g. N/A' },
              securityFSV: { type: Type.STRING, description: 'Security Value FSV from Remarks item 3 (#34#)' },
              justificationText: { type: Type.STRING, description: 'Full text of Justification of The Proposed Facility on page 1' },
              previousBoardDecisions: { type: Type.STRING, description: 'Previous Board Decision(s) line if visible' },
              natureOfBusiness: { type: Type.STRING, description: 'Section 1: Nature of Business (#54#)' },
              sec2Limit: { type: Type.STRING, description: 'Section 2 table Limit column (#56#)' },
              sec2Validity: { type: Type.STRING, description: 'Section 2 table Validity column (#58#)' },
              sec2Overdue: { type: Type.STRING, description: 'Section 2 table Overdue column if present' },
              sec2Recovery: { type: Type.STRING, description: 'Section 2 table Recovery column' },
              sec2Remarks: { type: Type.STRING, description: 'Section 2 table Remarks column' },
              section3AsOnDate: { type: Type.STRING, description: 'Date in Section 3 heading if present' },
              section3Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: { type: Type.STRING },
                    nameOfConcern: { type: Type.STRING },
                    bankOrNbfi: { type: Type.STRING },
                    natureOfFacility: { type: Type.STRING },
                    limit: { type: Type.STRING },
                    outstanding: { type: Type.STRING },
                    overdueEol: { type: Type.STRING },
                    expiryDate: { type: Type.STRING },
                    clStatus: { type: Type.STRING },
                  },
                },
              },
              section4CibDate: { type: Type.STRING, description: 'Date in Section 4 CIB Status heading' },
              section4Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    status: { type: Type.STRING },
                    unitBorrower: { type: Type.STRING },
                    unitGuarantor: { type: Type.STRING },
                    otherBorrower: { type: Type.STRING },
                    otherGuarantor: { type: Type.STRING },
                    remarks: { type: Type.STRING },
                  },
                },
              },
              section5StatusNote: { type: Type.STRING },
              section5Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: { type: Type.STRING },
                    typeOfCase: { type: Type.STRING },
                    caseNo: { type: Type.STRING },
                    filingDate: { type: Type.STRING },
                    amount: { type: Type.STRING },
                    previousDate: { type: Type.STRING },
                    nextDate: { type: Type.STRING },
                    presentStatus: { type: Type.STRING },
                  },
                },
              },
              section6StatusNote: { type: Type.STRING, description: 'e.g. N/A if written in Section 6 heading' },
              section6Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: { type: Type.STRING },
                    authority: { type: Type.STRING },
                    lastInspectionDate: { type: Type.STRING },
                    objectionStatus: { type: Type.STRING },
                    objections: { type: Type.STRING },
                    branchReply: { type: Type.STRING },
                  },
                },
              },
              section7PrimaryNote: { type: Type.STRING },
              section7PrimaryRows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    particulars: { type: Type.STRING },
                    value: { type: Type.STRING },
                  },
                },
              },
              section7PrimaryTotal: { type: Type.STRING },
              section7CollateralNote: { type: Type.STRING, description: 'e.g. N/A' },
              section7CollateralSchedules: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
              },
              section7ValuationRows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    schedule: { type: Type.STRING },
                    origBranchMv: { type: Type.STRING },
                    origBranchFsv: { type: Type.STRING },
                    origSurvMv: { type: Type.STRING },
                    origSurvFsv: { type: Type.STRING },
                    presBranchMv: { type: Type.STRING },
                    presBranchFsv: { type: Type.STRING },
                    presSurvMv: { type: Type.STRING },
                    presSurvFsv: { type: Type.STRING },
                    presAvMv: { type: Type.STRING },
                    presAvFsv: { type: Type.STRING },
                  },
                },
              },
              section7CoveragePct: { type: Type.STRING },
              section7Others: { type: Type.STRING },
              section8SourceOfRepayment: { type: Type.STRING },
              section9Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: { type: Type.STRING },
                    referenceNo: { type: Type.STRING },
                    date: { type: Type.STRING },
                    limit: { type: Type.STRING },
                    remarks: { type: Type.STRING },
                  },
                },
              },
              section10Particulars: {
                type: Type.OBJECT,
                properties: {
                  nameOfClient: { type: Type.STRING },
                  uniqueCustomerId: { type: Type.STRING },
                  awcdMsdAccount: { type: Type.STRING },
                  businessTypeNature: { type: Type.STRING },
                  businessManagementType: { type: Type.STRING },
                  clientType: { type: Type.STRING },
                  keyPersonPhone: { type: Type.STRING },
                  address: { type: Type.STRING },
                  officeBusinessAddress: { type: Type.STRING },
                  factoryAddress: { type: Type.STRING },
                  purposeOfInvestment: { type: Type.STRING },
                  limitOfInvestment: { type: Type.STRING },
                  modeOfInvestment: { type: Type.STRING },
                  lastVisitDate: { type: Type.STRING },
                  stockPosition: { type: Type.STRING },
                  fixedAsset: { type: Type.STRING },
                  businessNetWorth: { type: Type.STRING },
                  insuranceCoverage: { type: Type.STRING },
                  tradeLicenseValidity: { type: Type.STRING },
                  otherRegulatoryDocuments: { type: Type.STRING },
                  dateOfOriginalInvestment: { type: Type.STRING },
                  firstDisbursementManager: { type: Type.STRING },
                  firstDateOfClassification: { type: Type.STRING },
                },
              },
              section10Directors: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING },
                    age: { type: Type.STRING },
                    education: { type: Type.STRING },
                    designation: { type: Type.STRING },
                    sharePct: { type: Type.STRING },
                    experience: { type: Type.STRING },
                    netWorth: { type: Type.STRING },
                  },
                },
              },
              section10BankDirectors: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING },
                    bankNbfi: { type: Type.STRING },
                    designation: { type: Type.STRING },
                    sharePct: { type: Type.STRING },
                    numShares: { type: Type.STRING },
                  },
                },
              },
              section11FinancialsNote: { type: Type.STRING },
              section11Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: { type: Type.STRING },
                    grossSales: { type: Type.STRING },
                    grossProfit: { type: Type.STRING },
                    netProfit: { type: Type.STRING },
                    totalAsset: { type: Type.STRING },
                    totalLiabilities: { type: Type.STRING },
                    netWorth: { type: Type.STRING },
                  },
                },
              },
              section12BranchVisitReport: { type: Type.STRING, description: 'Section 12 Branch Visit Report text from page 4' },
              section15BranchBicRec: { type: Type.STRING, description: 'Recommendation of Branch Investment Committee from page 4' },
              section16BranchRec: { type: Type.STRING, description: 'Branch Recommendation from page 4' },
            },
          },
        },
      });

      const rawText = response.text ?? '{}';
      const parsed = JSON.parse(rawText);
      res.json({ data: parsed });
    } catch (error) {
      console.error('Error extracting Branch Proposal:', error);
      const message = error instanceof Error ? error.message : 'Failed to extract Branch Proposal.';
      res.status(500).json({ error: message });
    }
  });

  // AI "Generate" sections in IC Memo
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

      let systemInstruction =
        'You are a Senior Credit Analyst at the Special Asset Management Division (SAMD), Head Office, Sammilito Islami Bank PLC (Union Bank PLC Unit), Dhaka. Write formal, concise, objective banking memorandum prose strictly grounded in the provided document data and figures in Million BDT. Do not invent figures or facts not present in the context.';

      let userPrompt = '';

      if (section === 'justification') {
        userPrompt = `Scrutinize the Branch Proposal and the IC Memo data below, and write exactly 3 main justification lines for placeholder #51# under "Justification of The Proposed Facility".
Format as 3 concise, formal banking lines separated by newlines (e.g., "1. ...", "2. ...", "3. ..."), covering: (1) client's current business/financial position and genuine distress without willful default, (2) recovery & settlement math (down payment / principal / cost of fund coverage), and (3) reduction of classified/NPL exposure and protection of bank interest.
Keep each line crisp (1-2 sentences). Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'branchVisit') {
        userPrompt = `Scrutinize the Branch Proposal Section 12 (Branch Visit Report) and client particulars below, and write Section 12 "Branch Visit Report" for the Head Office IC Memo.
CRITICAL REQUIREMENT: Refer to the Branch in the THIRD PERSON (e.g., "The officials of Patiya Branch...", "Branch management visited...", never "we" or "our branch").
Include the visiting officials' names/designations if available, visit date, client name, premises inspected, and physical business findings accurately. Write 1 concise, formal paragraph. Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'samdObservation') {
        userPrompt = `Scrutinize the IC Memo figures (Principal, Charges, Cost of Fund, Surplus/Deficit, Down Payment, Waiver, Settlement Amount, CIB & Classification status) and the Branch Proposal below, and write Section 15 "Observation and Recommendation of SAMD".
Write exactly 3 main observation/recommendation lines separated by newlines (e.g., "1. ...", "2. ...", "3. ..."), addressing: (1) verification of outstanding liability, classification status, and absence of collateral/legal impediments, (2) Cost of Fund & waiver adequacy analysis (confirming recovery covers principal and cost of fund / BRPD Circular compliance), and (3) SAMD's formal recommendation to place before the Investment Committee for exit with waiver within the stipulated settlement period.
Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else {
        res.status(400).json({ error: 'Unknown section identifier.' });
        return;
      }

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: userPrompt,
        config: {
          systemInstruction,
          thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
        },
      });

      const text = (response.text ?? '').trim();
      res.json({ text });
    } catch (error) {
      console.error('Error generating IC Memo section:', error);
      const message = error instanceof Error ? error.message : 'Failed to generate section.';
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
