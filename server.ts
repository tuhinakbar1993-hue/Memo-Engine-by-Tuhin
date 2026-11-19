import 'dotenv/config';
import express from 'express';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, ThinkingLevel } from '@google/genai';

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
