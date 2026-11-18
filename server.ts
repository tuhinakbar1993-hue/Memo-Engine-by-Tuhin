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
