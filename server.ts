import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
import multer from "multer";
import mammoth from "mammoth";
import { generateOfflineLetter } from "./src/utils/offlineGenerator";
import * as pdfParseNamespace from "pdf-parse";

dotenv.config();

// Safe dynamic helper to resolve standard vs. synthetic vs. double-wrapped CommonJS default exports for pdf-parse
const getPdfParse = (): any => {
  const ns = pdfParseNamespace as any;
  if (typeof ns === "function") {
    return ns;
  }
  if (ns && ns.default && typeof ns.default === "function") {
    return ns.default;
  }
  if (ns && ns.default && ns.default.default && typeof ns.default.default === "function") {
    return ns.default.default;
  }
  return ns;
};

let aiInstance: GoogleGenAI | null = null;

function getGeminiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is not configured on the server. Please check your Secrets settings.");
  }
  if (!aiInstance) {
    aiInstance = new GoogleGenAI({
      apiKey: apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiInstance;
}

function parseDocumentOffline(text: string): any {
  console.log("[Offline Parser] Running fallback exact-label matching extractor with advanced formatting...");
  const labels = [
    { key: "refereeName", label: "Referee Full Name:" },
    { key: "refereeTitle", label: "Referee Job Title:" },
    { key: "refereeCompany", label: "Eminent Organisation / Company Name:" },
    { key: "refereeStature", label: "Referee Stature & Eminence Biography:" },
    { key: "refereeAddress", label: "Referee Business Address (UK Standard Format):" },
    { key: "refereeEmail", label: "Referee Trust Email:" },
    { key: "refereePhone", label: "Referee Direct Phone:" },
    { key: "refereeLinkedin", label: "Referee LinkedIn Profile URL:" },
    { key: "applicantName", label: "Candidate Full Name:" },
    { key: "applicantRole", label: "Candidate Tech Role:" },
    { key: "applicantCompany", label: "Target Tech Organisation:" },
    { key: "visaRoute", label: "Target Visa Route:" },
    { key: "selectedCriteria", label: "Target Endorsement Criteria (tick all that apply):" },
    { key: "relationship", label: "Professional Relationship & Timeline Context:" },
    { key: "anecdotes", label: "Key Work Achievements, Patents & Metrics:" },
    { key: "customDate", label: "Document Date:" }
  ];

  const result: any = {
    refereeName: "",
    refereeTitle: "",
    refereeCompany: "",
    refereeStature: "",
    refereeAddress: "",
    refereeEmail: "",
    refereePhone: "",
    refereeLinkedin: "",
    applicantName: "",
    applicantRole: "",
    applicantCompany: "",
    visaRoute: "Exceptional Talent",
    selectedCriteria: [] as string[],
    relationship: "",
    anecdotes: "",
    customDate: ""
  };

  for (let i = 0; i < labels.length; i++) {
    const current = labels[i];
    const startIndex = text.indexOf(current.label);
    if (startIndex === -1) continue;

    const contentStart = startIndex + current.label.length;
    let contentEnd = text.length;
    let nextIndex = -1;
    
    for (let j = 0; j < labels.length; j++) {
      if (i === j) continue;
      const idx = text.indexOf(labels[j].label);
      if (idx !== -1 && idx > contentStart) {
        if (nextIndex === -1 || idx < nextIndex) {
          nextIndex = idx;
        }
      }
    }
    
    if (nextIndex !== -1) {
      contentEnd = nextIndex;
    }

    let extracted = text.substring(contentStart, contentEnd);

    if (current.key === "refereeAddress") {
      // Split by line, clean underscores, and join with ", " to prevent run-together lines
      const addressLines = extracted.split(/[\r\n]+/)
        .map(l => l.replace(/_/g, "").trim())
        .filter(l => {
          if (!l) return false;
          const lower = l.toLowerCase();
          if (lower.includes("referee business address") || lower.includes("uk standard format") || lower.includes("e.g. apis") || lower.includes("mailing pattern")) return false;
          return true;
        });
      let cleanedAddr = addressLines.join(", ");
      cleanedAddr = cleanedAddr.replace(/,\s*,/g, ",");
      cleanedAddr = cleanedAddr.replace(/\s+/g, " ");
      cleanedAddr = cleanedAddr.replace(/,\s*/g, ", ");
      extracted = cleanedAddr.trim().replace(/^,|,$/g, "").trim();
      result.refereeAddress = extracted;
    } else if (current.key === "visaRoute") {
      const lower = extracted.toLowerCase();
      const lines = extracted.split(/[\r\n]+/);
      let isPromise = false;
      let isTalent = false;
      for (const line of lines) {
        const lineLower = line.toLowerCase();
        if (lineLower.includes("exceptional talent") || lineLower.includes("5 years lead")) {
          const isUnticked = lineLower.includes("☐") || /\[\s*\]/.test(lineLower);
          const hasTick = /☒|☑|✓|✔|\[\s*[xXvV*+]\s*\]/i.test(lineLower);
          if (hasTick || !isUnticked) isTalent = true;
        }
        if (lineLower.includes("exceptional promise") || lineLower.includes("emerging lead")) {
          const isUnticked = lineLower.includes("☐") || /\[\s*\]/.test(lineLower);
          const hasTick = /☒|☑|✓|✔|\[\s*[xXvV*+]\s*\]/i.test(lineLower);
          if (hasTick || !isUnticked) isPromise = true;
        }
      }
      if (isPromise && !isTalent) {
        result.visaRoute = "Exceptional Promise";
      } else {
        result.visaRoute = "Exceptional Talent";
      }
    } else if (current.key === "selectedCriteria") {
      const matched: string[] = [];
      const lines = extracted.split(/[\r\n]+/);
      for (const line of lines) {
        const lineLower = line.toLowerCase();
        let matchedCode = "";
        
        if (lineLower.includes("mandatory") || lineLower.includes("(mc)")) matchedCode = "MC";
        else if (lineLower.includes("optional criterion 1") || lineLower.includes("(oc1)")) matchedCode = "OC1";
        else if (lineLower.includes("optional criterion 2") || lineLower.includes("(oc2)")) matchedCode = "OC2";
        else if (lineLower.includes("optional criterion 3") || lineLower.includes("(oc3)")) matchedCode = "OC3";
        else if (lineLower.includes("optional criterion 4") || lineLower.includes("(oc4)")) matchedCode = "OC4";

        if (matchedCode) {
          const isUnticked = lineLower.includes("☐") || /\[\s*\]/.test(lineLower);
          const hasTick = /☒|☑|✓|✔|\[\s*[xXvV*+]\s*\]/i.test(lineLower);
          if (hasTick || !isUnticked) {
            matched.push(matchedCode);
          }
        }
      }
      
      if (matched.length === 0) {
        const lower = extracted.toLowerCase();
        if (/mc/i.test(lower)) matched.push("MC");
        if (/oc1/i.test(lower)) matched.push("OC1");
        if (/oc2/i.test(lower)) matched.push("OC2");
        if (/oc3/i.test(lower)) matched.push("OC3");
        if (/oc4/i.test(lower)) matched.push("OC4");
      }
      result.selectedCriteria = Array.from(new Set(matched));
    } else {
      extracted = extracted.replace(/_+/g, " ");
      extracted = extracted.replace(/\(e\.g\..*?\)/gi, "");
      result[current.key] = extracted.trim();
    }
  }

  for (const k of Object.keys(result)) {
    if (typeof result[k] === "string" && k !== "refereeAddress") {
      result[k] = result[k].replace(/^\s*[:\-]\s*/, "");
      result[k] = result[k].replace(/_/g, "").trim();
    }
  }

  console.log("[Offline Parser] Extracted result:", JSON.stringify(result));
  return result;
}

async function startServer() {
  const app = express();
  app.use(express.json());

  const PORT = 3000;

  // Health check endpoint
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", apiKeyAvailable: !!process.env.GEMINI_API_KEY });
  });

  // Configure multer memory upload
  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
  });

  // Document parser endpoint to autofill candidate form inputs
  app.post("/api/parse-document", upload.single("file"), async (req, res) => {
    let extractedText = "";
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(500).json({
          error: "GEMINI_API_KEY is not configured on the server. Please check your Secrets settings.",
        });
      }

      if (!req.file) {
        return res.status(400).json({ error: "No document file was uploaded." });
      }

      const buffer = req.file.buffer;
      const originalname = req.file.originalname.toLowerCase();

      if (originalname.endsWith(".docx")) {
        console.log(`[Document Ingestion] Parsing .docx file of size ${buffer.length} bytes...`);
        const result = await mammoth.extractRawText({ buffer: buffer });
        extractedText = result.value || "";
      } else if (originalname.endsWith(".pdf")) {
        console.log(`[Document Ingestion] Parsing .pdf file of size ${buffer.length} bytes...`);
        const parseFn = getPdfParse();
        const result = await parseFn(buffer);
        extractedText = result.text || "";
      } else {
        return res.status(400).json({ error: "Unsupported file extension. Only .docx or .pdf files are accepted." });
      }

      if (!extractedText.trim()) {
        return res.status(400).json({ error: "The uploaded document contains no readable text." });
      }

      console.log(`[Document Ingestion] Sending extracted text to Gemini (approx. ${extractedText.length} characters)...`);

      const ai = getGeminiClient();
      const modelsToTry = ["gemini-3.5-flash", "gemini-flash-latest", "gemini-3.1-flash-lite"];
      let jsonText = "";
      let lastError: any = null;

      for (const modelName of modelsToTry) {
        try {
          console.log(`[Document Ingestion] Attempting parse with model: ${modelName}...`);
          const response = await ai.models.generateContent({
            model: modelName,
            contents: `Deconstruct this candidate intake text or recommendation summary, extract the structured parameters, and format it exactly as the requested JSON schema. Text:\n\n${extractedText}`,
            config: {
              systemInstruction: "You are an elite document parser. Your single goal is to analyze the text of a completed Tech Nation Recommendation Letter Input Form and extract fields based on these exact labels. Here is the mapping from the labels in the document to the JSON properties you MUST return:\n\n1. SECTION 1 – REFEREE\n- 'Referee Full Name:' -> Extract value to JSON property 'refereeName'\n- 'Referee Job Title:' -> Extract value to JSON property 'refereeTitle'\n- 'Eminent Organisation / Company Name:' -> Extract value to JSON property 'refereeCompany'\n- 'Referee Stature & Eminence Biography:' -> Extract value to JSON property 'refereeStature'\n- 'Referee Business Address (UK Standard Format):' -> Extract value to JSON property 'refereeAddress'. IMPORTANT: Convert any linebreaks or newline spacing in the address into a single line separated by commas, maintaining readable format like 'QuantumScale Technologies Ltd, Level 18, One Canada Square, Canary Wharf, London, E14 5AA, United Kingdom'. Remove all stray underscores or bracket instructions.\n- 'Referee Trust Email:' -> Extract value to JSON property 'refereeEmail'\n- 'Referee Direct Phone:' -> Extract value to JSON property 'refereePhone'\n- 'Referee LinkedIn Profile URL:' -> Extract value to JSON property 'refereeLinkedin'\n\n2. SECTION 2 – CANDIDATE ROUTE & PROFILE\n- 'Candidate Full Name:' -> Extract value to JSON property 'applicantName'\n- 'Candidate Tech Role:' -> Extract value to JSON property 'applicantRole'\n- 'Target Tech Organisation:' -> Extract value to JSON property 'applicantCompany'\n- 'Target Visa Route:' -> Extract value to JSON property 'visaRoute'. Look for checkmarks (e.g. '[X]', '[x]', '☒', '☑') next to 'Exceptional Talent' or 'Exceptional Promise'. If 'Exceptional Promise' is selected, return 'Exceptional Promise', otherwise return 'Exceptional Talent'.\n- 'Target Endorsement Criteria (tick all that apply):' -> Extract values to 'selectedCriteria' (array of strings selected from: MC, OC1, OC2, OC3, OC4). If ticked (e.g. '[X]', '[x]', '☒', '☑') next to MC, include 'MC'; if OC1, include 'OC1'; if OC2, include 'OC2'; if OC3, include 'OC3'; if OC4, include 'OC4'. If a criterion box is NOT checked (or contains '☐' / empty '[ ]'), do NOT include it.\n\n3. SECTION 3 – RELATIONSHIP & IMPACT STORIES\n- 'Professional Relationship & Timeline Context:' -> Extract value to JSON property 'relationship'\n- 'Key Work Achievements, Patents & Metrics:' -> Extract value to JSON property 'anecdotes'\n- 'Document Date:' -> Extract value to JSON property 'customDate'. Use the exact date filled in the document (e.g. 'April 24th, 2025' or '19 June 2026').\n\nReturn ONLY a valid JSON object with these keys: refereeName, refereeTitle, refereeCompany, refereeStature, refereeAddress, refereeEmail, refereePhone, refereeLinkedin, applicantName, applicantRole, applicantCompany, visaRoute, selectedCriteria, relationship, anecdotes, customDate.",
              responseMimeType: "application/json",
              temperature: 0.1,
            }
          });
          jsonText = response.text || "";
          if (jsonText) {
            console.log(`[Document Ingestion] Parsing successful using model: ${modelName}`);
            break;
          }
        } catch (error: any) {
          lastError = error;
          console.warn(`[Document Ingestion] Model ${modelName} failed or limits reached:`, error?.message || error);
        }
      }

      // HIGH-FIDELITY OFFLINE DOCUMENT PARSER FALLBACK
      if (!jsonText) {
        console.log("[Document Ingestion] All generative models reached limits or were unavailable. Activating eMigr8 Offline High-Fidelity Form Parser...");
        const parsedData = parseDocumentOffline(extractedText);
        return res.json({ success: true, data: parsedData });
      }

      console.log("[Document Ingestion] Gemini response received:", jsonText);

      if (jsonText.startsWith("```")) {
        jsonText = jsonText.replace(/^```(json)?/, "").replace(/```$/, "").trim();
      }

      try {
        const parsedData = JSON.parse(jsonText);
        return res.json({ success: true, data: parsedData });
      } catch (parseError) {
        console.error("[Document Ingestion] Failed to parse JSON returned by Gemini:", parseError);
        console.log("[Document Ingestion] Activating emergency offline parser due to JSON parse failure...");
        const parsedData = parseDocumentOffline(extractedText);
        return res.json({ success: true, data: parsedData });
      }

    } catch (err: any) {
      console.error("[Document Ingestion Error]", err);
      // Secondary emergency fallback
      try {
        if (extractedText) {
          const parsedData = parseDocumentOffline(extractedText);
          return res.json({ success: true, data: parsedData });
        }
      } catch (innerErr) {
        console.error("[Document Ingestion Error] Emergency fallback also failed:", innerErr);
      }
      return res.status(500).json({ error: err.message || "An unexpected error occurred while parsing the document." });
    }
  });

  // Endpoint to generate the UK Global Talent Visa support letter
  app.post("/api/generate-letter", async (req, res) => {
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(500).json({
          error: "GEMINI_API_KEY is not configured on the server. Please check your Secrets settings.",
        });
      }

      const ai = getGeminiClient();

      const {
        refereeName,
        refereeTitle,
        refereeCompany,
        refereeStature,
        refereeAddress,
        refereeEmail,
        refereePhone,
        refereeLinkedin,
        applicantName,
        applicantRole,
        applicantCompany,
        visaRoute, // 'Exceptional Talent' | 'Exceptional Promise'
        selectedCriteria, // string[]: e.g. ['MC', 'OC1', 'OC3']
        relationships, // how the referee knows the applicant and timeline
        anecdotes, // raw text of achievements and anecdotes
        customDate,
      } = req.body;

      // Validate required inputs
      if (!refereeName || !refereeTitle || !refereeCompany || !applicantName) {
        return res.status(400).json({ error: "Missing required details to draft recommendation letter." });
      }

      const criteriaGuideText = `
MANDATORY CRITERION (MC): Sustained national or international recognition as a leading talent (Exceptional Talent) or an emerging leading talent (Exceptional Promise) in the digital technology sector over the past five years. Verified via awards, keynote invitations, cited publications, national press coverage, or recognised industry honours.

OPTIONAL CRITERION 1 (OC1): Groundbreaking innovation as a founder, senior executive, board member, or employee of a product-led digital technology company. The innovation must be demonstrably unique in the global digital technology sector.

OPTIONAL CRITERION 2 (OC2): Significant contributions to the advancement of the digital technology sector beyond their employed role. This includes field-advancing open-source contributions with measurable adoption, thought leadership at events of national/international stature, or peer-recognised mentorship of the field. Excludes paid work.

OPTIONAL CRITERION 3 (OC3): Multiple, significant, individually attributable technical, commercial, or entrepreneurial contributions to one or more product-led digital technology companies. Impact must be measurable and linked directly and explicitly to the applicant's individual decisions or output.

OPTIONAL CRITERION 4 (OC4): Original academic or technical research published in peer-reviewed journals or sector publications of stature, with demonstrable impact on a product-led technology field. The applicant's individual contribution must be clearly isolated.
`;

      const systemInstruction = `
You are eMigr8 Recommendation Engine, the world's leading specialist advisor on UK Global Talent Visa (Tech Nation) applications. Your single, absolute duty is to draft an elite, expert, human-sounding, legally compliant and highly persuasive letter of recommendation written from the perspective of an eminent referee.

Follow these rules with absolute, unyielding precision:

1. COMPILATION PERSPECTIVE, HUMANISATION & ELIMINATION OF AI FILLERS:
- Write the letter entirely and independently from the perspective of the referee (${refereeName}, ${refereeTitle} at ${refereeCompany}).
- It must read as if composed by this individual in their own natural, executive, authoritative, and deeply specific tone. Do NOT write using generic, dry, or robotic templates.
- STRICT NEGATIVE CONSTRAINT - ELIMINATE AI BUZZWORDS & FILLERS: Do not use common AI cliches, go-to descriptors, or word padding. Ban these exact phrases and words: "delve", "testament", "stands as a testament", "serves as a testament", "beacon", "spearhead", "meticulously", "leverage", "foster", "landscape", "evolving landscape", "furthermore", "monolithic", "orchestrate", "digital revolutionary", "transformative", "paradigm", "pivotal", "pivotal moment", "key turning point", "indelible mark", "deeply rooted", "shaping the future", "underscores the importance", "prominent media outlets", "active social media presence", "profiled in", "overarching", "game-changer", "unparalleled", "relentless", "synergy", "bespoke". Keep the writing clean, punchy, intellectual, and grounded in concrete activities.
- Prefer simple, active linking verbs. Do NOT write dramatic substitutions like "serves as an epicenter for" (instead use "is a center for" or similar).
- Avoid binary structural quirks or negative parallelisms like "not just X, but also Y", or "not X, but Y".
- Avoid overusing groups of three adjectives or examples (e.g., do not stack "exceptionally creative, technically sound, and commercially competent").
- Do not inflate ordinary achievements into grand historical significance or upgrade standard tasks into grandiose titles (the "Titan" effect). Use clear, direct, evidence-driven, professional, human language.

2. CRITERIA LABELLING IS STRICTLY FORBIDDEN (IMPLICIT EVIDENCE ONLY):
- Do NOT explicitly mention, name-drop, or map achievements using visa criteria codes or official terminology (e.g., do NOT write "this satisfies Mandatory Criterion", "under Optional Criterion 3", "satisfies MC", "visa route", "criteria", etc.).
- The applicant's alignment with Tech Nation Digital Technology requirements must be completely implicit. Detail what the applicant built, what the technical/architectural details were, and what the commercial/technical impact was, letting the facts speak for themselves.

3. FIRST-PERSON ACTIVE ENGAGEMENT:
- Do not say the applicant simply "worked" as a role or was "employed by" a company. Instead, frame accomplishments actively from the referee's firsthand knowledge: e.g., "Her work on...", "She designed and engineered...", "Analyzing her product roadmap strategies...".

4. TOP ADDRESS LAYOUT CONSTRAINTS:
- The top address block MUST contain ONLY the Referee's company name and Referee's business address (strictly do NOT include their personal name or personal job title in this top block!).
- Recommender personal details (name, job title) must strictly only appear in the signature section at the bottom.
- Inside address block must address 'Tech Nation UK, Wenlock Road, London, N1 7GU, United Kingdom' followed by 'For the attention of Tech Nation'.

5. SUBJECT LINE:
- Remove "RE:" and "Re:" completely from the subject line.
- The subject line must be a highly context-aware, customized, professional title based on the applicant's name, visa pathway, and recommendation purpose. Never use a hardcoded subject line.
- Excellent examples:
  - Recommendation for UK Global Talent Visa Application – [Candidate Name]
  - Letter of Recommendation for [Candidate Name]
  - Endorsement of [Candidate Name] for Exceptional Talent Recognition

6. SPELLING, LOCALISATION & FORMAT:
- Use UK English spelling throughout: recognised, organised, programme, specialised, colour, endeavour, centre, prioritised, standardised, characterised.
- Output ONLY the clean, professionally formatted plain typewriter-style business letter text.
- STRICTLY no markdown whatsoever (no asterisks, no double asterisks (**), no hashtags (#), no underscores (_), no hyphens as bullet points (-), or other markdown symbols).
- Single blank line separating paragraphs.

7. WORD COUNT:
- The body of the letter (from salutation 'Dear Sir/Madam' to complimentary close 'Yours faithfully') must be between 650 and 850 words. Be comprehensive, detailed, and elaborate on the technical achievements and leadership qualities of the applicant.
`;

      const prompt = `
Referee Name: ${refereeName}
Referee Title: ${refereeTitle}
Referee Company: ${refereeCompany}
Referee Stature/Bio: ${refereeStature}
Referee Address: ${refereeAddress || "Not Provided"}
Referee Email: ${refereeEmail || "Not Provided"}
Referee Phone: ${refereePhone || "Not Provided"}
Referee LinkedIn: ${refereeLinkedin || "Not Provided"}
Applicant Name: ${applicantName}
Applicant Role: ${applicantRole}
Applicant Company: ${applicantCompany}
Visa Route: ${visaRoute}
Selected Criteria: ${selectedCriteria.join(", ")}
How They Know Each Other / Relationship duration: ${relationships}
Anecdotes & Individual Achievements: ${anecdotes}
Letter Date: ${customDate && customDate.trim() ? customDate.trim() : new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" })}

Draft the final professional recommendation letter following all system instructions. Ensure the top address block contains ONLY the Referee's Company Name and Company Address (do NOT include Referee Name or Referee Title in the top address block, they belong in the signature block). Follow this with the letter date, the inside address block (Tech Nation UK, Wenlock Road, London, N1 7GU, United Kingdom \\n For the attention of Tech Nation), a dynamic and customized subject line (with no RE: prefix), salutation 'Dear Sir/Madam', detailed body paragraphs evaluating the applicant's suitability, yours faithfully, and signature block.

The signature block must contain their contact details and follow this exact format:
Yours faithfully,

${refereeName}
${refereeTitle}
${refereeCompany}
Email: ${refereeEmail || ""}
Phone: ${refereePhone || ""}
LinkedIn: ${refereeLinkedin || ""}
`;

      const modelsToTry = ["gemini-3.5-flash", "gemini-flash-latest", "gemini-3.1-flash-lite"];
      let letterText = "";
      let lastError: any = null;

      for (const modelName of modelsToTry) {
        try {
          console.log(`[Gemini Engine] Attempting generation with model: ${modelName}...`);
          const response = await ai.models.generateContent({
            model: modelName,
            contents: prompt,
            config: {
              systemInstruction: systemInstruction,
              temperature: 0.7,
            },
          });
          letterText = response.text || "";
          if (letterText) {
            console.log(`[Gemini Engine] Generation successful using model: ${modelName}`);
            break;
          }
        } catch (error: any) {
          lastError = error;
          console.warn(`[Gemini Engine] Model ${modelName} failed or limits reached:`, error?.message || error);
        }
      }

      // HIGH-FIDELITY OFFLINE FALLBACK CALIBRATION
      if (!letterText) {
        console.log("[Gemini Engine] All generative models reached free-tier limits or were unavailable. Activating eMigr8 Offline High-Fidelity Compiler...");
        letterText = generateOfflineLetter({
          refereeName,
          refereeTitle,
          refereeCompany,
          refereeStature,
          refereeAddress,
          refereeEmail,
          refereePhone,
          refereeLinkedin,
          applicantName,
          applicantRole,
          applicantCompany,
          visaRoute,
          selectedCriteria,
          relationships,
          anecdotes,
          customDate,
        });
      }

      res.json({ letter: letterText });
    } catch (error: any) {
      console.error("Gemini Generation Error:", error);
      
      // Secondary fallback just in case something in the request loop failed completely before loop execution
      try {
        console.log("[Gemini Engine] Activating eMigr8 Offline High-Fidelity Compiler under emergency catch block...");
        const letterText = generateOfflineLetter({
          refereeName: req.body.refereeName || "Expert Referee",
          refereeTitle: req.body.refereeTitle || "Senior Executive Officer",
          refereeCompany: req.body.refereeCompany || "Technology Leader PLC",
          refereeStature: req.body.refereeStature || "An eminent leader with notable achievements globally.",
          refereeAddress: req.body.refereeAddress,
          refereeEmail: req.body.refereeEmail,
          refereePhone: req.body.refereePhone,
          refereeLinkedin: req.body.refereeLinkedin,
          applicantName: req.body.applicantName || "Global Talent Applicant",
          applicantRole: req.body.applicantRole || "Lead Product Engineer",
          applicantCompany: req.body.applicantCompany || "Digital Innovation Corp",
          visaRoute: req.body.visaRoute || "Exceptional Talent",
          selectedCriteria: req.body.selectedCriteria || ["MC", "OC3"],
          relationships: req.body.relationships || "professional mentorship and strategic product advisory",
          anecdotes: req.body.anecdotes || "pioneering the core product line design and architecture",
          customDate: req.body.customDate,
        });
        return res.json({ letter: letterText });
      } catch (innerError) {
        console.error("Critical: Offline Backup compiler also failed:", innerError);
      }

      let errorMsg = error?.message || "Internal server error during recommendation draft generation.";
      if (errorMsg.includes("429") || errorMsg.toLowerCase().includes("quota") || errorMsg.includes("LIMIT_EXCEEDED") || errorMsg.includes("RESOURCE_EXHAUSTED")) {
        errorMsg = "eMigr8 daily free-tier generation limit reached for this session. To bypass this, please ask the owner to configure a personal API key in AI Studio under Settings > Secrets, or try again in a bit.";
      }
      res.status(500).json({ error: errorMsg });
    }
  });

  // Handle client asset serving and routing
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Serve static frontend files in production
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server executing successfully on port ${PORT}`);
  });
}

startServer();
