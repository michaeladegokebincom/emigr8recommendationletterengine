import { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import { 
  FileText, 
  User, 
  Briefcase, 
  HelpCircle, 
  AlertTriangle, 
  CheckCircle, 
  Copy, 
  Download, 
  Printer, 
  RefreshCw, 
  Sparkles, 
  ChevronRight, 
  BookOpen, 
  Globe2, 
  ArrowRight,
  ShieldCheck,
  FileCheck
} from "lucide-react";

import { PRESET_CASES } from "./presets";
import { RefereeDetails, ApplicantDetails } from "./types";

// DOCX Generation Imports
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";

const analyzeError = (errMsg: string) => {
  const lower = errMsg.toLowerCase();
  
  if (
    lower.includes("unexpected token") || 
    lower.includes("<!doctype") || 
    lower.includes("<html") || 
    lower.includes("invalid response") ||
    lower.includes("server returned") ||
    lower.includes("server-restart")
  ) {
    return {
      title: "Backend Server Cold-Start / Restarting",
      desc: "The eMigr8 server is currently warming up or recovering from a live reload / container boot.",
      reason: "When code alterations are made or the underlying container cold-starts, the Node.js server takes about 5 to 10 seconds to fully bind to its routing port. In the meantime, standard web traffic fallback returns the plain static HTML shell instead of a parsed JSON response.",
      action: "Please wait 5–10 seconds and click the blue 'Generate Recommendation Letter' button again to reconnect!"
    };
  }
  
  if (
    lower.includes("429") || 
    lower.includes("quota") || 
    lower.includes("limit_exceeded") || 
    lower.includes("resource_exhausted") ||
    lower.includes("limit reached")
  ) {
    return {
      title: "Google Gemini Free-Tier Quota Limit Reached",
      desc: "The daily API call quota for the shared, free-tier Google Gemini model has been temporarily exhausted.",
      reason: "To keep eMigr8 fully free and accessible, it uses a shared Google free-tier, which carries a rate limit of 20 requests per model per day. If multiple users are generating letters simultaneously, this limit can be exhausted quickly.",
      action: "To bypass this instantly and get unlimited high-speed generations, navigate to the top-right menu, click 'Settings > Secrets', add a custom secret named 'GEMINI_API_KEY' with your own Gemini API key, then click Generate again!"
    };
  }
  
  if (
    lower.includes("api_key") || 
    lower.includes("not configured") || 
    lower.includes("missing")
  ) {
    return {
      title: "Missing Server API Key",
      desc: "The Gemini API integration is missing the server-side credential key ('GEMINI_API_KEY').",
      reason: "The server requires a valid Gemini API secret to securely proxy your endorsement draft generation without exposing private developer keys to client web browsers.",
      action: "To fix this, go to Settings > Secrets in the AI Studio header menu, add a key named 'GEMINI_API_KEY' with your Google AI Studio API key, and restart the server, or wait/re-try."
    };
  }
  
  if (
    lower.includes("503") || 
    lower.includes("high demand") || 
    lower.includes("unavailable")
  ) {
    return {
      title: "Gemini Service Temporarily Overloaded",
      desc: "Google API servers are experiencing a transient spike in traffic or a momentary micro-outage.",
      reason: "A 503 error indicates that Google's servers are heavily loaded and temporarily rejected the request to safeguard their infrastructure. This typically resolves automatically in a few seconds.",
      action: "Please wait 10–15 seconds and click 'Generate' again to re-submit your letter request."
    };
  }

  // General fallback
  return {
    title: "Unexpected Generation System Alert",
    desc: "An unexpected error occurred while communicating with the AI generation server.",
    reason: errMsg,
    action: "Click the blue 'Generate' button to attempt to recompose the draft, or inspect your server console logs for full debugging telemetry."
  };
};

export default function App() {
  // Preset Selection
  const [selectedPresetId, setSelectedPresetId] = useState<string>("tech-leader-talent");

  // Form State
  const [referee, setReferee] = useState<RefereeDetails>(PRESET_CASES[0].referee);
  const [applicant, setApplicant] = useState<ApplicantDetails>(PRESET_CASES[0].applicant);

  // App State
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedLetter, setGeneratedLetter] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [activeFormTab, setActiveFormTab] = useState<"referee" | "applicant" | "anecdotes">("referee");
  const [copyAck, setCopyAck] = useState(false);

  // Ingest & Form Asset States
  const [isParsing, setIsParsing] = useState(false);
  const [parseStatus, setParseStatus] = useState<{ type: "success" | "error"; message: string } | null>(null);
  const [dragActive, setDragActive] = useState(false);

  // Sync state when Preset is changed
  const handleLoadPreset = (presetId: string) => {
    const preset = PRESET_CASES.find((p) => p.id === presetId);
    if (preset) {
      setSelectedPresetId(presetId);
      setReferee({ ...preset.referee });
      setApplicant({ ...preset.applicant });
      // Clear generated output when loading a new preset
      setGeneratedLetter("");
      setError("");
    }
  };

  // Form field Handlers
  const handleRefereeChange = (field: keyof RefereeDetails, value: string) => {
    setReferee((prev) => ({ ...prev, [field]: value }));
  };

  const handleApplicantChange = (field: keyof ApplicantDetails, value: any) => {
    setApplicant((prev) => ({ ...prev, [field]: value }));
  };

  const handleCriteriaToggle = (criterion: string) => {
    setApplicant((prev) => {
      const current = [...prev.selectedCriteria];
      if (current.includes(criterion)) {
        // Leave at least MC if toggled (Wait, let MC be toggleable but MC is usually mandatory)
        return {
          ...prev,
          selectedCriteria: current.filter((c) => c !== criterion),
        };
      } else {
        return {
          ...prev,
          selectedCriteria: [...current, criterion],
        };
      }
    });
  };

  const handleDownloadTemplate = () => {
    try {
      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "eMigr8 — Recommendation Letter Input Form",
                  bold: true,
                  size: 32, // 16pt
                  color: "1E3A8A",
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun("\nThis structured input form helps applicants and referees compile all the strategic facts and metrics needed to calibrate an elite, expert endorsement recommendation letter for Tech Nation.\n\n"),
              ],
            }),

            // SECTION 1 – REFEREE
            new Paragraph({
              children: [
                new TextRun({
                  text: "SECTION 1 – REFEREE",
                  bold: true,
                  size: 24, // 12pt
                  color: "1E3A8A",
                }),
              ],
            }),
            new Paragraph({ children: [new TextRun("\n")] }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Full Name:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Job Title:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Eminent Organisation / Company Name:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Stature & Eminence Biography:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n__________________________________________________\n__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Business Address (UK Standard Format):", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Trust Email:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee Direct Phone:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Referee LinkedIn Profile URL:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n\n")],
            }),

            // SECTION 2 – CANDIDATE ROUTE & PROFILE
            new Paragraph({
              children: [
                new TextRun({
                  text: "SECTION 2 – CANDIDATE ROUTE & PROFILE",
                  bold: true,
                  size: 24, // 12pt
                  color: "1E3A8A",
                }),
              ],
            }),
            new Paragraph({ children: [new TextRun("\n")] }),
            new Paragraph({
              children: [new TextRun({ text: "Candidate Full Name:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Candidate Tech Role:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Target Tech Organisation:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Target Visa Route:", bold: false })],
            }),
            new Paragraph({
              children: [
                new TextRun("☐ Exceptional Talent (5 Years Lead)\n"),
                new TextRun("☐ Exceptional Promise (Emerging Lead)\n\n"),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Target Endorsement Criteria (tick all that apply):", bold: false })],
            }),
            new Paragraph({
              children: [
                new TextRun("☐ Mandatory Criterion (MC) — Sustained recognition as a leader in the digital technology sector.\n"),
                new TextRun("☐ Optional Criterion 1 (OC1) — Groundbreaking track record of innovation in product-led digital technology companies.\n"),
                new TextRun("☐ Optional Criterion 2 (OC2) — Advancement of the technology sector through mentorship, technical education, and community initiatives.\n"),
                new TextRun("☐ Optional Criterion 3 (OC3) — Significant technical impact within product-led digital technology companies.\n"),
                new TextRun("☐ Optional Criterion 4 (OC4) — Original contributions to academic or industry research.\n\n\n"),
              ],
            }),

            // SECTION 3 – RELATIONSHIP & IMPACT STORIES
            new Paragraph({
              children: [
                new TextRun({
                  text: "SECTION 3 – RELATIONSHIP & IMPACT STORIES",
                  bold: true,
                  size: 24, // 12pt
                  color: "1E3A8A",
                }),
              ],
            }),
            new Paragraph({ children: [new TextRun("\n")] }),
            new Paragraph({
              children: [new TextRun({ text: "Professional Relationship & Timeline Context:", bold: false })],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "(e.g. I have known [Candidate Name] since...)\n", italics: true }),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n_________________________________________________________________________________\n\n"),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Key Work Achievements, Patents & Metrics:", bold: false })],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "(e.g. Designed and led...)\n", italics: true }),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n"),
                new TextRun("_________________________________________________________________________________\n\n"),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Document Date:", bold: false })],
            }),
            new Paragraph({
              children: [new TextRun("__________________________________________________\n\n")],
            }),
          ],
        }],
      });

      Packer.toBlob(doc).then((blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "eMigr8 — Recommendation Letter Input Form.docx";
        a.click();
        URL.revokeObjectURL(url);
      });
    } catch (err) {
      console.error(err);
    }
  };

  const handleFileUpload = async (file: File) => {
    const extension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
    if (extension !== ".docx" && extension !== ".pdf") {
      setParseStatus({
        type: "error",
        message: "Only Microsoft Word (.docx) and Adobe PDF (.pdf) documents are accepted for smart autofill.",
      });
      return;
    }

    setIsParsing(true);
    setParseStatus(null);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/parse-document", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errData: any = {};
        try {
          errData = await response.json();
        } catch (_) {}
        throw new Error(errData.error || "The doc analyzer endpoint failed to parse the document.");
      }

      const resData = await response.json();
      const parsed = resData.data;

      if (!parsed) {
        throw new Error("No parsed candidate properties returned from the server model.");
      }

      // Dynamically fill out the 3 eMigr8 tabs
      setReferee({
        name: parsed.refereeName || "",
        title: parsed.refereeTitle || "",
        company: parsed.refereeCompany || "",
        stature: parsed.refereeStature || "",
        address: parsed.refereeAddress || "",
        email: parsed.refereeEmail || "",
        phone: parsed.refereePhone || "",
        linkedin: parsed.refereeLinkedin || "",
      });

      setApplicant((prev) => ({
        ...prev,
        name: parsed.applicantName || "",
        role: parsed.applicantRole || "",
        company: parsed.applicantCompany || "",
        visaRoute: parsed.visaRoute === "Exceptional Promise" ? "Exceptional Promise" : "Exceptional Talent",
        selectedCriteria: Array.isArray(parsed.selectedCriteria) ? parsed.selectedCriteria : prev.selectedCriteria,
        relationship: parsed.relationship || "",
        anecdotes: parsed.anecdotes || prev.anecdotes,
        customDate: parsed.customDate || prev.customDate,
      }));

      setParseStatus({
        type: "success",
        message: `Successfully structuralized candidate data from "${file.name}"! Form inputs for applicant "${parsed.applicantName || 'Applicant'}" and referee "${parsed.refereeName || 'Referee'}" have been dynamically calibrated for Stage 1 evaluation.`,
      });
    } catch (err: any) {
      console.error("[Parser Ingestion Error]", err);
      setParseStatus({
        type: "error",
        message: err.message || "An unexpected error occurred while analyzing the document with Gemini.",
      });
    } finally {
      setIsParsing(false);
    }
  };

  // Submits to Server-Side endpoint
  const handleGenerateLetter = async () => {
    setIsGenerating(true);
    setError("");
    setGeneratedLetter("");

    try {
      const payload = {
        refereeName: referee.name,
        refereeTitle: referee.title,
        refereeCompany: referee.company,
        refereeStature: referee.stature,
        refereeAddress: referee.address,
        refereeEmail: referee.email,
        refereePhone: referee.phone,
        refereeLinkedin: referee.linkedin,
        applicantName: applicant.name,
        applicantRole: applicant.role,
        applicantCompany: applicant.company,
        visaRoute: applicant.visaRoute,
        selectedCriteria: applicant.selectedCriteria,
        relationships: applicant.relationship,
        anecdotes: applicant.anecdotes,
        customDate: applicant.customDate,
      };

      const response = await fetch("/api/generate-letter", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const contentType = response.headers.get("content-type");
      let data: any = {};
      if (contentType && contentType.includes("application/json")) {
        data = await response.json();
      } else {
        const textFallback = await response.text();
        if (textFallback.trim().startsWith("<!") || textFallback.toLowerCase().includes("<html")) {
          throw new Error("SERVER_RESTART_HTML: Received HTML instead of JSON. The backend server API might be recovering or offline.");
        }
        throw new Error(textFallback.slice(0, 100) || "The server returned an invalid response. Please try again.");
      }

      if (!response.ok) {
        throw new Error(data.error || "Failed to generate recommendation letter.");
      }

      setGeneratedLetter(data.letter);
    } catch (err: any) {
      setError(err?.message || "An unexpected network or generation error occurred.");
    } finally {
      setIsGenerating(false);
    }
  };

  // Auto-generate letter immediately on mount with default preset if empty
  useEffect(() => {
    if (!generatedLetter && !isGenerating) {
      handleGenerateLetter();
    }
  }, [selectedPresetId]);

  // Client-Side Letter Quality Audit System
  const getWordCount = () => {
    if (!generatedLetter) return 0;
    // Extract actual body word count or simple word split as approximation
    return generatedLetter.trim().split(/\s+/).length;
  };

  // Markdown audit finder
  const checkMarkdownErrors = () => {
    if (!generatedLetter) return { count: 0, details: [] as string[] };
    const details: string[] = [];
    if (/\*\*/.test(generatedLetter)) details.push("Double asterisks (**) used for bold");
    if (/(^|\s)\*[^\s*]/.test(generatedLetter)) details.push("Single asterisks (*) used for bullet lists or emphasis");
    if (/^#+/m.test(generatedLetter)) details.push("Hashtags (#) representing headings");
    if (/_/.test(generatedLetter)) details.push("Underscores (_) representing italics");
    if (/^\s*-\s+/m.test(generatedLetter)) details.push("Hyphens (-) representing bullet points");
    return {
      count: details.length,
      details,
    };
  };

  // US Spells Audit Finder
  const US_SPELLINGS = [
    { us: "recognized", uk: "recognised" },
    { us: "organized", uk: "organised" },
    { us: "program ", uk: "programme " },
    { us: "specialized", uk: "specialised" },
    { us: "color", uk: "colour" },
    { us: "endeavor", uk: "endeavour" },
    { us: "center", uk: "centre" },
    { us: "prioritized", uk: "prioritised" },
    { us: "standardized", uk: "standardised" },
    { us: "organization", uk: "organisation" },
    { us: "authorized", uk: "authorised" },
  ];

  const checkUSSpellings = () => {
    if (!generatedLetter) return { count: 0, details: [] as string[] };
    const list: string[] = [];
    US_SPELLINGS.forEach((pair) => {
      const regex = new RegExp(`\\b${pair.us}\\b`, "i");
      if (regex.test(generatedLetter)) {
        list.push(`American spelling: "${pair.us}" (should be "${pair.uk}")`);
      }
    });
    return {
      count: list.length,
      details: list,
    };
  };

  const checkCriteriaCompliance = () => {
    if (!generatedLetter) return [];

    const lowerText = generatedLetter.toLowerCase();
    const checked = applicant.selectedCriteria.map((critId) => {
      let passed = false;
      let matchedWord = "";
      let title = "";
      let desc = "";
      let keywords: string[] = [];
      let recommendation = "";

      switch (critId) {
        case "MC":
          title = "Mandatory Criterion (MC)";
          desc = "Sustained recognition as leader (Talent) or emerging leader (Promise)";
          keywords = ["lead", "leader", "leadership", "pioneer", "keynote", "expert", "eminent", "authority", "executive", "recognis", "prestigious", "award", "stature"];
          recommendation = "The letter lacks strong mentions of leadership, industry recognition, pioneering roles, or expert stature. Tech Nation MC requires active proof of being a digital technology leader. Ask your referee to weave in terms like 'pioneered', 'leadership role', or 'recognised authority'.";
          break;
        case "OC1":
          title = "Optional Criterion 1 (OC1)";
          desc = "Groundbreaking track record of innovation at product-led companies";
          keywords = ["innovat", "invent", "patent", "groundbreaking", "novel", "unique", "first-of-its-kind", "pioneered", "proprietary", "conceived", "original", "cutting-edge"];
          recommendation = "No direct evidence of groundbreaking innovation, patents, or novel product creation found. Tech Nation OC1 requires proof of innovative tech. Ask your referee to weave in terms like 'innovated', 'proprietary technology', 'patent pending', or 'original architecture'.";
          break;
        case "OC2":
          title = "Optional Criterion 2 (OC2)";
          desc = "Sector advancement outside work (Open source, mentorship, lectures)";
          keywords = ["mentor", "open-source", "open source", "github", "community", "conference", "public speak", "panelist", "lecture", "volunteer", "hackathon", "advancement", "advancing", "ecosystem"];
          recommendation = "Missing explicit references to contributions outside standard employment hours (such as open-source code, volunteer mentoring, public talks, or sector lectures). Tech Nation OC2 assesses your ecosystem impact. Emphasise references to 'mentorship', 'open-source repositories', or 'industry presentations'.";
          break;
        case "OC3":
          title = "Optional Criterion 3 (OC3)";
          desc = "Significant technical/commercial impact and contributions in product companies";
          keywords = ["growth", "revenue", "acquisition", "scale", "percent", "saved", "reduced", "million", "billion", "optimization", "optimised", "critical", "adopted", "impact", "metric", "scalable"];
          recommendation = "The text lacks concrete impact metrics, optimization figures, or evidence of adoption. Tech Nation OC3 demands proof of significant commercial or technical contributions. Include specific metrics like 'improved throughput by X%', 'scaled architecture', or 'generated revenue/savings'.";
          break;
        case "OC4":
          title = "Optional Criterion 4 (OC4)";
          desc = "Original peer-reviewed academic or technical research of stature";
          keywords = ["paper", "publication", "peer-review", "published", "journal", "research", "citation", "academic", "thesis", "conference paper", "h-index", "co-authored"];
          recommendation = "Letter lacks academic research signals (publications, journal papers, citations, or academic board approvals). Tech Nation OC4 is for formal researcher/academic pathways. Add details about your 'peer-reviewed publications', 'original research', or 'published paper'.";
          break;
        default:
          break;
      }

      // Check keywords
      for (const kw of keywords) {
        // Simple search instead of rigid boundary if parts of words match (e.g. "innovat" matching "innovative")
        if (lowerText.includes(kw.toLowerCase())) {
          passed = true;
          matchedWord = kw;
          break;
        }
      }

      return {
        id: critId,
        title,
        desc,
        passed,
        matchedWord,
        recommendation,
      };
    });

    return checked;
  };

  // Client-side Purge Utilities
  const handlePurgeMarkdown = () => {
    if (!generatedLetter) return;
    let text = generatedLetter;
    // Remove headers hashtaging
    text = text.replace(/^#+\s*(.*?)$/gm, "$1");
    // Remove bold double asterisk
    text = text.replace(/\*\*(.*?)\*\*/g, "$1");
    // Remove bold double underscore
    text = text.replace(/__(.*?)__/g, "$1");
    // Remove single asterisk italics
    text = text.replace(/\*(.*?)\*/g, "$1");
    // Remove underscores
    text = text.replace(/_(.*?)_/g, "$1");
    // Replace markdown bullets starting with hyphen or star with ordered numbers or simple spacing
    let listCounter = 1;
    text = text.replace(/^\s*[-*]\s+(.*?)$/gm, (match, content) => {
      const replacement = `${listCounter}. ${content}`;
      listCounter++;
      return replacement;
    });

    setGeneratedLetter(text);
  };

  const handleLocaliseUK = () => {
    if (!generatedLetter) return;
    let text = generatedLetter;
    US_SPELLINGS.forEach((pair) => {
      // Handle case insensitivity while preserving original word casing
      const regexLower = new RegExp(`\\b${pair.us}\\b`, "g");
      text = text.replace(regexLower, pair.uk);

      const capitalizedUs = pair.us.charAt(0).toUpperCase() + pair.us.slice(1);
      const capitalizedUk = pair.uk.charAt(0).toUpperCase() + pair.uk.slice(1);
      const regexCap = new RegExp(`\\b${capitalizedUs}\\b`, "g");
      text = text.replace(regexCap, capitalizedUk);
    });
    setGeneratedLetter(text);
  };

  const handleCopy = () => {
    if (!generatedLetter) return;
    navigator.clipboard.writeText(generatedLetter);
    setCopyAck(true);
    setTimeout(() => {
      setCopyAck(false);
    }, 2000);
  };

  const handleDownloadTxt = () => {
    if (!generatedLetter) return;
    const element = document.createElement("a");
    const file = new Blob([generatedLetter], { type: "text/plain;charset=utf-8" });
    element.href = URL.createObjectURL(file);
    element.download = `Tech_Nation_Recommendation_${applicant.name.replace(/\s+/g, "_")}.txt`;
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const handlePrint = () => {
    const printWindow = window.open("", "_blank");
    if (printWindow) {
      printWindow.document.write(`
        <html>
        <head>
          <title>Tech Nation Recommendation Letter - ${applicant.name}</title>
          <style>
            body { 
              font-family: 'Times New Roman', Times, serif; 
              line-height: 1.6; 
              padding: 40px; 
              color: #000;
              max-width: 800px;
              margin: 0 auto;
              font-size: 11pt;
            }
            pre { 
              white-space: pre-wrap; 
              font-family: 'Times New Roman', Times, serif; 
              text-align: justify;
              text-justify: inter-word;
            }
          </style>
        </head>
        <body onload="window.print()">
          <pre>${generatedLetter}</pre>
        </body>
        </html>
      `);
      printWindow.document.close();
    }
  };

  const markdownAudit = checkMarkdownErrors();
  const spellAudit = checkUSSpellings();
  const rawWordCount = getWordCount();

  // Highlight word count status
  const getWordCountStatus = () => {
    if (rawWordCount === 0) return { label: "Empty", bg: "bg-slate-100", text: "text-slate-600", dot: "bg-slate-400" };
    if (rawWordCount >= 650 && rawWordCount <= 850) {
      return { label: "Perfect Tech Nation Length (650-850 words)", bg: "bg-emerald-50", text: "text-emerald-700 border border-emerald-200", dot: "bg-emerald-500" };
    }
    if (rawWordCount < 650) {
      return { label: `Slightly Brief (${rawWordCount} words) - target is 650-850`, bg: "bg-amber-50", text: "text-amber-700 border border-amber-200", dot: "bg-amber-500" };
    }
    return { label: `Slightly Comprehensive (${rawWordCount} words) - target is 650-850`, bg: "bg-amber-50", text: "text-amber-700 border border-amber-200", dot: "bg-amber-500" };
  };

  const wordCountStatus = getWordCountStatus();

  return (
    <div className="min-h-screen bg-[#F3F4F6] text-[#111827] font-sans antialiased flex flex-col">
      {/* Dynamic Header */}
      <header className="sticky top-0 bg-white border-b border-gray-200 z-30 px-8 py-4">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-8 h-8 bg-[#1E3A8A] rounded-sm flex items-center justify-center text-white font-bold italic tracking-tighter shrink-0 select-none">e8</div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-sm font-semibold tracking-widest uppercase text-gray-500">
                  eMigr8 Recommendation Engine
                </h1>
                <span className="hidden sm:inline-block px-2 py-0.5 text-[10px] font-bold tracking-wider text-[#1E3A8A] bg-blue-50 border border-blue-200 rounded-sm uppercase">
                  Specialist Engine
                </span>
              </div>
              <p className="text-xs text-gray-400 mt-0.5 font-serif italic">
                Tech Nation Endorsement Support Calibration • Calibrated for Stage 1 Criteria
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3 self-end md:self-auto text-[10px] text-gray-450 font-mono bg-gray-50 px-3.5 py-1.5 rounded-sm border border-gray-200 select-none">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
            <span>Tech Nation Endorsement Framework v2026.1</span>
          </div>
        </div>
      </header>

      <main className="flex-1 w-full max-w-7xl mx-auto px-4 lg:px-8 py-6 flex flex-col gap-6">
        
        {/* Preset Scenarios Header */}
        <div className="bg-white rounded-sm p-6 border border-gray-200 shadow-[0_10px_40px_rgba(0,0,0,0.02)] flex flex-col xl:flex-row items-start xl:items-center justify-between gap-6">
          <div className="space-y-1.5 max-w-xl">
            <div className="flex items-center gap-2 text-[#1E3A8A] font-bold text-xs uppercase tracking-[0.15em]">
              <Sparkles className="w-4 h-4 text-[#1E3A8A] fill-[#1E3A8A]/10" />
              <span>Reference Calibration Blueprints</span>
            </div>
            <h2 className="text-lg font-serif italic text-gray-800">Load Verified Endorsement Scenarios</h2>
            <p className="text-xs text-gray-500 leading-relaxed font-serif">
              Select an endorsed candidate framework. Instantly populate the drafting engine with authentic narratives constructed around ARM partnerships, research labs, or venture backed breakthroughs.
            </p>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 w-full xl:w-auto xl:max-w-[700px] shrink-0">
            {PRESET_CASES.map((caseItem) => (
              <button
                key={caseItem.id}
                onClick={() => handleLoadPreset(caseItem.id)}
                className={`text-left p-4 rounded-sm transition-all duration-200 border text-xs flex flex-col justify-between cursor-pointer ${
                  selectedPresetId === caseItem.id
                    ? "bg-gray-50 text-[#1E3A8A] border-[#1E3A8A] ring-1 ring-[#1E3A8A]/10 shadow-xs"
                    : "bg-white hover:bg-gray-50/50 text-gray-700 border-gray-200"
                }`}
              >
                <div className="font-semibold text-gray-950 truncate tracking-wide">{caseItem.name}</div>
                <div className="text-[10px] text-gray-450 mt-1.5 line-clamp-2 leading-relaxed font-serif italic">
                  {caseItem.description}
                </div>
              </button>
            ))}
          </div>
        </div>

        {/* Primary Workspace */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
          
          {/* LEFT SIDE: Drafting Configuration */}
          <div className="lg:col-span-5 bg-white rounded-sm shadow-[0_10px_40px_rgba(0,0,0,0.02)] border border-gray-200 overflow-hidden flex flex-col">
            
            {/* Autofill & Document Importer Panel */}
            <div className="p-5 border-b border-gray-200 bg-gray-50/50 space-y-4">
              <div className="flex items-center justify-between">
                <span className="text-[10px] font-bold text-[#1E3A8A] uppercase tracking-[0.15em] flex items-center gap-1.5 font-sans">
                  <ShieldCheck className="w-3.5 h-3.5" />
                  Intake Form & Smart Autofill
                </span>
                <span className="text-[9px] font-mono text-gray-450 bg-white border border-gray-200 px-2 py-0.5 rounded-sm select-none">
                  Docx / PDF Analyzer
                </span>
              </div>

              {/* File Upload Dropzone */}
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragActive(true);
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  setDragActive(false);
                }}
                onDrop={async (e) => {
                  e.preventDefault();
                  setDragActive(false);
                  if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                    await handleFileUpload(e.dataTransfer.files[0]);
                  }
                }}
                className={`relative border border-dashed rounded-sm p-4 text-center transition-all duration-200 ${
                  dragActive 
                    ? "border-[#1E3A8A] bg-blue-50/20" 
                    : "border-gray-200 bg-white hover:bg-gray-50/30"
                }`}
              >
                <input
                  type="file"
                  id="doc-upload-raw"
                  accept=".docx,.pdf"
                  className="hidden"
                  onChange={async (e) => {
                    if (e.target.files && e.target.files[0]) {
                      await handleFileUpload(e.target.files[0]);
                    }
                  }}
                />
                
                {isParsing ? (
                  <div className="flex flex-col items-center justify-center py-2.5 space-y-2 select-none">
                    <RefreshCw className="w-5 h-5 text-[#1E3A8A] animate-spin" />
                    <div>
                      <p className="text-xs font-semibold text-gray-705">Extracting candidate data...</p>
                      <p className="text-[10px] text-gray-400 mt-0.5">Calibrating criteria & narrative blocks with Gemini</p>
                    </div>
                  </div>
                ) : (
                  <label htmlFor="doc-upload-raw" className="flex flex-col items-center justify-center cursor-pointer py-1.5 select-none">
                    <FileCheck className="w-5.5 h-5.5 text-gray-400 mb-2" />
                    <span className="text-xs font-bold text-[#1E3A8A] hover:underline">
                      Upload Completed Intake Form
                    </span>
                    <span className="text-[10px] text-gray-450 mt-1">
                      Drag and drop your .docx or .pdf resume/form to auto-populate
                    </span>
                  </label>
                )}
              </div>

              {/* Download Blank Template */}
              <button
                type="button"
                onClick={handleDownloadTemplate}
                className="w-full flex items-center justify-center gap-2 px-4 py-2.5 bg-[#1E3A8A] hover:bg-blue-900 border border-[#1E3A8A] text-white text-xs font-bold uppercase tracking-wider rounded-sm shadow-xs transition-colors cursor-pointer select-none"
              >
                <Download className="w-4 h-4" />
                Download Blank Intake Form (.docx)
              </button>

              {/* Toast Messages */}
              {parseStatus && (
                <div className={`p-3.5 rounded-sm border text-[11px] flex items-start gap-2.5 transition-all duration-300 ${
                  parseStatus.type === "success" 
                    ? "bg-emerald-50 border-emerald-200 text-emerald-900" 
                    : "bg-rose-50 border-rose-200 text-rose-900"
                }`}>
                  <div className="flex-1">
                    <span className={`font-bold uppercase tracking-wider text-[8px] block mb-0.5 ${
                      parseStatus.type === "success" ? "text-emerald-850" : "text-rose-850"
                    }`}>
                      {parseStatus.type === "success" ? "Analysis Complete" : "Processing Failed"}
                    </span>
                    <p className="leading-relaxed font-serif italic text-gray-750">{parseStatus.message}</p>
                  </div>
                  <button 
                    onClick={() => setParseStatus(null)}
                    className="text-gray-400 hover:text-gray-600 font-bold px-1.5 py-0.5 text-[9px]"
                  >
                    ✕
                  </button>
                </div>
              )}
            </div>

            {/* Tab Controller */}
            <div className="grid grid-cols-3 text-center border-b border-gray-200 bg-gray-50">
              <button
                onClick={() => setActiveFormTab("referee")}
                className={`py-4 px-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer ${
                  activeFormTab === "referee"
                    ? "border-[#1E3A8A] text-[#1E3A8A] bg-white font-semibold"
                    : "border-transparent text-gray-400 hover:text-gray-750 hover:bg-gray-50/80"
                }`}
              >
                1. Referee
              </button>
              <button
                onClick={() => setActiveFormTab("applicant")}
                className={`py-4 px-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer ${
                  activeFormTab === "applicant"
                    ? "border-[#1E3A8A] text-[#1E3A8A] bg-white font-semibold"
                    : "border-transparent text-gray-400 hover:text-gray-750 hover:bg-gray-50/80"
                }`}
              >
                2. Applicant
              </button>
              <button
                onClick={() => setActiveFormTab("anecdotes")}
                className={`py-4 px-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer ${
                  activeFormTab === "anecdotes"
                    ? "border-[#1E3A8A] text-[#1E3A8A] bg-white font-semibold"
                    : "border-transparent text-gray-400 hover:text-gray-750 hover:bg-gray-50/80"
                }`}
              >
                3. Evidence
              </button>
            </div>

            {/* Form Panels with subtle animations */}
            <div className="p-6 flex-1 min-h-[460px]">
              <AnimatePresence mode="wait">
                {activeFormTab === "referee" && (
                  <motion.div
                    key="referee-tab"
                    initial={{ opacity: 0, x: -6 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: 6 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-4"
                  >
                    <div>
                      <h3 className="text-xs font-bold uppercase tracking-[0.15em] text-[#1E3A8A] flex items-center gap-1.5 mb-1.5 pb-1 border-b border-gray-100 select-none">
                        <User className="w-3.5 h-3.5 text-[#1E3A8A]" />
                        Referee Credentials
                      </h3>
                      <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic mb-3">
                        Tech Nation expects referees of eminent stature (CEOs, CTOs, senior academics, or venture partners) who can independently vouch for your impact.
                      </p>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Referee Full Name
                        </label>
                        <input
                          type="text"
                          value={referee.name}
                          onChange={(e) => handleRefereeChange("name", e.target.value)}
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Referee Job Title
                        </label>
                        <input
                          type="text"
                          value={referee.title}
                          onChange={(e) => handleRefereeChange("title", e.target.value)}
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                        Eminent Organisation / Company Name
                      </label>
                      <input
                        type="text"
                        value={referee.company}
                        onChange={(e) => handleRefereeChange("company", e.target.value)}
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                        Referee Stature & Eminence Biography
                      </label>
                      <textarea
                        rows={3}
                        value={referee.stature}
                        onChange={(e) => handleRefereeChange("stature", e.target.value)}
                        placeholder="Detail outstanding research, fellowships, exit status, or advisory roles supporting their eminence status."
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] leading-relaxed transition-all"
                      />
                      <p className="text-[10px] text-gray-400 mt-1 italic font-serif">
                        Crucial: Must state publications, state honours, ARM, or venture partner affiliations.
                      </p>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                        Referee Business Address (UK Standard Format)
                      </label>
                      <input
                        type="text"
                        value={referee.address}
                        onChange={(e) => handleRefereeChange("address", e.target.value)}
                        placeholder="e.g. Apex Venture Labs, 88 Kingsway, London, WC2B 6AA, United Kingdom"
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                      />
                      <div className="mt-1.5 bg-blue-50/50 border border-blue-100 rounded-sm p-2.5 text-[10px] text-gray-500 leading-relaxed font-serif">
                        <span className="font-sans font-bold text-[#1E3A8A] uppercase tracking-wider text-[9px] block mb-0.5">UK Address Sample Format</span>
                        <span className="italic">Apex Venture Labs, 88 Kingsway, London, WC2B 6AA, United Kingdom</span>
                        <p className="mt-1 font-sans text-gray-400 leading-tight">Mailing pattern: [Building/Company], [Street], [City], [UK Postcode (e.g. EC4Y 1AY)], [Country].</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Referee Trust Email
                        </label>
                        <input
                          type="email"
                          value={referee.email}
                          onChange={(e) => handleRefereeChange("email", e.target.value)}
                          placeholder="e.g. a.sterling@apexlabs.com"
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Referee Direct Phone
                        </label>
                        <input
                          type="text"
                          value={referee.phone}
                          onChange={(e) => handleRefereeChange("phone", e.target.value)}
                          placeholder="e.g. +44 (0) 20 7946 0192"
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                        <p className="text-[10px] text-gray-400 mt-1 italic leading-tight">
                          Any format/length accepted.
                        </p>
                      </div>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                        Referee LinkedIn Profile URL
                      </label>
                      <input
                        type="url"
                        value={referee.linkedin || ""}
                        onChange={(e) => handleRefereeChange("linkedin", e.target.value)}
                        placeholder="e.g. https://www.linkedin.com/in/alistair-sterling-apex"
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                      />
                    </div>
                  </motion.div>
                )}

                {activeFormTab === "applicant" && (
                  <motion.div
                    key="applicant-tab"
                    initial={{ opacity: 0, x: -6 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: 6 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-4"
                  >
                    <div>
                      <h3 className="text-xs font-bold uppercase tracking-[0.15em] text-[#1E3A8A] flex items-center gap-1.5 mb-1.5 pb-1 border-b border-gray-100 select-none">
                        <Briefcase className="w-3.5 h-3.5 text-[#1E3A8A]" />
                        Candidate Profile & Route
                      </h3>
                      <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic mb-3">
                        Identify your target endorsement pathway and select the criteria the referee will structurally evaluate on your behalf.
                      </p>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Candidate Full Name
                        </label>
                        <input
                          type="text"
                          value={applicant.name}
                          onChange={(e) => handleApplicantChange("name", e.target.value)}
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Candidate Tech Role
                        </label>
                        <input
                          type="text"
                          value={applicant.role}
                          onChange={(e) => handleApplicantChange("role", e.target.value)}
                          placeholder="e.g. Lead Machine Learning Engineer"
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Target Tech Organisation
                        </label>
                        <input
                          type="text"
                          value={applicant.company}
                          onChange={(e) => handleApplicantChange("company", e.target.value)}
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Target Visa Route
                        </label>
                        <select
                          value={applicant.visaRoute}
                          onChange={(e) => handleApplicantChange("visaRoute", e.target.value)}
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all cursor-pointer"
                        >
                          <option value="Exceptional Talent">Exceptional Talent (5 Years Lead)</option>
                          <option value="Exceptional Promise">Exceptional Promise (Emerging Lead)</option>
                        </select>
                      </div>
                    </div>

                    {/* Criteria Picker */}
                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-2">
                        Target Endorsement Criteria (Choose MC + at least 2 OCs)
                      </label>
                      <div className="space-y-2">
                        {[
                          { id: "MC", title: "Mandatory Criterion (MC)", desc: "Sustained recognition as leader (Talent) or emerging leader (Promise)" },
                          { id: "OC1", title: "Optional Criterion 1 (OC1)", desc: "Groundbreaking track record of innovation at product-led companies" },
                          { id: "OC2", title: "Optional Criterion 2 (OC2)", desc: "Sector advancement outside work (Open source, mentorship, lectures)" },
                          { id: "OC3", title: "Optional Criterion 3 (OC3)", desc: "Significant technical/commercial impact and contributions in product-led digital technology companies" },
                          { id: "OC4", title: "Optional Criterion 4 (OC4)", desc: "Original peer-reviewed academic or technical research of stature" },
                        ].map((crit) => {
                          const isSelected = applicant.selectedCriteria.includes(crit.id);
                          return (
                            <button
                              key={crit.id}
                              type="button"
                              onClick={() => handleCriteriaToggle(crit.id)}
                              className={`w-full text-left p-3 rounded-sm border text-xs transition-all duration-150 flex items-start gap-2.5 cursor-pointer ${
                                isSelected
                                  ? "bg-gray-50 text-[#1E3A8A] border-[#1E3A8A] ring-1 ring-[#1E3A8A]/10 shadow-xs"
                                  : "bg-white hover:bg-gray-50/50 text-gray-700 border-gray-200"
                              }`}
                            >
                              <div className={`mt-0.5 w-3.5 h-3.5 rounded-sm border flex items-center justify-center shrink-0 ${
                                isSelected ? "bg-[#1E3A8A] border-[#1E3A8A] text-white" : "border-gray-300"
                              }`}>
                                {isSelected && <span className="text-[9px] font-bold">✓</span>}
                              </div>
                              <div className="space-y-0.5">
                                <div className="font-semibold text-gray-900">{crit.title}</div>
                                <div className={`text-[10px] font-serif italic ${isSelected ? "text-gray-500" : "text-gray-400"}`}>
                                  {crit.desc}
                                </div>
                              </div>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  </motion.div>
                )}

                {activeFormTab === "anecdotes" && (
                  <motion.div
                    key="anecdotes-tab"
                    initial={{ opacity: 0, x: -6 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: 6 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-4"
                  >
                    <div>
                      <h3 className="text-xs font-bold uppercase tracking-[0.15em] text-[#1E3A8A] flex items-center gap-1.5 mb-1.5 pb-1 border-b border-gray-100 select-none">
                        <BookOpen className="w-3.5 h-3.5 text-[#1E3A8A]" />
                        Relationship & Impact Stories
                      </h3>
                      <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic mb-3">
                        The Tech Nation endorsement panel thoroughly checks how the referee knows you to prove their independent assessment. Provide honest relationship metrics and key product milestones.
                      </p>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                        Professional Relationship & Timeline Context
                      </label>
                      <textarea
                        rows={2}
                        value={applicant.relationship}
                        onChange={(e) => handleApplicantChange("relationship", e.target.value)}
                        placeholder="Explain when and in what capacity they collaborated. e.g. mentored for 4 years, or acted as primary tech investor during seed stage."
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] leading-relaxed transition-all"
                      />
                    </div>

                    <div>
                      <div className="flex justify-between items-center mb-1.5">
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em]">
                          Key Work achievements, patents, and metrics
                        </label>
                        <span className="text-[10px] text-[#1E3A8A] font-bold bg-blue-50 px-2 py-0.5 rounded-sm border border-blue-150 uppercase tracking-wide">
                          Specific Stories Required
                        </span>
                      </div>
                      <textarea
                        rows={5}
                        value={applicant.anecdotes}
                        onChange={(e) => handleApplicantChange("anecdotes", e.target.value)}
                        placeholder="Detail key technical achievements. Keep standard names of architectures, patents (e.g. GB12344B), metrics (e.g. scaled to 10M requests, saved 40% CPU, 15k stars on GitHub)."
                        className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] leading-relaxed transition-all"
                      />
                      <p className="text-[10px] text-gray-400 mt-1.5 italic font-serif">
                        Tip: Include details of co-authored papers or specific commercial impact linked explicitly to your core code.
                      </p>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase tracking-[0.12em] mb-1.5">
                          Document Date Format
                        </label>
                        <input
                          type="text"
                          value={applicant.customDate}
                          onChange={(e) => handleApplicantChange("customDate", e.target.value)}
                          placeholder="e.g. 26 May 2026"
                          className="w-full text-xs px-3.5 py-2.5 bg-white border border-gray-200 rounded-sm focus:outline-hidden focus:ring-1 focus:ring-[#1E3A8A] focus:border-[#1E3A8A] transition-all"
                        />
                      </div>
                      <div className="flex items-end justify-end">
                        <div className="text-[10px] text-gray-500 bg-gray-50 px-3 py-2 border border-gray-200 rounded-sm uppercase tracking-wider font-mono select-none">
                          Preserves UK Orthography
                        </div>
                      </div>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {/* Run Generation CTA */}
            <div className="p-6 bg-gray-50 border-t border-gray-200 flex flex-col gap-4">
              <button
                type="button"
                onClick={handleGenerateLetter}
                disabled={isGenerating}
                className="w-full py-4 bg-[#1E3A8A] hover:bg-[#152e75] text-white rounded-sm text-xs font-bold uppercase tracking-widest transition duration-150 flex items-center justify-center gap-2.5 disabled:bg-gray-300 disabled:text-gray-500 disabled:cursor-not-allowed shadow-sm cursor-pointer select-none"
              >
                {isGenerating ? (
                  <>
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>Compiling Recommendation Context...</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="w-3.5 h-3.5 text-blue-200 fill-blue-200" />
                    <span>Generate Recommendation Letter</span>
                    <ArrowRight className="w-3.5 h-3.5 ml-1" />
                  </>
                )}
              </button>
              
              <div className="flex items-center gap-2 justify-center text-[10px] text-gray-400 tracking-wide select-none">
                <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                <span>Secure Server-Side API Proxy via Gemini 3.5 Flash</span>
              </div>
            </div>

          </div>

          {/* RIGHT SIDE: Letter Preview and Validation Panel */}
          <div className="lg:col-span-7 flex flex-col gap-4">
            
            {/* Validation & Audits Dashboard (Real-time Scanner) */}
            <div className="bg-white rounded-sm shadow-[0_10px_40px_rgba(0,0,0,0.02)] border border-gray-200 p-6">
              <h3 className="text-xs font-bold text-gray-400 uppercase tracking-[0.15em] mb-4 flex items-center gap-1.5 pb-1 border-b border-gray-100 select-none">
                <FileCheck className="w-4 h-4 text-[#1E3A8A]" />
                Tech Nation Letter Compliance Audits
              </h3>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {/* Word Count Metric */}
                <div className={`p-4 border rounded-sm flex flex-col justify-between ${wordCountStatus.bg} ${wordCountStatus.text}`}>
                  <div className="flex items-center justify-between text-[10px] font-bold uppercase tracking-wider">
                    <span>Target Word Count</span>
                    <span className="font-mono text-xs">{rawWordCount} / 850</span>
                  </div>
                  <div className="text-[11px] font-medium leading-relaxed mt-1.5">
                    {wordCountStatus.label}
                  </div>
                </div>

                {/* Markdown Scanner */}
                <div className={`p-4 border rounded-sm flex flex-col justify-between ${
                  markdownAudit.count === 0 
                    ? "bg-white text-emerald-800 border-emerald-200" 
                    : "bg-red-50/50 text-red-900 border-red-200"
                }`}>
                  <div className="flex items-center justify-between text-[10px] font-bold uppercase tracking-wider">
                    <span>No-Markdown Safe Scanner</span>
                    <span className="font-mono text-xs">{markdownAudit.count} Errors</span>
                  </div>
                  <div className="text-[11px] leading-relaxed mt-1.5 font-medium">
                    {markdownAudit.count === 0 ? (
                      "Passed markdown compliance (Plain-text only)"
                    ) : (
                      <div className="flex items-center justify-between gap-1.5">
                        <span className="truncate">Found markdown asterisks/bullets!</span>
                        <button
                          onClick={handlePurgeMarkdown}
                          className="bg-red-100/50 hover:bg-red-100 text-red-900 font-bold px-2 py-0.5 rounded-sm text-[9px] uppercase tracking-wide cursor-pointer text-nowrap select-none border border-red-200 transition"
                        >
                          Purge All
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {/* Spell check / Localisation Scanner */}
                <div className={`p-4 border rounded-sm flex flex-col justify-between ${
                  spellAudit.count === 0 
                    ? "bg-white text-emerald-800 border-emerald-200" 
                    : "bg-amber-50/50 text-amber-900 border-amber-200"
                }`}>
                  <div className="flex items-center justify-between text-[10px] font-bold uppercase tracking-wider">
                    <span>UK English Localiser</span>
                    <span className="font-mono text-xs">{spellAudit.count} Warnings</span>
                  </div>
                  <div className="text-[11px] leading-relaxed mt-1.5 font-medium">
                    {spellAudit.count === 0 ? (
                      "Orthography matches UK criteria"
                    ) : (
                      <div className="flex items-center justify-between gap-1.5">
                        <span className="truncate">US spelling (e.g. center, organized)</span>
                        <button
                          onClick={handleLocaliseUK}
                          className="bg-amber-100/50 hover:bg-amber-100 text-amber-900 font-bold px-2 py-0.5 rounded-sm text-[9px] uppercase tracking-wide cursor-pointer text-nowrap select-none border border-amber-200 transition"
                        >
                          Convert
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Warnings details expansion */}
              {(markdownAudit.count > 0 || spellAudit.count > 0) && (
                <div className="mt-4 bg-gray-50 rounded-sm p-4 border border-gray-200 text-[10px] text-gray-500 leading-relaxed font-mono">
                  <div className="font-bold text-gray-700 uppercase tracking-wide flex items-center gap-1.5 mb-1.5">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-500 animate-pulse" />
                    <span>Pending issues before endorsement download:</span>
                  </div>
                  {markdownAudit.details.map((d, idx) => (
                    <div key={`m-err-${idx}`} className="text-red-700 font-medium">
                      • {d} — This fails non-negotiable Plain-Text presentation parameters. Click "Purge All".
                    </div>
                  ))}
                  {spellAudit.details.map((d, idx) => (
                    <div key={`s-err-${idx}`} className="text-amber-800 font-semibold">
                      • {d} — Tech Nation assessors favor standard UK spelling styles. Click "Convert".
                    </div>
                  ))}
                </div>
              )}

              {/* Selected Criteria Coverage Checklist */}
              {generatedLetter && (
                <div className="mt-5 pt-4 border-t border-gray-100">
                  <div className="flex items-center justify-between mb-3">
                    <span className="text-[10px] font-bold text-gray-400 uppercase tracking-[0.15em] flex items-center gap-1.5">
                      <ShieldCheck className="w-3.5 h-3.5 text-[#1E3A8A]" />
                      Selected Criteria Clause Engagement
                    </span>
                    <span className="text-[9px] font-mono text-gray-400 bg-gray-50 border border-gray-200 px-2 py-0.5 rounded-sm select-none">
                      Keyword & Legal Signal Scanner
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                    {checkCriteriaCompliance().map((check) => (
                      <div 
                        key={check.id}
                        className={`p-3.5 border rounded-sm transition-all duration-200 ${
                          check.passed 
                            ? "bg-emerald-50/25 border-emerald-200 text-emerald-900" 
                            : "bg-rose-50/25 border-rose-200 text-rose-900"
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div className="space-y-0.5">
                            <h4 className={`text-xs font-bold leading-tight ${check.passed ? "text-emerald-800" : "text-rose-800"}`}>
                              {check.id}: {check.title}
                            </h4>
                            <p className="text-[10px] font-serif italic text-gray-400 leading-normal mt-0.5">
                              {check.desc}
                            </p>
                          </div>
                          <span className={`px-2 py-0.5 text-[9px] font-bold uppercase tracking-wider rounded-xs shrink-0 select-none ${
                            check.passed 
                              ? "bg-emerald-100 text-emerald-800 border border-emerald-200" 
                              : "bg-rose-100 text-rose-800 border border-rose-200"
                          }`}>
                            {check.passed ? "Verified" : "Attention"}
                          </span>
                        </div>

                        {!check.passed ? (
                          <div className="mt-2.5 bg-rose-50 border border-rose-100 rounded-sm p-3 text-[10px] leading-relaxed text-gray-700">
                            <span className="font-bold text-rose-850 uppercase tracking-wider text-[8px] block mb-0.5">Recommendation:</span>
                            {check.recommendation}
                          </div>
                        ) : (
                          <div className="mt-2.5 bg-emerald-50 border border-emerald-100 rounded-sm p-2 text-[10px] leading-normal text-emerald-800 flex items-center gap-1.5 font-mono">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0"></span>
                            <span className="truncate">Legal signal present: <strong className="font-bold uppercase font-sans text-emerald-950">"{check.matchedWord}"</strong></span>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
                    {/* Letter Output Room */}
            <div className="bg-white rounded-sm shadow-[0_10px_40px_rgba(0,0,0,0.02)] border border-gray-200 overflow-hidden flex flex-col min-h-[500px]">
              
              {/* Output Header Controls */}
              <div className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between flex-wrap gap-2">
                <div className="flex items-center gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-[#1E3A8A] select-none"></span>
                  <span className="text-[10px] font-bold uppercase tracking-widest text-gray-400 select-none">
                    Tech Nation Representative Recommendation Draft
                  </span>
                </div>

                <div className="flex items-center gap-1.5 text-xs">
                  <button
                    onClick={handleCopy}
                    disabled={!generatedLetter}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-white text-gray-750 border border-gray-200 rounded-sm hover:bg-gray-50/50 hover:text-gray-950 transition-all duration-150 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer font-bold text-[10px] uppercase tracking-wider"
                    title="Copy letter text to clipboard"
                  >
                    <Copy className="w-3.5 h-3.5 text-gray-450" />
                    <span>{copyAck ? "Copied!" : "Copy"}</span>
                  </button>

                  <button
                    onClick={handleDownloadTxt}
                    disabled={!generatedLetter}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-white text-gray-750 border border-gray-200 rounded-sm hover:bg-gray-50/50 hover:text-gray-950 transition-all duration-150 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer font-bold text-[10px] uppercase tracking-wider"
                    title="Download Plain Text Letter"
                  >
                    <Download className="w-3.5 h-3.5 text-gray-450" />
                    <span>TXT</span>
                  </button>

                  <button
                    onClick={handlePrint}
                    disabled={!generatedLetter}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-white text-gray-750 border border-gray-200 rounded-sm hover:bg-gray-50/50 hover:text-gray-950 transition-all duration-150 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer font-bold text-[10px] uppercase tracking-wider"
                    title="Print Recommendation"
                  >
                    <Printer className="w-3.5 h-3.5 text-gray-450" />
                    <span>Print</span>
                  </button>
                </div>
              </div>

              {/* Letter Rendering View */}
              <div className="relative p-6 sm:p-10 flex-1 bg-gray-50/30 flex flex-col min-h-[440px]">
                {isGenerating && (
                  <div className="absolute inset-0 bg-white/95 flex flex-col items-center justify-center p-6 text-center z-10 select-none">
                    <div className="p-4 bg-blue-50/80 rounded-full mb-4 text-[#1E3A8A] animate-bounce">
                      <Sparkles className="w-8 h-8 fill-blue-100" />
                    </div>
                    <div className="space-y-2 max-w-sm">
                      <h4 className="font-serif italic text-gray-800 text-sm">Drafting Endorsement Recommendation...</h4>
                      <p className="text-xs text-gray-400 leading-relaxed font-serif">
                        Fusing achievements and referee eminence into custom-tailored executive narratives aligned with official Stage 1 rules.
                      </p>
                    </div>

                    {/* Progress Indicator Items */}
                    <div className="mt-8 w-full max-w-xs space-y-2.5 text-[10px] font-bold text-gray-400 font-mono tracking-wider uppercase">
                      <div className="flex items-center gap-2 justify-start text-[#1E3A8A]">
                        <span className="w-1.5 h-1.5 rounded-full bg-[#1E3A8A] animate-pulse"></span>
                        <span>Calibrating with Tech Nation Standards...</span>
                      </div>
                      <div className="flex items-center gap-2 justify-start text-[#1E3A8A]">
                        <span className="w-1.5 h-1.5 rounded-full bg-[#1E3A8A] animate-pulse"></span>
                        <span>Adapting technical anecdotes...</span>
                      </div>
                      <div className="flex items-center gap-2 justify-start text-[#1E3A8A]">
                        <span className="w-1.5 h-1.5 rounded-full bg-[#1E3A8A] animate-pulse"></span>
                        <span>Eliminating illicit markdown formatting...</span>
                      </div>
                    </div>
                  </div>
                )}

                {error && (() => {
                  const diagnosis = analyzeError(error);
                  return (
                    <div className="absolute inset-x-4 top-4 bg-rose-50 border border-rose-300 rounded-sm p-4 text-xs z-20 shadow-md leading-relaxed font-sans max-h-[90%] overflow-y-auto">
                      <div className="flex items-start gap-2.5 mb-2.5">
                        <AlertTriangle className="w-4 h-4 text-rose-650 shrink-0 mt-0.5" />
                        <div className="flex-1">
                          <h4 className="font-bold text-rose-900 text-sm uppercase tracking-wide">
                            {diagnosis.title}
                          </h4>
                          <p className="text-rose-800 font-medium mt-1 leading-relaxed">
                            {diagnosis.desc}
                          </p>
                        </div>
                        <button 
                          onClick={() => setError("")}
                          className="text-rose-450 hover:text-rose-850 font-bold px-1.5 py-0.5 rounded hover:bg-rose-100 transition text-[10px] uppercase font-mono tracking-wider cursor-pointer select-none"
                          title="Dismiss this notification"
                        >
                          Dismiss
                        </button>
                      </div>

                      <div className="border-t border-rose-200/60 pt-2.5 mt-2 space-y-2">
                        <div>
                          <span className="font-bold text-rose-950 uppercase tracking-widest text-[9px] block">What Caused the Error?</span>
                          <p className="text-gray-650 font-serif leading-relaxed mt-0.5 pl-2.5 border-l-2 border-rose-400">
                            {diagnosis.reason}
                          </p>
                        </div>
                        <div className="bg-rose-100/50 p-2.5 rounded-sm border border-rose-200">
                          <span className="font-bold text-rose-950 uppercase tracking-widest text-[9px] block">How to Resolve This:</span>
                          <p className="text-gray-900 font-sans font-medium leading-relaxed mt-1">
                            {diagnosis.action}
                          </p>
                        </div>
                      </div>
                    </div>
                  );
                })()}

                {/* Plain-text Editable Render */}
                <textarea
                  className="w-full flex-1 min-h-[420px] max-h-[850px] p-8 sm:p-12 md:p-16 bg-white border border-gray-250 rounded-sm shadow-[0_4px_25px_rgba(0,0,0,0.02)] font-serif text-[13px] leading-relaxed text-gray-800 focus:outline-hidden focus:ring-1 focus:ring-gray-300 resize-y text-justify"
                  style={{ textAlign: "justify" }}
                  value={generatedLetter}
                  onChange={(e) => setGeneratedLetter(e.target.value)}
                  placeholder="The generated Tech Nation recommendation letter will populate here. Customize, edit, and audit the output letters organically."
                />

                <div className="mt-3.5 flex items-center justify-between text-[11px] text-gray-400 bg-gray-50 px-4 py-2.5 rounded-sm border border-gray-200">
                  <div className="flex items-center gap-2">
                    <HelpCircle className="w-3.5 h-3.5 text-[#1E3A8A]" />
                    <span>The letter above compiles to body limits. You can edit directly in this workspace to organically perfect sentences.</span>
                  </div>
                </div>
              </div>

            </div>

          </div>

        </div>

        {/* Guidance and Endorsement Criteria Reference Footer */}
        <div className="bg-white rounded-sm p-6 space-y-4 border border-gray-200 shadow-[0_10px_40px_rgba(0,0,0,0.02)]">
          <div className="flex items-center gap-2 text-[#1E3A8A] font-bold text-xs uppercase tracking-[0.15em] border-b border-gray-100 pb-2.5">
            <Globe2 className="w-4 h-4 text-[#1E3A8A]" />
            <span>Tech Nation Endorsement Criteria (Digital Technology Pathway)</span>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-5 gap-4 text-xs leading-relaxed">
            <div className="p-4 bg-gray-50/50 border border-gray-200 rounded-sm">
              <div className="font-semibold text-gray-900 mb-1.5 tracking-wide uppercase text-[10px]">Mandatory (MC)</div>
              <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic">
                Requires sustained recognition as leading talent (Talent) or emerging leading talent (Promise) over the last 5 years with public proof.
              </p>
            </div>

            <div className="p-4 bg-gray-50/50 border border-gray-200 rounded-sm">
              <div className="font-semibold text-gray-900 mb-1.5 tracking-wide uppercase text-[10px]">Innovation (OC1)</div>
              <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic">
                Recognises pioneering commercial innovation as a key employee, founder or executive. Highlight patented ideas and first-of-their-kind platforms.
              </p>
            </div>

            <div className="p-4 bg-gray-50/50 border border-gray-200 rounded-sm">
              <div className="font-semibold text-gray-900 mb-1.5 tracking-wide uppercase text-[10px]">External Contribution (OC2)</div>
              <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic">
                Contributions outside employment. Includes fields like open-source library architecture, key sector mentorship, or eminent speaking panels.
              </p>
            </div>

            <div className="p-4 bg-gray-50/50 border border-gray-200 rounded-sm">
              <div className="font-semibold text-gray-900 mb-1.5 tracking-wide uppercase text-[10px]">Significant Impact (OC3)</div>
              <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic">
                Directly attributable technical, commercial, or financial growth contributions at tech companies. Must be proven with exact figures.
              </p>
            </div>

            <div className="p-4 bg-gray-50/50 border border-gray-200 rounded-sm">
              <div className="font-semibold text-gray-900 mb-1.5 tracking-wide uppercase text-[10px]">Research Publications (OC4)</div>
              <p className="text-[11px] text-gray-500 leading-relaxed font-serif italic">
                Original academic or industrial research published in peer-reviewed journals. Proves advanced theoretical tech competence.
              </p>
            </div>
          </div>
        </div>

      </main>

      <footer className="bg-white border-t border-gray-200 py-6 mt-12 px-6">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4 text-[10px] text-gray-400 font-mono text-center md:text-left uppercase tracking-wider select-none">
          <div>
            &copy; {new Date().getFullYear()} eMigr8. Tech Nation Global Talent Endorsement Recommendation Generator. All rights reserved.
          </div>
          <div>
            Drafting Engine powered strictly by Google Generative AI in the United Kingdom
          </div>
        </div>
      </footer>
    </div>
  );
}
