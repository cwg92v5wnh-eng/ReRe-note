import "dotenv/config";
import express from "express";
import { GoogleGenAI } from "@google/genai";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
const port = Number(process.env.PORT) || 3000;
const maxContentLength = 8000;
const defaultModelName = "gemini-3-flash-preview";
const fallbackModelName = "gemini-2.5-flash";
const modelNames = (process.env.GEMINI_MODEL || defaultModelName)
  .split(",")
  .map((model) => model.trim())
  .filter(Boolean);
if (!modelNames.includes(fallbackModelName)) {
  modelNames.push(fallbackModelName);
}
const aiTimeoutMs = 45000;

app.use(express.json({ limit: "64kb" }));
app.use(express.static(__dirname));

app.get("/", (_request, response) => {
  response.sendFile(path.join(__dirname, "LectureNote.html"));
});

app.get("/api/config", (_request, response) => {
  const allowedOrigins = (process.env.MICROSOFT_ALLOWED_ORIGINS || "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean);

  response.json({
    microsoftClientId: process.env.MICROSOFT_CLIENT_ID || "",
    microsoftTenantId: process.env.MICROSOFT_TENANT_ID || "common",
    microsoftAllowedOrigins: allowedOrigins,
    authMode: process.env.AUTH_MODE || "local-and-microsoft",
  });
});

app.post("/api/ai/format-note", async (request, response) => {
  const { content } = request.body ?? {};

  if (typeof content !== "string") {
    response.status(400).json({ error: "content must be a string." });
    return;
  }

  const trimmedContent = content.trim();
  if (!trimmedContent) {
    response.status(400).json({ error: "content must not be empty." });
    return;
  }

  if (trimmedContent.length > maxContentLength) {
    response.status(400).json({ error: `content must be ${maxContentLength} characters or fewer.` });
    return;
  }

  if (!process.env.GEMINI_API_KEY) {
    response.status(500).json({ error: "GEMINI_API_KEY is not configured on the server." });
    return;
  }

  try {
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
    const prompt = [
      "あなたは日本語の講義ノートを整える編集者です。",
      "以下の未整理な講義メモを、読みやすく綺麗な日本語の講義ノートに書き直してください。",
      "",
      "要件:",
      "- 元の意味を保つ",
      "- 見出しで整理する",
      "- 必要に応じて箇条書きを使う",
      "- 重要語句はMarkdownの太字で強調する",
      "- 最後に短い「まとめ」セクションを追加する",
      "- 根拠のない事実や本文にない情報を追加しない",
      "- 余計な前置きや説明は付けず、整形済みノート本文だけを返す",
      "",
      "未整理メモ:",
      trimmedContent,
    ].join("\n");

    const { formatted, model } = await generateFormattedNote(ai, prompt);
    if (!formatted) {
      response.status(502).json({ error: "AI did not return formatted text." });
      return;
    }

    response.json({ formatted, model });
  } catch (error) {
    console.error("Gemini format-note error:", error);
    response.status(502).json({
      error: "AI formatting failed. Please try again later.",
      detail: aiErrorMessage(error),
    });
  }
});

async function generateFormattedNote(ai, prompt) {
  let lastError = null;

  for (const model of modelNames) {
    try {
      // Gemini model call. Replace GEMINI_MODEL in .env when switching models later.
      const geminiResponse = await withTimeout(
        ai.models.generateContent({
          model,
          contents: prompt,
        }),
        aiTimeoutMs,
      );
      return {
        formatted: extractGeminiText(geminiResponse),
        model,
      };
    } catch (error) {
      lastError = error;
      if (!isRetryableModelError(error)) {
        throw error;
      }
      console.warn(`Gemini model "${model}" failed; trying fallback if available.`, aiErrorMessage(error));
    }
  }

  throw lastError || new Error("No Gemini model could format the note.");
}

function extractGeminiText(geminiResponse) {
  if (typeof geminiResponse?.text === "string") {
    return geminiResponse.text.trim();
  }
  if (typeof geminiResponse?.text === "function") {
    return String(geminiResponse.text() || "").trim();
  }
  return "";
}

function withTimeout(promise, timeoutMs) {
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      setTimeout(() => reject(new Error("Gemini request timed out.")), timeoutMs);
    }),
  ]);
}

function isRetryableModelError(error) {
  const message = aiErrorMessage(error).toLowerCase();
  return (
    error?.status === 404 ||
    error?.status === 400 ||
    message.includes("not found") ||
    message.includes("model") ||
    message.includes("unsupported")
  );
}

function aiErrorMessage(error) {
  return String(error?.message || error || "Unknown AI error.");
}

app.listen(port, () => {
  console.log(`LectureNote server running at http://localhost:${port}`);
});


