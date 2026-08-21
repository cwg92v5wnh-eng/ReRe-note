import "dotenv/config";
import express from "express";
import { GoogleGenAI } from "@google/genai";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
const port = Number(process.env.PORT) || 3000;
const microsoftRefreshCookie = "lecture-note-ms-refresh";
const microsoftRefreshCookieMaxAge = 7 * 24 * 60 * 60;
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

app.set("trust proxy", 1);

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

app.post("/api/auth/microsoft/exchange", async (request, response) => {
  const { code, codeVerifier, redirectUri } = request.body ?? {};
  if (!isTrustedMicrosoftRequest(request, redirectUri)) {
    response.status(403).json({ error: "Microsoft login origin is not allowed." });
    return;
  }
  if (![code, codeVerifier, redirectUri].every((value) => typeof value === "string" && value.trim())) {
    response.status(400).json({ error: "Microsoft login information is incomplete." });
    return;
  }

  try {
    const token = await requestMicrosoftToken({
      code,
      code_verifier: codeVerifier,
      redirect_uri: redirectUri,
      grant_type: "authorization_code",
    });
    storeMicrosoftRefreshCookie(request, response, token.refresh_token);
    response.json(publicMicrosoftToken(token));
  } catch (error) {
    sendMicrosoftTokenError(response, error);
  }
});

app.post("/api/auth/microsoft/refresh", async (request, response) => {
  if (!isTrustedMicrosoftRequest(request)) {
    response.status(403).json({ error: "Microsoft refresh origin is not allowed." });
    return;
  }

  const refreshToken = readCookie(request, microsoftRefreshCookie);
  if (!refreshToken) {
    response.status(401).json({ error: "Microsoft session is not available." });
    return;
  }

  try {
    const token = await requestMicrosoftToken({
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    });
    storeMicrosoftRefreshCookie(request, response, token.refresh_token || refreshToken);
    response.json(publicMicrosoftToken(token));
  } catch (error) {
    if (error.status === 400 || error.status === 401) clearMicrosoftRefreshCookie(request, response);
    sendMicrosoftTokenError(response, error);
  }
});

app.post("/api/auth/microsoft/adopt", async (request, response) => {
  if (!isTrustedMicrosoftRequest(request)) {
    response.status(403).json({ error: "Microsoft refresh origin is not allowed." });
    return;
  }
  const refreshToken = typeof request.body?.refreshToken === "string" ? request.body.refreshToken : "";
  if (!refreshToken) {
    response.status(400).json({ error: "Microsoft refresh information is missing." });
    return;
  }

  try {
    const token = await requestMicrosoftToken({
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    });
    storeMicrosoftRefreshCookie(request, response, token.refresh_token || refreshToken);
    response.json(publicMicrosoftToken(token));
  } catch (error) {
    sendMicrosoftTokenError(response, error);
  }
});

app.post("/api/auth/microsoft/logout", (request, response) => {
  clearMicrosoftRefreshCookie(request, response);
  response.status(204).end();
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

async function requestMicrosoftToken(parameters) {
  const clientId = String(process.env.MICROSOFT_CLIENT_ID || "").trim();
  if (!clientId) {
    const error = new Error("MICROSOFT_CLIENT_ID is not configured.");
    error.status = 500;
    throw error;
  }

  const tenantId = sanitizeMicrosoftTenantId(process.env.MICROSOFT_TENANT_ID);
  const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      scope: "User.Read Files.ReadWrite.AppFolder offline_access",
      ...parameters,
    }),
  });
  const payload = await tokenResponse.json().catch(() => ({}));
  if (!tokenResponse.ok) {
    const error = new Error(payload.error_description || payload.error || "Microsoft token request failed.");
    error.status = tokenResponse.status;
    error.code = payload.error || "token_request_failed";
    throw error;
  }
  return payload;
}

function publicMicrosoftToken(token) {
  return {
    access_token: token.access_token || "",
    expires_in: Number(token.expires_in) || 0,
    scope: token.scope || "",
    token_type: token.token_type || "Bearer",
  };
}

function storeMicrosoftRefreshCookie(request, response, refreshToken) {
  if (!refreshToken) return;
  response.setHeader("Set-Cookie", serializeCookie(microsoftRefreshCookie, refreshToken, request, microsoftRefreshCookieMaxAge));
}

function clearMicrosoftRefreshCookie(request, response) {
  response.setHeader("Set-Cookie", serializeCookie(microsoftRefreshCookie, "", request, 0));
}

function serializeCookie(name, value, request, maxAge) {
  const parts = [
    `${name}=${encodeURIComponent(value)}`,
    "Path=/api/auth/microsoft",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${maxAge}`,
  ];
  if (request.secure || request.get("x-forwarded-proto") === "https") parts.push("Secure");
  return parts.join("; ");
}

function readCookie(request, name) {
  const prefix = `${name}=`;
  const match = String(request.headers.cookie || "")
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(prefix));
  if (!match) return "";
  try {
    return decodeURIComponent(match.slice(prefix.length));
  } catch {
    return "";
  }
}

function isTrustedMicrosoftRequest(request, redirectUri = "") {
  const origin = String(request.get("origin") || "").replace(/\/$/, "");
  const requestOrigin = `${request.protocol}://${request.get("host")}`.replace(/\/$/, "");
  const configuredOrigins = String(process.env.MICROSOFT_ALLOWED_ORIGINS || "")
    .split(",")
    .map((value) => value.trim().replace(/\/$/, ""))
    .filter(Boolean);
  const allowedOrigins = new Set([requestOrigin, ...configuredOrigins]);
  if (origin && !allowedOrigins.has(origin)) return false;
  if (!redirectUri) return true;
  try {
    const redirect = new URL(redirectUri);
    return allowedOrigins.has(redirect.origin) && redirect.pathname === "/auth-callback.html";
  } catch {
    return false;
  }
}

function sanitizeMicrosoftTenantId(value) {
  const tenant = String(value || "common").trim();
  return /^[a-z0-9.-]+$/i.test(tenant) ? tenant : "common";
}

function sendMicrosoftTokenError(response, error) {
  const status = error.status === 400 || error.status === 401 ? 401 : Number(error.status) || 502;
  response.status(status).json({
    error: error.code || "microsoft_token_failed",
    detail: String(error.message || "Microsoft token request failed.").slice(0, 240),
  });
}


