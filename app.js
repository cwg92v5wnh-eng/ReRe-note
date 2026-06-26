const STORAGE_KEY = "lecture-notes-v2";
const AUTH_USERS_KEY = "lecture-notes-users";
const AUTH_SESSION_KEY = "lecture-notes-session";
const AUTH_PROVIDER_KEY = "lecture-notes-auth-provider";
const MICROSOFT_AUTH_VERIFIER_KEY = "lecture-notes-ms-verifier";
const MICROSOFT_AUTH_STATE_KEY = "lecture-notes-ms-state";
const MICROSOFT_TOKEN_KEY = "lecture-notes-ms-token";
const MICROSOFT_GRAPH_SCOPES = ["User.Read", "Files.ReadWrite.AppFolder"];
const ONEDRIVE_STATE_FILE = "lecture-note-data.json";
const AUTOSAVE_DELAY = 180;
const ONEDRIVE_SAVE_DELAY = 1200;
const NOTEBOOK_INITIAL_LIMIT = 7;
const RECENT_INITIAL_LIMIT = 7;
const RECENT_MAX_LIMIT = 10;
const MAX_INDENT_LEVEL = 3;
const NOTEBOOK_THEMES = {
  green: { accent: "#4f8a67", soft: "rgba(159, 220, 181, 0.34)", border: "rgba(79, 138, 103, 0.34)", focus: "rgba(79, 138, 103, 0.12)" },
  blue: { accent: "#4a90e2", soft: "rgba(158, 198, 255, 0.36)", border: "rgba(74, 144, 226, 0.34)", focus: "rgba(74, 144, 226, 0.12)" },
  violet: { accent: "#8a63d2", soft: "rgba(199, 184, 255, 0.38)", border: "rgba(138, 99, 210, 0.34)", focus: "rgba(138, 99, 210, 0.12)" },
  orange: { accent: "#f2bb45", soft: "rgba(242, 193, 143, 0.38)", border: "rgba(191, 125, 45, 0.32)", focus: "rgba(242, 187, 69, 0.14)" },
  pink: { accent: "#d86d99", soft: "rgba(240, 170, 196, 0.38)", border: "rgba(216, 109, 153, 0.34)", focus: "rgba(216, 109, 153, 0.12)" },
  coral: { accent: "#ee6f61", soft: "rgba(238, 111, 97, 0.24)", border: "rgba(238, 111, 97, 0.34)", focus: "rgba(238, 111, 97, 0.12)" },
  gray: { accent: "#787a80", soft: "rgba(200, 200, 204, 0.36)", border: "rgba(120, 122, 128, 0.28)", focus: "rgba(120, 122, 128, 0.1)" },
};
const PAGE = document.body.dataset.page || document.documentElement.dataset.page;
let memoryStorageValue = null;
let activeUser = loadActiveUser();
let microsoftClientId = "";
let microsoftTenantId = "common";
let microsoftAllowedOrigins = [];
let authMode = "local-and-microsoft";

const state = loadState();
state.menuOpen = false;
state.saveTimer = null;
state.notebookLimit = NOTEBOOK_INITIAL_LIMIT;
state.recentLimit = RECENT_INITIAL_LIMIT;
state.pageInitialized = false;
state.notebookMenuId = null;
state.imageResize = null;
state.toolbarMoreOpen = false;
state.pasteRichOnce = false;
state.savedEditorRange = null;
state.oneDriveSaveTimer = null;
state.oneDriveSyncing = false;
state.applyingOneDriveState = false;

const els = {
  newNotebookBtn: document.getElementById("newNotebookBtn"),
  notebookCreateOverlay: document.getElementById("notebookCreateOverlay"),
  notebookCreateForm: document.getElementById("notebookCreateForm"),
  closeNotebookCreateBtn: document.getElementById("closeNotebookCreateBtn"),
  cancelNotebookCreateBtn: document.getElementById("cancelNotebookCreateBtn"),
  notebookNameInput: document.getElementById("notebookNameInput"),
  notebookColorChoices: document.getElementById("notebookColorChoices"),
  notebookIconChoices: document.getElementById("notebookIconChoices"),
  notebookCreatePreview: document.getElementById("notebookCreatePreview"),
  notebookPreviewTitle: document.getElementById("notebookPreviewTitle"),
  notebookCount: document.getElementById("notebookCount"),
  notebookList: document.getElementById("notebookList"),
  notebookItemTemplate: document.getElementById("notebookItemTemplate"),
  notebookSearchInput: document.getElementById("notebookSearchInput"),
  showMoreNotebooksBtn: document.getElementById("showMoreNotebooksBtn"),
  recentNoteList: document.getElementById("recentNoteList"),
  recentNoteTemplate: document.getElementById("recentNoteTemplate"),
  showMoreRecentBtn: document.getElementById("showMoreRecentBtn"),
  notebookTitle: document.getElementById("notebookTitle"),
  newNoteBtn: document.getElementById("newNoteBtn"),
  noteCount: document.getElementById("noteCount"),
  noteList: document.getElementById("noteList"),
  noteItemTemplate: document.getElementById("noteItemTemplate"),
  editorTitle: document.getElementById("editorTitle"),
  dateLabel: document.getElementById("dateLabel"),
  deleteNoteBtn: document.getElementById("deleteNoteBtn"),
  form: document.getElementById("noteForm"),
  bodyInput: document.getElementById("bodyInput"),
  sizeOneBtn: document.getElementById("sizeOneBtn"),
  sizeTwoBtn: document.getElementById("sizeTwoBtn"),
  sizeThreeBtn: document.getElementById("sizeThreeBtn"),
  boldBtn: document.getElementById("boldBtn"),
  underlineBtn: document.getElementById("underlineBtn"),
  pandaHighlightBtn: document.getElementById("pandaHighlightBtn"),
  strikeBtn: document.getElementById("strikeBtn"),
  highlightBtn: document.getElementById("highlightBtn"),
  imageInsertBtn: document.getElementById("imageInsertBtn"),
  richPasteBtn: document.getElementById("richPasteBtn"),
  imageInput: document.getElementById("imageInput"),
  freeNoteBtn: document.getElementById("freeNoteBtn"),
  toolbarMoreBtn: document.getElementById("toolbarMoreBtn"),
  toolbarMoreMenu: document.getElementById("toolbarMoreMenu"),
  aiFormatBtn: document.getElementById("aiFormatBtn"),
  aiFormatStatus: document.getElementById("aiFormatStatus"),
  textColorChoices: document.getElementById("textColorChoices"),
  authForm: document.getElementById("authForm"),
  authEmailInput: document.getElementById("authEmailInput"),
  authMessage: document.getElementById("authMessage"),
  registerBtn: document.getElementById("registerBtn"),
  microsoftLoginBtn: document.getElementById("microsoftLoginBtn"),
  logoutBtn: document.getElementById("logoutBtn"),
  currentUserLabel: document.getElementById("currentUserLabel"),
  accountMenuBtn: document.getElementById("accountMenuBtn"),
  accountOverlay: document.getElementById("accountOverlay"),
  accountDialogTitle: document.getElementById("accountDialogTitle"),
  closeAccountMenuBtn: document.getElementById("closeAccountMenuBtn"),
  switchAccountBtn: document.getElementById("switchAccountBtn"),
};

wireAuthEvents();
initializeMicrosoftAuth();

if (activeUser) {
  enterApp();
} else {
  renderAuthState();
}

function enterApp() {
  document.body.classList.add("is-authenticated");
  renderAuthState();

  if (state.pageInitialized) {
    if (PAGE === "notebooks") {
      ensureNotebookState();
      ensureDemoNotebookData();
      if (els.notebookSearchInput) els.notebookSearchInput.value = "";
      renderNotebookPage();
    } else {
      ensureNoteState();
      renderNotesPage();
    }
    return;
  }

  if (PAGE === "notebooks") {
    setupNotebookPage();
  } else {
    setupNotesPage();
  }
  state.pageInitialized = true;
}

function setupNotebookPage() {
  wireNotebookEvents();
  renderIconChoices();
  ensureNotebookState();
  ensureDemoNotebookData();
  if (els.notebookSearchInput) els.notebookSearchInput.value = "";
  renderNotebookPage();
}

function setupNotesPage() {
  wireNotesEvents();
  ensureNoteState();
  renderNotesPage();
}

function wireAuthEvents() {
  els.authForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    loginUser();
  });
  els.registerBtn?.addEventListener("click", registerUser);
  els.microsoftLoginBtn?.addEventListener("click", loginWithMicrosoft);
  els.logoutBtn?.addEventListener("click", logoutUser);
  els.accountMenuBtn?.addEventListener("click", openAccountMenu);
  els.closeAccountMenuBtn?.addEventListener("click", closeAccountMenu);
  els.switchAccountBtn?.addEventListener("click", switchMicrosoftAccount);
  els.accountOverlay?.addEventListener("click", (event) => {
    if (event.target === els.accountOverlay) closeAccountMenu();
  });
}

async function initializeMicrosoftAuth() {
  if (!els.microsoftLoginBtn) return;
  if (window.location.protocol === "file:") {
    setMicrosoftLoginUnavailable("Microsoftログインはローカルサーバーまたは公開サイトで利用できます。");
    return;
  }

  try {
    const response = await fetch("/api/config");
    const config = await response.json();
    authMode = normalizeAuthMode(config.authMode);
    applyAuthMode();

    if (!config.microsoftClientId) {
      setMicrosoftLoginUnavailable("Microsoftログイン設定がまだ入っていません。");
      return;
    }
    microsoftClientId = config.microsoftClientId;
    microsoftTenantId = sanitizeMicrosoftTenantId(config.microsoftTenantId);
    microsoftAllowedOrigins = normalizeMicrosoftAllowedOrigins(config.microsoftAllowedOrigins);

    if (!isMicrosoftOriginAllowed()) {
      setMicrosoftLoginUnavailable(`Microsoftログインの許可URLに ${window.location.origin} が入っていません。`);
      return;
    }

    els.microsoftLoginBtn.disabled = false;
    els.microsoftLoginBtn.textContent = "Microsoftでログイン";
    if (activeUser && loadAuthProvider() === "microsoft" && readMicrosoftToken()) {
      syncOneDrive({ preferCloud: true });
    }
  } catch {
    setMicrosoftLoginUnavailable("Microsoftログインを読み込めませんでした。");
  }
}

function setMicrosoftLoginUnavailable(message) {
  if (!els.microsoftLoginBtn) return;
  els.microsoftLoginBtn.disabled = true;
  els.microsoftLoginBtn.textContent = message;
}

function sanitizeMicrosoftTenantId(value) {
  const tenant = String(value || "common").trim();
  if (/^[a-z0-9.-]+$/i.test(tenant)) return tenant;
  return "common";
}

function normalizeMicrosoftAllowedOrigins(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((origin) => String(origin || "").trim().replace(/\/$/, ""))
    .filter(Boolean);
}

function normalizeAuthMode(value) {
  return value === "microsoft-only" ? "microsoft-only" : "local-and-microsoft";
}

function applyAuthMode() {
  const hideLocalAuth = authMode === "microsoft-only";
  const emailField = els.authEmailInput?.closest(".auth-field");
  const localButtons = els.authForm?.querySelector(".auth-buttons");
  if (emailField) emailField.hidden = hideLocalAuth;
  if (localButtons) localButtons.hidden = hideLocalAuth;
  if (els.authEmailInput) els.authEmailInput.required = !hideLocalAuth;
}

function isMicrosoftOriginAllowed() {
  if (!microsoftAllowedOrigins.length) return true;
  return microsoftAllowedOrigins.includes(window.location.origin.replace(/\/$/, ""));
}

function wireNotebookEvents() {
  els.newNotebookBtn?.addEventListener("click", openNotebookCreate);
  els.notebookCreateForm?.addEventListener("submit", submitNotebookCreate);
  els.closeNotebookCreateBtn?.addEventListener("click", closeNotebookCreate);
  els.cancelNotebookCreateBtn?.addEventListener("click", closeNotebookCreate);
  els.notebookNameInput?.addEventListener("input", updateNotebookCreatePreview);
  els.notebookColorChoices?.addEventListener("click", selectNotebookColor);
  els.notebookIconChoices?.addEventListener("click", selectNotebookIcon);
  els.notebookSearchInput?.addEventListener("input", () => {
    state.notebookLimit = NOTEBOOK_INITIAL_LIMIT;
    renderNotebookPage();
  });
  els.showMoreNotebooksBtn?.addEventListener("click", showMoreNotebooks);
  els.showMoreRecentBtn?.addEventListener("click", showMoreRecentNotes);
  document.addEventListener("click", handleOutsideNotebookMenuClick);
}

function wireNotesEvents() {
  els.newNoteBtn?.addEventListener("click", createNote);
  els.deleteNoteBtn?.addEventListener("click", deleteSelectedNote);
  els.bodyInput?.addEventListener("input", () => {
    updateCurrentNoteFromEditor();
    scheduleSave();
    saveEditorSelection();
    updateFormattingActiveState();
  });
  els.bodyInput?.addEventListener("keyup", () => {
    saveEditorSelection();
    updateFormattingActiveState();
  });
  els.bodyInput?.addEventListener("mouseup", () => {
    saveEditorSelection();
    updateFormattingActiveState();
  });
  els.pandaHighlightBtn?.addEventListener("pointerdown", preventToolbarFocusLoss);
  els.sizeOneBtn?.addEventListener("click", () => applyTextSize("1"));
  els.sizeTwoBtn?.addEventListener("click", () => applyTextSize("2"));
  els.sizeThreeBtn?.addEventListener("click", () => applyTextSize("3"));
  els.boldBtn?.addEventListener("click", () => toggleInlineFormat("bold"));
  els.underlineBtn?.addEventListener("click", () => toggleInlineFormat("underline"));
  els.pandaHighlightBtn?.addEventListener("click", applyPandaHighlight);
  els.strikeBtn?.addEventListener("click", () => {
    toggleInlineFormat("strike");
    closeToolbarMoreMenu();
  });
  els.highlightBtn?.addEventListener("click", () => {
    toggleHighlight();
    closeToolbarMoreMenu();
  });
  els.imageInsertBtn?.addEventListener("click", () => {
    els.imageInput?.click();
    closeToolbarMoreMenu();
  });
  els.richPasteBtn?.addEventListener("click", enableRichPasteOnce);
  els.imageInput?.addEventListener("change", handleImageFileSelect);
  els.bodyInput?.addEventListener("paste", handleEditorPaste);
  els.bodyInput?.addEventListener("click", handleEditorImageClick);
  els.bodyInput?.addEventListener("keydown", handleEditorEnterKey);
  els.bodyInput?.addEventListener("keydown", handleEditorIndentKey);
  els.bodyInput?.addEventListener("keydown", handleEditorImageDeleteKey);
  els.bodyInput?.addEventListener("pointerdown", handleImageResizeStart);
  document.addEventListener("pointermove", handleImageResizeMove);
  document.addEventListener("pointerup", handleImageResizeEnd);
  els.freeNoteBtn?.addEventListener("click", createFreeNoteAtCursor);
  els.toolbarMoreBtn?.addEventListener("click", toggleToolbarMoreMenu);
  els.aiFormatBtn?.addEventListener("click", () => {
    closeToolbarMoreMenu();
    formatCurrentNoteWithAI();
  });
  els.textColorChoices?.addEventListener("click", handleTextColorChoice);
  els.form?.addEventListener("keydown", handleShortcuts);
  document.addEventListener("keydown", handleGlobalKeys);
  document.addEventListener("click", handleOutsideMenuClick);
  document.addEventListener("click", handleOutsideToolbarMoreClick);
  window.addEventListener("beforeunload", saveBeforeUnload);
}

function handleShortcuts(event) {
  if (!(event.ctrlKey || event.metaKey)) return;
  const key = event.key.toLowerCase();
  if (key === "," || key === "<") {
    event.preventDefault();
    changeTextSizeByShortcut(-1);
    return;
  }
  if (key === "." || key === ">") {
    event.preventDefault();
    changeTextSizeByShortcut(1);
    return;
  }
  if (key === "s" || key === "enter") {
    event.preventDefault();
    saveSelectedNote();
    closeMenu();
    return;
  }
  if (key === "n") {
    event.preventDefault();
    createNote();
    return;
  }
  if (key === "f") {
    event.preventDefault();
    createFreeNoteAtCursor();
  }
}

function preventToolbarFocusLoss(event) {
  event.preventDefault();
}

function saveEditorSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput) return;
  const range = selection.getRangeAt(0);
  if (!els.bodyInput.contains(range.commonAncestorContainer)) return;
  state.savedEditorRange = range.cloneRange();
}

function restoreEditorSelection() {
  if (!state.savedEditorRange || !els.bodyInput) return false;
  const selection = window.getSelection();
  try {
    selection?.removeAllRanges();
    selection?.addRange(state.savedEditorRange);
    return true;
  } catch {
    state.savedEditorRange = null;
    return false;
  }
}

function changeTextSizeByShortcut(delta) {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  const current = Number(activeTextSize() || 1);
  const next = Math.max(1, Math.min(3, current + delta));
  applyTextSize(String(next));
}

function handleGlobalKeys(event) {
  if (event.key === "Escape") {
    closeMenu();
    closeToolbarMoreMenu();
    closeAccountMenu();
  }
}

function handleOutsideMenuClick(event) {
  if (!state.menuOpen) return;
  if (els.menuPanel?.contains(event.target) || els.menuButton?.contains(event.target)) return;
  closeMenu();
}

function handleOutsideToolbarMoreClick(event) {
  if (!state.toolbarMoreOpen) return;
  if (els.toolbarMoreMenu?.contains(event.target) || els.toolbarMoreBtn?.contains(event.target)) return;
  closeToolbarMoreMenu();
}

function saveBeforeUnload() {
  if (PAGE !== "notes" || !activeUser || !currentNote()) return;
  window.clearTimeout(state.saveTimer);
  updateCurrentNoteFromEditor({ render: false });
  saveToStorage();
}

function handleOutsideNotebookMenuClick(event) {
  if (!state.notebookMenuId) return;
  if (event.target.closest(".notebook-menu") || event.target.closest(".notebook-menu-button")) return;
  closeNotebookMenu();
}

function toggleMenu() {
  state.menuOpen ? closeMenu() : openMenu();
}

function openMenu() {
  state.menuOpen = true;
  if (els.menuPanel) els.menuPanel.hidden = false;
  if (els.menuButton) els.menuButton.setAttribute("aria-expanded", "true");
}

function closeMenu() {
  state.menuOpen = false;
  if (els.menuPanel) els.menuPanel.hidden = true;
  if (els.menuButton) els.menuButton.setAttribute("aria-expanded", "false");
}

function closeNotebookMenu() {
  state.notebookMenuId = null;
  renderNotebookPage();
}

function toggleToolbarMoreMenu(event) {
  event?.stopPropagation();
  state.toolbarMoreOpen ? closeToolbarMoreMenu() : openToolbarMoreMenu();
}

function openToolbarMoreMenu() {
  state.toolbarMoreOpen = true;
  if (els.toolbarMoreMenu) els.toolbarMoreMenu.hidden = false;
  if (els.toolbarMoreBtn) els.toolbarMoreBtn.setAttribute("aria-expanded", "true");
}

function closeToolbarMoreMenu() {
  state.toolbarMoreOpen = false;
  if (els.toolbarMoreMenu) els.toolbarMoreMenu.hidden = true;
  if (els.toolbarMoreBtn) els.toolbarMoreBtn.setAttribute("aria-expanded", "false");
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function loadUsers() {
  try {
    const users = JSON.parse(localStorage.getItem(AUTH_USERS_KEY) || "[]");
    return Array.isArray(users) ? users : [];
  } catch {
    return [];
  }
}

function saveUsers(users) {
  try {
    localStorage.setItem(AUTH_USERS_KEY, JSON.stringify(users));
  } catch {
    // The app can still work in-memory if localStorage is unavailable.
  }
}

function addLocalUser(email) {
  const users = loadUsers();
  if (!users.includes(email)) {
    users.push(email);
    saveUsers(users);
  }
}

function loadActiveUser() {
  try {
    return normalizeEmail(localStorage.getItem(AUTH_SESSION_KEY));
  } catch {
    return "";
  }
}

function loadAuthProvider() {
  try {
    return localStorage.getItem(AUTH_PROVIDER_KEY) || "local";
  } catch {
    return "local";
  }
}

function saveActiveUser(email, provider = "local") {
  activeUser = email;
  try {
    localStorage.setItem(AUTH_SESSION_KEY, email);
    localStorage.setItem(AUTH_PROVIDER_KEY, provider);
  } catch {
    // Session persistence is best-effort in local-only mode.
  }
}

function clearActiveUser() {
  activeUser = "";
  try {
    localStorage.removeItem(AUTH_SESSION_KEY);
    localStorage.removeItem(AUTH_PROVIDER_KEY);
  } catch {
    // Nothing else is needed if localStorage is blocked.
  }
}

function userStorageKey() {
  return activeUser ? `${STORAGE_KEY}:${activeUser}` : STORAGE_KEY;
}

function renderAuthState(message = "") {
  document.body.classList.toggle("is-authenticated", Boolean(activeUser));
  if (els.currentUserLabel) els.currentUserLabel.textContent = activeUser || "";
  if (els.accountDialogTitle) els.accountDialogTitle.textContent = activeUser || "";
  if (els.authMessage) els.authMessage.textContent = message;
  if (els.authEmailInput && activeUser) els.authEmailInput.value = activeUser;
}

function openAccountMenu() {
  if (!els.accountOverlay || !activeUser) return;
  els.accountOverlay.hidden = false;
  els.accountMenuBtn?.setAttribute("aria-expanded", "true");
  els.switchAccountBtn?.focus();
}

function closeAccountMenu() {
  if (!els.accountOverlay) return;
  els.accountOverlay.hidden = true;
  els.accountMenuBtn?.setAttribute("aria-expanded", "false");
}

async function switchMicrosoftAccount() {
  closeAccountMenu();
  clearActiveUser();
  clearMicrosoftToken();
  renderAuthState();
  await loginWithMicrosoft();
}

function replaceState(nextState) {
  state.notebooks = nextState.notebooks;
  state.notes = nextState.notes;
  state.selectedNotebookId = nextState.selectedNotebookId;
  state.selectedNoteId = nextState.selectedNoteId;
  state.notebookLimit = NOTEBOOK_INITIAL_LIMIT;
  state.recentLimit = RECENT_INITIAL_LIMIT;
}

function authEmailFromInput() {
  const email = normalizeEmail(els.authEmailInput?.value);
  if (!isValidEmail(email)) {
    renderAuthState("有効なメールアドレスを入力してください。");
    return "";
  }
  return email;
}

function registerUser() {
  if (authMode === "microsoft-only") return;
  const email = authEmailFromInput();
  if (!email) return;

  addLocalUser(email);

  saveActiveUser(email);
  replaceState(loadState());
  enterApp();
  saveToStorage();
}

function loginUser() {
  if (authMode === "microsoft-only") return;
  const email = authEmailFromInput();
  if (!email) return;

  const users = loadUsers();
  if (!users.includes(email)) {
    renderAuthState("先に登録してください。");
    return;
  }

  saveActiveUser(email);
  replaceState(loadState());
  enterApp();
}

async function loginWithMicrosoft() {
  if (!microsoftClientId) {
    renderAuthState("Microsoftログイン設定がまだ完了していません。");
    return;
  }

  try {
    els.microsoftLoginBtn.disabled = true;
    els.microsoftLoginBtn.textContent = "Microsoftへ接続中...";

    const verifier = randomBase64Url(64);
    const stateValue = randomBase64Url(24);
    sessionStorage.setItem(MICROSOFT_AUTH_VERIFIER_KEY, verifier);
    sessionStorage.setItem(MICROSOFT_AUTH_STATE_KEY, stateValue);

    const popup = window.open(
      await buildMicrosoftAuthUrl(verifier, stateValue),
      "lecture-note-microsoft-login",
      "width=520,height=680,popup=yes",
    );
    if (!popup) {
      throw new Error("popup_blocked");
    }

    const authResult = await waitForMicrosoftAuthMessage(stateValue);
    const token = await exchangeMicrosoftCode(authResult.code, verifier);
    const profile = await fetchMicrosoftProfile(token.access_token);
    saveMicrosoftToken(token);
    await completeMicrosoftLogin(profile);
  } catch (error) {
    renderAuthState(readableMicrosoftAuthError(error));
  } finally {
    if (els.microsoftLoginBtn) {
      els.microsoftLoginBtn.disabled = false;
      els.microsoftLoginBtn.textContent = "Microsoftでログイン";
    }
  }
}

async function buildMicrosoftAuthUrl(verifier, stateValue) {
  const challenge = await sha256Base64Url(verifier);
  const params = new URLSearchParams({
    client_id: microsoftClientId,
    response_type: "code",
    redirect_uri: microsoftRedirectUri(),
    response_mode: "query",
    scope: MICROSOFT_GRAPH_SCOPES.join(" "),
    state: stateValue,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account",
  });
  return `${microsoftAuthorityBaseUrl()}/oauth2/v2.0/authorize?${params.toString()}`;
}

function waitForMicrosoftAuthMessage(expectedState) {
  return new Promise((resolve, reject) => {
    const timeout = window.setTimeout(() => {
      window.removeEventListener("message", handleMessage);
      reject(new Error("login_timeout"));
    }, 120000);

    function handleMessage(event) {
      if (event.origin !== window.location.origin) return;
      const data = event.data || {};
      if (data.type !== "lecture-note:microsoft-auth") return;

      window.clearTimeout(timeout);
      window.removeEventListener("message", handleMessage);

      if (data.state !== expectedState) {
        reject(new Error("state_mismatch"));
        return;
      }
      if (data.error) {
        reject(new Error(data.errorDescription || data.error));
        return;
      }
      if (!data.code) {
        reject(new Error("missing_code"));
        return;
      }
      resolve(data);
    }

    window.addEventListener("message", handleMessage);
  });
}

async function exchangeMicrosoftCode(code, verifier) {
  const response = await fetch(`${microsoftAuthorityBaseUrl()}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: microsoftClientId,
      scope: MICROSOFT_GRAPH_SCOPES.join(" "),
      code,
      redirect_uri: microsoftRedirectUri(),
      grant_type: "authorization_code",
      code_verifier: verifier,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || "token_exchange_failed");
  }
  return payload;
}

function microsoftAuthorityBaseUrl() {
  return `https://login.microsoftonline.com/${encodeURIComponent(microsoftTenantId)}`;
}

async function fetchMicrosoftProfile(accessToken) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName", {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const profile = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(profile.error?.message || "profile_fetch_failed");
  }
  return profile;
}

async function completeMicrosoftLogin(profile) {
  const email = normalizeEmail(profile?.mail || profile?.userPrincipalName);
  if (!email) {
    renderAuthState("Microsoftアカウントのメールアドレスを取得できませんでした。");
    return;
  }

  addLocalUser(email);
  saveActiveUser(email, "microsoft");
  replaceState(loadState());
  enterApp();
  renderAuthState();
  await syncOneDrive({ preferCloud: true });
}

function readableMicrosoftAuthError(error) {
  const message = String(error?.message || error || "");
  console.warn("Microsoft login error:", error);
  if (message.includes("AADSTS70002") || message.includes("client_secret")) {
    return `Microsoftアプリ登録で、リダイレクトURIを「シングルページ アプリケーション（SPA）」として登録してください: ${microsoftRedirectUri()}`;
  }
  if (message.includes("cancel") || message.includes("access_denied")) return "Microsoftログインがキャンセルされました。";
  if (message.includes("popup_blocked") || message.includes("popup")) {
    return "ポップアップがブロックされました。ブラウザ設定でこのサイトのポップアップを許可してください。";
  }
  if (message.includes("redirect_uri")) {
    return `Microsoftアプリ登録のリダイレクトURIを確認してください。登録するURL: ${microsoftRedirectUri()}`;
  }
  if (message.includes("AADSTS9002326") || message.includes("single-page application")) {
    return "Microsoftアプリ登録のプラットフォームを「シングルページ アプリケーション」にしてください。";
  }
  if (message.includes("unauthorized_client")) {
    return "Microsoftアプリ登録で、個人用Microsoftアカウントを許可しているか確認してください。";
  }
  if (message.includes("Failed to fetch") || message.includes("NetworkError")) {
    return "Microsoftのトークン取得に接続できませんでした。公開URLとSPA設定を確認してください。";
  }
  if (message.includes("login_timeout")) return "Microsoftログインがタイムアウトしました。もう一度お試しください。";
  return `Microsoftログインに失敗しました: ${shortErrorMessage(message)}`;
}
function shortErrorMessage(message) {
  return String(message || "詳細不明").replace(/\s+/g, " ").slice(0, 180);
}

function microsoftRedirectUri() {
  return `${window.location.origin}/auth-callback.html`;
}

function saveMicrosoftToken(token) {
  try {
    sessionStorage.setItem(MICROSOFT_TOKEN_KEY, JSON.stringify({
      accessToken: token.access_token || "",
      expiresAt: Date.now() + Number(token.expires_in || 0) * 1000,
      scope: token.scope || "",
    }));
  } catch {
    // Microsoft login can still be used for the current page session.
  }
}

function readMicrosoftToken() {
  try {
    const token = JSON.parse(sessionStorage.getItem(MICROSOFT_TOKEN_KEY) || "null");
    if (!token?.accessToken || Number(token.expiresAt) <= Date.now() + 30000) return null;
    return token;
  } catch {
    return null;
  }
}

function clearMicrosoftToken() {
  try {
    sessionStorage.removeItem(MICROSOFT_TOKEN_KEY);
    sessionStorage.removeItem(MICROSOFT_AUTH_VERIFIER_KEY);
    sessionStorage.removeItem(MICROSOFT_AUTH_STATE_KEY);
  } catch {
    // Best-effort cleanup.
  }
}

function randomBase64Url(byteLength) {
  const bytes = new Uint8Array(byteLength);
  crypto.getRandomValues(bytes);
  return base64UrlEncode(bytes);
}

async function sha256Base64Url(value) {
  const bytes = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return base64UrlEncode(new Uint8Array(digest));
}

function base64UrlEncode(bytes) {
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function logoutUser() {
  window.clearTimeout(state.oneDriveSaveTimer);
  closeAccountMenu();
  clearActiveUser();
  clearMicrosoftToken();
  closeMenu();
  renderAuthState();
}

function loadState() {
  const raw = readStoredState();
  if (!raw) return createInitialState();

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return migrateLegacyState(parsed);

    const notebooks = Array.isArray(parsed.notebooks) ? parsed.notebooks.map(normalizeNotebook) : [];
    const defaultNotebookId = notebooks[0]?.id || "";
    const notes = Array.isArray(parsed.notes)
      ? parsed.notes.map((note) => normalizeNote(note, defaultNotebookId))
      : [];
    const normalized = {
      notebooks,
      notes,
      selectedNotebookId: parsed.selectedNotebookId || notebooks[0]?.id || null,
      selectedNoteId: parsed.selectedNoteId || null,
    };
    if (looksLikeLegacyDefaultSeed(normalized)) {
      return createInitialState();
    }
    ensureStateIntegrity(normalized);
    return normalized;
  } catch {
    return createInitialState();
  }
}

function createInitialState() {
  const seed = buildSampleSeed();
  return {
    notebooks: seed.notebooks,
    notes: seed.notes,
    selectedNotebookId: seed.notebooks[0].id,
    selectedNoteId: latestNoteForNotebookIn(seed.notes, seed.notebooks[0].id)?.id ?? null,
  };
}

function looksLikeLegacyDefaultSeed(stateLike) {
  return (
    stateLike.notebooks.length === 1 &&
    stateLike.notes.length === 1 &&
    stateLike.notebooks[0]?.title === "ノートブック 1" &&
    stateLike.notes[0]?.title === "講義ノートの使い方"
  );
}

function buildSampleSeed() {
  const notebookDefs = [
    {
      title: "哲学",
      color: "violet",
      notes: [
        { title: "存在論のメモ", body: "存在とは何か。\n\n本質と実在の関係を整理する。", tags: ["存在論", "概念"], important: true, offset: 120000 },
        { title: "デカルトの要点", body: "我思う、ゆえに我あり。\n\n確実性の出発点。", tags: ["デカルト", "近代"], important: false, offset: 240000 },
      ],
    },
    {
      title: "情報社会",
      color: "blue",
      notes: [
        { title: "SNSと公共性", body: "匿名性と責任。\n\n拡散の速さと文化の分断。", tags: ["SNS", "メディア"], important: false, offset: 100000 },
        { title: "アルゴリズムの影響", body: "おすすめ機能が注意を奪う。\n\n行動履歴と最適化。", tags: ["AI", "推薦"], important: true, offset: 200000 },
      ],
    },
    {
      title: "ゼミ",
      color: "green",
      notes: [
        { title: "研究テーマの確認", body: "先行研究を3本読む。\n\n比較軸を先に決める。", tags: ["研究", "進捗"], important: false, offset: 80000 },
        { title: "発表スライド案", body: "導入、仮説、手法、結果、考察。\n\n図を先に置く。", tags: ["発表", "資料"], important: true, offset: 170000 },
      ],
    },
    {
      title: "フランス語",
      color: "orange",
      notes: [
        { title: "動詞活用の確認", body: "etre / avoir / aller。\n\n不規則活用をまとめる。", tags: ["語彙", "活用"], important: false, offset: 70000 },
        { title: "会話表現", body: "Bonjour, comment ca va?\n\n授業中に使うフレーズ。", tags: ["会話", "基礎"], important: false, offset: 140000 },
      ],
    },
    {
      title: "試験対策",
      color: "pink",
      notes: [
        { title: "頻出問題", body: "定義問題と比較問題を中心に復習する。\n\n過去問を1周。", tags: ["試験", "復習"], important: true, offset: 60000 },
        { title: "暗記カード", body: "用語を短く。\n\n質問と答えを1行でまとめる。", tags: ["暗記", "まとめ"], important: false, offset: 110000 },
      ],
    },
  ];

  const notebookStates = [];
  const notes = [];
  const now = Date.now();
  let notebookOffset = 0;

  for (const def of notebookDefs) {
    const notebook = {
      id: crypto.randomUUID(),
      title: def.title,
      createdAt: now - notebookOffset,
      updatedAt: now - notebookOffset,
      color: def.color,
    };
    notebookStates.push(notebook);

    def.notes.forEach((noteDef, index) => {
      const timestamp = now - noteDef.offset - notebookOffset;
      notes.push({
        id: crypto.randomUUID(),
        notebookId: notebook.id,
        title: noteDef.title,
        subject: notebook.title,
        date: formatDateInput(timestamp),
        body: noteDef.body,
        bodyHtml: "",
        tags: noteDef.tags,
        important: noteDef.important,
        createdAt: timestamp,
        updatedAt: timestamp + index * 1000,
      });
    });

    notebookOffset += 360000;
  }

  return { notebooks: notebookStates, notes };
}

function ensureStateIntegrity(currentState) {
  if (!currentState.notebooks.length && !currentState.notes.length) {
    const seed = buildSampleSeed();
    currentState.notebooks = seed.notebooks;
    currentState.notes = seed.notes;
    currentState.selectedNotebookId = seed.notebooks[0]?.id ?? null;
    currentState.selectedNoteId = latestNoteForNotebookIn(seed.notes, seed.notebooks[0]?.id ?? "")?.id ?? null;
    return;
  }

  if (!currentState.notebooks.length) {
    currentState.notebooks.push(createNotebookObject("繝弱・繝医ヶ繝・け 1"));
  }

  const notebookIds = new Set(currentState.notebooks.map((notebook) => notebook.id));
  const fallbackNotebookId = currentState.notebooks[0].id;
  currentState.notes = currentState.notes.map((note) => {
    const normalized = normalizeNote(note, fallbackNotebookId);
    if (!notebookIds.has(normalized.notebookId)) {
      normalized.notebookId = fallbackNotebookId;
    }
    return normalized;
  });

  if (!currentState.selectedNotebookId || !notebookIds.has(currentState.selectedNotebookId)) {
    currentState.selectedNotebookId = fallbackNotebookId;
  }

  const notebookNotes = notesForNotebookIn(currentState.notes, currentState.selectedNotebookId);
  if (!currentState.selectedNoteId || !notebookNotes.some((note) => note.id === currentState.selectedNoteId)) {
    currentState.selectedNoteId = latestNoteForNotebookIn(currentState.notes, currentState.selectedNotebookId)?.id ?? null;
  }
}

function latestNoteForNotebookIn(notes, notebookId) {
  return notes
    .filter((note) => note.notebookId === notebookId)
    .slice()
    .sort((a, b) => b.updatedAt - a.updatedAt || b.createdAt - a.createdAt)[0] ?? null;
}

function saveToStorage() {
  if (!activeUser) return;

  const payload = JSON.stringify({
    version: 1,
    savedAt: new Date().toISOString(),
    notebooks: state.notebooks,
    notes: state.notes,
    selectedNotebookId: state.selectedNotebookId,
    selectedNoteId: state.selectedNoteId,
  });

  try {
    localStorage.setItem(userStorageKey(), payload);
  } catch {
    memoryStorageValue = payload;
  }

  scheduleOneDriveSave();
}

function scheduleOneDriveSave() {
  if (state.applyingOneDriveState || loadAuthProvider() !== "microsoft" || !readMicrosoftToken()) return;
  window.clearTimeout(state.oneDriveSaveTimer);
  setOneDriveStatus("保存待ち");
  state.oneDriveSaveTimer = window.setTimeout(() => {
    uploadStateToOneDrive();
  }, ONEDRIVE_SAVE_DELAY);
}

async function syncOneDrive({ preferCloud = false } = {}) {
  if (state.oneDriveSyncing || loadAuthProvider() !== "microsoft") return;
  const token = readMicrosoftToken();
  if (!token) {
    setOneDriveStatus("再ログインが必要です", true);
    return;
  }

  state.oneDriveSyncing = true;
  setOneDriveBusy(true, "同期中...");
  try {
    const cloudPayload = await downloadStateFromOneDrive(token.accessToken);
    if (cloudPayload && preferCloud) {
      if (shouldKeepLocalStateOverCloud(cloudPayload)) {
        await uploadStateToOneDrive(token.accessToken, { manageBusyState: false });
        setOneDriveStatus("ローカルの内容を保護しました");
        return;
      }
      applyOneDriveState(cloudPayload);
      setOneDriveStatus("同期済み");
    } else {
      await uploadStateToOneDrive(token.accessToken, { manageBusyState: false });
    }
  } catch (error) {
    console.warn("OneDrive sync error:", error);
    setOneDriveStatus(oneDriveErrorMessage(error), true);
  } finally {
    state.oneDriveSyncing = false;
    setOneDriveBusy(false);
  }
}

async function downloadStateFromOneDrive(accessToken) {
  const response = await fetch(oneDriveContentUrl(), {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (response.status === 404) return null;
  if (!response.ok) throw await createOneDriveError(response);

  const payload = await response.json();
  if (!payload || !Array.isArray(payload.notebooks) || !Array.isArray(payload.notes)) {
    throw new Error("OneDriveの保存データを読み取れませんでした。");
  }
  return payload;
}

async function uploadStateToOneDrive(accessToken = readMicrosoftToken()?.accessToken, options = {}) {
  if (!accessToken) {
    setOneDriveStatus("再ログインが必要です", true);
    return;
  }

  const manageBusyState = options.manageBusyState !== false;
  if (manageBusyState) setOneDriveBusy(true, "保存中...");
  window.clearTimeout(state.oneDriveSaveTimer);

  try {
    const response = await fetch(oneDriveContentUrl(), {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json; charset=utf-8",
      },
      body: serializedState(),
    });
    if (!response.ok) throw await createOneDriveError(response);
    setOneDriveStatus("OneDriveに保存済み");
  } catch (error) {
    console.warn("OneDrive save error:", error);
    setOneDriveStatus(oneDriveErrorMessage(error), true);
  } finally {
    if (manageBusyState) setOneDriveBusy(false);
  }
}

function applyOneDriveState(payload) {
  state.applyingOneDriveState = true;
  try {
    localStorage.setItem(userStorageKey(), JSON.stringify(payload));
    replaceState(loadState());
    if (PAGE === "notebooks") {
      ensureNotebookState();
      renderNotebookPage();
    } else {
      ensureNoteState();
      renderNotesPage();
    }
  } finally {
    state.applyingOneDriveState = false;
  }
}

function shouldKeepLocalStateOverCloud(cloudPayload) {
  const localPayload = readLocalStatePayload();
  if (!localPayload) return false;

  const localNotebookCount = safeArray(localPayload.notebooks).length;
  const cloudNotebookCount = safeArray(cloudPayload.notebooks).length;
  const localNoteCount = safeArray(localPayload.notes).length;
  const cloudNoteCount = safeArray(cloudPayload.notes).length;

  const cloudHasLessContent = cloudNotebookCount < localNotebookCount || cloudNoteCount < localNoteCount;
  if (!cloudHasLessContent) return false;

  const localSavedAt = statePayloadTimestamp(localPayload);
  const cloudSavedAt = statePayloadTimestamp(cloudPayload);
  return localSavedAt >= cloudSavedAt;
}

function readLocalStatePayload() {
  const raw = readStoredState();
  if (!raw) return null;
  try {
    const payload = JSON.parse(raw);
    if (!payload || !Array.isArray(payload.notebooks) || !Array.isArray(payload.notes)) return null;
    return payload;
  } catch {
    return null;
  }
}

function statePayloadTimestamp(payload) {
  const savedAt = Date.parse(payload?.savedAt || "");
  if (Number.isFinite(savedAt)) return savedAt;
  const notebookTimes = safeArray(payload?.notebooks).map((item) => Number(item.updatedAt) || 0);
  const noteTimes = safeArray(payload?.notes).map((item) => Number(item.updatedAt) || 0);
  return Math.max(0, ...notebookTimes, ...noteTimes);
}

function safeArray(value) {
  return Array.isArray(value) ? value : [];
}
function serializedState() {
  return JSON.stringify({
    version: 1,
    savedAt: new Date().toISOString(),
    notebooks: state.notebooks,
    notes: state.notes,
    selectedNotebookId: state.selectedNotebookId,
    selectedNoteId: state.selectedNoteId,
  });
}

function oneDriveContentUrl() {
  return `https://graph.microsoft.com/v1.0/me/drive/special/approot:/${encodeURIComponent(ONEDRIVE_STATE_FILE)}:/content`;
}

async function createOneDriveError(response) {
  const payload = await response.json().catch(() => ({}));
  const message = payload.error?.message || `OneDrive request failed (${response.status})`;
  const error = new Error(message);
  error.status = response.status;
  return error;
}

function oneDriveErrorMessage(error) {
  if (error?.status === 401) return "再ログインが必要です";
  if (error?.status === 403) return "OneDriveの保存権限がありません";
  return "同期に失敗しました";
}

function setOneDriveBusy(isBusy, message = "") {
  if (message) setOneDriveStatus(message);
}

function setOneDriveStatus(message, isError = false) {
  document.body.dataset.oneDriveStatus = isError ? "error" : message;
}

function readStoredState() {
  if (!activeUser) return null;

  try {
    const raw = localStorage.getItem(userStorageKey());
    if (raw) {
      memoryStorageValue = raw;
    }
    return raw || memoryStorageValue;
  } catch {
    return memoryStorageValue;
  }
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function sortNotes(a, b) {
  if (a.important !== b.important) return Number(b.important) - Number(a.important);
  return b.updatedAt - a.updatedAt;
}

function normalizeNotebook(notebook) {
  return {
    id: notebook.id || crypto.randomUUID(),
    title: notebook.title || "ノートブック",
    createdAt: Number(notebook.createdAt) || Number(notebook.updatedAt) || Date.now(),
    updatedAt: Number(notebook.updatedAt) || Date.now(),
    color: notebook.color || "",
    icon: notebook.icon || "book",
    important: Boolean(notebook.important),
  };
}

function createNotebookObject(title, color = "green", icon = "book") {
  return {
    id: crypto.randomUUID(),
    title,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    color,
    icon,
    important: false,
  };
}

function touchNotebook(notebookId, timestamp = Date.now()) {
  const notebook = state.notebooks.find((item) => item.id === notebookId);
  if (!notebook) return;

  const notebookNotes = state.notes.filter((note) => note.notebookId === notebookId);
  notebook.updatedAt = notebookNotes.length
    ? Math.max(timestamp, ...notebookNotes.map((note) => note.updatedAt))
    : timestamp;
}

function normalizeNote(note, fallbackNotebookId = "") {
  const body = String(note.body || "");
  return {
    id: note.id || crypto.randomUUID(),
    notebookId: note.notebookId || fallbackNotebookId,
    title: note.title || deriveTitle(body),
    date: note.date || today(),
    body,
    tags: Array.isArray(note.tags)
      ? note.tags.filter(Boolean)
      : String(note.tags || "")
          .split(",")
          .map((tag) => tag.trim())
          .filter(Boolean),
    important: Boolean(note.important),
    createdAt: Number(note.createdAt) || Number(note.updatedAt) || Date.now(),
    updatedAt: Number(note.updatedAt) || Date.now(),
  };
}

function createNoteObject(notebookId, title = "新しいノート") {
  return {
    id: crypto.randomUUID(),
    notebookId,
    title,
    date: today(),
    body: "",
    tags: [],
    important: false,
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
}

function ensureNotebookState() {
  if (!state.notebooks.length) {
    const notebook = createNotebookObject("ノートブック 1");
    state.notebooks.push(notebook);
    saveToStorage();
  }
}

function ensureDemoNotebookData() {
  const hasNotebookContent = state.notebooks.some((notebook) => notesForNotebookIn(state.notes, notebook.id).length > 0);
  if (state.notebooks.length > 0 && state.notes.length > 0 && hasNotebookContent) return;

  const seed = buildSampleSeed();
  state.notebooks = seed.notebooks;
  state.notes = seed.notes;
  state.selectedNotebookId = seed.notebooks[0]?.id ?? null;
  state.selectedNoteId = latestNoteForNotebookIn(seed.notes, seed.notebooks[0]?.id ?? "")?.id ?? null;
  saveToStorage();
}

function ensureNoteState() {
  if (!state.notebooks.length) {
    ensureNotebookState();
  }

  const requestedNotebookId = getNotebookIdFromQuery();
  if (requestedNotebookId && state.notebooks.some((notebook) => notebook.id === requestedNotebookId)) {
    state.selectedNotebookId = requestedNotebookId;
  } else if (!state.selectedNotebookId || !state.notebooks.some((notebook) => notebook.id === state.selectedNotebookId)) {
    state.selectedNotebookId = state.notebooks[0].id;
  }

  const requestedNoteId = getNoteIdFromQuery();

  if (!state.notes.length || !state.notes.some((note) => note.notebookId === state.selectedNotebookId)) {
    const note = createNoteObject(state.selectedNotebookId);
    state.notes.unshift(note);
    state.selectedNoteId = note.id;
    saveToStorage();
  }

  if (requestedNoteId && state.notes.some((note) => note.id === requestedNoteId && note.notebookId === state.selectedNotebookId)) {
    state.selectedNoteId = requestedNoteId;
  }

  if (!state.selectedNoteId || !state.notes.some((note) => note.id === state.selectedNoteId && note.notebookId === state.selectedNotebookId)) {
    state.selectedNoteId = latestNoteForNotebook(state.selectedNotebookId)?.id ?? null;
  }
}

function getNotebookIdFromQuery() {
  const params = new URLSearchParams(window.location.search);
  return params.get("notebookId");
}

function getNoteIdFromQuery() {
  const params = new URLSearchParams(window.location.search);
  return params.get("noteId");
}

function notesForNotebookIn(notes, notebookId) {
  return notes.filter((note) => note.notebookId === notebookId).sort(sortNotes);
}

function notesForNotebook(notebookId) {
  return notesForNotebookIn(state.notes, notebookId);
}

function currentNotebook() {
  return state.notebooks.find((notebook) => notebook.id === state.selectedNotebookId) ?? null;
}

function currentNote() {
  return state.notes.find((note) => note.id === state.selectedNoteId) ?? null;
}

function latestNoteForNotebook(notebookId) {
  return latestNoteForNotebookIn(state.notes, notebookId);
}

function createNotebook() {
  const title = (els.notebookNameInput?.value || "").trim() || "講義ノート";
  const color = selectedCreateOption(els.notebookColorChoices, "color") || "green";
  const icon = selectedCreateOption(els.notebookIconChoices, "icon") || "book";
  const notebook = createNotebookObject(title, color, icon);
  state.notebooks = [notebook, ...state.notebooks];
  state.selectedNotebookId = notebook.id;
  state.selectedNoteId = null;
  saveToStorage();
  renderNotebookPage();
  closeNotebookCreate();
}

function openNotebookCreate() {
  if (!els.notebookCreateOverlay) {
    createNotebook();
    return;
  }

  renderIconChoices();
  if (els.notebookNameInput) els.notebookNameInput.value = "講義ノート";
  setCreateChoice(els.notebookColorChoices, "color", "green");
  setCreateChoice(els.notebookIconChoices, "icon", "book");
  updateNotebookCreatePreview();
  els.notebookCreateOverlay.hidden = false;
  document.body.classList.add("is-creating-notebook");
  window.requestAnimationFrame(() => {
    els.notebookNameInput?.focus();
    els.notebookNameInput?.select();
  });
}

function closeNotebookCreate() {
  if (els.notebookCreateOverlay) els.notebookCreateOverlay.hidden = true;
  document.body.classList.remove("is-creating-notebook");
}

function submitNotebookCreate(event) {
  event.preventDefault();
  createNotebook();
}

function selectedCreateOption(container, key) {
  return container?.querySelector(".is-selected")?.dataset[key] || "";
}

function setCreateChoice(container, key, value) {
  if (!container) return;
  container.querySelectorAll("button").forEach((button) => {
    const selected = button.dataset[key] === value;
    button.classList.toggle("is-selected", selected);
    if (key === "color") button.textContent = selected ? "✓" : "";
    if (key === "icon") button.innerHTML = notebookIconSvg(button.dataset.icon);
  });
}

function selectNotebookColor(event) {
  const button = event.target.closest("button[data-color]");
  if (!button) return;
  setCreateChoice(els.notebookColorChoices, "color", button.dataset.color);
  updateNotebookCreatePreview();
}

function selectNotebookIcon(event) {
  const button = event.target.closest("button[data-icon]");
  if (!button) return;
  setCreateChoice(els.notebookIconChoices, "icon", button.dataset.icon);
  updateNotebookCreatePreview();
}

function renderIconChoices() {
  els.notebookIconChoices?.querySelectorAll("button[data-icon]").forEach((button) => {
    button.innerHTML = notebookIconSvg(button.dataset.icon);
  });
  const previewIcon = els.notebookCreatePreview?.querySelector(".create-preview-icon");
  if (previewIcon) previewIcon.innerHTML = notebookIconSvg(previewIcon.dataset.icon || "book");
}

function notebookIconSvg(icon) {
  const paths = {
    book: '<path d="M4 5.5A3.5 3.5 0 0 1 7.5 2H20v18H7.5A3.5 3.5 0 0 0 4 23.5V5.5Z"/><path d="M4 5.5A3.5 3.5 0 0 1 7.5 2H20"/><path d="M8 6h8"/>',
    doc: '<path d="M7 3h7l5 5v13H7V3Z"/><path d="M14 3v6h5"/><path d="M10 13h6"/><path d="M10 17h5"/>',
    bookmark: '<path d="M7 3h10v18l-5-3-5 3V3Z"/>',
    cap: '<path d="M3 9.5 12 5l9 4.5-9 4.5-9-4.5Z"/><path d="M7 12v4.5c2.6 2 7.4 2 10 0V12"/><path d="M21 10v5"/>',
    idea: '<path d="M9 18h6"/><path d="M10 22h4"/><path d="M8 10a4 4 0 1 1 8 0c0 2.5-2 3.3-2.5 6h-3C10 13.3 8 12.5 8 10Z"/>',
    case: '<path d="M4 8h16v12H4V8Z"/><path d="M9 8V5h6v3"/><path d="M4 13h16"/><path d="M10 13v2h4v-2"/>',
    folder: '<path d="M3 6h7l2 2h9v11H3V6Z"/><path d="M3 10h18"/>',
    star: '<path d="m12 3 2.8 5.7 6.2.9-4.5 4.4 1.1 6.2L12 17.3l-5.6 2.9 1.1-6.2L3 9.6l6.2-.9L12 3Z"/>',
  };
  return `<svg class="notebook-svg-icon" viewBox="0 0 24 24" aria-hidden="true">${paths[icon] || paths.book}</svg>`;
}

function updateNotebookCreatePreview() {
  const title = (els.notebookNameInput?.value || "").trim() || "講義ノート";
  const color = selectedCreateOption(els.notebookColorChoices, "color") || "green";
  const icon = selectedCreateOption(els.notebookIconChoices, "icon") || "book";
  if (els.notebookPreviewTitle) els.notebookPreviewTitle.textContent = title;
  if (els.notebookCreatePreview) els.notebookCreatePreview.dataset.kind = color;
  const previewIcon = els.notebookCreatePreview?.querySelector(".create-preview-icon");
  if (previewIcon) {
    previewIcon.dataset.icon = icon;
    previewIcon.innerHTML = notebookIconSvg(icon);
  }
}

function navigateToNotes(notebookId, noteId = null) {
  state.selectedNotebookId = notebookId;
  const url = new URL("./notes.html", window.location.href);
  url.searchParams.set("notebookId", notebookId);
  const initialNote = noteId || latestNoteForNotebook(notebookId)?.id;
  if (initialNote) url.searchParams.set("noteId", initialNote);
  window.location.href = url.toString();
}

function selectNotebook(notebookId) {
  state.selectedNotebookId = notebookId;
  saveToStorage();
  renderNotebookPage();
}

function createNote() {
  const notebookId = state.selectedNotebookId || state.notebooks[0].id;
  const note = createNoteObject(notebookId);
  state.notes = [note, ...state.notes];
  state.selectedNoteId = note.id;
  touchNotebook(notebookId, note.createdAt);
  saveToStorage();
  renderNotesPage();
  focusBody();
}

function deleteSelectedNote() {
  const current = currentNote();
  if (!current) return;
  deleteNoteById(current.id, { ensureReplacement: true });
}

function deleteNoteById(noteId, options = {}) {
  const current = state.notes.find((note) => note.id === noteId);
  if (!current) return;

  const ok = window.confirm(`「${current.title || "無題"}」を削除しますか？`);
  if (!ok) return;

  state.notes = state.notes.filter((note) => note.id !== current.id);
  if (state.selectedNoteId === current.id) {
    const notes = notesForNotebook(current.notebookId);
    state.selectedNoteId = notes[0]?.id ?? null;
  }
  if (options.ensureReplacement && !state.selectedNoteId) {
    const note = createNoteObject(state.selectedNotebookId);
    state.notes.unshift(note);
    state.selectedNoteId = note.id;
  }
  touchNotebook(current.notebookId);
  saveToStorage();
  renderCurrentPage();
}

function saveSelectedNote() {
  const changed = updateCurrentNoteFromEditor({ render: false });
  const note = currentNote();
  if (!note) return;

  saveToStorage();
  if (changed) renderNotesPage();
}

async function formatCurrentNoteWithAI() {
  const note = currentNote();
  if (!note || !els.bodyInput) return;

  if (window.location.protocol === "file:") {
    setAIFormatStatus("AI機能はローカルサーバーまたは公開サイトで利用できます。", "error");
    return;
  }

  const content = editorPlainText().trim();
  if (!content) {
    setAIFormatStatus("整える本文がありません。", "error");
    return;
  }

  if (content.length > 8000) {
    setAIFormatStatus("本文が長すぎます。8000文字以内にしてください。", "error");
    return;
  }

  setAIFormatLoading(true);
  setAIFormatStatus("整えています...", "loading");

  try {
    const response = await fetch("/api/ai/format-note", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ content }),
    });
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(readableAIError(payload));
    }

    replaceEditorWithMarkdown(payload.formatted || "");
    updateCurrentNoteFromEditor({ render: false });
    saveToStorage();
    renderNotesPage(false);
    setAIFormatStatus("整えました。", "success");
  } catch (error) {
    setAIFormatStatus(readableNetworkError(error), "error");
  } finally {
    setAIFormatLoading(false);
  }
}

function readableAIError(payload = {}) {
  const message = `${payload.error || ""} ${payload.detail || ""}`.toLowerCase();
  if (message.includes("gemini_api_key")) return "Gemini APIキーが未設定です。";
  if (message.includes("api key")) return "Gemini APIキーを確認してください。";
  if (message.includes("model") || message.includes("not found") || message.includes("unsupported")) {
    return "AIモデルが利用できません。モデル設定を確認してください。";
  }
  if (message.includes("timed out")) return "AIの応答が遅すぎます。少し短くして再試行してください。";
  return payload.error || "AI整形に失敗しました。";
}

function readableNetworkError(error) {
  const message = String(error?.message || "");
  if (message.includes("Failed to fetch") || message.includes("NetworkError")) {
    return "AIサーバーに接続できません。ローカルサーバーまたは公開サイトで開いているか確認してください。";
  }
  return message || "AI整形に失敗しました。";
}

function setAIFormatLoading(isLoading) {
  if (!els.aiFormatBtn) return;
  els.aiFormatBtn.disabled = isLoading;
  els.aiFormatBtn.textContent = isLoading ? "整え中..." : "AIで整える";
}

function setAIFormatStatus(message, kind = "") {
  if (!els.aiFormatStatus) return;
  els.aiFormatStatus.textContent = message;
  els.aiFormatStatus.dataset.kind = kind;
}

function replaceEditorWithPlainText(text) {
  els.bodyInput.innerHTML = "";
  els.bodyInput.appendChild(createPlainTextFragment(text));
}

function replaceEditorWithMarkdown(markdown) {
  els.bodyInput.innerHTML = "";
  els.bodyInput.appendChild(createMarkdownFragment(markdown));
}

function updateCurrentNoteFromEditor(options = {}) {
  const note = currentNote();
  if (!note) return false;

  const nextBody = getEditorBody();
  const nextTitle = deriveTitleFromEditor();
  if (note.body === nextBody) return false;

  note.body = nextBody;
  note.title = nextTitle;
  note.updatedAt = Date.now();
  state.notes.sort(sortNotes);
  touchNotebook(note.notebookId, note.updatedAt);
  if (options.render !== false) {
    renderNotesPage(false);
  }
  return true;
}

function scheduleSave() {
  window.clearTimeout(state.saveTimer);
  state.saveTimer = window.setTimeout(() => {
    saveSelectedNote();
  }, AUTOSAVE_DELAY);
}

function toggleInlineFormat(format) {
  if (!els.bodyInput || els.bodyInput.disabled) return;

  const commands = {
    bold: "bold",
    underline: "underline",
    strike: "strikeThrough",
  };
  applyRichCommand(commands[format]);
}

function applyTextColor(color) {
  if (!els.bodyInput || els.bodyInput.disabled || !color) return;
  applyRichCommand("foreColor", color);
}

function applyPandaHighlight() {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  restoreEditorSelection();
  const active = isPandaHighlightSelectionActive();
  els.bodyInput.focus();
  if (active) {
    removePandaHighlightFromSelection();
  } else {
    wrapSelectionWithElement("span", { class: "panda-highlight" });
  }
  updateCurrentNoteFromEditor();
  saveToStorage();
  updateFormattingActiveState();
}

function isPandaHighlightSelectionActive() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput?.contains(selection.anchorNode)) {
    return false;
  }
  return Boolean(activePandaHighlightNode()) || pandaHighlightNodesInRange(selection.getRangeAt(0)).length > 0;
}

function removePandaHighlightFromSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput?.contains(selection.anchorNode)) return;
  const range = selection.getRangeAt(0);
  const nodes = pandaHighlightNodesInRange(range);
  if (!nodes.length) {
    const node = activePandaHighlightNode();
    if (node) nodes.push(node);
  }
  nodes.forEach(unwrapElement);
}

function pandaHighlightNodesInRange(range) {
  const root = els.bodyInput;
  const nodes = [];
  root.querySelectorAll(".panda-highlight, span, font").forEach((node) => {
    if (!isPandaHighlightElement(node)) return;
    try {
      if (range.intersectsNode(node)) nodes.push(node);
    } catch {
      // Ignore detached nodes while the browser updates the selection.
    }
  });
  return nodes;
}

function isPandaHighlightElement(node) {
  if (!node || node === els.bodyInput) return false;
  if (node.classList?.contains("panda-highlight")) return true;
  const background = normalizeColorValue(node.style?.backgroundColor || "");
  const color = normalizeColorValue(node.style?.color || "");
  return (background === "#111111" || background === "#000000") && color === "#ffffff";
}

function unwrapElement(element) {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) {
    parent.insertBefore(element.firstChild, element);
  }
  element.remove();
}

function toggleHighlight() {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  const active = Boolean(activeHighlightNode());
  applyRichCommand("hiliteColor", active ? "transparent" : "#fff3a3");
}

function handleTextColorChoice(event) {
  const button = event.target.closest("button[data-color]");
  if (!button) return;
  applyTextColor(button.dataset.color);
}

function handleImageFileSelect(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  insertImageFile(file);
  event.target.value = "";
}

function handleEditorPaste(event) {
  const imageFile = Array.from(event.clipboardData?.files || []).find((file) => file.type.startsWith("image/"));
  event.preventDefault();
  if (imageFile) {
    insertImageFile(imageFile);
    return;
  }

  const shouldKeepFormatting = state.pasteRichOnce || (event.shiftKey && (event.ctrlKey || event.metaKey));
  state.pasteRichOnce = false;
  updateRichPasteButtonState();

  if (shouldKeepFormatting) {
    const html = event.clipboardData?.getData("text/html") || "";
    if (html) {
      insertRichHtmlAtSelection(html);
      return;
    }
  }

  const plainText = event.clipboardData?.getData("text/plain") || "";
  if (!plainText) return;
  insertPlainTextAtSelection(plainText);
}

function enableRichPasteOnce() {
  state.pasteRichOnce = !state.pasteRichOnce;
  updateRichPasteButtonState();
  closeToolbarMoreMenu();
  els.bodyInput?.focus();
  setAIFormatStatus(state.pasteRichOnce ? "次の貼り付けは書式を維持します。" : "", "");
}

function updateRichPasteButtonState() {
  if (!els.richPasteBtn) return;
  els.richPasteBtn.classList.toggle("is-active", state.pasteRichOnce);
  els.richPasteBtn.setAttribute("aria-pressed", String(state.pasteRichOnce));
}

function insertRichHtmlAtSelection(html) {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  const fragment = createSafeRichFragment(html);
  insertFragmentAtEditorSelection(fragment);
  updateCurrentNoteFromEditor();
  saveToStorage();
}

function handleEditorIndentKey(event) {
  if (event.key !== "Tab") return;
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  if (event.target.closest?.(".free-note")) return;

  event.preventDefault();
  adjustSelectedParagraphIndent(event.shiftKey ? -1 : 1);
}

function handleEditorEnterKey(event) {
  if (event.key !== "Enter" || event.shiftKey) return;
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  if (event.target.closest?.(".free-note")) return;

  const currentBlock = currentEditorBlock();
  if (currentBlock && Number(currentBlock.dataset.indentLevel || 0) > 0 && isEmptyEditorBlock(currentBlock)) {
    event.preventDefault();
    currentBlock.dataset.indentLevel = "0";
    resetParagraphTextSize(currentBlock);
    placeCaretAtEnd(currentBlock);
    updateCurrentNoteFromEditor({ render: false });
    saveToStorage();
    updateFormattingActiveState();
    return;
  }

  window.requestAnimationFrame(() => {
    const nextBlock = currentEditorBlock();
    if (!nextBlock) return;
    resetParagraphTextSize(nextBlock);
    updateCurrentNoteFromEditor({ render: false });
    saveToStorage();
    updateFormattingActiveState();
  });
}

function adjustSelectedParagraphIndent(delta) {
  const blocks = selectedEditorBlocks();
  if (!blocks.length) return;

  blocks.forEach((block) => {
    const currentLevel = Number(block.dataset.indentLevel || 0);
    const nextLevel = Math.max(0, Math.min(MAX_INDENT_LEVEL, currentLevel + delta));
    block.dataset.indentLevel = String(nextLevel);
  });

  updateCurrentNoteFromEditor({ render: false });
  saveToStorage();
}

function selectedEditorBlocks() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput?.contains(selection.anchorNode)) return [];

  if (selection.isCollapsed) {
    const block = closestEditorBlock(selection.anchorNode) || ensureCurrentParagraphBlock();
    return block ? [block] : [];
  }

  const range = selection.getRangeAt(0);
  const blocks = new Set();
  const startBlock = closestEditorBlock(range.startContainer);
  const endBlock = closestEditorBlock(range.endContainer);
  if (startBlock) blocks.add(startBlock);
  if (endBlock) blocks.add(endBlock);

  els.bodyInput.querySelectorAll("p, div, li, h1, h2, h3, h4, h5, h6").forEach((node) => {
    if (!isEditableParagraphBlock(node)) return;
    if (rangeIntersectsNode(range, node)) blocks.add(node);
  });

  return Array.from(blocks);
}

function currentEditorBlock() {
  const selection = window.getSelection();
  if (!selection || !selection.anchorNode || !els.bodyInput?.contains(selection.anchorNode)) return null;
  return closestEditorBlock(selection.anchorNode);
}

function closestEditorBlock(node) {
  if (!node || !els.bodyInput?.contains(node)) return null;
  const element = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  const block = element?.closest?.("p, div, li, h1, h2, h3, h4, h5, h6");
  return isEditableParagraphBlock(block) ? block : null;
}

function ensureCurrentParagraphBlock() {
  els.bodyInput.focus();
  document.execCommand("formatBlock", false, "div");
  const selection = window.getSelection();
  return selection?.anchorNode ? closestEditorBlock(selection.anchorNode) : null;
}

function isEditableParagraphBlock(node) {
  if (!node || node === els.bodyInput || !els.bodyInput?.contains(node)) return false;
  if (node.closest(".free-note, .note-image, .note-image-spacer")) return false;
  return true;
}

function isEmptyEditorBlock(block) {
  if (!block) return false;
  const text = block.textContent.replace(/[\u200b\u00a0]/g, "").trim();
  return !text && !block.querySelector("img, .note-image");
}

function resetParagraphTextSize(block) {
  if (!block) return;
  delete block.dataset.size;
  block.querySelectorAll("[data-size]").forEach((node) => {
    node.removeAttribute("data-size");
  });
}

function placeCaretAtEnd(node) {
  const range = document.createRange();
  range.selectNodeContents(node);
  range.collapse(false);
  const selection = window.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}

function rangeIntersectsNode(range, node) {
  try {
    return range.intersectsNode(node);
  } catch {
    return false;
  }
}

function insertPlainTextAtSelection(text) {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  const fragment = createPlainTextFragment(text);
  insertFragmentAtEditorSelection(fragment);
  updateCurrentNoteFromEditor();
  saveToStorage();
}

function createPlainTextFragment(text) {
  const fragment = document.createDocumentFragment();
  String(text).replace(/\r\n/g, "\n").split("\n").forEach((line, index) => {
    if (index > 0) fragment.appendChild(document.createElement("br"));
    fragment.appendChild(document.createTextNode(line));
  });
  return fragment;
}

function createSafeRichFragment(html) {
  const template = document.createElement("template");
  template.innerHTML = sanitizeEditorHtml(html);
  const fragment = document.createDocumentFragment();
  Array.from(template.content.childNodes).forEach((node) => {
    appendSanitizedRichNode(fragment, node);
  });
  return fragment;
}

function appendSanitizedRichNode(parent, node) {
  if (node.nodeType === Node.TEXT_NODE) {
    parent.appendChild(document.createTextNode(node.textContent || ""));
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;

  const tag = node.tagName.toLowerCase();
  if (tag === "br") {
    parent.appendChild(document.createElement("br"));
    return;
  }

  if (["script", "style", "iframe", "object", "embed", "meta", "link"].includes(tag)) return;

  const mappedTag = richPasteTagName(tag);
  const element = document.createElement(mappedTag);
  copySafeRichStyles(node, element);

  Array.from(node.childNodes).forEach((child) => appendSanitizedRichNode(element, child));

  if (mappedTag === "p" && !element.childNodes.length) {
    element.appendChild(document.createElement("br"));
  }
  parent.appendChild(element);
}

function richPasteTagName(tag) {
  const allowed = {
    b: "strong",
    strong: "strong",
    u: "u",
    s: "s",
    strike: "s",
    del: "s",
    em: "span",
    i: "span",
    font: "span",
    span: "span",
    p: "p",
    div: "p",
    h1: "div",
    h2: "div",
    h3: "div",
    ul: "ul",
    ol: "ol",
    li: "li",
    blockquote: "p",
    code: "span",
    pre: "p",
  };
  return allowed[tag] || "span";
}

function copySafeRichStyles(source, target) {
  const tag = source.tagName.toLowerCase();
  const pastedTextSize = richPasteTextSize(source);
  if (pastedTextSize) target.dataset.size = pastedTextSize;
  if (tag === "blockquote") target.dataset.indentLevel = "2";

  const style = source.style;
  const fontWeight = style.fontWeight || "";
  if (fontWeight === "bold" || Number(fontWeight) >= 600) {
    target.style.fontWeight = "800";
  }
  if (style.textDecorationLine.includes("underline") || style.textDecoration.includes("underline")) {
    target.style.textDecoration = "underline";
  }
  if (style.textDecorationLine.includes("line-through") || style.textDecoration.includes("line-through")) {
    target.style.textDecoration = "line-through";
  }
  if (style.color) target.style.color = style.color;
  if (style.backgroundColor && style.backgroundColor !== "transparent") {
    target.style.backgroundColor = style.backgroundColor;
  }
  if (style.textTransform === "uppercase") {
    target.style.textTransform = "uppercase";
  }
}

function richPasteTextSize(source) {
  const existingSize = source.dataset?.size;
  if (["1", "2", "3"].includes(existingSize)) return existingSize;

  const tag = source.tagName.toLowerCase();
  if (tag === "h1" || tag === "h2") return "3";
  if (tag === "h3") return "2";

  const legacyFontSize = source.getAttribute("size");
  if (tag === "font" && legacyFontSize) {
    const numericSize = Number(legacyFontSize);
    if (numericSize >= 5) return "3";
    if (numericSize === 4) return "2";
    if (numericSize > 0) return "1";
  }

  const pixelSize = cssFontSizeToPixels(source.style?.fontSize);
  if (!pixelSize) return "";
  if (pixelSize >= 22) return "3";
  if (pixelSize >= 17.5) return "2";
  return "1";
}

function cssFontSizeToPixels(value) {
  const fontSize = String(value || "").trim().toLowerCase();
  if (!fontSize) return 0;

  const keywordSizes = {
    "xx-small": 9,
    "x-small": 10,
    small: 13,
    medium: 16,
    large: 18,
    "x-large": 24,
    "xx-large": 32,
    "xxx-large": 48,
  };
  if (keywordSizes[fontSize]) return keywordSizes[fontSize];

  const match = fontSize.match(/^(-?\d*\.?\d+)(px|pt|rem|em|%)$/);
  if (!match) return 0;

  const amount = Number(match[1]);
  const unit = match[2];
  if (!Number.isFinite(amount) || amount <= 0) return 0;
  if (unit === "px") return amount;
  if (unit === "pt") return amount * (4 / 3);
  if (unit === "rem" || unit === "em") return amount * 16;
  if (unit === "%") return amount * 0.16;
  return 0;
}

function createMarkdownFragment(markdown) {
  const fragment = document.createDocumentFragment();
  const lines = String(markdown || "").replace(/\r\n/g, "\n").split("\n");
  let list = null;

  lines.forEach((rawLine) => {
    const line = rawLine.trimEnd();
    const trimmed = line.trim();

    if (!trimmed) {
      list = null;
      fragment.appendChild(document.createElement("br"));
      return;
    }

    const heading = trimmed.match(/^(#{1,3})\s+(.+)$/);
    if (heading) {
      list = null;
      const block = document.createElement("div");
      block.dataset.size = heading[1].length === 1 ? "3" : "2";
      block.dataset.indentLevel = "0";
      appendInlineMarkdown(block, heading[2]);
      fragment.appendChild(block);
      return;
    }

    const bullet = trimmed.match(/^[-*]\s+(.+)$/);
    if (bullet) {
      if (!list) {
        list = document.createElement("ul");
        fragment.appendChild(list);
      }
      const item = document.createElement("li");
      item.dataset.indentLevel = "2";
      appendInlineMarkdown(item, bullet[1]);
      list.appendChild(item);
      return;
    }

    list = null;
    const paragraph = document.createElement("p");
    paragraph.dataset.indentLevel = "1";
    appendInlineMarkdown(paragraph, trimmed);
    fragment.appendChild(paragraph);
  });

  return fragment;
}

function appendInlineMarkdown(parent, text) {
  String(text || "").split(/(\*\*[^*]+\*\*)/g).forEach((part) => {
    const bold = part.match(/^\*\*([^*]+)\*\*$/);
    if (bold) {
      const strong = document.createElement("strong");
      strong.textContent = bold[1];
      parent.appendChild(strong);
      return;
    }
    parent.appendChild(document.createTextNode(part));
  });
}

async function insertImageFile(file) {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  if (!file.type.startsWith("image/")) return;

  const dataUrl = await imageFileToDataUrl(file);
  insertImageAtSelection(dataUrl, file.name || "image");
}

function imageFileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resizeImageDataUrl(String(reader.result || ""), resolve, reject));
    reader.addEventListener("error", reject);
    reader.readAsDataURL(file);
  });
}

function resizeImageDataUrl(dataUrl, resolve, reject) {
  const image = new Image();
  image.addEventListener("load", () => {
    const maxWidth = 1200;
    const scale = Math.min(1, maxWidth / image.naturalWidth);
    const width = Math.round(image.naturalWidth * scale);
    const height = Math.round(image.naturalHeight * scale);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    canvas.getContext("2d").drawImage(image, 0, 0, width, height);
    resolve(canvas.toDataURL("image/jpeg", 0.82));
  });
  image.addEventListener("error", reject);
  image.src = dataUrl;
}

function insertImageAtSelection(dataUrl, altText = "image") {
  els.bodyInput.focus();
  const figure = document.createElement("figure");
  figure.className = "note-image";
  figure.contentEditable = "false";

  const image = document.createElement("img");
  image.src = dataUrl;
  image.alt = altText;
  image.loading = "lazy";

  figure.appendChild(image);
  ensureImageResizeHandle(figure);
  insertNodeAtEditorSelection(figure);
  updateCurrentNoteFromEditor();
  saveToStorage();
}

function prepareEditorImages() {
  if (!els.bodyInput) return;
  els.bodyInput.querySelectorAll(".note-image").forEach((figure) => {
    figure.contentEditable = "false";
    ensureImageResizeHandle(figure);
    normalizeImageSpacer(figure);
  });
}

function prepareFreeNotes() {
  if (!els.bodyInput) return;
  els.bodyInput.querySelectorAll(".free-note").forEach((note) => {
    note.contentEditable = "true";
    note.setAttribute("data-placeholder", "繝｡繝｢");
  });
}

function createFreeNoteAtCursor() {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  const freeNote = document.createElement("aside");
  freeNote.className = "free-note";
  freeNote.contentEditable = "true";
  freeNote.dataset.placeholder = "繝｡繝｢";
  freeNote.style.top = `${getCurrentEditorLineTop()}px`;
  freeNote.appendChild(document.createElement("br"));
  els.bodyInput.appendChild(freeNote);
  focusFreeNote(freeNote);
  updateCurrentNoteFromEditor();
  saveToStorage();
}

function getCurrentEditorLineTop() {
  const fallbackTop = 24;
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput.contains(selection.anchorNode)) return fallbackTop;

  const range = selection.getRangeAt(0).cloneRange();
  let rect = range.getBoundingClientRect();
  let marker = null;

  if (!rect || (!rect.top && !rect.height)) {
    marker = document.createElement("span");
    marker.textContent = "\u200b";
    range.insertNode(marker);
    rect = marker.getBoundingClientRect();
  }

  const editorRect = els.bodyInput.getBoundingClientRect();
  const top = Math.max(fallbackTop, rect.top - editorRect.top + els.bodyInput.scrollTop);
  marker?.remove();
  return Math.round(top);
}

function focusFreeNote(freeNote) {
  window.requestAnimationFrame(() => {
    freeNote.focus();
    const range = document.createRange();
    range.selectNodeContents(freeNote);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  });
}

function normalizeImageSpacer(figure) {
  const next = figure.nextElementSibling;
  if (next?.classList.contains("note-image-spacer")) return;
  if (next?.tagName === "P" && !next.textContent.trim() && !next.querySelector("img")) {
    next.className = "note-image-spacer";
    return;
  }
  if (!next) {
    figure.after(createImageSpacer());
  }
}

function createImageSpacer() {
  const spacer = document.createElement("div");
  spacer.className = "note-image-spacer";
  spacer.appendChild(document.createElement("br"));
  return spacer;
}

function ensureImageResizeHandle(figure) {
  if (figure.querySelector(".note-image-resize-handle")) return;
  const handle = document.createElement("span");
  handle.className = "note-image-resize-handle";
  handle.setAttribute("aria-hidden", "true");
  figure.appendChild(handle);
}

function handleEditorImageClick(event) {
  const figure = event.target.closest?.(".note-image");
  selectEditorImage(figure);
}

function handleEditorImageDeleteKey(event) {
  if (event.key !== "Backspace" && event.key !== "Delete") return;
  const selectedImage = els.bodyInput?.querySelector(".note-image.is-selected");
  if (!selectedImage) return;
  event.preventDefault();
  deleteEditorImage(selectedImage);
}

function deleteEditorImage(figure) {
  if (!figure) return;
  const spacer = figure.nextElementSibling;
  figure.remove();
  if (spacer?.classList.contains("note-image-spacer") && !spacer.textContent.trim()) {
    spacer.remove();
  }
  updateCurrentNoteFromEditor();
  saveToStorage();
}

function selectEditorImage(figure) {
  els.bodyInput?.querySelectorAll(".note-image.is-selected").forEach((item) => {
    if (item !== figure) item.classList.remove("is-selected");
  });
  if (figure) figure.classList.add("is-selected");
}

function handleImageResizeStart(event) {
  const handle = event.target.closest?.(".note-image-resize-handle");
  if (!handle) return;

  const figure = handle.closest(".note-image");
  if (!figure) return;
  event.preventDefault();
  event.stopPropagation();
  selectEditorImage(figure);

  state.imageResize = {
    figure,
    startX: event.clientX,
    startWidth: figure.getBoundingClientRect().width,
    maxWidth: els.bodyInput.getBoundingClientRect().width - 56,
  };
  document.body.classList.add("is-resizing-image");
}

function handleImageResizeMove(event) {
  if (!state.imageResize) return;
  event.preventDefault();

  const nextWidth = Math.max(140, Math.min(state.imageResize.maxWidth, state.imageResize.startWidth + event.clientX - state.imageResize.startX));
  state.imageResize.figure.style.width = `${Math.round(nextWidth)}px`;
}

function handleImageResizeEnd() {
  if (!state.imageResize) return;

  updateCurrentNoteFromEditor();
  saveToStorage();
  state.imageResize = null;
  document.body.classList.remove("is-resizing-image");
}

function insertNodeAtEditorSelection(node) {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput.contains(selection.anchorNode)) {
    els.bodyInput.appendChild(node);
  } else {
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(node);
  }

  const spacer = createImageSpacer();
  node.after(spacer);

  const nextRange = document.createRange();
  nextRange.setStart(spacer, 0);
  nextRange.collapse(true);
  selection?.removeAllRanges();
  selection?.addRange(nextRange);
}

function insertFragmentAtEditorSelection(fragment) {
  els.bodyInput.focus();
  const lastNode = fragment.lastChild;
  if (!lastNode) return;

  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !els.bodyInput.contains(selection.anchorNode)) {
    els.bodyInput.appendChild(fragment);
  } else {
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(fragment);
  }

  const nextRange = document.createRange();
  nextRange.setStartAfter(lastNode);
  nextRange.collapse(true);
  selection?.removeAllRanges();
  selection?.addRange(nextRange);
}

function applyTextSize(size) {
  if (!els.bodyInput || els.bodyInput.getAttribute("aria-disabled") === "true") return;
  wrapSelectionWithElement("span", { "data-size": size });
}

function applyRichCommand(command, value = null) {
  if (!command) return;
  els.bodyInput.focus();
  document.execCommand(command, false, value);
  updateCurrentNoteFromEditor();
  saveToStorage();
  updateFormattingActiveState();
}

function wrapSelectionWithElement(tagName, attributes = {}) {
  els.bodyInput.focus();
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || selection.isCollapsed) return null;

  const range = selection.getRangeAt(0);
  if (!els.bodyInput.contains(range.commonAncestorContainer)) return null;

  const wrapper = document.createElement(tagName);
  Object.entries(attributes).forEach(([key, value]) => wrapper.setAttribute(key, value));

  try {
    wrapper.appendChild(range.extractContents());
    range.insertNode(wrapper);
    selection.removeAllRanges();
    const nextRange = document.createRange();
    nextRange.selectNodeContents(wrapper);
    selection.addRange(nextRange);
    updateCurrentNoteFromEditor();
    saveToStorage();
    updateFormattingActiveState();
    return wrapper;
  } catch {
    // Keep editor stable if the browser cannot wrap the current complex selection.
    return null;
  }
}

function updateFormattingActiveState() {
  if (!els.bodyInput || document.activeElement !== els.bodyInput) return;

  setFormatActive(els.boldBtn, document.queryCommandState("bold"));
  setFormatActive(els.underlineBtn, document.queryCommandState("underline"));
  setFormatActive(els.strikeBtn, document.queryCommandState("strikeThrough"));
  setFormatActive(els.highlightBtn, Boolean(activeHighlightNode()));
  setFormatActive(els.pandaHighlightBtn, isPandaHighlightSelectionActive());
  updateColorChipActiveState();
  const size = activeTextSize();
  setFormatActive(els.sizeOneBtn, size === "1");
  setFormatActive(els.sizeTwoBtn, size === "2");
  setFormatActive(els.sizeThreeBtn, size === "3");
}

function updateColorChipActiveState() {
  if (!els.textColorChoices) return;
  const activeColor = normalizeColorValue(document.queryCommandValue("foreColor"));
  els.textColorChoices.querySelectorAll("button[data-color]").forEach((button) => {
    const active = normalizeColorValue(button.dataset.color) === activeColor;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });
}

function normalizeColorValue(value) {
  const color = String(value || "").trim().toLowerCase();
  if (!color) return "";
  if (color.startsWith("#")) return color;
  const rgb = color.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (!rgb) return color;
  return `#${rgb.slice(1).map((part) => Number(part).toString(16).padStart(2, "0")).join("")}`;
}

function activeTextSize() {
  const selection = window.getSelection();
  if (!selection || !selection.anchorNode || !els.bodyInput?.contains(selection.anchorNode)) return "";
  const node = selection.anchorNode.nodeType === Node.ELEMENT_NODE ? selection.anchorNode : selection.anchorNode.parentElement;
  return node?.closest?.("[data-size]")?.dataset.size || "";
}

function activeHighlightNode() {
  const selection = window.getSelection();
  if (!selection || !selection.anchorNode || !els.bodyInput?.contains(selection.anchorNode)) return null;
  let node = selection.anchorNode.nodeType === Node.ELEMENT_NODE ? selection.anchorNode : selection.anchorNode.parentElement;
  while (node && node !== els.bodyInput) {
    const background = node.style?.backgroundColor || "";
    if (background && background !== "transparent") return node;
    node = node.parentElement;
  }
  return null;
}

function activePandaHighlightNode() {
  const selection = window.getSelection();
  if (!selection || !selection.anchorNode || !els.bodyInput?.contains(selection.anchorNode)) return null;
  let node = selection.anchorNode.nodeType === Node.ELEMENT_NODE ? selection.anchorNode : selection.anchorNode.parentElement;
  while (node && node !== els.bodyInput) {
    if (isPandaHighlightElement(node)) return node;
    node = node.parentElement;
  }
  return null;
}

function getEditorBody() {
  return sanitizeEditorHtml(cleanEditorChrome(els.bodyInput?.innerHTML || ""));
}

function cleanEditorChrome(html) {
  const container = document.createElement("div");
  container.innerHTML = html;
  container.querySelectorAll(".note-image-resize-handle").forEach((handle) => handle.remove());
  container.querySelectorAll(".note-image.is-selected").forEach((figure) => figure.classList.remove("is-selected"));
  return container.innerHTML;
}

function sanitizeEditorHtml(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+="[^"]*"/gi, "")
    .replace(/\son\w+='[^']*'/gi, "");
}

function editorPlainText(html = getEditorBody()) {
  const container = document.createElement("div");
  container.innerHTML = String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(div|p|li|h[1-6])>/gi, "\n");
  container.querySelectorAll(".free-note").forEach((note) => note.remove());
  return (container.textContent || "").replace(/\u00a0/g, " ");
}

function deriveTitleFromEditor() {
  return deriveTitle(editorPlainText());
}

function setFormatActive(button, active) {
  if (!button) return;
  button.classList.toggle("is-active", active);
  button.setAttribute("aria-pressed", String(active));
}

function deriveTitle(body) {
  const firstLine = String(body || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(Boolean);
  if (!firstLine) return "無題のノート";

  const firstSentence = firstLine.match(/^.*?[。！？?](?=\s|$|[\u3000]|.)/)?.[0];
  return (firstSentence || firstLine).trim() || "無題のノート";
}

function formatDateTime(timestamp) {
  const date = new Date(timestamp);
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${year}/${month}/${day} ${hours}:${minutes}`;
}

function formatRelativeTime(timestamp) {
  const diff = Date.now() - timestamp;
  const minutes = Math.floor(diff / 60000);
  if (minutes < 1) return "たった今";
  if (minutes < 60) return `${minutes}分前`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}時間前`;
  const days = Math.floor(hours / 24);
  return `${days}日前`;
}

function notebookColorKey(notebook) {
  if (!notebook) return "green";
  if (notebook.color) return notebook.color;
  const index = state.notebooks.findIndex((item) => item.id === notebook.id);
  return ["violet", "blue", "green", "orange", "pink"][Math.abs(index) % 5];
}

function notebookTheme(notebook) {
  return NOTEBOOK_THEMES[notebookColorKey(notebook)] || NOTEBOOK_THEMES.green;
}

function applyCurrentNotebookTheme() {
  if (PAGE !== "notes") return;
  const theme = notebookTheme(currentNotebook());
  const root = document.documentElement;
  root.style.setProperty("--accent", theme.accent);
  root.style.setProperty("--accent-soft", theme.soft);
  root.style.setProperty("--accent-border", theme.border);
  root.style.setProperty("--accent-focus", theme.focus);
}

function notebookTags(notebookId) {
  const tags = notesForNotebook(notebookId)
    .flatMap((note) => note.tags || [])
    .filter(Boolean);
  return [...new Set(tags)].slice(0, 4);
}

function renderNotebookPage() {
  const query = (els.notebookSearchInput?.value || "").trim().toLowerCase();
  const notebooks = state.notebooks
    .slice()
    .sort((a, b) => Number(b.important) - Number(a.important) || b.updatedAt - a.updatedAt)
    .filter((notebook) => {
      if (!query) return true;
      const haystack = [notebook.title, ...notesForNotebook(notebook.id).map((note) => note.title), ...notesForNotebook(notebook.id).map((note) => note.body)].join(" ").toLowerCase();
      return haystack.includes(query);
    });

  if (els.notebookCount) els.notebookCount.textContent = String(notebooks.length);
  if (els.notebookList) els.notebookList.innerHTML = "";
  const visibleNotebooks = notebooks.slice(0, state.notebookLimit);

  if (!notebooks.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "ノートブックがありません。";
    els.notebookList.appendChild(empty);
  } else {
    const fragment = document.createDocumentFragment();
    visibleNotebooks.forEach((notebook) => fragment.appendChild(renderNotebookItem(notebook)));
    els.notebookList.appendChild(fragment);
  }

  if (els.showMoreNotebooksBtn) {
    els.showMoreNotebooksBtn.hidden = notebooks.length <= NOTEBOOK_INITIAL_LIMIT;
    els.showMoreNotebooksBtn.textContent = state.notebookLimit >= notebooks.length ? "閉じる" : "さらに表示";
    els.showMoreNotebooksBtn.classList.toggle("is-open", state.notebookLimit >= notebooks.length);
  }

  renderRecentNotes();
}

function showMoreNotebooks() {
  state.notebookLimit = state.notebookLimit >= state.notebooks.length ? NOTEBOOK_INITIAL_LIMIT : Number.MAX_SAFE_INTEGER;
  renderNotebookPage();
}

function showMoreRecentNotes() {
  state.recentLimit = state.recentLimit >= RECENT_MAX_LIMIT ? RECENT_INITIAL_LIMIT : RECENT_MAX_LIMIT;
  renderRecentNotes();
}

function renderNotebookItem(notebook) {
  const card = document.createElement("div");
  card.className = "notebook-card";
  card.dataset.kind = notebookColorKey(notebook);
  card.setAttribute("role", "button");
  card.tabIndex = 0;
  card.classList.toggle("is-active", notebook.id === state.selectedNotebookId);
  card.classList.toggle("is-important", Boolean(notebook.important));

  const accent = document.createElement("div");
  accent.className = "notebook-cover-accent";

  const coverIcon = document.createElement("div");
  coverIcon.className = "notebook-cover-icon";
  coverIcon.dataset.icon = notebook.icon || "book";
  coverIcon.innerHTML = notebookIconSvg(notebook.icon);

  const actions = document.createElement("div");
  actions.className = "notebook-card-actions";

  const importantMark = document.createElement("span");
  importantMark.className = "notebook-important-mark";
  importantMark.textContent = "★";

  const menuButton = document.createElement("button");
  menuButton.type = "button";
  menuButton.className = "notebook-menu-button";
  menuButton.setAttribute("aria-label", "ノートブックメニュー");
  menuButton.setAttribute("aria-expanded", String(state.notebookMenuId === notebook.id));
  menuButton.textContent = "…";

  const menu = document.createElement("div");
  menu.className = "notebook-menu";
  menu.hidden = state.notebookMenuId !== notebook.id;

  const importantButton = document.createElement("button");
  importantButton.type = "button";
  importantButton.className = "notebook-menu-item notebook-important";
  importantButton.textContent = notebook.important ? "重要を解除" : "重要";

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "notebook-menu-item danger notebook-delete";
  deleteButton.textContent = "削除";

  menu.appendChild(importantButton);
  menu.appendChild(deleteButton);
  actions.appendChild(importantMark);
  actions.appendChild(menuButton);
  actions.appendChild(menu);

  const info = document.createElement("div");
  info.className = "notebook-card-info";

  const title = document.createElement("h3");
  title.className = "notebook-card-title";
  title.textContent = notebook.title;

  const meta = document.createElement("div");
  meta.className = "notebook-card-meta";

  const count = document.createElement("span");
  count.className = "notebook-meta-item";
  count.textContent = `${notesForNotebook(notebook.id).length} notes`;

  const line = document.createElement("div");
  line.className = "notebook-card-line";

  card.appendChild(accent);
  card.appendChild(coverIcon);
  card.appendChild(actions);
  meta.appendChild(count);
  info.appendChild(title);
  info.appendChild(line);
  info.appendChild(meta);
  card.appendChild(info);

  menuButton.addEventListener("click", (event) => {
    event.stopPropagation();
    state.notebookMenuId = state.notebookMenuId === notebook.id ? null : notebook.id;
    renderNotebookPage();
  });

  importantButton.addEventListener("click", (event) => {
    event.stopPropagation();
    state.notebookMenuId = null;
    toggleNotebookImportant(notebook.id);
  });

  deleteButton.addEventListener("click", (event) => {
    event.stopPropagation();
    deleteNotebookById(notebook.id);
  });

  card.addEventListener("click", (event) => {
    if (event.target.closest(".notebook-card-actions")) return;
    state.notebookMenuId = null;
    selectNotebook(notebook.id);
    navigateToNotes(notebook.id);
  });

  card.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    if (event.target.closest(".notebook-card-actions")) return;
    event.preventDefault();
    state.notebookMenuId = null;
    selectNotebook(notebook.id);
    navigateToNotes(notebook.id);
  });

  return card;
}

function toggleNotebookImportant(notebookId) {
  const notebook = state.notebooks.find((item) => item.id === notebookId);
  if (!notebook) return;

  notebook.important = !notebook.important;
  notebook.updatedAt = Date.now();
  saveToStorage();
  renderNotebookPage();
}

function deleteNotebookById(notebookId) {
  const notebook = state.notebooks.find((item) => item.id === notebookId);
  if (!notebook) return;

  const ok = window.confirm("Delete this notebook?");
  if (!ok) return;

  state.notebooks = state.notebooks.filter((item) => item.id !== notebookId);
  state.notes = state.notes.filter((note) => note.notebookId !== notebookId);
  state.notebookMenuId = null;

  if (!state.notebooks.length) {
    const nextNotebook = createNotebookObject("ノートブック 1");
    const nextNote = createNoteObject(nextNotebook.id);
    state.notebooks = [nextNotebook];
    state.notes = [nextNote];
    state.selectedNotebookId = nextNotebook.id;
    state.selectedNoteId = nextNote.id;
  } else if (state.selectedNotebookId === notebookId) {
    const nextNotebook = state.notebooks.slice().sort((a, b) => b.updatedAt - a.updatedAt)[0];
    state.selectedNotebookId = nextNotebook.id;
    state.selectedNoteId = latestNoteForNotebook(nextNotebook.id)?.id ?? null;
  }

  saveToStorage();
  renderNotebookPage();
}

function renderRecentNotes() {
  if (!els.recentNoteList) return;
  els.recentNoteList.innerHTML = "";

  const recentNotes = state.notes
    .slice()
    .sort((a, b) => b.updatedAt - a.updatedAt);
  const recent = recentNotes.slice(0, state.recentLimit);

  if (!recent.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "最近のノートはありません。";
    els.recentNoteList.appendChild(empty);
    if (els.showMoreRecentBtn) els.showMoreRecentBtn.hidden = true;
    return;
  }

  const fragment = document.createDocumentFragment();
  recent.forEach((note) => fragment.appendChild(renderRecentNoteItem(note)));
  els.recentNoteList.appendChild(fragment);

  if (els.showMoreRecentBtn) {
    els.showMoreRecentBtn.hidden = recentNotes.length <= RECENT_INITIAL_LIMIT;
    els.showMoreRecentBtn.textContent = state.recentLimit >= RECENT_MAX_LIMIT ? "閉じる" : "さらに表示";
    els.showMoreRecentBtn.classList.toggle("is-open", state.recentLimit >= RECENT_MAX_LIMIT);
  }
}

function renderRecentNoteItem(note) {
  const node = els.recentNoteTemplate.content.firstElementChild.cloneNode(true);
  node.querySelector(".recent-note-title").textContent = note.title || "無題のノート";
  node.querySelector(".recent-note-time").textContent = formatRelativeTime(note.updatedAt);
  node.querySelector(".recent-note-star").style.opacity = note.important ? "1" : "0.25";

  const tags = node.querySelector(".recent-note-tags");
  const notebook = state.notebooks.find((item) => item.id === note.notebookId);
  if (notebook) {
    const chip = document.createElement("span");
    chip.className = "recent-chip";
    chip.dataset.kind = notebookColorKey(notebook);
    chip.textContent = notebook.title;
    tags.appendChild(chip);
  }

  node.addEventListener("click", () => {
    navigateToNotes(note.notebookId, note.id);
  });

  return node;
}

function renderNotesPage(updateEditor = true) {
  applyCurrentNotebookTheme();

  if (els.notebookTitle) {
    els.notebookTitle.textContent = currentNotebook()?.title || "ノート";
  }

  const notes = notesForNotebook(state.selectedNotebookId);
  if (els.noteCount) els.noteCount.textContent = String(notes.length);
  if (els.noteList) els.noteList.innerHTML = "";

  if (!notes.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "ノートがありません。";
    els.noteList.appendChild(empty);
  } else {
    const fragment = document.createDocumentFragment();
    notes.forEach((note) => fragment.appendChild(renderNoteItem(note)));
    els.noteList.appendChild(fragment);
  }

  if (updateEditor) renderEditor();
}

function renderNoteItem(note) {
  const node = els.noteItemTemplate.content.firstElementChild.cloneNode(true);
  node.classList.toggle("is-active", note.id === state.selectedNoteId);
  node.classList.toggle("is-important", note.important);

  node.querySelector(".note-item-title").textContent = deriveTitle(note.title || note.body);
  node.querySelector(".note-item-date").textContent = note.date || today();

  node.addEventListener("click", () => {
    state.selectedNoteId = note.id;
    renderNotesPage();
  });

  const star = node.querySelector(".note-item-star");
  star.addEventListener("click", (event) => {
    event.stopPropagation();
    toggleImportant(note.id);
  });

  return node;
}

function renderEditor() {
  if (!els.editorTitle) return;

  const note = currentNote();
  const editable = Boolean(note);

  if (els.deleteNoteBtn) els.deleteNoteBtn.disabled = !editable;
  if (els.bodyInput) {
    els.bodyInput.contentEditable = String(editable);
    els.bodyInput.setAttribute("aria-disabled", String(!editable));
  }

  if (!note) {
    els.editorTitle.textContent = currentNotebook()?.title || "ノートを選択";
    if (els.dateLabel) els.dateLabel.textContent = "";
    clearEditor();
    return;
  }

  els.editorTitle.textContent = note.title || "無題のノート";
  if (els.dateLabel) els.dateLabel.textContent = note.date || today();
  els.bodyInput.innerHTML = note.body;
  prepareEditorImages();
  prepareFreeNotes();
  updateEditorHeader();
  updateFormattingActiveState();
}

function updateEditorHeader() {
  const note = currentNote();
  if (!note) return;
  els.editorTitle.textContent = note.title || "無題のノート";
  if (els.dateLabel) els.dateLabel.textContent = note.date || today();
}

function toggleImportant(noteId) {
  const note = state.notes.find((item) => item.id === noteId);
  if (!note) return;

  note.important = !note.important;
  state.notes.sort(sortNotes);
  saveToStorage();
  renderCurrentPage();
}

function toggleCurrentImportant() {
  const note = currentNote();
  if (!note) return;
  note.important = !note.important;
  state.notes.sort(sortNotes);
  saveToStorage();
  renderNotesPage();
}

function clearEditor() {
  if (els.bodyInput) els.bodyInput.innerHTML = "";
}

function focusBody() {
  window.requestAnimationFrame(() => {
    els.bodyInput?.focus();
  });
}

function renderCurrentPage() {
  if (PAGE === "notebooks") {
    renderNotebookPage();
    return;
  }
  renderNotesPage();
}































