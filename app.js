/* ====== PUZZLE TRACKER APP LOGIC ====== */

const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email";
const OWNER_EMAIL = "mzbazeball@gmail.com";
const SHEET_RANGE_READ = "Sheet1!A2:I";
const SHEET_RANGE_APPEND = "Sheet1!A1";

let tokenClient;
let refreshTokenClient;
let accessToken = null;
let tokenExpiresAt = 0;
let gapiInited = false;
let allPuzzles = []; // cached rows from sheet
let currentGroups = []; // groups computed for the current grouping
let currentGroupKey = null; // 'pieces' | 'brand' | 'yearmonth'
let currentGroupValue = null; // selected category label, e.g. "1000 pieces"
let currentViewMode = "grid"; // 'grid' | 'list'
let currentSortOrder = "desc"; // 'asc' | 'desc' (date order within a category)
let allSortKey = "date";  // 'date' | 'pieces' (sort key for View All)
let allSortOrder = "desc"; // 'asc' | 'desc' (sort order for View All)
let navStack = ["screen-home"];
let isOwner = false;
let editingPuzzle = null; // puzzle object being edited, or null when adding new
let currentModalPuzzle = null; // puzzle currently shown in the detail modal

// Bulk photo add state
let bulkFiles = [];
let bulkIndex = 0;
let bulkPuzzles = [];

// Modal carousel state
let modalPhotoIds = [];
let modalPhotoIndex = 0;

// Brand category sort
let brandSortMode = "alpha"; // 'alpha' | 'countDesc' | 'countAsc'

// Parse imageFileId — handles old plain-string IDs and new JSON-array format
function parseImageIds(raw) {
  if (!raw) return [];
  if (typeof raw === "string" && raw.startsWith("[")) {
    try { return JSON.parse(raw).filter(Boolean); } catch (e) { return [raw]; }
  }
  return raw ? [raw] : [];
}

// Populate brand/artist datalists from loaded puzzles
function populateAutocomplete() {
  const brands = [...new Set(allPuzzles.map(p => p.brand).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  const artists = [...new Set(allPuzzles.map(p => p.artist).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  document.getElementById("brand-list").innerHTML = brands.map(b => `<option value="${escapeHtml(b)}">`).join("");
  document.getElementById("artist-list").innerHTML = artists.map(a => `<option value="${escapeHtml(a)}">`).join("");
}

const GROUP_TITLES = {
  pieces: "Piece Count",
  brand: "Brand",
  year: "Year",
  yearmonth: "Year and Month"
};

const TOKEN_STORAGE_KEY = "puzzleTrackerToken_v2"; // bumped: added userinfo.email scope
const TOKEN_EXPIRY_BUFFER_MS = 2 * 60 * 1000; // refresh 2 min early

// ---------- Token caching ----------

function saveToken(token, expiresInSeconds) {
  const expiresAt = Date.now() + expiresInSeconds * 1000;
  tokenExpiresAt = expiresAt;
  try {
    localStorage.setItem(TOKEN_STORAGE_KEY, JSON.stringify({ token, expiresAt }));
  } catch (e) {
    console.warn("Could not save token to localStorage", e);
  }
}

function loadStoredToken() {
  try {
    const raw = localStorage.getItem(TOKEN_STORAGE_KEY);
    if (!raw) return null;
    const data = JSON.parse(raw);
    if (!data.token || !data.expiresAt) return null;
    if (Date.now() > data.expiresAt - TOKEN_EXPIRY_BUFFER_MS) return null; // expired/expiring
    tokenExpiresAt = data.expiresAt;
    return data.token;
  } catch (e) {
    return null;
  }
}

function clearStoredToken() {
  tokenExpiresAt = 0;
  try {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
  } catch (e) {
    // ignore
  }
}

// Silently get a fresh access token without interrupting the user or
// navigating away from the current screen. Relies on the user still
// having an active Google session in this browser.
function silentRefreshToken() {
  return new Promise((resolve, reject) => {
    if (!refreshTokenClient) {
      refreshTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.CLIENT_ID,
        scope: SCOPES,
        callback: () => {} // overwritten per-call below
      });
    }
    refreshTokenClient.callback = (resp) => {
      if (resp.error) {
        reject(new Error(resp.error));
        return;
      }
      accessToken = resp.access_token;
      gapi.client.setToken({ access_token: accessToken });
      saveToken(accessToken, resp.expires_in || 3600);
      resolve(accessToken);
    };
    refreshTokenClient.requestAccessToken({ prompt: "" });
  });
}

// Call before any API request. Refreshes the token in the background
// if it's expired or about to expire, so long sessions (bulk photo
// uploads, leaving the tab open, etc.) don't hit a sign-in wall.
async function ensureFreshToken() {
  if (!accessToken || Date.now() > tokenExpiresAt - TOKEN_EXPIRY_BUFFER_MS) {
    try {
      await silentRefreshToken();
    } catch (err) {
      console.warn("Silent token refresh failed", err);
      throw err;
    }
  }
}

// ---------- Google API loading ----------

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

async function initGoogle() {
  if (!CONFIG.CLIENT_ID || CONFIG.CLIENT_ID.startsWith("YOUR_")) {
    document.getElementById("configWarning").innerHTML =
      '<div class="config-warning">App is not configured yet. Open <code>config.js</code> and follow SETUP_INSTRUCTIONS.md to add your Client ID, Spreadsheet ID, and Drive Folder ID.</div>';
    return;
  }

  await Promise.all([
    loadScript("https://apis.google.com/js/api.js"),
    loadScript("https://accounts.google.com/gsi/client")
  ]);

  await new Promise((resolve) => gapi.load("client", resolve));
  await gapi.client.init({
    discoveryDocs: [
      "https://sheets.googleapis.com/$discovery/rest?version=v4",
      "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"
    ]
  });
  gapiInited = true;

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.CLIENT_ID,
    scope: SCOPES,
    callback: (resp) => {
      if (resp.error) {
        console.error(resp);
        return;
      }
      accessToken = resp.access_token;
      gapi.client.setToken({ access_token: accessToken });
      saveToken(accessToken, resp.expires_in || 3600);
      onSignedIn();
    }
  });

  renderSignInButton();

  // Use a cached token if we still have a valid one
  const cached = loadStoredToken();
  if (cached) {
    accessToken = cached;
    gapi.client.setToken({ access_token: accessToken });
    onSignedIn();
    return;
  }

  // No cached token — handle hash URL for public visitors, then show home.
  const hashNav = await navigateToHash();
  if (!hashNav) {
    navStack = ["screen-home"];
    showScreen("screen-home", false);
  }
}

function renderSignInButton() {
  const btn = document.createElement("button");
  btn.textContent = "Sign in with Google";
  btn.className = "submit-btn";
  btn.style.maxWidth = "280px";
  btn.onclick = () => tokenClient.requestAccessToken({ prompt: "" });
  const wrap = document.getElementById("signinBtn");
  wrap.innerHTML = "";
  wrap.appendChild(btn);
}

async function onSignedIn() {
  document.getElementById("signOutBtn").style.visibility = "visible";
  await applyAccessControls();
  // Hash URL takes priority over session restore
  const hashNav = await navigateToHash();
  if (hashNav) return;
  const restored = await restoreSession();
  if (!restored) {
    navStack = ["screen-home"];
    showScreen("screen-home", false);
  }
}

async function applyAccessControls() {
  const trackBtn = document.getElementById("btnTrack");
  const bulkBtn = document.getElementById("btnBulkPhotos");
  const ownerSignInBtn = document.getElementById("btnOwnerSignIn");
  isOwner = true;
  trackBtn.style.display = "";
  bulkBtn.style.display = "";
  ownerSignInBtn.style.display = "none";
}

document.getElementById("signOutBtn").addEventListener("click", () => {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  gapi.client.setToken(null);
  clearStoredToken();
  clearSession();
  isOwner = false;
  document.getElementById("signOutBtn").style.visibility = "hidden";
  document.getElementById("btnTrack").style.display = "none";
  document.getElementById("btnBulkPhotos").style.display = "none";
  document.getElementById("btnOwnerSignIn").style.display = "";
  navStack = ["screen-home"];
  showScreen("screen-home", false);
  renderSignInButton();
});

document.getElementById("btnOwnerSignIn").addEventListener("click", () => {
  showScreen("screen-signin");
});

// ---------- Session persistence ----------

const SESSION_KEY = "puzzleTrackerSession_v1";

function saveSession() {
  try {
    sessionStorage.setItem(SESSION_KEY, JSON.stringify({
      screen: navStack[navStack.length - 1],
      navStack,
      currentGroupKey,
      currentGroupValue,
      currentViewMode,
      currentSortOrder,
      allSortKey,
      allSortOrder
    }));
  } catch(e) {}
}

function clearSession() {
  try { sessionStorage.removeItem(SESSION_KEY); } catch(e) {}
}

async function restoreSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    if (!raw) return false;
    const s = JSON.parse(raw);
    const screen = s.screen;
    // Don't restore form or sign-in screens
    if (!screen || ["screen-track","screen-bulk-photos","screen-signin","screen-home"].includes(screen)) return false;

    // Restore state variables
    navStack = s.navStack || ["screen-home", screen];
    currentGroupKey = s.currentGroupKey || null;
    currentGroupValue = s.currentGroupValue || null;
    currentViewMode = s.currentViewMode || "grid";
    currentSortOrder = s.currentSortOrder || "desc";
    allSortKey = s.allSortKey || "date";
    allSortOrder = s.allSortOrder || "desc";

    setViewMode(currentViewMode);

    if (screen === "screen-view-results") {
      showScreen(screen, false);
      document.getElementById("headerTitle").textContent = currentGroupValue || "Tracked Puzzles";
      document.getElementById("sortRow").style.display = currentGroupKey === "year" ? "flex" : "none";
      if (currentGroupKey === "all") {
        document.getElementById("sortAllDateDesc").classList.toggle("active", allSortKey === "date" && allSortOrder === "desc");
        document.getElementById("sortAllDateAsc").classList.toggle("active", allSortKey === "date" && allSortOrder === "asc");
        document.getElementById("sortAllPiecesDesc").classList.toggle("active", allSortKey === "pieces" && allSortOrder === "desc");
        document.getElementById("sortAllPiecesAsc").classList.toggle("active", allSortKey === "pieces" && allSortOrder === "asc");
        await loadAndRenderAll();
      } else if (currentGroupKey === "wooden") {
        await loadAndRenderWooden();
      } else {
        allPuzzles = await fetchAllPuzzles();
        currentGroups = groupPuzzles(allPuzzles, currentGroupKey);
        renderResults();
      }
      return true;
    }

    if (screen === "screen-view-categories") {
      showScreen(screen, false);
      document.getElementById("headerTitle").textContent = GROUP_TITLES[currentGroupKey] || "View Tracked Puzzles";
      await loadAndRenderCategories();
      return true;
    }

    if (screen === "screen-view-options" || screen === "screen-view-date-options") {
      showScreen(screen, false);
      return true;
    }
  } catch(e) {
    console.warn("Could not restore session", e);
  }
  return false;
}

// ---------- Hash Routing ----------

function buildHash() {
  const screen = navStack[navStack.length - 1] || "screen-home";
  if (screen === "screen-home") return "";
  if (screen === "screen-search") return "search";
  if (screen === "screen-view-options") return "view";
  if (screen === "screen-view-date-options") return "view/date";
  if (screen === "screen-view-categories") return "view/" + (currentGroupKey || "");
  if (screen === "screen-view-results") {
    if (currentGroupKey === "all") return "view/all";
    if (currentGroupKey === "wooden") return "view/wooden";
    if (currentGroupValue) return "view/" + currentGroupKey + "/" + encodeURIComponent(currentGroupValue);
    return "view/" + (currentGroupKey || "");
  }
  return "";
}

function updateHash() {
  const hash = buildHash();
  history.replaceState(null, "", hash ? "#" + hash : location.pathname + location.search);
}

function setModalHash(puzzleId) {
  history.replaceState(null, "", "#puzzle/" + encodeURIComponent(puzzleId));
}

function clearModalHash() {
  const hash = buildHash();
  history.replaceState(null, "", hash ? "#" + hash : location.pathname + location.search);
}

async function navigateToHash() {
  const raw = location.hash.slice(1);
  if (!raw) return false;

  const parts = raw.split("/");
  const section = parts[0];

  if (section === "search") {
    navStack = ["screen-home", "screen-search"];
    showScreen("screen-search", false);
    return true;
  }

  // Ensure puzzles are loaded (works via API key without auth)
  if (!allPuzzles.length) {
    try { allPuzzles = await fetchAllPuzzles(); } catch (e) { return false; }
  }

  if (section === "puzzle") {
    const id = decodeURIComponent(parts[1] || "");
    const puzzle = allPuzzles.find(p => p.id === id);
    if (!puzzle) return false;
    navStack = ["screen-home"];
    showScreen("screen-home", false);
    openModal(puzzle);
    return true;
  }

  if (section === "view") {
    const groupKey = parts[1];
    const groupValue = parts[2] ? decodeURIComponent(parts[2]) : null;

    if (!groupKey || groupKey === "options") {
      navStack = ["screen-home", "screen-view-options"];
      showScreen("screen-view-options", false);
      return true;
    }

    if (groupKey === "date") {
      navStack = ["screen-home", "screen-view-options", "screen-view-date-options"];
      showScreen("screen-view-date-options", false);
      return true;
    }

    currentGroupKey = groupKey;

    if (groupKey === "all") {
      navStack = ["screen-home", "screen-view-options", "screen-view-results"];
      const items = [...allPuzzles].sort((a, b) => a.date < b.date ? 1 : -1);
      currentGroups = [{ label: "All Tracked Puzzles", items }];
      currentGroupValue = "All Tracked Puzzles";
      showScreen("screen-view-results", false);
      document.getElementById("headerTitle").textContent = "All Tracked Puzzles";
      renderResults();
      return true;
    }

    if (groupKey === "wooden") {
      navStack = ["screen-home", "screen-view-options", "screen-view-results"];
      const items = allPuzzles.filter(p => p.wooden).sort((a, b) => a.date < b.date ? 1 : -1);
      currentGroups = [{ label: "Wooden Puzzles", items }];
      currentGroupValue = "Wooden Puzzles";
      showScreen("screen-view-results", false);
      document.getElementById("headerTitle").textContent = "Wooden Puzzles";
      renderResults();
      return true;
    }

    currentGroups = groupPuzzles(allPuzzles, groupKey);
    const dateGroups = ["year", "yearmonth"];

    if (groupValue) {
      currentGroupValue = groupValue;
      navStack = dateGroups.includes(groupKey)
        ? ["screen-home", "screen-view-options", "screen-view-date-options", "screen-view-categories", "screen-view-results"]
        : ["screen-home", "screen-view-options", "screen-view-categories", "screen-view-results"];
      showScreen("screen-view-results", false);
      document.getElementById("headerTitle").textContent = groupValue;
      renderResults();
    } else {
      navStack = dateGroups.includes(groupKey)
        ? ["screen-home", "screen-view-options", "screen-view-date-options", "screen-view-categories"]
        : ["screen-home", "screen-view-options", "screen-view-categories"];
      document.getElementById("headerTitle").textContent = GROUP_TITLES[groupKey] || "View Tracked Puzzles";
      renderCategories();
      showScreen("screen-view-categories", false);
    }
    return true;
  }

  return false;
}

// ---------- Navigation ----------

function showScreen(id, pushHistory = true) {
  document.querySelectorAll(".screen").forEach((el) => el.classList.remove("active"));
  document.getElementById(id).classList.add("active");

  const titles = {
    "screen-signin": "Puzzle Tracker",
    "screen-home": "Puzzle Tracker",
    "screen-track": "Track a Puzzle",
    "screen-view-options": "View Tracked Puzzles",
    "screen-view-date-options": "View by Date",
    "screen-view-categories": "View Tracked Puzzles",
    "screen-view-results": "Tracked Puzzles",
    "screen-bulk-photos": "Bulk Add Photos",
    "screen-search": "Search Puzzles"
  };
  document.getElementById("headerTitle").textContent = titles[id] || "Puzzle Tracker";

  document.getElementById("backBtn").style.visibility = id === "screen-home" ? "hidden" : "visible";

  if (pushHistory) {
    navStack.push(id);
  }

  saveSession();
  updateHash();
}

document.getElementById("backBtn").addEventListener("click", () => {
  if (navStack.length > 1) {
    navStack.pop();
    const prev = navStack[navStack.length - 1];
    showScreen(prev, false);
    if (prev === "screen-view-categories") {
      document.getElementById("headerTitle").textContent = GROUP_TITLES[currentGroupKey] || "View Tracked Puzzles";
    } else if (prev === "screen-view-results") {
      document.getElementById("headerTitle").textContent = currentGroupValue || "Tracked Puzzles";
    }
  }
});

document.getElementById("btnTrack").addEventListener("click", () => {
  resetTrackForm();
  populateAutocomplete();
  showScreen("screen-track");
});

document.getElementById("btnView").addEventListener("click", () => {
  showScreen("screen-view-options");
});

document.getElementById("btnSearch").addEventListener("click", async () => {
  showScreen("screen-search");
  const input = document.getElementById("searchInput");
  input.value = "";
  document.getElementById("searchResults").innerHTML = "";
  // Pre-load puzzles in the background so search feels instant
  if (!allPuzzles.length) {
    document.getElementById("searchResults").innerHTML = '<div class="status-msg">Loading…</div>';
    try {
      allPuzzles = await fetchAllPuzzles();
    } catch(err) {
      document.getElementById("searchResults").innerHTML = '<div class="status-msg">Error loading puzzles.</div>';
      return;
    }
  }
  document.getElementById("searchResults").innerHTML = '<div class="empty-msg">Start typing to search.</div>';
  // Focus the input after a brief delay (needed on mobile)
  setTimeout(() => input.focus(), 120);
});

document.getElementById("btnViewByDate").addEventListener("click", () => {
  showScreen("screen-view-date-options");
});

document.getElementById("btnBulkPhotos").addEventListener("click", async () => {
  showScreen("screen-bulk-photos");
  resetBulkPhotos();
  document.getElementById("bulkProgress").textContent = "Loading puzzles...";
  try {
    bulkPuzzles = await fetchAllPuzzles();
    document.getElementById("bulkProgress").textContent = "Choose photos to begin.";
  } catch (err) {
    console.error(err);
    document.getElementById("bulkProgress").textContent = "Error loading puzzles. Try again.";
  }
});

document.querySelectorAll(".option-btn[data-group]").forEach((btn) => {
  btn.addEventListener("click", async () => {
    currentGroupKey = btn.dataset.group;

    if (currentGroupKey === "all") {
      currentGroupValue = "All Tracked Puzzles";
      allSortKey = "date";
      allSortOrder = "desc";
      ["sortAllDateDesc","sortAllDateAsc","sortAllPiecesDesc","sortAllPiecesAsc"].forEach(id =>
        document.getElementById(id).classList.remove("active"));
      document.getElementById("sortAllDateDesc").classList.add("active");
      showScreen("screen-view-results");
      document.getElementById("headerTitle").textContent = currentGroupValue;
      document.getElementById("sortRow").style.display = "none";
      await loadAndRenderAll();
      return;
    }

    if (currentGroupKey === "wooden") {
      currentGroupValue = "Wooden Puzzles";
      showScreen("screen-view-results");
      document.getElementById("headerTitle").textContent = currentGroupValue;
      document.getElementById("sortRow").style.display = "none";
      await loadAndRenderWooden();
      return;
    }

    currentGroupValue = null;
    showScreen("screen-view-categories");
    document.getElementById("headerTitle").textContent = GROUP_TITLES[currentGroupKey] || "View Tracked Puzzles";
    await loadAndRenderCategories();
  });
});

document.getElementById("categoriesContainer").addEventListener("click", (e) => {
  const btn = e.target.closest(".option-btn[data-category]");
  if (!btn) return;
  currentGroupValue = btn.dataset.category;
  showScreen("screen-view-results");
  document.getElementById("headerTitle").textContent = currentGroupValue;
  currentSortOrder = "desc";
  document.getElementById("sortRow").style.display = currentGroupKey === "year" ? "flex" : "none";
  setSortOrder("desc");
});

document.getElementById("toggleGrid").addEventListener("click", () => {
  setViewMode("grid");
});
document.getElementById("toggleList").addEventListener("click", () => {
  setViewMode("list");
});

document.getElementById("sortDesc").addEventListener("click", () => {
  setSortOrder("desc");
});
document.getElementById("sortAsc").addEventListener("click", () => {
  setSortOrder("asc");
});

function setAllSort(key, order) {
  allSortKey = key;
  allSortOrder = order;
  ["sortAllDateDesc","sortAllDateAsc","sortAllPiecesDesc","sortAllPiecesAsc"].forEach(id => {
    document.getElementById(id).classList.remove("active");
  });
  const activeId = "sortAll" + (key === "date" ? "Date" : "Pieces") + (order === "desc" ? "Desc" : "Asc");
  document.getElementById(activeId).classList.add("active");
  renderResults();
}

document.getElementById("sortAllDateDesc").addEventListener("click", () => setAllSort("date", "desc"));
document.getElementById("sortAllDateAsc").addEventListener("click",  () => setAllSort("date", "asc"));
document.getElementById("sortAllPiecesDesc").addEventListener("click", () => setAllSort("pieces", "desc"));
document.getElementById("sortAllPiecesAsc").addEventListener("click",  () => setAllSort("pieces", "asc"));

function setViewMode(mode) {
  currentViewMode = mode;
  document.getElementById("toggleGrid").classList.toggle("active", mode === "grid");
  document.getElementById("toggleList").classList.toggle("active", mode === "list");
  renderResults();
}

function setSortOrder(order) {
  currentSortOrder = order;
  document.getElementById("sortDesc").classList.toggle("active", order === "desc");
  document.getElementById("sortAsc").classList.toggle("active", order === "asc");
  renderResults();
}

// ---------- Track Form ----------

function resetTrackForm() {
  editingPuzzle = null;
  document.getElementById("trackForm").reset();
  document.getElementById("photoPreviewWrap").innerHTML = "";
  document.getElementById("trackStatus").textContent = "";
  document.getElementById("f-date").value = new Date().toISOString().slice(0, 10);
  document.getElementById("saveBtn").textContent = "Save Puzzle";
}

document.getElementById("f-photo").addEventListener("change", (e) => {
  const files = Array.from(e.target.files);
  const wrap = document.getElementById("photoPreviewWrap");
  wrap.innerHTML = "";
  files.forEach(file => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = document.createElement("img");
      img.src = ev.target.result;
      img.className = "photo-thumb-preview";
      wrap.appendChild(img);
    };
    reader.readAsDataURL(file);
  });
});

document.getElementById("trackForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const saveBtn = document.getElementById("saveBtn");
  const status = document.getElementById("trackStatus");
  saveBtn.disabled = true;
  status.textContent = "Saving...";

  try {
    const date = document.getElementById("f-date").value;
    const brand = document.getElementById("f-brand").value.trim();
    const title = document.getElementById("f-title").value.trim();
    const artist = document.getElementById("f-artist").value.trim();
    const pieces = document.getElementById("f-pieces").value;
    const wooden = document.getElementById("f-wooden").checked;
    const notCompleted = document.getElementById("f-notcompleted").checked;
    const photoFiles = Array.from(document.getElementById("f-photo").files);
    const baseName = `${date}_${brand}_${title}`.replace(/[^a-zA-Z0-9_-]/g, "_");

    if (editingPuzzle) {
      // Keep existing photos, append any newly selected ones
      let existingIds = parseImageIds(editingPuzzle.imageFileId);
      if (photoFiles.length > 0) {
        status.textContent = "Uploading photo(s)...";
        const newIds = [];
        for (let i = 0; i < photoFiles.length; i++) {
          const fid = await uploadPhotoToDrive(photoFiles[i], `${baseName}_${Date.now()}_${i}`);
          newIds.push(fid);
        }
        existingIds = existingIds.concat(newIds);
      }
      const imageFileId = existingIds.length === 1 ? existingIds[0] : JSON.stringify(existingIds);

      status.textContent = "Saving changes...";
      const row = [editingPuzzle.id, date, brand, title, artist, pieces, wooden ? "TRUE" : "FALSE", notCompleted ? "TRUE" : "FALSE", imageFileId];
      await updatePuzzleRow(editingPuzzle.rowNumber, row);

      status.textContent = "Saved!";
      resetTrackForm();
      setTimeout(() => { showScreen("screen-home", false); navStack = ["screen-home"]; }, 600);
    } else {
      let imageFileId = "";
      if (photoFiles.length > 0) {
        status.textContent = "Uploading photo(s)...";
        const ids = [];
        for (let i = 0; i < photoFiles.length; i++) {
          const fid = await uploadPhotoToDrive(photoFiles[i], `${baseName}_${i}`);
          ids.push(fid);
        }
        imageFileId = ids.length === 1 ? ids[0] : JSON.stringify(ids);
      }

      status.textContent = "Saving puzzle details...";
      const id = Date.now().toString();
      await appendPuzzleRow([id, date, brand, title, artist, pieces, wooden ? "TRUE" : "FALSE", notCompleted ? "TRUE" : "FALSE", imageFileId]);

      status.textContent = "Saved!";
      resetTrackForm();
      setTimeout(() => { showScreen("screen-home", false); navStack = ["screen-home"]; }, 600);
    }
  } catch (err) {
    console.error(err);
    status.textContent = "Error saving puzzle. Please try again.";
  } finally {
    saveBtn.disabled = false;
  }
});

// ---------- Google Drive upload ----------

async function uploadPhotoToDrive(file, baseName) {
  await ensureFreshToken();

  const metadata = {
    name: baseName + "." + (file.type.split("/")[1] || "jpg"),
    parents: [CONFIG.DRIVE_FOLDER_ID],
    mimeType: file.type
  };

  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", file);

  const uploadResp = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
    method: "POST",
    headers: { Authorization: "Bearer " + accessToken },
    body: form
  });
  const uploadData = await uploadResp.json();
  const fileId = uploadData.id;

  // Make file viewable via link so it can be displayed as an <img>
  await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + accessToken,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ role: "reader", type: "anyone" })
  });

  return fileId;
}

function driveImageUrl(fileId, size = 400) {
  if (!fileId) return null;
  return `https://drive.google.com/thumbnail?id=${fileId}&sz=w${size}`;
}

// ---------- Google Sheets read/write ----------

async function appendPuzzleRow(row) {
  await ensureFreshToken();
  await gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: SHEET_RANGE_APPEND,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    resource: { values: [row] }
  });
}

async function updatePuzzleRow(rowNumber, row) {
  await ensureFreshToken();
  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `Sheet1!A${rowNumber}:I${rowNumber}`,
    valueInputOption: "USER_ENTERED",
    resource: { values: [row] }
  });
}

async function deletePhotoFromDrive(fileId) {
  try {
    await ensureFreshToken();
    await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
      method: "DELETE",
      headers: { Authorization: "Bearer " + accessToken }
    });
  } catch (err) {
    console.warn("Could not delete old photo from Drive", err);
  }
}

async function fetchAllPuzzles() {
  let rows;
  if (accessToken) {
    // Signed-in (owner) path: use OAuth-authenticated Sheets API
    await ensureFreshToken();
    const resp = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: SHEET_RANGE_READ
    });
    rows = resp.result.values || [];
  } else {
    // Anonymous viewer path: public read via API key (no sign-in required).
    // Requires the Sheet to be shared "Anyone with the link - Viewer" and
    // a Sheets API key restricted to this site (see SETUP_INSTRUCTIONS.md).
    if (!CONFIG.API_KEY || CONFIG.API_KEY.startsWith("YOUR_")) {
      throw new Error("Public viewing is not configured yet (missing API_KEY in config.js).");
    }
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${encodeURIComponent(SHEET_RANGE_READ)}?key=${CONFIG.API_KEY}`;
    const resp = await fetch(url);
    if (!resp.ok) {
      throw new Error("Could not load puzzles (" + resp.status + ")");
    }
    const data = await resp.json();
    rows = data.values || [];
  }
  return rows
    .map((r, idx) => ({ r, rowNumber: idx + 2 }))
    .filter(({ r }) => r.length > 0 && r[0])
    .map(({ r, rowNumber }) => ({
      id: r[0] || "",
      date: r[1] || "",
      brand: r[2] || "",
      title: r[3] || "",
      artist: r[4] || "",
      pieces: r[5] || "",
      wooden: (r[6] || "").toString().toUpperCase() === "TRUE",
      notCompleted: (r[7] || "").toString().toUpperCase() === "TRUE",
      imageFileId: r[8] || "",
      rowNumber
    }));
}

// ---------- Results rendering ----------

async function loadAndRenderAll() {
  const container = document.getElementById("resultsContainer");
  container.innerHTML = '<div class="status-msg">Loading...</div>';
  try {
    allPuzzles = await fetchAllPuzzles();
    const items = [...allPuzzles].sort((a, b) => (a.date < b.date ? 1 : -1));
    currentGroups = [{ label: "All Tracked Puzzles", items }];
    renderResults();
  } catch (err) {
    console.error(err);
    container.innerHTML = '<div class="status-msg">Error loading puzzles.</div>';
  }
}

async function loadAndRenderWooden() {
  const container = document.getElementById("resultsContainer");
  container.innerHTML = '<div class="status-msg">Loading...</div>';
  try {
    allPuzzles = await fetchAllPuzzles();
    const items = allPuzzles
      .filter((p) => p.wooden)
      .sort((a, b) => (a.date < b.date ? 1 : -1));
    currentGroups = [{ label: "Wooden Puzzles", items }];
    renderResults();
  } catch (err) {
    console.error(err);
    container.innerHTML = '<div class="status-msg">Error loading puzzles.</div>';
  }
}

async function loadAndRenderCategories() {
  const container = document.getElementById("categoriesContainer");
  container.innerHTML = '<div class="status-msg">Loading...</div>';
  try {
    allPuzzles = await fetchAllPuzzles();
    renderCategories();
  } catch (err) {
    console.error(err);
    container.innerHTML = '<div class="status-msg">Error loading puzzles.</div>';
  }
}

function renderCategories() {
  const container = document.getElementById("categoriesContainer");
  container.innerHTML = "";

  if (!allPuzzles.length) {
    container.innerHTML = '<div class="empty-msg">No puzzles tracked yet.</div>';
    return;
  }

  currentGroups = groupPuzzles(allPuzzles, currentGroupKey);

  // Show/hide brand sort row
  const brandSortRow = document.getElementById("brandSortRow");
  if (currentGroupKey === "brand") {
    brandSortRow.style.display = "";
    // Apply current sort mode
    let sorted = [...currentGroups];
    if (brandSortMode === "countDesc") {
      sorted.sort((a, b) => b.items.length - a.items.length);
    } else if (brandSortMode === "countAsc") {
      sorted.sort((a, b) => a.items.length - b.items.length);
    }
    // Update active button
    ["sortBrandAlpha","sortBrandCountDesc","sortBrandCountAsc"].forEach(id => {
      document.getElementById(id).classList.remove("active");
    });
    const activeId = brandSortMode === "countDesc" ? "sortBrandCountDesc"
                   : brandSortMode === "countAsc"  ? "sortBrandCountAsc"
                   : "sortBrandAlpha";
    document.getElementById(activeId).classList.add("active");
    sorted.forEach((group) => {
      const btn = document.createElement("button");
      btn.className = "option-btn";
      btn.dataset.category = group.label;
      btn.textContent = `${group.label} (${group.items.length})`;
      container.appendChild(btn);
    });
  } else {
    brandSortRow.style.display = "none";
    currentGroups.forEach((group) => {
      const btn = document.createElement("button");
      btn.className = "option-btn";
      btn.dataset.category = group.label;
      btn.textContent = `${group.label} (${group.items.length})`;
      container.appendChild(btn);
    });
  }
}

["sortBrandAlpha","sortBrandCountDesc","sortBrandCountAsc"].forEach(id => {
  document.getElementById(id).addEventListener("click", () => {
    brandSortMode = id === "sortBrandCountDesc" ? "countDesc"
                 : id === "sortBrandCountAsc"  ? "countAsc"
                 : "alpha";
    renderCategories();
  });
});

const PIECE_RANGES = [
  { min: 0, max: 99, label: "0-99" },
  { min: 100, max: 299, label: "100-299" },
  { min: 300, max: 499, label: "300-499" },
  { min: 500, max: 749, label: "500-749" },
  { min: 750, max: 999, label: "750-999" },
  { min: 1000, max: 1499, label: "1000-1499" },
  { min: 1500, max: Infinity, label: "1500+" }
];

function pieceRangeFor(pieces) {
  const n = parseInt(pieces, 10);
  if (isNaN(n)) return null;
  for (const r of PIECE_RANGES) {
    if (n >= r.min && n <= r.max) return r;
  }
  return null;
}

function groupPuzzles(puzzles, key) {
  const groups = {};
  puzzles.forEach((p) => {
    let groupName;
    if (key === "pieces") {
      const range = pieceRangeFor(p.pieces);
      groupName = range ? { sortKey: range.min, label: range.label } : { sortKey: -1, label: "Unknown piece count" };
    } else if (key === "brand") {
      groupName = p.brand || "Unknown brand";
    } else if (key === "year") {
      if (p.date) {
        const d = new Date(p.date + "T00:00:00");
        if (!isNaN(d)) {
          groupName = { sortKey: p.date.slice(0, 4), label: p.date.slice(0, 4) };
        } else {
          groupName = { sortKey: "0000", label: "Unknown date" };
        }
      } else {
        groupName = { sortKey: "0000", label: "Unknown date" };
      }
    } else {
      // yearmonth
      if (p.date) {
        const d = new Date(p.date + "T00:00:00");
        if (!isNaN(d)) {
          groupName = d.toLocaleString("en-US", { year: "numeric", month: "long" });
          groupName = { sortKey: p.date.slice(0, 7), label: groupName };
        } else {
          groupName = { sortKey: "0000-00", label: "Unknown date" };
        }
      } else {
        groupName = { sortKey: "0000-00", label: "Unknown date" };
      }
    }

    let label, sortKey;
    if (typeof groupName === "object") {
      label = groupName.label;
      sortKey = groupName.sortKey;
    } else {
      label = groupName;
      sortKey = groupName;
    }

    if (!groups[label]) groups[label] = { label, sortKey, items: [] };
    groups[label].items.push(p);
  });

  let groupArr = Object.values(groups);

  if (key === "pieces") {
    groupArr.sort((a, b) => b.sortKey - a.sortKey); // largest piece counts first
  } else if (key === "brand") {
    groupArr.sort((a, b) => a.label.localeCompare(b.label));
  } else {
    groupArr.sort((a, b) => (a.sortKey < b.sortKey ? 1 : -1)); // newest first
  }

  // sort items within each group by date desc
  groupArr.forEach((g) => {
    g.items.sort((a, b) => (a.date < b.date ? 1 : -1));
  });

  return groupArr;
}

function renderResults() {
  const container = document.getElementById("resultsContainer");
  const countEl = document.getElementById("resultsCount");
  container.innerHTML = "";

  const group = currentGroups.find((g) => g.label === currentGroupValue);
  let items = group ? group.items : [];

  // Show/hide sort row and count banner for "View All"
  const allSortRow = document.getElementById("allSortRow");
  if (currentGroupKey === "all" && items.length) {
    allSortRow.style.display = "";
    countEl.textContent = `${items.length} puzzle${items.length === 1 ? "" : "s"} tracked`;
    countEl.style.display = "";
    // Apply all-view sort
    items = [...items].sort((a, b) => {
      if (allSortKey === "pieces") {
        const pa = parseInt(a.pieces, 10);
        const pb = parseInt(b.pieces, 10);
        const na = isNaN(pa) ? (allSortOrder === "desc" ? -Infinity : Infinity) : pa;
        const nb = isNaN(pb) ? (allSortOrder === "desc" ? -Infinity : Infinity) : pb;
        return allSortOrder === "desc" ? nb - na : na - nb;
      } else {
        if (a.date < b.date) return allSortOrder === "desc" ? 1 : -1;
        if (a.date > b.date) return allSortOrder === "desc" ? -1 : 1;
        return 0;
      }
    });
  } else {
    allSortRow.style.display = "none";
    countEl.style.display = "none";
  }

  if (currentGroupKey === "year") {
    items = [...items].sort((a, b) => {
      if (a.date < b.date) return currentSortOrder === "asc" ? -1 : 1;
      if (a.date > b.date) return currentSortOrder === "asc" ? 1 : -1;
      return 0;
    });
  }

  if (!items.length) {
    container.innerHTML = '<div class="empty-msg">No puzzles in this category.</div>';
    return;
  }

  const wrap = document.createElement("div");
  wrap.className = currentViewMode === "grid" ? "grid" : "";

  items.forEach((p) => {
    const item = document.createElement("div");
    item.className = currentViewMode === "grid" ? "grid-item" : "list-item";

    const imgUrl = driveImageUrl(parseImageIds(p.imageFileId)[0] || "", currentViewMode === "grid" ? 400 : 100);
    let imgHtml;
    if (imgUrl) {
      imgHtml = `<img src="${imgUrl}" alt="${escapeHtml(p.title)}" loading="lazy">`;
    } else {
      imgHtml = `<div class="noimg">No photo</div>`;
    }

    const metaHtml = `
      <div class="meta">
        <div>${escapeHtml(p.date)}</div>
        <div class="b">${escapeHtml(p.brand)}</div>
        <div>${escapeHtml(p.pieces)} pieces</div>
      </div>`;

    item.innerHTML = imgHtml + metaHtml;
    item.addEventListener("click", () => openModal(p));
    wrap.appendChild(item);
  });

  container.appendChild(wrap);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ---------- Search ----------

document.getElementById("searchInput").addEventListener("input", (e) => {
  const query = e.target.value.trim().toLowerCase();
  const container = document.getElementById("searchResults");

  if (!query) {
    container.innerHTML = '<div class="empty-msg">Start typing to search.</div>';
    return;
  }

  const matches = allPuzzles.filter((p) =>
    (p.title  && p.title.toLowerCase().includes(query))  ||
    (p.brand  && p.brand.toLowerCase().includes(query))  ||
    (p.artist && p.artist.toLowerCase().includes(query))
  );

  if (!matches.length) {
    container.innerHTML = '<div class="empty-msg">No puzzles found.</div>';
    return;
  }

  // Sort by date descending
  const sorted = [...matches].sort((a, b) => (a.date < b.date ? 1 : -1));

  container.innerHTML = "";
  const resultCount = document.createElement("div");
  resultCount.className = "results-count";
  resultCount.textContent = `${sorted.length} result${sorted.length === 1 ? "" : "s"}`;
  container.appendChild(resultCount);

  sorted.forEach((p) => {
    const item = document.createElement("div");
    item.className = "list-item";
    const imgUrl = driveImageUrl(parseImageIds(p.imageFileId)[0] || "", 100);
    const imgHtml = imgUrl
      ? `<img src="${imgUrl}" alt="" loading="lazy">`
      : `<div class="noimg">No photo</div>`;
    const sub = [p.artist, p.pieces ? p.pieces + " pieces" : ""].filter(Boolean).join(" · ");
    item.innerHTML = `${imgHtml}
      <div class="meta">
        <div>${escapeHtml(p.date)}</div>
        <div class="b">${escapeHtml(p.brand)}</div>
        <div>${escapeHtml(p.title)}</div>
        ${sub ? `<div style="color:#777;font-size:0.8rem;">${escapeHtml(sub)}</div>` : ""}
      </div>`;
    item.addEventListener("click", () => openModal(p));
    container.appendChild(item);
  });
});

// ---------- Modal ----------

function showModalPhoto(index) {
  modalPhotoIndex = index;
  const ids = modalPhotoIds;
  const modalImg = document.getElementById("modalImg");
  const modalNoImg = document.getElementById("modalNoImg");
  const photoPrev = document.getElementById("photoPrev");
  const photoNext = document.getElementById("photoNext");
  const photoDots = document.getElementById("photoDots");

  if (ids.length === 0) {
    modalImg.style.display = "none";
    modalNoImg.style.display = "flex";
    photoPrev.style.display = "none";
    photoNext.style.display = "none";
    photoDots.innerHTML = "";
  } else {
    modalImg.src = driveImageUrl(ids[index], 800);
    modalImg.style.display = "block";
    modalNoImg.style.display = "none";
    const multi = ids.length > 1;
    photoPrev.style.display = (multi && index > 0) ? "flex" : "none";
    photoNext.style.display = (multi && index < ids.length - 1) ? "flex" : "none";
    photoDots.innerHTML = multi
      ? ids.map((_, i) => `<span class="photo-dot${i === index ? " active" : ""}"></span>`).join("")
      : "";
  }
}

function openModal(p) {
  currentModalPuzzle = p;
  modalPhotoIds = parseImageIds(p.imageFileId);
  showModalPhoto(0);

  document.getElementById("modalDate").textContent = p.date;
  document.getElementById("modalBrand").textContent = p.brand;
  document.getElementById("modalTitle").textContent = p.title;
  document.getElementById("modalArtist").textContent = p.artist || "—";
  document.getElementById("modalPieces").textContent = p.pieces;
  document.getElementById("modalWooden").textContent = p.wooden ? "Yes" : "No";
  document.getElementById("modalCompleted").textContent = p.notCompleted ? "No" : "Yes";

  document.getElementById("modalEdit").style.display = isOwner ? "" : "none";
  document.getElementById("modalDelete").style.display = isOwner ? "" : "none";

  document.getElementById("modalOverlay").classList.add("active");
  if (p.id) setModalHash(p.id);
}

document.getElementById("photoPrev").addEventListener("click", () => {
  if (modalPhotoIndex > 0) showModalPhoto(modalPhotoIndex - 1);
});
document.getElementById("photoNext").addEventListener("click", () => {
  if (modalPhotoIndex < modalPhotoIds.length - 1) showModalPhoto(modalPhotoIndex + 1);
});

// Touch swipe support
(function() {
  const modal = document.getElementById("modalOverlay");
  let touchStartX = 0;
  modal.addEventListener("touchstart", e => { touchStartX = e.touches[0].clientX; }, { passive: true });
  modal.addEventListener("touchend", e => {
    const dx = e.changedTouches[0].clientX - touchStartX;
    if (Math.abs(dx) < 40) return;
    if (dx < 0 && modalPhotoIndex < modalPhotoIds.length - 1) showModalPhoto(modalPhotoIndex + 1);
    if (dx > 0 && modalPhotoIndex > 0) showModalPhoto(modalPhotoIndex - 1);
  }, { passive: true });
})();

function closeModal() {
  document.getElementById("modalOverlay").classList.remove("active");
  clearModalHash();
}

document.getElementById("modalCopyLink")?.addEventListener("click", () => {
  const url = location.href; // already set to #puzzle/[id]
  navigator.clipboard.writeText(url).then(() => {
    const btn = document.getElementById("modalCopyLink");
    const orig = btn.textContent;
    btn.textContent = "✓ Copied!";
    setTimeout(() => { btn.textContent = orig; }, 2000);
  }).catch(() => {
    prompt("Copy this link:", url);
  });
});
document.getElementById("modalClose").addEventListener("click", closeModal);
document.getElementById("modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") closeModal();
});

document.getElementById("modalEdit").addEventListener("click", () => {
  if (!currentModalPuzzle) return;
  startEditPuzzle(currentModalPuzzle);
  closeModal();
});

async function deletePuzzleRow(rowNumber) {
  await ensureFreshToken();
  // Fetch sheetId (numeric) for the first sheet
  const meta = await gapi.client.sheets.spreadsheets.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    fields: "sheets.properties"
  });
  const sheetId = meta.result.sheets[0].properties.sheetId;
  await gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    resource: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId,
            dimension: "ROWS",
            startIndex: rowNumber - 1, // 0-indexed
            endIndex: rowNumber
          }
        }
      }]
    }
  });
}

document.getElementById("modalDelete").addEventListener("click", async () => {
  if (!currentModalPuzzle) return;
  const p = currentModalPuzzle;
  const label = [p.brand, p.title, p.date].filter(Boolean).join(" — ");
  if (!confirm(`Delete "${label}"?\n\nThis cannot be undone.`)) return;

  document.getElementById("modalOverlay").classList.remove("active");

  try {
    // Delete photo(s) from Drive first (best-effort)
    for (const fid of parseImageIds(p.imageFileId)) {
      await deletePhotoFromDrive(fid).catch(() => {});
    }
    // Delete the Sheet row
    await deletePuzzleRow(p.rowNumber);
    // Remove from local cache and re-render
    allPuzzles = allPuzzles.filter((x) => x.rowNumber !== p.rowNumber);
    // Adjust rowNumbers for rows that shifted up
    allPuzzles.forEach((x) => { if (x.rowNumber > p.rowNumber) x.rowNumber--; });
    currentGroups = currentGroups.map((g) => ({
      ...g,
      items: g.items.filter((x) => x.rowNumber !== p.rowNumber)
    }));
    renderResults();
  } catch (err) {
    console.error(err);
    alert("Error deleting puzzle. Please try again.");
  }
});

function startEditPuzzle(p) {
  editingPuzzle = p;

  document.getElementById("trackForm").reset();
  document.getElementById("trackStatus").textContent = "";
  document.getElementById("f-date").value = p.date || "";
  document.getElementById("f-brand").value = p.brand || "";
  document.getElementById("f-title").value = p.title || "";
  document.getElementById("f-artist").value = p.artist || "";
  document.getElementById("f-pieces").value = p.pieces || "";
  document.getElementById("f-wooden").checked = !!p.wooden;
  document.getElementById("f-notcompleted").checked = !!p.notCompleted;

  const wrap = document.getElementById("photoPreviewWrap");
  wrap.innerHTML = "";
  parseImageIds(p.imageFileId).forEach(fid => {
    const img = document.createElement("img");
    img.src = driveImageUrl(fid, 400);
    img.className = "photo-thumb-preview";
    wrap.appendChild(img);
  });

  document.getElementById("saveBtn").textContent = "Update Puzzle";
  populateAutocomplete();
  showScreen("screen-track");
  document.getElementById("headerTitle").textContent = "Edit Puzzle";
}

// ---------- Bulk Add Photos ----------

function resetBulkPhotos() {
  bulkFiles = [];
  bulkIndex = 0;
  document.getElementById("bulk-photo-input").value = "";
  document.getElementById("bulkPhotoPreview").style.display = "none";
  document.getElementById("bulkPhotoPreview").src = "";
  document.getElementById("bulk-search").value = "";
  document.getElementById("bulkSearchField").style.display = "none";
  document.getElementById("bulkSkip").style.display = "none";
  document.getElementById("bulkResultsList").innerHTML = "";
  document.getElementById("bulkProgress").textContent = "Choose photos to begin.";
}

document.getElementById("bulk-photo-input").addEventListener("change", (e) => {
  bulkFiles = Array.from(e.target.files);
  bulkIndex = 0;
  showBulkPhoto();
});

function showBulkPhoto() {
  const preview = document.getElementById("bulkPhotoPreview");
  const progress = document.getElementById("bulkProgress");
  const searchField = document.getElementById("bulkSearchField");
  const skipBtn = document.getElementById("bulkSkip");
  const search = document.getElementById("bulk-search");

  if (!bulkFiles.length) {
    preview.style.display = "none";
    searchField.style.display = "none";
    skipBtn.style.display = "none";
    document.getElementById("bulkResultsList").innerHTML = "";
    progress.textContent = "Choose photos to begin.";
    return;
  }

  if (bulkIndex >= bulkFiles.length) {
    preview.style.display = "none";
    searchField.style.display = "none";
    skipBtn.style.display = "none";
    document.getElementById("bulkResultsList").innerHTML = "";
    progress.textContent = `All done! Processed ${bulkFiles.length} photo${bulkFiles.length === 1 ? "" : "s"}.`;
    return;
  }

  const file = bulkFiles[bulkIndex];
  const reader = new FileReader();
  reader.onload = (ev) => {
    preview.src = ev.target.result;
    preview.style.display = "block";
  };
  reader.readAsDataURL(file);

  progress.textContent = `Photo ${bulkIndex + 1} of ${bulkFiles.length} — find the matching puzzle below`;
  searchField.style.display = "";
  skipBtn.style.display = "";
  search.value = "";
  renderBulkResults("");
}

document.getElementById("bulk-search").addEventListener("input", (e) => {
  renderBulkResults(e.target.value.trim().toLowerCase());
});

function renderBulkResults(query) {
  const container = document.getElementById("bulkResultsList");
  container.innerHTML = "";

  if (!query) {
    container.innerHTML = '<div class="empty-msg">Type to search for the matching puzzle.</div>';
    return;
  }

  const matches = bulkPuzzles.filter((p) =>
    (p.title && p.title.toLowerCase().includes(query)) ||
    (p.brand && p.brand.toLowerCase().includes(query))
  ).slice(0, 25);

  if (!matches.length) {
    container.innerHTML = '<div class="empty-msg">No matching puzzles.</div>';
    return;
  }

  matches.forEach((p) => {
    const item = document.createElement("div");
    item.className = "list-item";
    const hasPhoto = p.imageFileId ? " — already has a photo (will be replaced)" : "";
    item.innerHTML = `
      <div class="meta">
        <div>${escapeHtml(p.date)}</div>
        <div class="b">${escapeHtml(p.brand)} — ${escapeHtml(p.title)}</div>
        <div>${escapeHtml(p.pieces)} pieces${hasPhoto}</div>
      </div>`;
    item.addEventListener("click", () => assignBulkPhoto(p));
    container.appendChild(item);
  });
}

async function assignBulkPhoto(puzzle) {
  const progress = document.getElementById("bulkProgress");
  const file = bulkFiles[bulkIndex];

  progress.textContent = `Uploading photo ${bulkIndex + 1} of ${bulkFiles.length}...`;
  document.getElementById("bulkSearchField").style.display = "none";
  document.getElementById("bulkSkip").style.display = "none";
  document.getElementById("bulkResultsList").innerHTML = "";

  try {
    const newFileId = await uploadPhotoToDrive(
      file,
      `${puzzle.date}_${puzzle.brand}_${puzzle.title}`.replace(/[^a-zA-Z0-9_-]/g, "_")
    );

    if (puzzle.imageFileId) {
      await deletePhotoFromDrive(puzzle.imageFileId);
    }

    const row = [
      puzzle.id,
      puzzle.date,
      puzzle.brand,
      puzzle.title,
      puzzle.artist,
      puzzle.pieces,
      puzzle.wooden ? "TRUE" : "FALSE",
      puzzle.notCompleted ? "TRUE" : "FALSE",
      newFileId
    ];
    await updatePuzzleRow(puzzle.rowNumber, row);
    puzzle.imageFileId = newFileId;

    bulkIndex++;
    showBulkPhoto();
  } catch (err) {
    console.error(err);
    progress.textContent = "Error uploading photo. Try again or skip.";
    document.getElementById("bulkSearchField").style.display = "";
    document.getElementById("bulkSkip").style.display = "";
    renderBulkResults(document.getElementById("bulk-search").value.trim().toLowerCase());
  }
}

document.getElementById("bulkSkip").addEventListener("click", () => {
  bulkIndex++;
  showBulkPhoto();
});

// ---------- Init ----------

resetTrackForm();
initGoogle();
