/* ====== PUZZLE TRACKER APP LOGIC ====== */

const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email";
const OWNER_EMAIL = "mzbazeball@gmail.com";
const SHEET_RANGE_READ = "Sheet1!A2:I";
const SHEET_RANGE_APPEND = "Sheet1!A1";

let tokenClient;
let accessToken = null;
let gapiInited = false;
let allPuzzles = []; // cached rows from sheet
let currentGroups = []; // groups computed for the current grouping
let currentGroupKey = null; // 'pieces' | 'brand' | 'yearmonth'
let currentGroupValue = null; // selected category label, e.g. "1000 pieces"
let currentViewMode = "grid"; // 'grid' | 'list'
let navStack = ["screen-home"];

const GROUP_TITLES = {
  pieces: "Piece Count",
  brand: "Brand",
  yearmonth: "Year and Month"
};

const TOKEN_STORAGE_KEY = "puzzleTrackerToken_v2"; // bumped: added userinfo.email scope
const TOKEN_EXPIRY_BUFFER_MS = 2 * 60 * 1000; // refresh 2 min early

// ---------- Token caching ----------

function saveToken(token, expiresInSeconds) {
  const expiresAt = Date.now() + expiresInSeconds * 1000;
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
    return data.token;
  } catch (e) {
    return null;
  }
}

function clearStoredToken() {
  try {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
  } catch (e) {
    // ignore
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

  // Otherwise try silent sign-in (works if user previously granted access this session/browser)
  tokenClient.requestAccessToken({ prompt: "" });
}

function renderSignInButton() {
  const btn = document.createElement("button");
  btn.textContent = "Sign in with Google";
  btn.className = "submit-btn";
  btn.style.maxWidth = "280px";
  btn.onclick = () => tokenClient.requestAccessToken({ prompt: "consent" });
  const wrap = document.getElementById("signinBtn");
  wrap.innerHTML = "";
  wrap.appendChild(btn);
}

async function onSignedIn() {
  document.getElementById("signOutBtn").style.visibility = "visible";
  showScreen("screen-home", false);
  await applyAccessControls();
}

async function applyAccessControls() {
  const trackBtn = document.getElementById("btnTrack");
  try {
    const resp = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
      headers: { Authorization: "Bearer " + accessToken }
    });
    const data = await resp.json();
    if (data.email && data.email.toLowerCase() === OWNER_EMAIL.toLowerCase()) {
      trackBtn.style.display = "";
    } else {
      trackBtn.style.display = "none";
    }
  } catch (err) {
    console.error("Could not verify account email", err);
    // Fail safe: hide the track button if we can't confirm the account
    trackBtn.style.display = "none";
  }
}

document.getElementById("signOutBtn").addEventListener("click", () => {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  gapi.client.setToken(null);
  clearStoredToken();
  document.getElementById("signOutBtn").style.visibility = "hidden";
  navStack = ["screen-home"];
  showScreen("screen-signin", false);
  renderSignInButton();
});

// ---------- Navigation ----------

function showScreen(id, pushHistory = true) {
  document.querySelectorAll(".screen").forEach((el) => el.classList.remove("active"));
  document.getElementById(id).classList.add("active");

  const titles = {
    "screen-signin": "Puzzle Tracker",
    "screen-home": "Puzzle Tracker",
    "screen-track": "Track a Puzzle",
    "screen-view-options": "View Tracked Puzzles",
    "screen-view-categories": "View Tracked Puzzles",
    "screen-view-results": "Tracked Puzzles"
  };
  document.getElementById("headerTitle").textContent = titles[id] || "Puzzle Tracker";

  document.getElementById("backBtn").style.visibility = id === "screen-home" || id === "screen-signin" ? "hidden" : "visible";

  if (pushHistory) {
    navStack.push(id);
  }
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
  showScreen("screen-track");
});

document.getElementById("btnView").addEventListener("click", () => {
  showScreen("screen-view-options");
});

document.querySelectorAll(".option-btn[data-group]").forEach((btn) => {
  btn.addEventListener("click", async () => {
    currentGroupKey = btn.dataset.group;

    if (currentGroupKey === "all") {
      currentGroupValue = "All Tracked Puzzles";
      showScreen("screen-view-results");
      document.getElementById("headerTitle").textContent = currentGroupValue;
      await loadAndRenderAll();
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
  renderResults();
});

document.getElementById("toggleGrid").addEventListener("click", () => {
  setViewMode("grid");
});
document.getElementById("toggleList").addEventListener("click", () => {
  setViewMode("list");
});

function setViewMode(mode) {
  currentViewMode = mode;
  document.getElementById("toggleGrid").classList.toggle("active", mode === "grid");
  document.getElementById("toggleList").classList.toggle("active", mode === "list");
  renderResults();
}

// ---------- Track Form ----------

function resetTrackForm() {
  document.getElementById("trackForm").reset();
  document.getElementById("photoPreview").style.display = "none";
  document.getElementById("photoPreview").src = "";
  document.getElementById("trackStatus").textContent = "";
  document.getElementById("f-date").value = new Date().toISOString().slice(0, 10);
}

document.getElementById("f-photo").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const img = document.getElementById("photoPreview");
    img.src = ev.target.result;
    img.style.display = "block";
  };
  reader.readAsDataURL(file);
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
    const photoFile = document.getElementById("f-photo").files[0];

    let imageUrl = "";
    if (photoFile) {
      status.textContent = "Uploading photo...";
      imageUrl = await uploadPhotoToDrive(photoFile, `${date}_${brand}_${title}`.replace(/[^a-zA-Z0-9_-]/g, "_"));
    }

    status.textContent = "Saving puzzle details...";
    const id = Date.now().toString();
    await appendPuzzleRow([id, date, brand, title, artist, pieces, wooden ? "TRUE" : "FALSE", notCompleted ? "TRUE" : "FALSE", imageUrl]);

    status.textContent = "Saved!";
    resetTrackForm();
    setTimeout(() => {
      showScreen("screen-home", false);
      navStack = ["screen-home"];
    }, 600);
  } catch (err) {
    console.error(err);
    status.textContent = "Error saving puzzle. Please try again.";
  } finally {
    saveBtn.disabled = false;
  }
});

// ---------- Google Drive upload ----------

async function uploadPhotoToDrive(file, baseName) {
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
  await gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: SHEET_RANGE_APPEND,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    resource: { values: [row] }
  });
}

async function fetchAllPuzzles() {
  const resp = await gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: SHEET_RANGE_READ
  });
  const rows = resp.result.values || [];
  return rows
    .filter((r) => r.length > 0 && r[0])
    .map((r) => ({
      id: r[0] || "",
      date: r[1] || "",
      brand: r[2] || "",
      title: r[3] || "",
      artist: r[4] || "",
      pieces: r[5] || "",
      wooden: (r[6] || "").toString().toUpperCase() === "TRUE",
      notCompleted: (r[7] || "").toString().toUpperCase() === "TRUE",
      imageFileId: r[8] || ""
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

  currentGroups.forEach((group) => {
    const btn = document.createElement("button");
    btn.className = "option-btn";
    btn.dataset.category = group.label;
    btn.textContent = `${group.label} (${group.items.length})`;
    container.appendChild(btn);
  });
}

function groupPuzzles(puzzles, key) {
  const groups = {};
  puzzles.forEach((p) => {
    let groupName;
    if (key === "pieces") {
      groupName = p.pieces ? `${p.pieces} pieces` : "Unknown piece count";
    } else if (key === "brand") {
      groupName = p.brand || "Unknown brand";
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
    groupArr.sort((a, b) => {
      const an = parseInt(a.label) || -1;
      const bn = parseInt(b.label) || -1;
      return bn - an;
    });
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
  container.innerHTML = "";

  const group = currentGroups.find((g) => g.label === currentGroupValue);
  const items = group ? group.items : [];

  if (!items.length) {
    container.innerHTML = '<div class="empty-msg">No puzzles in this category.</div>';
    return;
  }

  const wrap = document.createElement("div");
  wrap.className = currentViewMode === "grid" ? "grid" : "";

  items.forEach((p) => {
    const item = document.createElement("div");
    item.className = currentViewMode === "grid" ? "grid-item" : "list-item";

    const imgUrl = driveImageUrl(p.imageFileId, currentViewMode === "grid" ? 400 : 100);
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

// ---------- Modal ----------

function openModal(p) {
  const modalImg = document.getElementById("modalImg");
  const modalNoImg = document.getElementById("modalNoImg");
  const imgUrl = driveImageUrl(p.imageFileId, 800);

  if (imgUrl) {
    modalImg.src = imgUrl;
    modalImg.style.display = "block";
    modalNoImg.style.display = "none";
  } else {
    modalImg.style.display = "none";
    modalNoImg.style.display = "flex";
  }

  document.getElementById("modalDate").textContent = p.date;
  document.getElementById("modalBrand").textContent = p.brand;
  document.getElementById("modalTitle").textContent = p.title;
  document.getElementById("modalArtist").textContent = p.artist || "—";
  document.getElementById("modalPieces").textContent = p.pieces;
  document.getElementById("modalWooden").textContent = p.wooden ? "Yes" : "No";
  document.getElementById("modalCompleted").textContent = p.notCompleted ? "No" : "Yes";

  document.getElementById("modalOverlay").classList.add("active");
}

document.getElementById("modalClose").addEventListener("click", () => {
  document.getElementById("modalOverlay").classList.remove("active");
});
document.getElementById("modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") {
    document.getElementById("modalOverlay").classList.remove("active");
  }
});

// ---------- Init ----------

resetTrackForm();
initGoogle();
