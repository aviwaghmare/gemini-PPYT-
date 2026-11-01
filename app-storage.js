/* ==========================================================================
   POWERPOINT LITE — DEVICE STORAGE, TABS & FILE MANAGER (www/app-storage.js)
   ========================================================================== */

(function () {
  "use strict";

  const DB_NAME = "PowerPointLiteDB";
  const DB_VERSION = 1;
  const STORE_FILES = "presentations";
  const MAX_RECENT_FILES = 7;

  let dbInstance = null;
  let allFilesCache = [];
  let activeTab = "recent";
  let searchQuery = "";
  let selectedFileId = null;
  let renameMode = "rename";

  function openDatabase() {
    return new Promise((resolve, reject) => {
      if (dbInstance) return resolve(dbInstance);
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_FILES)) {
          const store = db.createObjectStore(STORE_FILES, { keyPath: "id" });
          store.createIndex("lastOpened", "lastOpened", { unique: false });
          store.createIndex("createdAt", "createdAt", { unique: false });
        }
      };

      request.onsuccess = (e) => {
        dbInstance = e.target.result;
        resolve(dbInstance);
      };

      request.onerror = (e) => reject(e.target.error);
    });
  }

  async function getAllStoredFiles() {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, "readonly");
      const store = tx.objectStore(STORE_FILES);
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  async function saveFileRecord(record) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, "readwrite");
      const store = tx.objectStore(STORE_FILES);
      const req = store.put(record);
      req.onsuccess = () => resolve(record);
      req.onerror = () => reject(req.error);
    });
  }

  async function deleteFileRecord(id) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, "readwrite");
      const store = tx.objectStore(STORE_FILES);
      const req = store.delete(id);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }

  function detectFolderLocation(file, customPath) {
    const pathStr = (customPath || file.webkitRelativePath || file.path || file.name || "").toLowerCase();
    if (pathStr.includes("whatsapp") || pathStr.includes("wa ") || pathStr.includes("group")) {
      return "WhatsApp";
    }
    if (pathStr.includes("telegram")) return "Telegram";
    if (pathStr.includes("download")) return "Download";
    if (pathStr.includes("document")) return "Documents";
    if (pathStr.includes("bluetooth")) return "Bluetooth";
    return "WhatsApp";
  }

  async function ensureSamplePresentationExists() {
    const existing = await getAllStoredFiles();
    if (existing.length > 0) {
      allFilesCache = existing;
      return;
    }

    if (typeof JSZip !== "undefined" && window.PPTXEngine && window.PPTXEngine.createSampleCacheMemoryPPTX) {
      try {
        const sampleBlob = await window.PPTXEngine.createSampleCacheMemoryPPTX();
        const now = Date.now();
        const demoRecord = {
          id: "sample-cache-memory-ppt",
          name: "GORUP 06 Cac...ECA F-4050.pptx",
          folder: "WhatsApp",
          size: sampleBlob.size,
          createdAt: now - 3600 * 1000 * 24,
          lastOpened: now,
          isRecent: true,
          isFavourite: false,
          blob: sampleBlob,
          customSlidesData: null
        };
        await saveFileRecord(demoRecord);
        allFilesCache = [demoRecord];
      } catch (err) {
        console.warn("Could not seed sample PPTX:", err);
      }
    }
  }

  async function refreshFileListUI() {
    allFilesCache = await getAllStoredFiles();

    const listEl = document.getElementById("ppt-file-list");
    const emptyEl = document.getElementById("empty-list-state");
    const emptyMsgEl = document.getElementById("empty-state-msg");
    const sectionTitleEl = document.getElementById("list-section-title");

    if (!listEl) return;

    let filtered = [...allFilesCache];

    if (searchQuery.trim() !== "") {
      const q = searchQuery.toLowerCase().trim();
      filtered = filtered.filter(
        (f) => f.name.toLowerCase().includes(q) || (f.folder && f.folder.toLowerCase().includes(q))
      );
    }

    if (activeTab === "recent") {
      sectionTitleEl.textContent = "Last Week";
      filtered = filtered
        .filter((f) => f.isRecent && f.lastOpened > 0)
        .sort((a, b) => b.lastOpened - a.lastOpened)
        .slice(0, MAX_RECENT_FILES);
      emptyMsgEl.textContent = "No recent presentations opened yet.";
    } else if (activeTab === "open") {
      sectionTitleEl.textContent = "All Device PPTs (Newest First)";
      filtered = filtered.sort((a, b) => (b.createdAt || b.lastOpened) - (a.createdAt || a.lastOpened));
      emptyMsgEl.textContent = "No presentations found on device. Tap below to load your .pptx files.";
    } else if (activeTab === "favourite") {
      sectionTitleEl.textContent = "Favourite Presentations";
      filtered = filtered
        .filter((f) => f.isFavourite)
        .sort((a, b) => b.lastOpened - a.lastOpened);
      emptyMsgEl.textContent = "No starred presentations yet. Tap ⋮ on any file to add it to Favourite.";
    }

    listEl.innerHTML = "";

    if (filtered.length === 0) {
      emptyEl.classList.remove("hidden");
      return;
    }

    emptyEl.classList.add("hidden");

    filtered.forEach((fileItem) => {
      const row = document.createElement("div");
      row.className = "ppt-file-item";
      row.dataset.id = fileItem.id;

      const cleanDisplayName = fileItem.name.replace(/\.pptx$/i, "");
      const starBadge = fileItem.isFavourite ? `<span class="star-badge-inline">★</span>` : "";

      row.innerHTML = `
        <div class="file-item-icon" aria-hidden="true">
          <svg viewBox="0 0 48 48" width="28" height="28">
            <circle cx="26" cy="24" r="18" fill="#e85b35" />
            <path d="M26 6 A18 18 0 0 1 44 24 L26 24 Z" fill="#ff9f2e" />
            <rect x="4" y="20" width="21" height="21" rx="4.5" fill="#c42b1c" />
            <text x="14.5" y="36" fill="#ffffff" font-family="sans-serif" font-weight="800" font-size="15" text-anchor="middle">P</text>
          </svg>
        </div>
        <div class="file-item-info">
          <div class="file-item-name">${escapeHtml(cleanDisplayName)}</div>
          <div class="file-item-meta">
            ${starBadge}
            <span>${escapeHtml(fileItem.folder || "WhatsApp")}</span>
          </div>
        </div>
        <button class="file-item-more-btn" data-more-id="${fileItem.id}" aria-label="File options">
          <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
            <circle cx="12" cy="5" r="2"></circle>
            <circle cx="12" cy="12" r="2"></circle>
            <circle cx="12" cy="19" r="2"></circle>
          </svg>
        </button>
      `;

      row.addEventListener("click", (e) => {
        if (e.target.closest(".file-item-more-btn")) return;
        openPresentationById(fileItem.id);
      });

      const moreBtn = row.querySelector(".file-item-more-btn");
      moreBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        openFileActionSheet(fileItem.id);
      });

      listEl.appendChild(row);
    });
  }

  async function handlePickedFiles(fileList, sourceFolderHint) {
    if (!fileList || fileList.length === 0) return;

    showLoading("Loading presentation...");
    let lastAddedId = null;

    try {
      for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        if (!file.name.toLowerCase().endsWith(".pptx")) {
          showToast("Please select a valid .pptx PowerPoint file");
          continue;
        }

        const arrayBuffer = await file.arrayBuffer();
        const blob = new Blob([arrayBuffer], {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        });

        const existing = allFilesCache.find(
          (item) => item.name.toLowerCase() === file.name.toLowerCase()
        );

        const now = Date.now();
        const record = {
          id: existing ? existing.id : "ppt-" + now + "-" + Math.random().toString(36).slice(2, 7),
          name: file.name,
          folder: sourceFolderHint || detectFolderLocation(file),
          size: file.size,
          createdAt: file.lastModified || now,
          lastOpened: now,
          isRecent: true,
          isFavourite: existing ? existing.isFavourite : false,
          blob: blob,
          customSlidesData: null
        };

        await saveFileRecord(record);
        lastAddedId = record.id;
      }

      await enforceRecentSevenLimit();
      await refreshFileListUI();

      if (fileList.length === 1 && lastAddedId) {
        await openPresentationById(lastAddedId);
      } else if (lastAddedId) {
        showToast(`${fileList.length} presentations added`);
      }
    } catch (err) {
      console.error("Error importing PPTX:", err);
      showToast("Could not read PPTX file");
    } finally {
      hideLoading();
    }
  }

  async function enforceRecentSevenLimit() {
    const all = await getAllStoredFiles();
    const recentSorted = all
      .filter((f) => f.isRecent)
      .sort((a, b) => b.lastOpened - a.lastOpened);

    if (recentSorted.length > MAX_RECENT_FILES) {
      for (let i = MAX_RECENT_FILES; i < recentSorted.length; i++) {
        recentSorted[i].isRecent = false;
        await saveFileRecord(recentSorted[i]);
      }
    }
  }

  async function openPresentationById(fileId) {
    const all = await getAllStoredFiles();
    const target = all.find((f) => f.id === fileId);
    if (!target) {
      showToast("File not found");
      return;
    }

    showLoading("Opening slides...");
    try {
      target.lastOpened = Date.now();
      target.isRecent = true;
      await saveFileRecord(target);
      await enforceRecentSevenLimit();

      if (window.PPTXEngine && window.PPTXEngine.loadPresentationRecord) {
        await window.PPTXEngine.loadPresentationRecord(target);
      }

      switchScreen("screen-reader");
      refreshFileListUI();
    } catch (err) {
      console.error("Failed to open presentation:", err);
      showToast("Error rendering presentation slides");
    } finally {
      hideLoading();
    }
  }

  function openFileActionSheet(fileId) {
    const file = allFilesCache.find((f) => f.id === fileId);
    if (!file) return;

    selectedFileId = fileId;
    document.getElementById("sheet-selected-filename").textContent = file.name.replace(/\.pptx$/i, "");
    document.getElementById("sheet-selected-filepath").textContent = file.folder || "WhatsApp";
    document.getElementById("label-file-star").textContent = file.isFavourite
      ? "Remove from Favourite"
      : "Add to Favourite";

    showBottomSheet("sheet-file-actions");
  }

  async function toggleSelectedFileStar() {
    const file = allFilesCache.find((f) => f.id === selectedFileId);
    if (!file) return;

    file.isFavourite = !file.isFavourite;
    await saveFileRecord(file);
    closeAllModals();
    await refreshFileListUI();
    showToast(file.isFavourite ? "Added to Favourite ★" : "Removed from Favourite");
  }

  async function shareSelectedFile() {
    const file = allFilesCache.find((f) => f.id === selectedFileId);
    closeAllModals();
    if (!file || !file.blob) return;

    const fileName = file.name.endsWith(".pptx") ? file.name : file.name + ".pptx";
    const shareFileObj = new File([file.blob], fileName, {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });

    if (navigator.canShare && navigator.canShare({ files: [shareFileObj] })) {
      try {
        await navigator.share({
          title: fileName,
          files: [shareFileObj]
        });
        return;
      } catch (e) {
        if (e.name !== "AbortError") console.warn("Share fallback:", e);
      }
    }

    triggerBlobDownload(file.blob, fileName);
    showToast("Sharing / Downloading " + fileName);
  }

  function promptRenameSelectedFile(mode = "rename") {
    const file = allFilesCache.find((f) => f.id === selectedFileId);
    if (!file) return;

    renameMode = mode;
    closeAllModals();

    const modalTitle = document.getElementById("rename-modal-title");
    const inputEl = document.getElementById("input-rename-value");

    modalTitle.textContent = mode === "saveAs" ? "Save Presentation As" : "Rename Presentation";
    const baseName = file.name.replace(/\.pptx$/i, "");
    inputEl.value = mode === "saveAs" ? baseName + " (Copy)" : baseName;

    document.getElementById("modal-backdrop").classList.remove("hidden");
    document.getElementById("modal-rename-file").classList.remove("hidden");
    setTimeout(() => inputEl.focus(), 80);
  }

  async function confirmRenameOrSaveAs() {
    const inputEl = document.getElementById("input-rename-value");
    const rawName = (inputEl.value || "").trim();
    if (!rawName) {
      showToast("Name cannot be empty");
      return;
    }

    const finalName = rawName.toLowerCase().endsWith(".pptx") ? rawName : rawName + ".pptx";
    const file = allFilesCache.find((f) => f.id === selectedFileId);
    if (!file) return;

    if (renameMode === "saveAs") {
      const now = Date.now();
      const copyRecord = {
        ...file,
        id: "ppt-" + now + "-" + Math.random().toString(36).slice(2, 6),
        name: finalName,
        createdAt: now,
        lastOpened: now,
        isRecent: true
      };
      await saveFileRecord(copyRecord);
      selectedFileId = copyRecord.id;
      document.getElementById("reader-file-title").textContent = finalName.replace(/\.pptx$/i, "");
      triggerBlobDownload(copyRecord.blob, finalName);
      showToast("Saved copy as " + finalName);
    } else {
      file.name = finalName;
      await saveFileRecord(file);
      showToast("Renamed to " + finalName.replace(/\.pptx$/i, ""));
    }

    closeAllModals();
    await refreshFileListUI();
  }

  async function deleteSelectedFile() {
    const file = allFilesCache.find((f) => f.id === selectedFileId);
    if (!file) return;

    await deleteFileRecord(file.id);
    closeAllModals();
    await refreshFileListUI();
    showToast("File deleted");
  }

  function switchScreen(screenId) {
    document.querySelectorAll(".screen").forEach((s) => s.classList.remove("active"));
    const target = document.getElementById(screenId);
    if (target) target.classList.add("active");
  }

  function showBottomSheet(sheetId) {
    document.getElementById("modal-backdrop").classList.remove("hidden");
    document.getElementById(sheetId).classList.remove("hidden");
  }

  function closeAllModals() {
    document.getElementById("modal-backdrop").classList.add("hidden");
    document.querySelectorAll(".bottom-sheet, .dialog-modal").forEach((el) => {
      el.classList.add("hidden");
    });
  }

  let toastTimer = null;
  function showToast(msg) {
    const toast = document.getElementById("app-toast");
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.add("hidden"), 2400);
  }

  function showLoading(msg) {
    const overlay = document.getElementById("loading-overlay");
    const text = document.getElementById("loading-text");
    if (text && msg) text.textContent = msg;
    if (overlay) overlay.classList.remove("hidden");
  }

  function hideLoading() {
    const overlay = document.getElementById("loading-overlay");
    if (overlay) overlay.classList.add("hidden");
  }

  function triggerBlobDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 500);
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function initHomeEvents() {
    const savedTheme = localStorage.getItem("ppt_app_theme") || "dark";
    document.documentElement.setAttribute("data-theme", savedTheme);
    const radioInput = document.querySelector(`input[name="app-theme"][value="${savedTheme}"]`);
    if (radioInput) radioInput.checked = true;

    document.getElementById("btn-open-theme-settings").addEventListener("click", () => {
      document.getElementById("modal-backdrop").classList.remove("hidden");
      document.getElementById("modal-theme-settings").classList.remove("hidden");
    });

    document.querySelectorAll('input[name="app-theme"]').forEach((radio) => {
      radio.addEventListener("change", (e) => {
        const theme = e.target.value;
        document.documentElement.setAttribute("data-theme", theme);
        localStorage.setItem("ppt_app_theme", theme);
      });
    });

    document.getElementById("btn-close-theme-modal").addEventListener("click", closeAllModals);
    document.getElementById("modal-backdrop").addEventListener("click", closeAllModals);

    const searchBar = document.getElementById("home-search-bar");
    const searchInput = document.getElementById("home-search-input");

    document.getElementById("btn-home-search-toggle").addEventListener("click", () => {
      searchBar.classList.toggle("hidden");
      if (!searchBar.classList.contains("hidden")) {
        searchInput.focus();
      } else {
        searchInput.value = "";
        searchQuery = "";
        refreshFileListUI();
      }
    });

    document.getElementById("btn-home-search-clear").addEventListener("click", () => {
      searchInput.value = "";
      searchQuery = "";
      searchBar.classList.add("hidden");
      refreshFileListUI();
    });

    searchInput.addEventListener("input", (e) => {
      searchQuery = e.target.value;
      refreshFileListUI();
    });

    const filePicker = document.getElementById("native-file-picker");

    document.querySelectorAll(".home-tabs .tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".home-tabs .tab-btn").forEach((b) => {
          b.classList.remove("active");
          b.setAttribute("aria-selected", "false");
        });
        btn.classList.add("active");
        btn.setAttribute("aria-selected", "true");

        activeTab = btn.dataset.tab;
        refreshFileListUI();
      });
    });

    document.getElementById("btn-browse-device-ppt").addEventListener("click", () => filePicker.click());
    document.getElementById("btn-empty-pick-file").addEventListener("click", () => filePicker.click());

    filePicker.addEventListener("change", (e) => {
      handlePickedFiles(e.target.files);
      filePicker.value = "";
    });

    document.getElementById("action-file-star").addEventListener("click", toggleSelectedFileStar);
    document.getElementById("action-file-share").addEventListener("click", shareSelectedFile);
    document.getElementById("action-file-rename").addEventListener("click", () => promptRenameSelectedFile("rename"));
    document.getElementById("action-file-delete").addEventListener("click", deleteSelectedFile);

    document.getElementById("btn-rename-cancel").addEventListener("click", closeAllModals);
    document.getElementById("btn-rename-confirm").addEventListener("click", confirmRenameOrSaveAs);

    document.getElementById("btn-reader-close").addEventListener("click", () => {
      switchScreen("screen-home");
      refreshFileListUI();
    });

    document.getElementById("btn-open-reader-more").addEventListener("click", () => {
      showBottomSheet("sheet-reader-more");
    });
  }

  window.AppStorage = {
    openDatabase,
    getAllStoredFiles,
    saveFileRecord,
    deleteFileRecord,
    refreshFileListUI,
    openPresentationById,
    handlePickedFiles,
    promptRenameSelectedFile,
    setSelectedFileId: (id) => { selectedFileId = id; },
    getSelectedFileId: () => selectedFileId,
    switchScreen,
    showBottomSheet,
    closeAllModals,
    showToast,
    showLoading,
    hideLoading,
    triggerBlobDownload
  };

  window.addEventListener("DOMContentLoaded", async () => {
    initHomeEvents();
    await ensureSamplePresentationExists();
    await refreshFileListUI();
  });
})();
