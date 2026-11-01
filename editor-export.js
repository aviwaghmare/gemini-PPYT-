/* ==========================================================================
   POWERPOINT LITE — SIMPLE SLIDE EDITOR & PDF/PRINT EXPORTER (www/editor-export.js)
   ========================================================================== */

(function () {
  "use strict";

  let workingSlidesClone = [];
  let undoHistoryStack = [];
  let editorSlideIdx = 0;
  let selectedElementId = null;
  let clipboardElement = null;

  let isDraggingEl = false;
  let dragStartX = 0;
  let dragStartY = 0;
  let elInitialX = 0;
  let elInitialY = 0;

  function openEditorAtSlide(slideIndex) {
    const originalSlides = window.PPTXEngine.getActiveSlidesData();
    if (!originalSlides || originalSlides.length === 0) return;

    workingSlidesClone = JSON.parse(JSON.stringify(originalSlides));
    undoHistoryStack = [JSON.stringify(workingSlidesClone)];

    editorSlideIdx =
      typeof slideIndex === "number"
        ? Math.max(0, Math.min(workingSlidesClone.length - 1, slideIndex))
        : window.PPTXEngine.getCurrentVisibleSlideIndex() || 0;

    selectedElementId = null;
    hideContextPill();

    window.AppStorage.switchScreen("screen-editor");
    renderEditorStage();
    renderEditorFilmstrip();
  }

  function pushUndoSnapshot() {
    const snap = JSON.stringify(workingSlidesClone);
    if (undoHistoryStack[undoHistoryStack.length - 1] !== snap) {
      undoHistoryStack.push(snap);
      if (undoHistoryStack.length > 25) {
        undoHistoryStack.shift();
      }
    }
  }

  function triggerUndo() {
    if (undoHistoryStack.length <= 1) {
      window.AppStorage.showToast("Nothing left to undo");
      return;
    }
    undoHistoryStack.pop();
    const prevSnap = undoHistoryStack[undoHistoryStack.length - 1];
    workingSlidesClone = JSON.parse(prevSnap);
    selectedElementId = null;
    hideContextPill();
    renderEditorStage();
    renderEditorFilmstrip();
    window.AppStorage.showToast("Undone");
  }

  async function saveEditsAndReturn() {
    syncActiveDOMTextToModel();
    window.AppStorage.showLoading("Saving changes...");

    try {
      window.PPTXEngine.setActiveSlidesData(JSON.parse(JSON.stringify(workingSlidesClone)));
      await persistCurrentPresentationToStorage(false);

      window.PPTXEngine.renderVerticalReaderSlides();
      window.AppStorage.switchScreen("screen-reader");
      window.AppStorage.showToast("Presentation saved ✓");
    } catch (err) {
      console.error("Failed to save edits:", err);
      window.AppStorage.showToast("Error saving changes");
    } finally {
      window.AppStorage.hideLoading();
    }
  }

  function discardEditsAndReturn() {
    selectedElementId = null;
    hideContextPill();
    window.AppStorage.switchScreen("screen-reader");
    window.AppStorage.showToast("Edits discarded");
  }

  function renderEditorStage() {
    const stageWrapper = document.getElementById("editor-slide-stage-wrapper");
    if (!stageWrapper) return;

    const currentSlide = workingSlidesClone[editorSlideIdx];
    if (!currentSlide) return;

    stageWrapper.innerHTML = "";
    const editableCard = window.PPTXEngine.buildSlideDOMElement(currentSlide, true);
    editableCard.id = "editor-slide-canvas";

    const domElements = editableCard.querySelectorAll(".ppt-element");
    domElements.forEach((domEl) => {
      const elId = domEl.dataset.elementId;
      const modelEl = currentSlide.elements.find((item) => item.id === elId);
      if (!modelEl) return;

      if (elId === selectedElementId) {
        domEl.classList.add("selected-element");
      }

      if (modelEl.type === "text") {
        domEl.contentEditable = "true";
        domEl.spellcheck = false;

        domEl.addEventListener("input", () => {
          modelEl.text = domEl.innerText;
        });

        domEl.addEventListener("blur", () => {
          modelEl.text = domEl.innerText;
          pushUndoSnapshot();
          renderEditorFilmstrip();
        });
      }

      domEl.addEventListener("click", (e) => {
        e.stopPropagation();
        selectSlideElement(elId, domEl, modelEl);
      });

      attachElementDragHandlers(domEl, modelEl, editableCard);
    });

    editableCard.addEventListener("click", () => {
      syncActiveDOMTextToModel();
      selectedElementId = null;
      editableCard.querySelectorAll(".ppt-element").forEach((el) => el.classList.remove("selected-element"));
      hideContextPill();
    });

    stageWrapper.appendChild(editableCard);
  }

  function selectSlideElement(elId, domEl, modelEl) {
    syncActiveDOMTextToModel();
    selectedElementId = elId;

    const stage = document.getElementById("editor-slide-canvas");
    if (stage) {
      stage.querySelectorAll(".ppt-element").forEach((el) => el.classList.remove("selected-element"));
    }
    domEl.classList.add("selected-element");

    document.getElementById("btn-fmt-bold").classList.toggle("active", !!modelEl.bold);
    document.getElementById("btn-fmt-italic").classList.toggle("active", !!modelEl.italic);
    document.getElementById("btn-fmt-underline").classList.toggle("active", !!modelEl.underline);

    if (modelEl.color) {
      const colorInput = document.getElementById("input-fmt-color");
      if (colorInput && /^#[0-9a-f]{6}$/i.test(modelEl.color)) {
        colorInput.value = modelEl.color;
      }
      const colorLetter = document.querySelector(".color-letter");
      if (colorLetter) colorLetter.style.borderBottomColor = modelEl.color;
    }

    showContextPill();
  }

  function attachElementDragHandlers(domEl, modelEl, cardContainer) {
    domEl.addEventListener(
      "touchstart",
      (e) => {
        if (e.touches.length !== 1) return;
        if (selectedElementId !== modelEl.id) return;

        isDraggingEl = false;
        dragStartX = e.touches[0].clientX;
        dragStartY = e.touches[0].clientY;
        elInitialX = modelEl.x;
        elInitialY = modelEl.y;
      },
      { passive: true }
    );

    domEl.addEventListener(
      "touchmove",
      (e) => {
        if (selectedElementId !== modelEl.id || e.touches.length !== 1) return;
        const dx = e.touches[0].clientX - dragStartX;
        const dy = e.touches[0].clientY - dragStartY;

        if (!isDraggingEl && Math.hypot(dx, dy) > 10) {
          isDraggingEl = true;
        }

        if (isDraggingEl) {
          const rect = cardContainer.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            const deltaXPct = (dx / rect.width) * 100;
            const deltaYPct = (dy / rect.height) * 100;

            modelEl.x = Math.max(0, Math.min(92, Math.round((elInitialX + deltaXPct) * 10) / 10));
            modelEl.y = Math.max(0, Math.min(92, Math.round((elInitialY + deltaYPct) * 10) / 10));

            domEl.style.left = `${modelEl.x}%`;
            domEl.style.top = `${modelEl.y}%`;
          }
        }
      },
      { passive: true }
    );

    domEl.addEventListener("touchend", () => {
      if (isDraggingEl) {
        isDraggingEl = false;
        pushUndoSnapshot();
        renderEditorFilmstrip();
      }
    });
  }

  function syncActiveDOMTextToModel() {
    const stage = document.getElementById("editor-slide-canvas");
    const slide = workingSlidesClone[editorSlideIdx];
    if (!stage || !slide) return;

    stage.querySelectorAll(".ppt-text-box").forEach((domBox) => {
      const id = domBox.dataset.elementId;
      const item = slide.elements.find((e) => e.id === id);
      if (item) {
        item.text = domBox.innerText;
      }
    });
  }

  function renderEditorFilmstrip() {
    const track = document.getElementById("editor-thumbnails-track");
    if (!track) return;

    track.innerHTML = "";
    workingSlidesClone.forEach((slideObj, idx) => {
      const thumbItem = document.createElement("div");
      thumbItem.className = "editor-thumb-item" + (idx === editorSlideIdx ? " active" : "");

      const previewWrap = document.createElement("div");
      previewWrap.className = "editor-thumb-preview";

      const miniSlide = window.PPTXEngine.buildSlideDOMElement(slideObj, false);
      miniSlide.style.width = "100%";
      miniSlide.style.height = "100%";
      miniSlide.style.boxShadow = "none";
      miniSlide.querySelectorAll(".ppt-text-box").forEach((t) => {
        t.style.fontSize = "3.5px";
        t.style.lineHeight = "1.15";
      });

      previewWrap.appendChild(miniSlide);

      const numBar = document.createElement("div");
      numBar.className = "editor-thumb-number";
      numBar.textContent = String(idx + 1);

      thumbItem.appendChild(previewWrap);
      thumbItem.appendChild(numBar);

      thumbItem.addEventListener("click", () => {
        syncActiveDOMTextToModel();
        editorSlideIdx = idx;
        selectedElementId = null;
        hideContextPill();
        renderEditorStage();
        renderEditorFilmstrip();
      });

      track.appendChild(thumbItem);
    });
  }

  function getSelectedModelElement() {
    const slide = workingSlidesClone[editorSlideIdx];
    if (!slide || !selectedElementId) return null;
    return slide.elements.find((e) => e.id === selectedElementId) || null;
  }

  function applyTextFormatting(prop, toggleOrValue) {
    syncActiveDOMTextToModel();
    const el = getSelectedModelElement();
    if (!el || el.type !== "text") {
      window.AppStorage.showToast("Tap a text box on the slide first");
      return;
    }

    if (prop === "bold") el.bold = !el.bold;
    else if (prop === "italic") el.italic = !el.italic;
    else if (prop === "underline") el.underline = !el.underline;
    else if (prop === "color") el.color = toggleOrValue;
    else if (prop === "sizeDelta") {
      el.fontSize = Math.max(8, Math.min(44, (el.fontSize || 14) + toggleOrValue));
    }

    pushUndoSnapshot();
    renderEditorStage();
    renderEditorFilmstrip();
  }

  function addNewTextBoxToSlide() {
    syncActiveDOMTextToModel();
    const slide = workingSlidesClone[editorSlideIdx];
    if (!slide) return;

    const newId = "el-txt-" + Date.now();
    const newTextEl = {
      id: newId,
      type: "text",
      x: 18,
      y: 35,
      w: 64,
      h: 18,
      text: "Tap to edit text",
      fontSize: 16,
      color: "#111111",
      bold: true,
      italic: false,
      underline: false,
      bgColor: "transparent"
    };

    slide.elements.push(newTextEl);
    selectedElementId = newId;
    pushUndoSnapshot();
    renderEditorStage();
    renderEditorFilmstrip();
    showContextPill();
  }

  function handleEditorImageSelection(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target.result;
      const slide = workingSlidesClone[editorSlideIdx];
      if (!slide) return;

      const selectedEl = getSelectedModelElement();
      if (selectedEl && selectedEl.type === "image") {
        selectedEl.src = dataUrl;
        window.AppStorage.showToast("Image replaced");
      } else {
        const newImgId = "el-img-" + Date.now();
        slide.elements.push({
          id: newImgId,
          type: "image",
          x: 25,
          y: 25,
          w: 45,
          h: 45,
          src: dataUrl
        });
        selectedElementId = newImgId;
        window.AppStorage.showToast("Image inserted");
      }

      pushUndoSnapshot();
      renderEditorStage();
      renderEditorFilmstrip();
    };
    reader.readAsDataURL(file);
  }

  function showContextPill() {
    const pill = document.getElementById("editor-context-pill");
    if (pill) pill.classList.remove("hidden");
  }

  function hideContextPill() {
    const pill = document.getElementById("editor-context-pill");
    if (pill) pill.classList.add("hidden");
  }

  function cutSelectedElement() {
    syncActiveDOMTextToModel();
    const slide = workingSlidesClone[editorSlideIdx];
    const el = getSelectedModelElement();
    if (!slide || !el) return;

    clipboardElement = JSON.parse(JSON.stringify(el));
    slide.elements = slide.elements.filter((item) => item.id !== el.id);
    selectedElementId = null;
    hideContextPill();
    pushUndoSnapshot();
    renderEditorStage();
    renderEditorFilmstrip();
    window.AppStorage.showToast("Cut to clipboard");
  }

  function copySelectedElement() {
    syncActiveDOMTextToModel();
    const el = getSelectedModelElement();
    if (!el) return;

    clipboardElement = JSON.parse(JSON.stringify(el));
    window.AppStorage.showToast("Copied");
  }

  function pasteClipboardElement() {
    if (!clipboardElement) {
      window.AppStorage.showToast("Clipboard is empty");
      return;
    }
    const slide = workingSlidesClone[editorSlideIdx];
    if (!slide) return;

    const pasted = JSON.parse(JSON.stringify(clipboardElement));
    pasted.id = "el-paste-" + Date.now();
    pasted.x = Math.min(85, (pasted.x || 10) + 4);
    pasted.y = Math.min(85, (pasted.y || 10) + 4);

    slide.elements.push(pasted);
    selectedElementId = pasted.id;
    pushUndoSnapshot();
    renderEditorStage();
    renderEditorFilmstrip();
    window.AppStorage.showToast("Pasted");
  }

  function deleteSelectedElement() {
    const slide = workingSlidesClone[editorSlideIdx];
    if (!slide || !selectedElementId) return;

    slide.elements = slide.elements.filter((item) => item.id !== selectedElementId);
    selectedElementId = null;
    hideContextPill();
    pushUndoSnapshot();
    renderEditorStage();
    renderEditorFilmstrip();
    window.AppStorage.showToast("Deleted element");
  }

  async function persistCurrentPresentationToStorage(triggerDownloadFile = false) {
    const record = window.PPTXEngine.getActiveFileRecord();
    const slides = window.PPTXEngine.getActiveSlidesData();
    if (!record || !slides) return;

    const updatedBlob = await window.PPTXEngine.packSlidesIntoPPTXBlob(slides);
    record.blob = updatedBlob;
    record.customSlidesData = JSON.parse(JSON.stringify(slides));
    record.lastOpened = Date.now();

    await window.AppStorage.saveFileRecord(record);

    if (triggerDownloadFile) {
      const fileName = record.name.endsWith(".pptx") ? record.name : record.name + ".pptx";
      window.AppStorage.triggerBlobDownload(updatedBlob, fileName);
    }
  }

  async function exportAndShareAsPDF() {
    window.AppStorage.closeAllModals();
    const slides = window.PPTXEngine.getActiveSlidesData();
    const record = window.PPTXEngine.getActiveFileRecord();
    if (!slides || slides.length === 0) return;

    window.AppStorage.showLoading("Converting slides to PDF...");

    try {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        orientation: "landscape",
        unit: "px",
        format: [960, 540]
      });

      const offscreenHost = document.createElement("div");
      offscreenHost.style.position = "fixed";
      offscreenHost.style.left = "-9999px";
      offscreenHost.style.top = "0";
      offscreenHost.style.width = "960px";
      offscreenHost.style.height = "540px";
      document.body.appendChild(offscreenHost);

      for (let i = 0; i < slides.length; i++) {
        offscreenHost.innerHTML = "";
        const slideCard = window.PPTXEngine.buildSlideDOMElement(slides[i], false);
        slideCard.style.width = "960px";
        slideCard.style.height = "540px";

        slides[i].elements.forEach((el) => {
          if (el.type === "text") {
            const domTxt = slideCard.querySelector(`[data-element-id="${el.id}"]`);
            if (domTxt) {
              domTxt.style.fontSize = `${Math.round((el.fontSize || 14) * 1.65)}px`;
            }
          }
        });

        offscreenHost.appendChild(slideCard);

        const canvas = await html2canvas(slideCard, {
          scale: 1.5,
          useCORS: true,
          backgroundColor: slides[i].background || "#ffffff"
        });

        const imgData = canvas.toDataURL("image/jpeg", 0.92);
        if (i > 0) pdf.addPage([960, 540], "landscape");
        pdf.addImage(imgData, "JPEG", 0, 0, 960, 540);
      }

      document.body.removeChild(offscreenHost);

      const baseName = (record ? record.name : "Presentation").replace(/\.pptx$/i, "");
      const pdfFileName = `${baseName}.pdf`;
      const pdfBlob = pdf.output("blob");
      const pdfFileObj = new File([pdfBlob], pdfFileName, { type: "application/pdf" });

      window.AppStorage.hideLoading();

      if (navigator.canShare && navigator.canShare({ files: [pdfFileObj] })) {
        try {
          await navigator.share({
            title: pdfFileName,
            files: [pdfFileObj]
          });
          return;
        } catch (shareErr) {
          if (shareErr.name === "AbortError") return;
        }
      }

      window.AppStorage.triggerBlobDownload(pdfBlob, pdfFileName);
      window.AppStorage.showToast("Saved PDF: " + pdfFileName);
    } catch (err) {
      console.error("PDF Export failed:", err);
      window.AppStorage.hideLoading();
      window.AppStorage.showToast("Could not generate PDF");
    }
  }

  async function handleSaveMenuAction() {
    window.AppStorage.closeAllModals();
    window.AppStorage.showLoading("Saving presentation...");
    try {
      await persistCurrentPresentationToStorage(false);
      window.AppStorage.showToast("Saved successfully ✓");
    } finally {
      window.AppStorage.hideLoading();
    }
  }

  async function handleSaveAsMenuAction() {
    window.AppStorage.closeAllModals();
    await persistCurrentPresentationToStorage(false);
    const record = window.PPTXEngine.getActiveFileRecord();
    if (record) {
      window.AppStorage.setSelectedFileId(record.id);
      window.AppStorage.promptRenameSelectedFile("saveAs");
    }
  }

  function handlePrintMenuAction() {
    window.AppStorage.closeAllModals();
    window.AppStorage.showToast("Opening Printer Service...");
    setTimeout(() => {
      window.print();
    }, 250);
  }

  window.addEventListener("DOMContentLoaded", () => {
    document.getElementById("btn-enter-edit-mode").addEventListener("click", () => {
      openEditorAtSlide(window.PPTXEngine.getCurrentVisibleSlideIndex());
    });

    document.getElementById("btn-editor-save-done").addEventListener("click", saveEditsAndReturn);
    document.getElementById("btn-editor-discard").addEventListener("click", discardEditsAndReturn);
    document.getElementById("btn-editor-undo").addEventListener("click", triggerUndo);
    document.getElementById("btn-editor-play").addEventListener("click", () => {
      syncActiveDOMTextToModel();
      window.PPTXEngine.setActiveSlidesData(JSON.parse(JSON.stringify(workingSlidesClone)));
      window.GesturesPlayer.launchFullscreenPlayer(editorSlideIdx);
    });

    document.getElementById("btn-fmt-bold").addEventListener("click", () => applyTextFormatting("bold"));
    document.getElementById("btn-fmt-italic").addEventListener("click", () => applyTextFormatting("italic"));
    document.getElementById("btn-fmt-underline").addEventListener("click", () => applyTextFormatting("underline"));

    document.getElementById("input-fmt-color").addEventListener("input", (e) => {
      const colorLetter = document.querySelector(".color-letter");
      if (colorLetter) colorLetter.style.borderBottomColor = e.target.value;
      applyTextFormatting("color", e.target.value);
    });

    document.getElementById("btn-fmt-size-down").addEventListener("click", () => applyTextFormatting("sizeDelta", -2));
    document.getElementById("btn-fmt-size-up").addEventListener("click", () => applyTextFormatting("sizeDelta", 2));
    document.getElementById("btn-fmt-add-text").addEventListener("click", addNewTextBoxToSlide);

    const imgInput = document.getElementById("editor-image-picker");
    document.getElementById("btn-fmt-image").addEventListener("click", () => imgInput.click());
    imgInput.addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        handleEditorImageSelection(e.target.files[0]);
      }
      imgInput.value = "";
    });

    document.getElementById("btn-ctx-cut").addEventListener("click", cutSelectedElement);
    document.getElementById("btn-ctx-copy").addEventListener("click", copySelectedElement);
    document.getElementById("btn-ctx-paste").addEventListener("click", pasteClipboardElement);
    document.getElementById("btn-ctx-delete").addEventListener("click", deleteSelectedElement);

    document.getElementById("action-export-pdf").addEventListener("click", exportAndShareAsPDF);
    document.getElementById("action-save-ppt").addEventListener("click", handleSaveMenuAction);
    document.getElementById("action-save-as-ppt").addEventListener("click", handleSaveAsMenuAction);
    document.getElementById("action-print-ppt").addEventListener("click", handlePrintMenuAction);
  });

  window.EditorExport = {
    openEditorAtSlide,
    exportAndShareAsPDF
  };
})();
