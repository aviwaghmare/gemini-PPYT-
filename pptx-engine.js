/* ==========================================================================
   POWERPOINT LITE — OPENXML .PPTX PARSER, RENDERER & FIND ENGINE (www/pptx-engine.js)
   ========================================================================== */

(function () {
  "use strict";

  let activeFileRecord = null;
  let activeSlidesData = [];
  let currentVisibleSlideIndex = 0;

  let findMatches = [];
  let currentFindMatchIdx = -1;

  const DEFAULT_SLIDE_WIDTH_EMU = 9144000;
  const DEFAULT_SLIDE_HEIGHT_EMU = 5143500;

  async function loadPresentationRecord(fileRecord) {
    activeFileRecord = fileRecord;
    window.AppStorage.setSelectedFileId(fileRecord.id);

    const titleEl = document.getElementById("reader-file-title");
    if (titleEl) {
      titleEl.textContent = fileRecord.name.replace(/\.pptx$/i, "");
    }

    if (fileRecord.customSlidesData && Array.isArray(fileRecord.customSlidesData) && fileRecord.customSlidesData.length > 0) {
      activeSlidesData = JSON.parse(JSON.stringify(fileRecord.customSlidesData));
    } else if (fileRecord.blob) {
      activeSlidesData = await parsePPTXBlob(fileRecord.blob);
    } else {
      activeSlidesData = getBuiltInCacheMemorySlides();
    }

    currentVisibleSlideIndex = 0;
    renderVerticalReaderSlides();
    closeFindBar();
  }

  async function parsePPTXBlob(blob) {
    try {
      const zip = await JSZip.loadAsync(blob);

      const customStateFile = zip.file("ppt/ppt_lite_state.json");
      if (customStateFile) {
        const stateJson = await customStateFile.async("string");
        const parsedState = JSON.parse(stateJson);
        if (Array.isArray(parsedState) && parsedState.length > 0) {
          return parsedState;
        }
      }

      let slideWidthEMU = DEFAULT_SLIDE_WIDTH_EMU;
      let slideHeightEMU = DEFAULT_SLIDE_HEIGHT_EMU;

      const presXmlFile = zip.file("ppt/presentation.xml");
      if (presXmlFile) {
        const presXmlStr = await presXmlFile.async("string");
        const parser = new DOMParser();
        const presDoc = parser.parseFromString(presXmlStr, "application/xml");
        const sldSz = presDoc.getElementsByTagName("p:sldSz")[0] || presDoc.getElementsByTagName("sldSz")[0];
        if (sldSz) {
          const cx = parseInt(sldSz.getAttribute("cx"), 10);
          const cy = parseInt(sldSz.getAttribute("cy"), 10);
          if (cx > 0 && cy > 0) {
            slideWidthEMU = cx;
            slideHeightEMU = cy;
          }
        }
      }

      const slideFileNames = Object.keys(zip.files)
        .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
        .sort((a, b) => {
          const numA = parseInt(a.match(/slide(\d+)\.xml/i)[1], 10);
          const numB = parseInt(b.match(/slide(\d+)\.xml/i)[1], 10);
          return numA - numB;
        });

      if (slideFileNames.length === 0) {
        return getBuiltInCacheMemorySlides();
      }

      const parsedSlides = [];
      const domParser = new DOMParser();

      for (let i = 0; i < slideFileNames.length; i++) {
        const slidePath = slideFileNames[i];
        const slideNum = i + 1;
        const slideXmlStr = await zip.file(slidePath).async("string");
        const slideDoc = domParser.parseFromString(slideXmlStr, "application/xml");

        const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
        const relMap = {};
        if (zip.file(relsPath)) {
          const relsStr = await zip.file(relsPath).async("string");
          const relsDoc = domParser.parseFromString(relsStr, "application/xml");
          const relNodes = relsDoc.getElementsByTagName("Relationship");
          for (let r = 0; r < relNodes.length; r++) {
            const id = relNodes[r].getAttribute("Id");
            const target = relNodes[r].getAttribute("Target");
            if (id && target) {
              relMap[id] = target.replace(/^\.\.\//, "ppt/");
            }
          }
        }

        const slideElements = [];
        const spTree = slideDoc.getElementsByTagName("p:spTree")[0] || slideDoc.getElementsByTagName("spTree")[0] || slideDoc;

        const shapes = spTree.getElementsByTagName("p:sp");
        for (let s = 0; s < shapes.length; s++) {
          const sp = shapes[s];
          const box = extractBoundingBoxPct(sp, slideWidthEMU, slideHeightEMU);
          const paragraphs = extractParagraphsFromShape(sp);

          if (paragraphs.length > 0) {
            const combinedText = paragraphs.map((p) => p.text).join("\n").trim();
            if (combinedText.length > 0) {
              const firstPara = paragraphs[0];
              slideElements.push({
                id: `el-s${slideNum}-txt-${s}`,
                type: "text",
                x: box.x,
                y: box.y,
                w: box.w,
                h: box.h,
                text: combinedText,
                fontSize: firstPara.fontSize || 15,
                color: firstPara.color || "#1f2937",
                bold: firstPara.bold || false,
                italic: firstPara.italic || false,
                underline: firstPara.underline || false,
                bgColor: box.bgColor || "transparent"
              });
            }
          } else if (box.bgColor && box.bgColor !== "transparent") {
            slideElements.push({
              id: `el-s${slideNum}-shp-${s}`,
              type: "shape",
              x: box.x,
              y: box.y,
              w: box.w,
              h: box.h,
              bgColor: box.bgColor
            });
          }
        }

        const pics = spTree.getElementsByTagName("p:pic");
        for (let p = 0; p < pics.length; p++) {
          const pic = pics[p];
          const box = extractBoundingBoxPct(pic, slideWidthEMU, slideHeightEMU);
          const blip = pic.getElementsByTagName("a:blip")[0] || pic.getElementsByTagName("blip")[0];
          if (blip) {
            const embedId = blip.getAttribute("r:embed") || blip.getAttribute("embed");
            const mediaPath = relMap[embedId];
            if (mediaPath && zip.file(mediaPath)) {
              const base64Data = await zip.file(mediaPath).async("base64");
              const ext = mediaPath.split(".").pop().toLowerCase();
              const mime = ext === "png" ? "image/png" : ext === "svg" ? "image/svg+xml" : "image/jpeg";
              slideElements.push({
                id: `el-s${slideNum}-img-${p}`,
                type: "image",
                x: box.x,
                y: box.y,
                w: box.w,
                h: box.h,
                src: `data:${mime};base64,${base64Data}`
              });
            }
          }
        }

        if (slideElements.length === 0) {
          const rawTextNodes = slideDoc.getElementsByTagName("a:t");
          const lines = [];
          for (let t = 0; t < rawTextNodes.length; t++) {
            if (rawTextNodes[t].textContent.trim()) {
              lines.push(rawTextNodes[t].textContent.trim());
            }
          }
          if (lines.length > 0) {
            slideElements.push({
              id: `el-s${slideNum}-fallback`,
              type: "text",
              x: 6,
              y: 10,
              w: 88,
              h: 80,
              text: lines.join("\n\n"),
              fontSize: 16,
              color: "#111111",
              bold: false
            });
          }
        }

        parsedSlides.push({
          index: slideNum,
          background: "#ffffff",
          elements: slideElements
        });
      }

      return parsedSlides;
    } catch (err) {
      console.warn("Standard OpenXML parse fallback triggered:", err);
      return getBuiltInCacheMemorySlides();
    }
  }

  function extractBoundingBoxPct(node, slideW, slideH) {
    let x = 8, y = 10, w = 84, h = 22, bgColor = "transparent";

    const xfrm = node.getElementsByTagName("a:xfrm")[0] || node.getElementsByTagName("xfrm")[0];
    if (xfrm) {
      const off = xfrm.getElementsByTagName("a:off")[0] || xfrm.getElementsByTagName("off")[0];
      const ext = xfrm.getElementsByTagName("a:ext")[0] || xfrm.getElementsByTagName("ext")[0];
      if (off) {
        const ox = parseInt(off.getAttribute("x"), 10);
        const oy = parseInt(off.getAttribute("y"), 10);
        if (!isNaN(ox)) x = Math.max(0, Math.min(95, (ox / slideW) * 100));
        if (!isNaN(oy)) y = Math.max(0, Math.min(95, (oy / slideH) * 100));
      }
      if (ext) {
        const ex = parseInt(ext.getAttribute("cx"), 10);
        const ey = parseInt(ext.getAttribute("cy"), 10);
        if (!isNaN(ex) && ex > 0) w = Math.max(5, Math.min(100 - x, (ex / slideW) * 100));
        if (!isNaN(ey) && ey > 0) h = Math.max(4, Math.min(100 - y, (ey / slideH) * 100));
      }
    }

    const spPr = node.getElementsByTagName("p:spPr")[0];
    if (spPr) {
      const solidFill = spPr.getElementsByTagName("a:solidFill")[0];
      if (solidFill) {
        const srgb = solidFill.getElementsByTagName("a:srgbClr")[0];
        if (srgb && srgb.getAttribute("val")) {
          bgColor = "#" + srgb.getAttribute("val");
        }
      }
    }

    return { x, y, w, h, bgColor };
  }

  function extractParagraphsFromShape(spNode) {
    const paragraphs = [];
    const pNodes = spNode.getElementsByTagName("a:p");

    for (let i = 0; i < pNodes.length; i++) {
      const p = pNodes[i];
      const runs = p.getElementsByTagName("a:r");
      let paraText = "";
      let fontSize = 14;
      let color = "#1f2937";
      let bold = false;
      let italic = false;
      let underline = false;

      for (let r = 0; r < runs.length; r++) {
        const run = runs[r];
        const rPr = run.getElementsByTagName("a:rPr")[0];
        if (rPr) {
          const sz = parseInt(rPr.getAttribute("sz"), 10);
          if (!isNaN(sz) && sz > 0) {
            fontSize = Math.round(sz / 100);
          }
          if (rPr.getAttribute("b") === "1") bold = true;
          if (rPr.getAttribute("i") === "1") italic = true;
          if (rPr.getAttribute("u") && rPr.getAttribute("u") !== "none") underline = true;

          const srgb = rPr.getElementsByTagName("a:srgbClr")[0];
          if (srgb && srgb.getAttribute("val")) {
            color = "#" + srgb.getAttribute("val");
          }
        }
        const tNode = run.getElementsByTagName("a:t")[0];
        if (tNode) {
          paraText += tNode.textContent;
        }
      }

      if (!paraText) {
        const directT = p.getElementsByTagName("a:t");
        for (let t = 0; t < directT.length; t++) {
          paraText += directT[t].textContent;
        }
      }

      if (paraText.trim().length > 0) {
        paragraphs.push({ text: paraText, fontSize, color, bold, italic, underline });
      }
    }

    return paragraphs;
  }

  function renderVerticalReaderSlides() {
    const listContainer = document.getElementById("reader-slides-list");
    const counterBadge = document.getElementById("slide-counter-badge");
    if (!listContainer) return;

    listContainer.innerHTML = "";
    const total = activeSlidesData.length || 1;

    activeSlidesData.forEach((slideObj, idx) => {
      const wrapper = document.createElement("div");
      wrapper.className = "slide-card-container";
      wrapper.dataset.slideIndex = idx;

      const card = buildSlideDOMElement(slideObj, false);

      let lastTap = 0;
      card.addEventListener("touchend", (e) => {
        const now = Date.now();
        if (now - lastTap < 280) {
          e.preventDefault();
          if (window.EditorExport && window.EditorExport.openEditorAtSlide) {
            window.EditorExport.openEditorAtSlide(idx);
          }
        }
        lastTap = now;
      });

      wrapper.appendChild(card);
      listContainer.appendChild(wrapper);
    });

    if (counterBadge) {
      counterBadge.textContent = `Slide ${currentVisibleSlideIndex + 1} of ${total}`;
    }

    const scrollViewport = document.getElementById("reader-scroll-viewport");
    if (scrollViewport) {
      scrollViewport.onscroll = () => {
        const containers = listContainer.querySelectorAll(".slide-card-container");
        const viewportCenter = scrollViewport.scrollTop + scrollViewport.clientHeight * 0.38;

        containers.forEach((c, idx) => {
          const top = c.offsetTop;
          const bottom = top + c.offsetHeight;
          if (viewportCenter >= top && viewportCenter <= bottom) {
            currentVisibleSlideIndex = idx;
            if (counterBadge) {
              counterBadge.textContent = `Slide ${idx + 1} of ${total}`;
            }
          }
        });
      };
    }
  }

  function buildSlideDOMElement(slideObj, isEditable = false) {
    const card = document.createElement("div");
    card.className = "slide-card" + (isEditable ? " editable-stage" : "");
    card.style.backgroundColor = slideObj.background || "#ffffff";

    (slideObj.elements || []).forEach((el) => {
      const domEl = document.createElement("div");
      domEl.className = `ppt-element ppt-${el.type}-box`;
      domEl.dataset.elementId = el.id;
      domEl.style.left = `${el.x}%`;
      domEl.style.top = `${el.y}%`;
      domEl.style.width = `${el.w}%`;
      domEl.style.height = `${el.h}%`;

      if (el.bgColor && el.bgColor !== "transparent") {
        domEl.style.backgroundColor = el.bgColor;
      }
      if (el.borderRadius) {
        domEl.style.borderRadius = el.borderRadius;
      }

      if (el.type === "text") {
        const scaledSize = Math.max(8, Math.min(28, (el.fontSize || 14) * 0.68));
        domEl.style.fontSize = `${scaledSize}px`;
        domEl.style.color = el.color || "#111111";
        domEl.style.fontWeight = el.bold ? "700" : "400";
        domEl.style.fontStyle = el.italic ? "italic" : "normal";
        domEl.style.textDecoration = el.underline ? "underline" : "none";
        domEl.style.whiteSpace = "pre-wrap";
        domEl.textContent = el.text || "";
      } else if (el.type === "image") {
        const img = document.createElement("img");
        img.src = el.src;
        img.alt = "Slide visual";
        domEl.appendChild(img);
      } else if (el.type === "shape" && el.htmlContent) {
        domEl.innerHTML = el.htmlContent;
      }

      card.appendChild(domEl);
    });

    return card;
  }

  function openFindBar() {
    const bar = document.getElementById("slide-find-bar");
    const input = document.getElementById("slide-find-input");
    if (!bar || !input) return;

    bar.classList.remove("hidden");
    input.focus();
  }

  function closeFindBar() {
    const bar = document.getElementById("slide-find-bar");
    const input = document.getElementById("slide-find-input");
    const countEl = document.getElementById("slide-find-count");
    if (bar) bar.classList.add("hidden");
    if (input) input.value = "";
    if (countEl) countEl.textContent = "0/0";

    clearAllFindHighlights();
    findMatches = [];
    currentFindMatchIdx = -1;
  }

  function clearAllFindHighlights() {
    const listContainer = document.getElementById("reader-slides-list");
    if (!listContainer) return;

    const textBoxes = listContainer.querySelectorAll(".ppt-text-box");
    textBoxes.forEach((box) => {
      if (box.dataset.rawText !== undefined) {
        box.textContent = box.dataset.rawText;
      }
    });
  }

  function executeSlideFind(queryStr) {
    clearAllFindHighlights();
    findMatches = [];
    currentFindMatchIdx = -1;

    const countEl = document.getElementById("slide-find-count");
    const q = (queryStr || "").trim();
    if (!q) {
      if (countEl) countEl.textContent = "0/0";
      return;
    }

    const listContainer = document.getElementById("reader-slides-list");
    if (!listContainer) return;

    const regex = new RegExp(`(${escapeRegExp(q)})`, "gi");
    const textBoxes = listContainer.querySelectorAll(".ppt-text-box");

    textBoxes.forEach((box) => {
      const raw = box.dataset.rawText !== undefined ? box.dataset.rawText : box.textContent;
      box.dataset.rawText = raw;

      if (regex.test(raw)) {
        const safeHtml = escapeHtml(raw).replace(
          regex,
          `<mark class="ppt-find-match">$1</mark>`
        );
        box.innerHTML = safeHtml;
      }
    });

    findMatches = Array.from(listContainer.querySelectorAll("mark.ppt-find-match"));
    if (findMatches.length > 0) {
      currentFindMatchIdx = 0;
      focusActiveFindMatch();
    } else if (countEl) {
      countEl.textContent = "0/0";
    }
  }

  function focusActiveFindMatch() {
    const countEl = document.getElementById("slide-find-count");
    if (findMatches.length === 0) {
      if (countEl) countEl.textContent = "0/0";
      return;
    }

    findMatches.forEach((m) => m.classList.remove("active-match"));
    const activeMark = findMatches[currentFindMatchIdx];
    if (activeMark) {
      activeMark.classList.add("active-match");
      activeMark.scrollIntoView({ behavior: "smooth", block: "center" });
    }

    if (countEl) {
      countEl.textContent = `${currentFindMatchIdx + 1}/${findMatches.length}`;
    }
  }

  function getBuiltInCacheMemorySlides() {
    return [
      {
        index: 1,
        background: "#ffffff",
        elements: [
          {
            id: "s1-banner",
            type: "shape",
            x: 2, y: 3, w: 80, h: 12,
            bgColor: "#0c7bb3",
            htmlContent: `<div style="width:100%;height:100%;background:repeating-linear-gradient(45deg,#0c7bb3,#0c7bb3 12px,#ffffff 12px,#ffffff 20px);opacity:0.9;"></div>`
          },
          {
            id: "s1-badge",
            type: "text",
            x: 84, y: 5, w: 14, h: 8,
            text: "MIT CSN",
            fontSize: 13, color: "#0c5a8a", bold: true
          },
          {
            id: "s1-title",
            type: "text",
            x: 5, y: 18, w: 48, h: 20,
            text: "CACHE\nMEMORY",
            fontSize: 28, color: "#125b9a", bold: true
          },
          {
            id: "s1-intro-hdr",
            type: "text",
            x: 4, y: 39, w: 30, h: 6,
            text: "INTRODUCTION",
            fontSize: 11, color: "#c0392b", bold: true
          },
          {
            id: "s1-body",
            type: "text",
            x: 4, y: 45, w: 50, h: 24,
            text: "A computer's performance depends on both its processing speed and how quickly it can supply instructions and data to the processor. Even a fast processor may spend time waiting for main memory.\n\nCache memory helps reduce this waiting time by keeping useful instructions and data close to the processor.",
            fontSize: 11, color: "#222222", bold: false
          },
          {
            id: "s1-credits",
            type: "text",
            x: 5, y: 71, w: 48, h: 24,
            text: "Subject: Digital Electronics and Computer Architecture (DECA)\nPresented by:  Avinash waghmare [F-50]\n                        Aakash shinde [F-67]\n                        Yuvraj Vaidya [F-40]\n                        Gaurav yedke [F-59]",
            fontSize: 10.5, color: "#555555", bold: false
          },
          {
            id: "s1-graphic",
            type: "shape",
            x: 56, y: 20, w: 40, h: 72,
            bgColor: "transparent",
            htmlContent: `
              <svg viewBox="0 0 220 160" width="100%" height="100%">
                <rect x="135" y="18" width="68" height="118" rx="4" fill="#dfe6e9" stroke="#636e72" stroke-width="2"/>
                <line x1="135" y1="55" x2="203" y2="55" stroke="#636e72" stroke-width="2"/>
                <line x1="135" y1="92" x2="203" y2="92" stroke="#636e72" stroke-width="2"/>
                <rect x="25" y="65" width="85" height="52" rx="4" fill="#74b9ff" stroke="#0984e3" stroke-width="2"/>
                <circle cx="68" cy="48" r="12" fill="#fab1a0"/>
                <rect x="52" y="62" width="32" height="45" rx="6" fill="#0984e3"/>
                <text x="169" y="40" font-size="9" fill="#2d3436" text-anchor="middle" font-weight="bold">CACHE</text>
                <text x="169" y="77" font-size="9" fill="#2d3436" text-anchor="middle" font-weight="bold">RAM</text>
                <text x="169" y="114" font-size="9" fill="#2d3436" text-anchor="middle" font-weight="bold">DISK</text>
              </svg>
            `
          }
        ]
      },
      {
        index: 2,
        background: "#ffffff",
        elements: [
          {
            id: "s2-title",
            type: "text",
            x: 4, y: 6, w: 72, h: 9,
            text: "FOUNDATION: CPU, RAM AND THE NEED FOR CACHE",
            fontSize: 16, color: "#b85d3b", bold: true
          },
          {
            id: "s2-sub1",
            type: "text",
            x: 4, y: 16, w: 90, h: 12,
            text: "WHAT IS A COMPUTER?\nA computer is an electronic system that accepts input, processes data according to instructions, stores information, and produces output.",
            fontSize: 10.5, color: "#1c4e80", bold: true
          },
          {
            id: "s2-col1",
            type: "text",
            x: 4, y: 31, w: 50, h: 42,
            text: "TWO ESSENTIAL COMPONENTS\n\n1. CPU — Central Processing Unit\nThe CPU executes program instructions, performs calculations, and controls computer operations. It needs a continuous supply of instructions and data.\n\n2. RAM — Random Access Memory\nMain memory, usually implemented using DRAM, temporarily holds programs and data currently in use. It is volatile: its contents are lost when power is removed.",
            fontSize: 10.5, color: "#2c3e50", bold: false
          },
          {
            id: "s2-bottom",
            type: "text",
            x: 4, y: 74, w: 52, h: 22,
            text: "WHY IS ANOTHER MEMORY NEEDED?\nAccessing main memory takes longer than many CPU operations. If the CPU repeatedly waits for RAM, overall execution slows down. Cache reduces the average time needed to obtain instructions and data.",
            fontSize: 10, color: "#7f3b18", bold: false
          },
          {
            id: "s2-pyramid",
            type: "shape",
            x: 58, y: 30, w: 38, h: 64,
            bgColor: "transparent",
            htmlContent: `
              <svg viewBox="0 0 240 165" width="100%" height="100%">
                <text x="120" y="14" font-size="10" fill="#125b9a" font-weight="bold" text-anchor="middle">MEMORY HIERARCHY</text>
                <polygon points="120,22 92,50 148,50" fill="#d35400"/>
                <text x="120" y="44" font-size="8" fill="#fff" text-anchor="middle">Registers</text>
                <polygon points="90,53 66,80 174,80 150,53" fill="#6c5ce7"/>
                <text x="120" y="70" font-size="9" fill="#fff" text-anchor="middle" font-weight="bold">Cache</text>
                <polygon points="64,83 38,112 202,112 176,83" fill="#0984e3"/>
                <text x="120" y="101" font-size="9" fill="#fff" text-anchor="middle" font-weight="bold">Main Memory (RAM)</text>
                <polygon points="36,115 12,144 228,144 204,115" fill="#2d3436"/>
                <text x="120" y="133" font-size="9" fill="#fff" text-anchor="middle">Secondary Storage</text>
              </svg>
            `
          }
        ]
      },
      {
        index: 3,
        background: "#ffffff",
        elements: [
          {
            id: "s3-title",
            type: "text",
            x: 5, y: 6, w: 80, h: 9,
            text: "CACHE MEMORY: DEFINITION AND BASIC PRINCIPLE",
            fontSize: 16, color: "#6c4298", bold: true
          },
          {
            id: "s3-def",
            type: "text",
            x: 5, y: 16, w: 90, h: 14,
            text: "ACADEMIC DEFINITION\nCache memory is a small, high-speed memory that stores copies of selected instructions and data from a lower memory level to reduce the processor's average memory access time.",
            fontSize: 11, color: "#1b4f72", bold: false
          },
          {
            id: "s3-simple",
            type: "text",
            x: 5, y: 33, w: 42, h: 22,
            text: "SIMPLE EXPLANATION\nImagine studying at a desk. The bookshelf represents RAM, while the pages kept on your desk represent cache. Keeping useful material nearby saves repeated trips to the bookshelf.",
            fontSize: 10.5, color: "#a04000", bold: false
          },
          {
            id: "s3-desk-img",
            type: "shape",
            x: 12, y: 54, w: 26, h: 28,
            bgColor: "transparent",
            htmlContent: `
              <svg viewBox="0 0 140 100" width="100%" height="100%">
                <rect x="75" y="10" width="48" height="75" fill="#8d5524" stroke="#5c3818" stroke-width="2"/>
                <rect x="80" y="16" width="38" height="14" fill="#f1c40f"/>
                <rect x="80" y="34" width="38" height="14" fill="#e74c3c"/>
                <rect x="80" y="52" width="38" height="14" fill="#3498db"/>
                <rect x="18" y="56" width="75" height="8" fill="#a0522d"/>
                <line x1="25" y1="64" x2="25" y2="92" stroke="#5c3818" stroke-width="4"/>
                <line x1="85" y1="64" x2="85" y2="92" stroke="#5c3818" stroke-width="4"/>
                <circle cx="52" cy="34" r="10" fill="#ffdbac"/>
                <rect x="38" y="44" width="28" height="14" rx="4" fill="#1e3799"/>
              </svg>
            `
          },
          {
            id: "s3-locality",
            type: "text",
            x: 50, y: 33, w: 46, h: 48,
            text: "WHY DOES THIS WORK?\nPrograms often reuse recently accessed information or access nearby memory locations. This behaviour is called locality of reference.\n\n1. TEMPORAL LOCALITY — REUSE OVER TIME\nRecently accessed instructions or data are likely to be accessed again soon.\nExample: The same instructions execute repeatedly inside a loop.\n\n2. SPATIAL LOCALITY — NEARBY LOCATIONS\nMemory locations near a recently accessed location are likely to be accessed soon.\nExample: Reading consecutive elements of an array.",
            fontSize: 10.5, color: "#1f2937", bold: false
          },
          {
            id: "s3-footer",
            type: "text",
            x: 5, y: 84, w: 90, h: 12,
            text: "CACHE BLOCK / CACHE LINE\nData is normally transferred into cache in fixed-size blocks containing neighbouring bytes. A cache line is the cache entry that holds one such block, along with identification and status information.",
            fontSize: 9.8, color: "#125b9a", bold: false
          }
        ]
      },
      {
        index: 4,
        background: "#ffffff",
        elements: [
          {
            id: "s4-title",
            type: "text",
            x: 5, y: 7, w: 85, h: 10,
            text: "WORKING OF CACHE: HIT AND MISS",
            fontSize: 17, color: "#0c5a8a", bold: true
          },
          {
            id: "s4-sub",
            type: "text",
            x: 5, y: 18, w: 50, h: 8,
            text: "BASIC READ OPERATION",
            fontSize: 12, color: "#125b9a", bold: true
          },
          {
            id: "s4-step1",
            type: "text",
            x: 5, y: 30, w: 28, h: 26,
            bgColor: "#1565c0",
            text: "Step 1: The CPU requests an instruction or data item using a memory address.",
            fontSize: 11, color: "#ffffff", bold: true
          },
          {
            id: "s4-step2",
            type: "text",
            x: 36, y: 30, w: 28, h: 26,
            bgColor: "#1565c0",
            text: "Step 2: The cache checks whether the requested block is currently stored.",
            fontSize: 11, color: "#ffffff", bold: true
          },
          {
            id: "s4-step3-hit",
            type: "text",
            x: 67, y: 26, w: 28, h: 28,
            bgColor: "#6a1b9a",
            text: "Step 3A — Cache Hit:\nIf the block is present and valid, the cache supplies the requested information immediately to the CPU.",
            fontSize: 10.5, color: "#ffffff", bold: true
          },
          {
            id: "s4-step3-miss",
            type: "text",
            x: 67, y: 60, w: 28, h: 28,
            bgColor: "#c62828",
            text: "Step 3B — Cache Miss:\nIf not found in cache, the block is fetched from Main Memory (RAM) and copied into cache.",
            fontSize: 10.5, color: "#ffffff", bold: true
          }
        ]
      }
    ];
  }

  async function createSampleCacheMemoryPPTX() {
    const zip = new JSZip();
    const slides = getBuiltInCacheMemorySlides();
    zip.file("ppt/ppt_lite_state.json", JSON.stringify(slides));
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`
    );
    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });
  }

  async function packSlidesIntoPPTXBlob(slidesArray) {
    const zip = new JSZip();
    zip.file("ppt/ppt_lite_state.json", JSON.stringify(slidesArray));
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`
    );
    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });
  }

  function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  window.addEventListener("DOMContentLoaded", () => {
    const btnOpenFind = document.getElementById("btn-open-slide-find");
    const btnCloseFind = document.getElementById("btn-find-close");
    const findInput = document.getElementById("slide-find-input");
    const btnPrev = document.getElementById("btn-find-prev");
    const btnNext = document.getElementById("btn-find-next");

    if (btnOpenFind) btnOpenFind.addEventListener("click", openFindBar);
    if (btnCloseFind) btnCloseFind.addEventListener("click", closeFindBar);
    if (findInput) {
      findInput.addEventListener("input", (e) => {
        executeSlideFind(e.target.value);
      });
    }
    if (btnNext) {
      btnNext.addEventListener("click", () => {
        if (findMatches.length === 0) return;
        currentFindMatchIdx = (currentFindMatchIdx + 1) % findMatches.length;
        focusActiveFindMatch();
      });
    }
    if (btnPrev) {
      btnPrev.addEventListener("click", () => {
        if (findMatches.length === 0) return;
        currentFindMatchIdx = (currentFindMatchIdx - 1 + findMatches.length) % findMatches.length;
        focusActiveFindMatch();
      });
    }
  });

  window.PPTXEngine = {
    loadPresentationRecord,
    parsePPTXBlob,
    renderVerticalReaderSlides,
    buildSlideDOMElement,
    createSampleCacheMemoryPPTX,
    packSlidesIntoPPTXBlob,
    getActiveFileRecord: () => activeFileRecord,
    getActiveSlidesData: () => activeSlidesData,
    setActiveSlidesData: (newSlides) => { activeSlidesData = newSlides; },
    getCurrentVisibleSlideIndex: () => currentVisibleSlideIndex,
    setCurrentVisibleSlideIndex: (idx) => { currentVisibleSlideIndex = idx; }
  };
})();
