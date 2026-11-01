/* ==========================================================================
   POWERPOINT LITE — 60FPS PINCH-ZOOM & HORIZONTAL FULLSCREEN PLAYER (www/gestures-player.js)
   ========================================================================== */

(function () {
  "use strict";

  let readerScale = 1;
  let readerPanX = 0;
  let readerStartDist = 0;
  let readerInitialScale = 1;
  let isReaderPinching = false;
  let readerTouchStartX = 0;
  let readerInitialPanX = 0;

  function initVerticalReaderZoom() {
    const viewport = document.getElementById("reader-scroll-viewport");
    const slidesList = document.getElementById("reader-slides-list");
    if (!viewport || !slidesList) return;

    viewport.addEventListener(
      "touchstart",
      (e) => {
        if (e.touches.length === 2) {
          isReaderPinching = true;
          readerStartDist = getTouchDistance(e.touches[0], e.touches[1]);
          readerInitialScale = readerScale;
        } else if (e.touches.length === 1 && readerScale > 1.02) {
          readerTouchStartX = e.touches[0].clientX;
          readerInitialPanX = readerPanX;
        }
      },
      { passive: true }
    );

    viewport.addEventListener(
      "touchmove",
      (e) => {
        if (isReaderPinching && e.touches.length === 2) {
          e.preventDefault();
          const currentDist = getTouchDistance(e.touches[0], e.touches[1]);
          if (readerStartDist > 0) {
            const factor = currentDist / readerStartDist;
            readerScale = Math.max(1, Math.min(3.5, readerInitialScale * factor));
            if (readerScale <= 1.02) {
              readerPanX = 0;
            }
            applyReaderTransform(slidesList);
          }
        } else if (!isReaderPinching && e.touches.length === 1 && readerScale > 1.02) {
          const dx = e.touches[0].clientX - readerTouchStartX;
          const maxPan = ((readerScale - 1) * viewport.clientWidth) / 2;
          readerPanX = Math.max(-maxPan, Math.min(maxPan, readerInitialPanX + dx));
          applyReaderTransform(slidesList);
        }
      },
      { passive: false }
    );

    viewport.addEventListener("touchend", (e) => {
      if (e.touches.length < 2) {
        isReaderPinching = false;
        if (readerScale < 1.04) {
          readerScale = 1;
          readerPanX = 0;
          applyReaderTransform(slidesList);
        }
      }
    });
  }

  function applyReaderTransform(slidesList) {
    slidesList.style.transform = `translate3d(${readerPanX}px, 0, 0) scale(${readerScale})`;
  }

  function resetReaderZoom() {
    readerScale = 1;
    readerPanX = 0;
    const slidesList = document.getElementById("reader-slides-list");
    if (slidesList) {
      slidesList.style.transform = "translate3d(0, 0, 0) scale(1)";
    }
  }

  let fsCurrentIndex = 0;
  let fsScale = 1;
  let fsPanX = 0;
  let fsPanY = 0;
  let fsPinchStartDist = 0;
  let fsInitialScale = 1;
  let fsIsPinching = false;
  let fsTouchStartX = 0;
  let fsTouchStartY = 0;
  let fsInitialPanX = 0;
  let fsInitialPanY = 0;
  let fsHudTimer = null;

  async function launchFullscreenPlayer(startIndex) {
    const slides = window.PPTXEngine.getActiveSlidesData();
    if (!slides || slides.length === 0) return;

    fsCurrentIndex =
      typeof startIndex === "number"
        ? startIndex
        : window.PPTXEngine.getCurrentVisibleSlideIndex() || 0;

    window.AppStorage.switchScreen("screen-fullscreen-player");
    const playerScreen = document.getElementById("screen-fullscreen-player");

    try {
      if (document.documentElement.requestFullscreen) {
        await document.documentElement.requestFullscreen();
      }
    } catch (err) {}

    let lockedNativeLandscape = false;
    try {
      if (screen.orientation && screen.orientation.lock) {
        await screen.orientation.lock("landscape");
        lockedNativeLandscape = true;
      }
    } catch (err) {
      lockedNativeLandscape = false;
    }

    if (!lockedNativeLandscape && window.innerHeight > window.innerWidth) {
      playerScreen.classList.add("force-landscape-css");
    } else {
      playerScreen.classList.remove("force-landscape-css");
    }

    renderFullscreenSlide(fsCurrentIndex);
    showFullscreenHudTemporarily();
  }

  async function exitFullscreenPlayer() {
    const playerScreen = document.getElementById("screen-fullscreen-player");
    if (playerScreen) {
      playerScreen.classList.remove("force-landscape-css");
    }

    try {
      if (screen.orientation && screen.orientation.unlock) {
        screen.orientation.unlock();
      }
    } catch (err) {}

    try {
      if (document.fullscreenElement && document.exitFullscreen) {
        await document.exitFullscreen();
      }
    } catch (err) {}

    window.AppStorage.switchScreen("screen-reader");
  }

  function renderFullscreenSlide(index, slideDirection = 0) {
    const slides = window.PPTXEngine.getActiveSlidesData();
    if (!slides || slides.length === 0) return;

    fsCurrentIndex = Math.max(0, Math.min(slides.length - 1, index));
    window.PPTXEngine.setCurrentVisibleSlideIndex(fsCurrentIndex);

    fsScale = 1;
    fsPanX = 0;
    fsPanY = 0;

    const wrapper = document.getElementById("fs-slide-wrapper");
    const indicator = document.getElementById("fs-slide-indicator");
    if (!wrapper) return;

    wrapper.innerHTML = "";
    const slideCard = window.PPTXEngine.buildSlideDOMElement(slides[fsCurrentIndex], false);
    wrapper.appendChild(slideCard);

    if (indicator) {
      indicator.textContent = `${fsCurrentIndex + 1} / ${slides.length}`;
    }

    if (slideDirection !== 0) {
      const offsetPx = slideDirection > 0 ? 55 : -55;
      wrapper.style.transition = "none";
      wrapper.style.opacity = "0.5";
      wrapper.style.transform = getFullscreenTransformString(offsetPx, 0, 1);

      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          wrapper.style.transition = "transform 0.22s cubic-bezier(0.2, 0.9, 0.3, 1), opacity 0.22s ease";
          wrapper.style.opacity = "1";
          wrapper.style.transform = getFullscreenTransformString(0, 0, 1);
        });
      });
    } else {
      wrapper.style.transition = "none";
      wrapper.style.opacity = "1";
      wrapper.style.transform = getFullscreenTransformString(0, 0, 1);
    }
  }

  function getFullscreenTransformString(px, py, scale) {
    const playerScreen = document.getElementById("screen-fullscreen-player");
    const isCssRotated =
      playerScreen &&
      playerScreen.classList.contains("force-landscape-css") &&
      window.innerHeight > window.innerWidth;

    if (isCssRotated) {
      return `rotate(90deg) translate3d(${px}px, ${py}px, 0) scale(${scale})`;
    }
    return `translate3d(${px}px, ${py}px, 0) scale(${scale})`;
  }

  function applyFullscreenTransform() {
    const wrapper = document.getElementById("fs-slide-wrapper");
    if (!wrapper) return;
    wrapper.style.transition = "none";
    wrapper.style.transform = getFullscreenTransformString(fsPanX, fsPanY, fsScale);
  }

  function initFullscreenTouchGestures() {
    const stage = document.getElementById("fs-touch-stage");
    if (!stage) return;

    let lastTapTime = 0;

    stage.addEventListener(
      "touchstart",
      (e) => {
        showFullscreenHudTemporarily();

        if (e.touches.length === 2) {
          fsIsPinching = true;
          fsPinchStartDist = getTouchDistance(e.touches[0], e.touches[1]);
          fsInitialScale = fsScale;
        } else if (e.touches.length === 1) {
          fsIsPinching = false;
          fsTouchStartX = e.touches[0].clientX;
          fsTouchStartY = e.touches[0].clientY;
          fsInitialPanX = fsPanX;
          fsInitialPanY = fsPanY;
        }
      },
      { passive: true }
    );

    stage.addEventListener(
      "touchmove",
      (e) => {
        e.preventDefault();

        if (fsIsPinching && e.touches.length === 2) {
          const dist = getTouchDistance(e.touches[0], e.touches[1]);
          if (fsPinchStartDist > 0) {
            const factor = dist / fsPinchStartDist;
            fsScale = Math.max(1, Math.min(4.0, fsInitialScale * factor));
            if (fsScale <= 1.03) {
              fsPanX = 0;
              fsPanY = 0;
            }
            applyFullscreenTransform();
          }
        } else if (!fsIsPinching && e.touches.length === 1 && fsScale > 1.05) {
          const rawDx = e.touches[0].clientX - fsTouchStartX;
          const rawDy = e.touches[0].clientY - fsTouchStartY;

          const playerScreen = document.getElementById("screen-fullscreen-player");
          const isCssRotated =
            playerScreen &&
            playerScreen.classList.contains("force-landscape-css") &&
            window.innerHeight > window.innerWidth;

          const dx = isCssRotated ? rawDy : rawDx;
          const dy = isCssRotated ? -rawDx : rawDy;

          const maxPanX = ((fsScale - 1) * window.innerWidth) * 0.45;
          const maxPanY = ((fsScale - 1) * window.innerHeight) * 0.45;

          fsPanX = Math.max(-maxPanX, Math.min(maxPanX, fsInitialPanX + dx));
          fsPanY = Math.max(-maxPanY, Math.min(maxPanY, fsInitialPanY + dy));
          applyFullscreenTransform();
        }
      },
      { passive: false }
    );

    stage.addEventListener("touchend", (e) => {
      const now = Date.now();
      if (e.changedTouches.length === 1 && !fsIsPinching) {
        const moveDist = Math.hypot(
          e.changedTouches[0].clientX - fsTouchStartX,
          e.changedTouches[0].clientY - fsTouchStartY
        );
        if (moveDist < 12 && now - lastTapTime < 260) {
          if (fsScale > 1.05) {
            fsScale = 1;
            fsPanX = 0;
            fsPanY = 0;
          } else {
            fsScale = 2.0;
          }
          applyFullscreenTransform();
          lastTapTime = 0;
          return;
        }
        if (moveDist < 12) {
          lastTapTime = now;
        }
      }

      if (!fsIsPinching && fsScale <= 1.05 && e.changedTouches.length === 1) {
        const rawDx = e.changedTouches[0].clientX - fsTouchStartX;
        const rawDy = e.changedTouches[0].clientY - fsTouchStartY;

        const playerScreen = document.getElementById("screen-fullscreen-player");
        const isCssRotated =
          playerScreen &&
          playerScreen.classList.contains("force-landscape-css") &&
          window.innerHeight > window.innerWidth;

        const primaryDelta = isCssRotated ? rawDy : rawDx;
        const secondaryDelta = isCssRotated ? rawDx : rawDy;

        if (Math.abs(primaryDelta) > 48 && Math.abs(primaryDelta) > Math.abs(secondaryDelta) * 1.2) {
          const slides = window.PPTXEngine.getActiveSlidesData();
          if (primaryDelta < 0 && fsCurrentIndex < slides.length - 1) {
            renderFullscreenSlide(fsCurrentIndex + 1, 1);
          } else if (primaryDelta > 0 && fsCurrentIndex > 0) {
            renderFullscreenSlide(fsCurrentIndex - 1, -1);
          }
        }
      }

      if (e.touches.length === 0) {
        fsIsPinching = false;
        if (fsScale < 1.05) {
          fsScale = 1;
          fsPanX = 0;
          fsPanY = 0;
          applyFullscreenTransform();
        }
      }
    });
  }

  function showFullscreenHudTemporarily() {
    const hud = document.getElementById("fs-hud-overlay");
    if (!hud) return;
    hud.classList.remove("hud-hidden");
    clearTimeout(fsHudTimer);
    fsHudTimer = setTimeout(() => {
      hud.classList.add("hud-hidden");
    }, 2200);
  }

  function getTouchDistance(t1, t2) {
    const dx = t1.clientX - t2.clientX;
    const dy = t1.clientY - t2.clientY;
    return Math.hypot(dx, dy);
  }

  window.addEventListener("DOMContentLoaded", () => {
    initVerticalReaderZoom();
    initFullscreenTouchGestures();

    const btnPlayReader = document.getElementById("btn-launch-fullscreen-play");
    if (btnPlayReader) {
      btnPlayReader.addEventListener("click", () => {
        resetReaderZoom();
        launchFullscreenPlayer(window.PPTXEngine.getCurrentVisibleSlideIndex());
      });
    }

    const btnExitFs = document.getElementById("btn-exit-fullscreen");
    if (btnExitFs) {
      btnExitFs.addEventListener("click", exitFullscreenPlayer);
    }

    window.addEventListener("keydown", (e) => {
      const playerScreen = document.getElementById("screen-fullscreen-player");
      if (!playerScreen || !playerScreen.classList.contains("active")) return;

      if (e.key === "Escape") {
        exitFullscreenPlayer();
      } else if (e.key === "ArrowRight") {
        renderFullscreenSlide(fsCurrentIndex + 1, 1);
      } else if (e.key === "ArrowLeft") {
        renderFullscreenSlide(fsCurrentIndex - 1, -1);
      }
    });
  });

  window.GesturesPlayer = {
    launchFullscreenPlayer,
    exitFullscreenPlayer,
    resetReaderZoom
  };
})();
