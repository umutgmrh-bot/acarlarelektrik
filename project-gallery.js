(() => {
  const AUTO_MS = 3200;

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function initGallery(gallery) {
    if (!gallery || gallery.classList.contains("pg-slider")) return;

    const slides = Array.from(gallery.querySelectorAll(".project-detail-photo"));
    if (!slides.length) return;

    gallery.classList.add("pg-slider");

    const viewport = document.createElement("div");
    viewport.className = "pg-viewport";

    const track = document.createElement("div");
    track.className = "pg-track";

    // Precompute blurred background per slide (uses the same image)
    slides.forEach((fig) => {
      const img = fig.querySelector("img");
      const src = img && img.getAttribute("src") ? String(img.getAttribute("src")) : "";
      if (src) fig.style.setProperty("--pg-slide-bg", `url("${src.replace(/"/g, '\\"')}")`);
    });

    // Move slides into the track
    slides.forEach((fig) => {
      track.appendChild(fig);
    });

    viewport.appendChild(track);
    gallery.appendChild(viewport);

    let idx = 0;
    let timer = null;
    let paused = false;

    const prefersReducedMotion =
      window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

    const hasMultiple = slides.length > 1;

    const btnPrev = document.createElement("button");
    btnPrev.type = "button";
    btnPrev.className = "pg-btn pg-btn-prev";
    btnPrev.setAttribute("aria-label", "Önceki");
    btnPrev.innerHTML = "‹";

    const btnNext = document.createElement("button");
    btnNext.type = "button";
    btnNext.className = "pg-btn pg-btn-next";
    btnNext.setAttribute("aria-label", "Sonraki");
    btnNext.innerHTML = "›";

    const dotsWrap = document.createElement("div");
    dotsWrap.className = "pg-dots";

    const dots = slides.map((_, i) => {
      const dot = document.createElement("button");
      dot.type = "button";
      dot.className = "pg-dot";
      dot.setAttribute("aria-label", `Görsel ${i + 1}`);
      dot.setAttribute("aria-current", i === 0 ? "true" : "false");
      dotsWrap.appendChild(dot);
      dot.addEventListener("click", () => go(i));
      return dot;
    });

    function setDots(active) {
      dots.forEach((d, i) => d.setAttribute("aria-current", i === active ? "true" : "false"));
    }

    function go(nextIndex) {
      idx = clamp(nextIndex, 0, slides.length - 1);
      track.style.transform = `translateX(${-idx * 100}%)`;
      setDots(idx);
    }

    function next() {
      go((idx + 1) % slides.length);
    }

    function prev() {
      go((idx - 1 + slides.length) % slides.length);
    }

    function stop() {
      if (timer) window.clearInterval(timer);
      timer = null;
    }

    function start() {
      if (!hasMultiple || prefersReducedMotion) return;
      stop();
      timer = window.setInterval(() => {
        if (!paused) next();
      }, AUTO_MS);
    }

    function pause() {
      paused = true;
    }

    function resume() {
      paused = false;
    }

    if (hasMultiple) {
      gallery.appendChild(btnPrev);
      gallery.appendChild(btnNext);
      gallery.appendChild(dotsWrap);

      btnPrev.addEventListener("click", () => {
        prev();
      });
      btnNext.addEventListener("click", () => {
        next();
      });

      gallery.addEventListener("mouseenter", pause, { passive: true });
      gallery.addEventListener("mouseleave", resume, { passive: true });
      gallery.addEventListener("focusin", pause);
      gallery.addEventListener("focusout", resume);

      window.addEventListener(
        "keydown",
        (e) => {
          // Only when the slider is near viewport focus (avoid hijacking globally too much)
          const active = document.activeElement;
          if (!gallery.contains(active)) return;
          if (e.key === "ArrowLeft") prev();
          if (e.key === "ArrowRight") next();
        },
        { passive: true }
      );
    }

    // Initial render and autoplay
    go(0);
    start();
  }

  function initAll() {
    document.querySelectorAll(".project-detail-gallery").forEach(initGallery);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initAll, { once: true });
  } else {
    initAll();
  }
})();

