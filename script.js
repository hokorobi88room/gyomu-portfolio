"use strict";
/* 綻 流夢 ポートフォリオ — 素のJS(ライブラリ不使用)。
   テーマ切替 / ヘッダー影 / スクロールreveal / 数値カウントアップ / 作品フィルタ */
(function () {
  var root = document.documentElement;
  var reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* ---- テーマ切替(手動トグル。既定はOSに追従) ---- */
  function currentTheme() {
    var t = root.getAttribute("data-theme");
    if (t) return t;
    return matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }
  document.querySelectorAll("#themeToggle, .js-theme-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var next = currentTheme() === "dark" ? "light" : "dark";
      root.setAttribute("data-theme", next);
      try { localStorage.setItem("theme", next); } catch (e) {}
    });
  });

  /* ---- ヘッダー: スクロールで境界線 ---- */
  var header = document.getElementById("siteHeader");
  if (header) {
    var onScroll = function () {
      header.classList.toggle("is-scrolled", window.scrollY > 8);
    };
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });
  }

  /* ---- 数値カウントアップ(視界に入ったら1回だけ) ---- */
  function animateCount(el) {
    var target = parseFloat(el.getAttribute("data-count"));
    var dec = parseInt(el.getAttribute("data-decimals") || "0", 10);
    if (isNaN(target)) return;
    if (reduce) { el.textContent = target.toFixed(dec); return; }
    var dur = 850, start = null;
    function fmt(v) {
      return dec > 0 ? v.toFixed(dec) : Math.round(v).toLocaleString("ja-JP");
    }
    function step(ts) {
      if (start === null) start = ts;
      var p = Math.min((ts - start) / dur, 1);
      var eased = 1 - Math.pow(1 - p, 3); // easeOutCubic
      el.textContent = fmt(target * eased);
      if (p < 1) requestAnimationFrame(step);
      else el.textContent = fmt(target);
    }
    requestAnimationFrame(step);
  }

  /* ---- スクロールreveal + カウント起動 ---- */
  var reveals = document.querySelectorAll(".reveal");
  var counters = document.querySelectorAll("[data-count]");
  if (reduce || !("IntersectionObserver" in window)) {
    reveals.forEach(function (el) { el.classList.add("is-in"); });
    counters.forEach(function (el) {
      var t = parseFloat(el.getAttribute("data-count"));
      var d = parseInt(el.getAttribute("data-decimals") || "0", 10);
      if (!isNaN(t)) el.textContent = d > 0 ? t.toFixed(d) : t.toLocaleString("ja-JP");
    });
  } else {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (!e.isIntersecting) return;
        e.target.classList.add("is-in");
        e.target.querySelectorAll && e.target.querySelectorAll("[data-count]").forEach(animateCount);
        io.unobserve(e.target);
      });
    }, { threshold: 0.14, rootMargin: "0px 0px -8% 0px" });
    reveals.forEach(function (el) { io.observe(el); });
    // reveal外に置かれたカウンタも拾う
    counters.forEach(function (el) {
      if (!el.closest(".reveal")) {
        var solo = new IntersectionObserver(function (es) {
          es.forEach(function (e) { if (e.isIntersecting) { animateCount(e.target); solo.unobserve(e.target); } });
        }, { threshold: 0.6 });
        solo.observe(el);
      }
    });
  }

  /* ---- 作品フィルタ(タブ) ---- */
  var tabs = document.querySelectorAll("[data-filter]");
  var cards = document.querySelectorAll("[data-tags]");
  if (tabs.length && cards.length) {
    function apply(f) {
      cards.forEach(function (c) {
        var show = f === "all" || (" " + c.getAttribute("data-tags") + " ").indexOf(" " + f + " ") > -1;
        c.hidden = !show;
      });
    }
    tabs.forEach(function (tab) {
      tab.addEventListener("click", function () {
        tabs.forEach(function (t) { t.classList.remove("is-active"); t.setAttribute("aria-selected", "false"); });
        tab.classList.add("is-active");
        tab.setAttribute("aria-selected", "true");
        apply(tab.getAttribute("data-filter"));
      });
    });
  }
})();
