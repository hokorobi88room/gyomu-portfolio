"use strict";
/* 診断型ファネル solve.js — ライブラリ不使用。
   ページを自動判定: #dxHero=診断ページ / #rxRoot=結果ページ */
(function () {
  var root = document.documentElement;
  var reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* ---- テーマ切替(本体サイトと同じlocalStorageキーを共有) ---- */
  function currentTheme() {
    var t = root.getAttribute("data-theme");
    if (t) return t;
    return matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }
  document.querySelectorAll(".js-theme-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var next = currentTheme() === "dark" ? "light" : "dark";
      root.setAttribute("data-theme", next);
      try { localStorage.setItem("theme", next); } catch (e) {}
    });
  });

  /* ============ 診断ページ ============ */
  var hero = document.getElementById("dxHero");
  if (hero) initDiagnosis();

  function initDiagnosis() {
    var q1 = document.getElementById("dxQ1");
    var q2 = document.getElementById("dxQ2");
    var progress = document.getElementById("dxProgress");
    var dot1 = document.getElementById("dot1");
    var dot2 = document.getElementById("dot2");
    var picked = null;

    function show(stepEl) {
      [hero, q1, q2].forEach(function (s) { s.classList.remove("active"); });
      stepEl.classList.add("active");
      var atHero = stepEl === hero;
      progress.hidden = atHero;
      dot1.classList.toggle("on", stepEl === q1 || stepEl === q2);
      dot2.classList.toggle("on", stepEl === q2);
      // フォーカスを見出しへ(アクセシビリティ)
      var h = stepEl.querySelector("h1,h2");
      if (h) { h.setAttribute("tabindex", "-1"); h.focus({ preventScroll: true }); }
      window.scrollTo({ top: 0, behavior: reduce ? "auto" : "smooth" });
    }

    document.getElementById("dxStart").addEventListener("click", function () { show(q1); });
    document.getElementById("dxBack").addEventListener("click", function () { show(q1); });

    q1.querySelectorAll(".dx-card").forEach(function (card) {
      card.addEventListener("click", function () {
        picked = card.getAttribute("data-p");
        show(q2);
      });
    });

    q2.querySelectorAll(".dx-chip").forEach(function (chip) {
      chip.addEventListener("click", function () {
        if (!picked) { show(q1); return; }
        var h = chip.getAttribute("data-h");
        location.href = "result.html?p=" + encodeURIComponent(picked) + "&h=" + encodeURIComponent(h);
      });
    });
  }

  /* ============ 結果ページ ============ */
  var rxRoot = document.getElementById("rxRoot");
  var gPersona = null, gData = null, gPopupShown = false, gTriggersSet = false, gLastFocus = null;
  if (rxRoot) initResult(rxRoot);

  function qp(name) {
    return new URLSearchParams(location.search).get(name);
  }
  function slot(name) { return rxRoot.querySelector('[data-slot="' + name + '"]'); }
  function esc(s) { return String(s == null ? "" : s); }

  function computeLoss(hRaw, rate) {
    var h = parseFloat(hRaw);
    if (isNaN(h) || h <= 0) h = 5;
    var weekBig = h === 0.5 ? "&lt;1" : (h === 15 ? "15+" : String(h));
    var yearHours = Math.round(h * 52);
    var bizDays = Math.round(yearHours / 8 * 10) / 10;
    var yen = yearHours * rate;
    var man = yen / 10000;
    var yenMan = man >= 10 ? Math.round(man) + "万円" : (Math.round(man * 10) / 10) + "万円";
    var monthsFloat = bizDays / 20;
    var months = monthsFloat >= 1 ? "丸" + Math.floor(monthsFloat) + "か月以上" : Math.round(bizDays) + "営業日ぶん";
    return { weekBig: weekBig, yearHours: yearHours, bizDays: bizDays, yenMan: yenMan, months: months };
  }

  function initResult(rx) {
    var err = document.getElementById("rxError");
    fetch("personas.json").then(function (r) { return r.json(); }).then(function (data) {
      var pid = qp("p");
      var persona = data.personas[pid] || data.personas["other"];
      if (!persona) throw new Error("no persona");
      render(rx, persona, pid, computeLoss(qp("h"), data.hourlyRateAssumption), data);
      rx.hidden = false;
      var h1 = slot("h1"); if (h1) h1.focus({ preventScroll: true });
    }).catch(function () {
      if (err) err.hidden = false;
    });
  }

  function render(rx, persona, pid, loss, data) {
    var res = persona.result;
    slot("kicker").textContent = res.kicker;
    slot("h1").innerHTML = res.h1;
    slot("sub").textContent = res.sub;

    // 共感
    var emp = slot("empathy");
    res.empathy.forEach(function (p) {
      var el = document.createElement("p"); el.textContent = p; emp.appendChild(el);
    });

    // 損失フロー
    var flow = slot("lossflow");
    flow.innerHTML =
      cell(loss.weekBig, "時間/週", "あなたの申告") + arrow() +
      cell(loss.yearHours.toLocaleString("ja-JP"), "時間/年", "年間の合計") + arrow() +
      cell(loss.bizDays, "日", "営業日にすると");
    slot("lossyen").innerHTML = "仮にあなたの時間を時給 " + data.hourlyRateAssumption.toLocaleString("ja-JP") +
      "円とすると、年間 <b>約" + loss.yenMan + "</b>。";
    slot("lossnarr").textContent = res.lossNarrative.replace("{months}", loss.months);
    slot("dignity").textContent = res.dignity;

    // 解決策
    var sol = res.solution;
    slot("claim").innerHTML = hl(sol.claim);
    slot("reason").textContent = sol.reason;
    slot("evidence").textContent = sol.evidence;
    slot("restate").textContent = sol.restate;
    var ifs = slot("ifs");
    sol.ifs.forEach(function (t) {
      var d = document.createElement("div"); d.className = "rx-if"; d.textContent = t.replace(/^もし、?/, "");
      ifs.appendChild(d);
    });
    if (pid === "other") {
      var wl = document.createElement("p"); wl.className = "rx-works-link";
      wl.innerHTML = '近い実例を <a href="../index.html#works">9作品のギャラリー</a> から探せます。';
      ifs.parentNode.appendChild(wl);
    }

    // デモ(Step4で renderDemo が中身を作る。type=none は非表示)
    var demowrap = slot("demowrap");
    if (res.demo && res.demo.type !== "none" && typeof renderDemo === "function") {
      renderDemo(demowrap, res.demo, persona, pid);
    } else {
      demowrap.hidden = true;
    }

    // FAQ(既存 .faq-item を流用)
    var faq = slot("faq");
    res.faq.forEach(function (item) {
      var d = document.createElement("details"); d.className = "faq-item";
      var s = document.createElement("summary"); s.textContent = item.q;
      var a = document.createElement("div"); a.className = "faq-item__a"; a.textContent = item.a;
      d.appendChild(s); d.appendChild(a); faq.appendChild(d);
    });

    // 価格
    slot("price").innerHTML =
      '<span class="rx-price__label">当方の目安</span>' +
      '<span class="rx-price__val">' + esc(res.price) + '</span>' +
      '<span class="rx-price__note">' + esc(res.priceNote) + '</span>';

    // インラインCTA(ポップアップを閉じても残る主導線)
    var cta = slot("inlinecta");
    var html = '<a class="btn btn--primary" href="' + esc(data.coconalaProfile) + '" target="_blank" rel="noopener">あなたの作業を、そのまま送る（無料）</a>';
    if (persona.kit) html += '<a class="btn btn--ghost" href="' + esc(persona.kit) + '" download>体験キットを持ち帰る（Zip）</a>';
    html += '<span class="hint">ご提案・お見積りまで無料です。</span>';
    cta.innerHTML = html;

    // ポップアップ用にコンテキストを保持しトリガーを仕込む
    gPersona = persona; gData = data;
    setupPopupTriggers();
  }

  function cell(n, u, k) {
    return '<div class="rx-loss__cell"><span class="n num-font">' + n + '<span class="u">' + u + '</span></span><span class="k">' + k + '</span></div>';
  }
  function arrow() { return '<span class="rx-loss__arrow num-font">→</span>'; }

  /* ============ 未完デモ(ツァイガルニク装置) ============ */
  function renderDemo(wrap, cfg, persona, pid) {
    wrap.hidden = false;
    var demo = document.createElement("div");
    demo.className = "demo";
    demo.innerHTML =
      '<div class="demo__bar"><span class="demo__dots"><i></i><i></i><i></i></span>' +
      '<span class="demo__title">LIVE DEMO — サンプルデータ（架空）</span></div>' +
      '<div class="demo__btnwrap"><button class="btn btn--primary demo__run" type="button">' + esc(cfg.buttonLabel) + '</button>' +
      '<div class="demo__count" hidden></div></div>' +
      '<div class="demo__stage" hidden></div>' +
      '<div class="demo__lock" hidden></div>';
    wrap.appendChild(demo);

    var btn = demo.querySelector(".demo__run");
    var stage = demo.querySelector(".demo__stage");
    var count = demo.querySelector(".demo__count");
    var lockEl = demo.querySelector(".demo__lock");
    var ran = false;

    btn.addEventListener("click", function () {
      if (ran) return; ran = true;
      btn.disabled = true;
      btn.textContent = cfg.runningLabel || "処理中…";
      stage.hidden = false; count.hidden = false;
      var fn = DEMOS[cfg.type] || DEMOS.table;
      fn(stage, count, cfg, function done() {
        btn.textContent = "体験はここまで";
        lockEl.hidden = false;
        lockEl.innerHTML =
          '<div class="demo__lock-fade"></div>' +
          '<div class="demo__lock-box"><div class="demo__lock-title">' + esc(cfg.lockTitle) + '</div>' +
          '<div class="demo__lock-reason">' + esc(cfg.lockReason) + '</div></div>';
        if (!reduce) lockEl.querySelector(".demo__lock-box").scrollIntoView({ block: "nearest", behavior: "smooth" });
        if (typeof firePopup === "function") firePopup("demo");
      });
    });
  }

  // 各アイテムを step 間隔で流し込む共通ドライバ
  function stream(items, stopAt, addOne, count, label, total, onDone) {
    var i = 0, n = Math.min(stopAt, items.length);
    function tick() {
      if (i >= n) {
        count.innerHTML = label(stopAt) + ' <b>残り ' + (total - stopAt).toLocaleString("ja-JP") + '</b>';
        onDone(); return;
      }
      addOne(items[i], i);
      i++;
      count.innerHTML = label(i);
      if (reduce) tick(); else setTimeout(tick, 260);
    }
    if (reduce) { while (i < n) { addOne(items[i], i); i++; } count.innerHTML = label(n) + ' <b>残り ' + (total - stopAt).toLocaleString("ja-JP") + '</b>'; onDone(); }
    else tick();
  }

  var DEMOS = {
    table: function (stage, count, cfg, onDone) {
      var t = document.createElement("table"); t.className = "demo-table";
      t.innerHTML = "<thead><tr>" + cfg.columns.map(function (c) { return "<th>" + esc(c) + "</th>"; }).join("") + "</tr></thead><tbody></tbody>";
      stage.appendChild(t);
      var tb = t.querySelector("tbody");
      stream(cfg.rows, cfg.stopAt, function (row) {
        var tr = document.createElement("tr"); tr.className = "demo-row-in";
        tr.innerHTML = row.map(function (v) { return "<td>" + esc(v) + "</td>"; }).join("");
        tb.appendChild(tr);
      }, count, function (n) { return "収集 " + n + " 件"; }, cfg.total, function () {
        // ぼかしのプレビュー行(残りの雰囲気)
        var extra = cfg.rows.slice(cfg.stopAt, cfg.stopAt + 2);
        extra.forEach(function (row) {
          var tr = document.createElement("tr"); tr.className = "demo-blur";
          tr.innerHTML = row.map(function (v) { return "<td>" + esc(v) + "</td>"; }).join("");
          tb.appendChild(tr);
        });
        onDone();
      });
    },
    invoice: function (stage, count, cfg, onDone) {
      var g = document.createElement("div"); g.className = "demo-invoices"; stage.appendChild(g);
      stream(cfg.cards, cfg.stopAt, function (c) {
        var d = document.createElement("div"); d.className = "demo-inv demo-row-in";
        d.innerHTML = '<div class="demo-inv__no">' + esc(c.no) + '</div><div class="demo-inv__to">' + esc(c.to) + '</div><div class="demo-inv__amt num-font">' + esc(c.amount) + '</div>';
        g.appendChild(d);
      }, count, function (n) { return "作成 " + n + " 社"; }, cfg.total, onDone);
    },
    reception: function (stage, count, cfg, onDone) {
      var box = document.createElement("div"); box.className = "demo-recep"; stage.appendChild(box);
      stream(cfg.entries, cfg.stopAt, function (e) {
        var d = document.createElement("div"); d.className = "demo-rec demo-row-in";
        var steps = e.steps.map(function (s) {
          var stop = /停止|到達/.test(s);
          return '<span' + (stop ? ' class="stop"' : '') + '>' + esc(s) + '</span>';
        }).join("");
        d.innerHTML = '<div class="demo-rec__head"><b>' + esc(e.name) + '</b><span class="t">' + esc(e.when) + '</span></div><div class="demo-rec__steps">' + steps + '</div>';
        box.appendChild(d);
      }, count, function (n) { return "自動処理 " + n + " 件"; }, cfg.total, onDone);
    },
    grid: function (stage, count, cfg, onDone) {
      var g = document.createElement("div"); g.className = "demo-grid"; stage.appendChild(g);
      var ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>';
      stream(cfg.tiles, cfg.stopAt, function (t) {
        var d = document.createElement("div"); d.className = "demo-tile";
        d.innerHTML = '<div class="demo-tile__badge">' + ICON + '</div><div class="demo-tile__body"><div class="demo-tile__name">' + esc(t.name) + '</div><div class="demo-tile__role">' + esc(t.role) + '</div></div>';
        g.appendChild(d);
      }, count, function (n) { return "生成 " + n + " 枚"; }, cfg.total, onDone);
    },
    transcribe: function (stage, count, cfg, onDone) {
      var box = document.createElement("div"); box.className = "demo-trans"; stage.appendChild(box);
      var lines = cfg.lines, li = 0;
      count.innerHTML = "変換中… 00:00";
      function nextLine() {
        if (li >= lines.length) {
          count.innerHTML = 'サンプル冒頭（約30%）まで <b>残りは録音の続き</b>';
          onDone(); return;
        }
        var ln = lines[li];
        var row = document.createElement("div"); row.className = "demo-tline";
        row.innerHTML = '<span class="meta">' + esc(ln.t) + '</span><span class="body"><span class="who">' + esc(ln.who) + '</span><span class="txt"></span><span class="demo-caret"></span></span>';
        box.appendChild(row);
        var txt = row.querySelector(".txt"), caret = row.querySelector(".demo-caret");
        var full = ln.text, ci = 0;
        count.innerHTML = "変換中… " + esc(ln.t);
        if (reduce) { txt.textContent = full; caret.remove(); li++; nextLine(); return; }
        (function type() {
          if (ci < full.length) { ci = Math.min(ci + 2, full.length); txt.textContent = full.slice(0, ci); setTimeout(type, 34); }
          else { caret.remove(); li++; setTimeout(nextLine, 170); }
        })();
      }
      nextLine();
    }
  };
  function hl(s) { // 数字/秒/「◯」を強調
    return esc(s).replace(/(10秒|数秒|100枚|ボタン1回|24時間)/g, '<span class="hl">$1</span>');
  }

  /* ============ 依頼ポップアップ(第5段) ============ */
  function setupPopupTriggers() {
    if (gTriggersSet) return; gTriggersSet = true;

    // (b) 滞在90秒 かつ スクロール70%到達
    var deep = false, dwelled = false;
    function maybe() { if (deep && dwelled) firePopup("dwell"); }
    window.addEventListener("scroll", function () {
      var sc = document.documentElement;
      var pct = (window.scrollY + window.innerHeight) / sc.scrollHeight;
      if (pct >= 0.7) { deep = true; maybe(); }
    }, { passive: true });
    setTimeout(function () { dwelled = true; maybe(); }, 90000);

    // (c) exit intent(デスクトップのみ)
    if (matchMedia("(pointer:fine)").matches) {
      document.addEventListener("mouseout", function (e) {
        if (!e.relatedTarget && e.clientY <= 0) firePopup("exit");
      });
    }
  }

  function firePopup(source) {
    if (gPopupShown) return;
    try { if (sessionStorage.getItem("solve_popup_shown")) return; } catch (e) {}
    if (!gPersona || !gData) return;
    gPopupShown = true;
    try { sessionStorage.setItem("solve_popup_shown", "1"); } catch (e) {}
    buildPopup(gPersona, gData);
    openPopup();
  }

  function buildPopup(persona, data) {
    var box = document.getElementById("popupContent");
    var spoken = persona.spoken || "その作業";
    var sub = persona.kit
      ? '<a href="' + esc(persona.kit) + '" download id="popupKit">まずは体験キットを持ち帰る（Zip・登録不要）</a>'
      : '<a href="../index.html#works">まずは近い実例（9作品）を見る</a>';
    box.innerHTML =
      '<div class="popup__kicker">ここまでは、サンプルの話。</div>' +
      '<div class="popup__body" id="popupTitle">' +
      '<p class="lead">ここまで試したということは、一度は想像したはずです。<br>「これが、自分のあの作業で動いたら」と。</p>' +
      '<p>その想像は、だいたい実現できます。あなたが選んだ「' + esc(spoken) + '」は、私がいちばん数をこなしてきた種類の仕事だからです。</p>' +
      '<p>やり方は、いまの作業内容をそのまま送るだけ。翻訳するのは私の仕事です。ご提案と概算まで無料です。副業のため同時にお受けできる件数に限りがあるので、迷っているなら、相談だけ先に置いておいてください。</p>' +
      '<p>決めるのは、提案と金額を見てからで大丈夫です。</p>' +
      '</div>' +
      '<div class="popup__buttons">' +
      '<a class="btn btn--primary" id="popupGo" href="' + esc(data.coconalaProfile) + '" target="_blank" rel="noopener">あなたの作業を、そのまま送る（無料）</a>' +
      '</div>' +
      '<div class="popup__sub">' + sub + '</div>';

    var go = document.getElementById("popupGo");
    if (go) go.addEventListener("click", function () { showThanks(); });
    var kit = document.getElementById("popupKit");
    if (kit) kit.addEventListener("click", function () { showToast("READMEの1ページ目から読んでください。5分で動きます。"); });
  }

  function showThanks() {
    var box = document.getElementById("popupContent");
    box.innerHTML =
      '<div class="popup__thanks">' +
      '<h3>ありがとうございます。</h3>' +
      '<p>ココナラが開きます。メッセージ1通目は「診断結果のURL」を貼るだけでも通じます。お返事は24時間以内を心がけています。楽しみにしていてください。</p>' +
      '<button class="btn btn--ghost" data-close type="button">閉じる</button></div>';
    box.querySelector("[data-close]").addEventListener("click", closePopup);
  }

  var trapHandler = null;
  function openPopup() {
    var popup = document.getElementById("popup");
    gLastFocus = document.activeElement;
    popup.classList.add("open");
    document.body.style.overflow = "hidden";
    popup.querySelectorAll("[data-close]").forEach(function (b) { b.addEventListener("click", closePopup); });
    // フォーカスを最初のボタンへ
    var focusables = popup.querySelectorAll("button, a[href]");
    if (focusables.length) focusables[Math.min(1, focusables.length - 1)].focus();
    // ESC + フォーカストラップ
    trapHandler = function (e) {
      if (e.key === "Escape") { closePopup(); return; }
      if (e.key !== "Tab") return;
      var f = popup.querySelectorAll("button, a[href]");
      if (!f.length) return;
      var first = f[0], last = f[f.length - 1];
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", trapHandler);
  }

  function closePopup() {
    var popup = document.getElementById("popup");
    popup.classList.remove("open");
    document.body.style.overflow = "";
    if (trapHandler) { document.removeEventListener("keydown", trapHandler); trapHandler = null; }
    if (gLastFocus && gLastFocus.focus) gLastFocus.focus();
  }

  function showToast(msg) {
    var t = document.getElementById("toast");
    if (!t) return;
    t.textContent = msg; t.classList.add("show");
    setTimeout(function () { t.classList.remove("show"); }, 4200);
  }
})();
