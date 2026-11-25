(function () {
  "use strict";

  var LABELS = { 1: "Poor", 2: "Fair", 3: "Average", 4: "Good", 5: "Excellent" };
  var smallScreen = window.matchMedia("(max-width: 640px)");
  var finePointer = window.matchMedia("(pointer: fine)");

  // Phones keep pages in memory (back/forward cache). Coming back with Back would show
  // stale buttons ("Please wait…") or an old timer, so reload to get the current state.
  window.addEventListener("pageshow", function (e) {
    if (e.persisted) window.location.reload();
  });

  // Session countdown: based on the clock, because phones pause timers when the screen locks
  var timer = document.getElementById("session-timer");
  if (timer) {
    var endAt = Date.now() + (parseInt(timer.getAttribute("data-seconds"), 10) || 1800) * 1000;
    var tick = function () {
      var left = Math.max(0, Math.round((endAt - Date.now()) / 1000));
      var m = Math.floor(left / 60), s = left % 60;
      timer.textContent = (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
      if (left <= 0) { window.location.href = "/"; }
    };
    tick();
    setInterval(tick, 1000);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) tick(); });
  }

  // Autofocus only with a mouse; on phones it would pop the keyboard and jump the page
  var af = document.querySelector("[data-autofocus]");
  if (af && finePointer.matches) af.focus();

  // Show / hide password (admin login)
  document.querySelectorAll(".pw-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var inp = document.getElementById(btn.getAttribute("data-target"));
      var show = inp.type === "password";
      inp.type = show ? "text" : "password";
      btn.textContent = show ? "Hide" : "Show";
    });
  });

  // Rating form: answered counter, progress bar, highlighting
  var form = document.getElementById("fbform");
  if (form) {
    var counter = document.getElementById("answered");
    var bar = document.getElementById("answered-bar");
    var total = parseInt(counter.getAttribute("data-total"), 10);
    var items = form.querySelectorAll("[data-q]");
    var refresh = function () {
      var n = 0;
      items.forEach(function (it) {
        var picked = it.querySelector("input:checked");
        it.classList.toggle("answered", !!picked);
        var label = it.querySelector(".q-picked");
        if (label) label.textContent = picked ? "Selected: " + picked.value + " – " + LABELS[picked.value] : "";
        if (picked) { it.classList.remove("miss"); n += 1; }
      });
      counter.textContent = n + " of " + total + " answered";
      if (bar) bar.style.width = (100 * n / total) + "%";
    };
    form.addEventListener("change", refresh);
    refresh();

    var clear = form.querySelector(".js-clear");
    if (clear) clear.addEventListener("click", function () {
      form.querySelectorAll("input[type=radio]").forEach(function (r) { r.checked = false; });
      refresh();
    });

    form.addEventListener("submit", function (e) {
      var first = null;
      items.forEach(function (it) {
        if (!it.querySelector("input:checked")) { it.classList.add("miss"); first = first || it; }
      });
      if (first) {
        e.preventDefault();
        e.stopImmediatePropagation();
        first.scrollIntoView({ behavior: "smooth", block: "center" });
        first.classList.add("shake");
        setTimeout(function () { first.classList.remove("shake"); }, 500);
      }
    });

    var ta = document.getElementById("comments"), cc = document.getElementById("cc");
    if (ta && cc) ta.addEventListener("input", function () { cc.textContent = ta.value.length; });

    // On phones, hide the bottom bar while the keyboard is open so it does not cover the text box
    form.addEventListener("focusin", function (e) {
      if (smallScreen.matches && e.target.matches("textarea, input[type=text]")) document.body.classList.add("typing");
    });
    form.addEventListener("focusout", function () { document.body.classList.remove("typing"); });
  }

  // Required-field check for login forms
  document.querySelectorAll("form.js-validate:not(#fbform)").forEach(function (f) {
    f.addEventListener("submit", function (e) {
      var bad = null;
      f.querySelectorAll("[required]").forEach(function (el) {
        var empty = el.type === "radio"
          ? !f.querySelector("input[name='" + el.name + "']:checked")
          : !el.value.trim();
        el.classList.toggle("field-err", empty);
        if (empty && !bad) bad = el;
      });
      if (bad) { e.preventDefault(); e.stopImmediatePropagation(); bad.focus(); }
    });
    f.addEventListener("input", function (e) { e.target.classList.remove("field-err"); });
  });

  // Final submit: declaration + confirmation
  document.querySelectorAll("form.js-final").forEach(function (f) {
    var box = f.querySelector("input[name=declaration]");
    var msg = f.querySelector(".declare-err");
    box.addEventListener("change", function () { if (msg && box.checked) msg.hidden = true; });
    f.addEventListener("submit", function (e) {
      if (!box.checked) {
        e.preventDefault(); e.stopImmediatePropagation();
        var card = box.closest(".card");
        if (card) {
          card.scrollIntoView({ behavior: "smooth", block: "center" });
          card.classList.add("shake"); setTimeout(function () { card.classList.remove("shake"); }, 500);
        }
        if (msg) msg.hidden = false;
        return;
      }
      if (!window.confirm("Submit your feedback now? It cannot be changed after submission.")) {
        e.preventDefault(); e.stopImmediatePropagation();
      }
    });
  });

  document.querySelectorAll("form.js-confirm").forEach(function (f) {
    f.addEventListener("submit", function (e) {
      if (!window.confirm(f.getAttribute("data-msg"))) { e.preventDefault(); e.stopImmediatePropagation(); }
    });
  });

  // Upload: show chosen file name
  var file = document.getElementById("datafile");
  if (file) file.addEventListener("change", function () {
    var name = document.getElementById("dz-name");
    if (name && file.files.length) {
      name.textContent = file.files[0].name;
      file.closest(".dropzone").classList.add("has-file");
    }
  });

  // Prevent double submission (registered last so validation handlers run first)
  document.querySelectorAll("form.js-once").forEach(function (f) {
    f.addEventListener("submit", function () {
      f.querySelectorAll("button[type=submit]").forEach(function (b) {
        b.disabled = true;
        b.textContent = f.classList.contains("js-upload") ? "Importing… please wait" : "Please wait…";
      });
    });
  });
})();
