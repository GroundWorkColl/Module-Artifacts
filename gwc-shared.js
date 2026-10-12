/* GroundWork Collaborative shared tools script.
   Adds a "Save my work" button to each tool:
   - Download as Word
   - Print or save as PDF
   - Back up and restore saved data (for tools that save in the browser)
   Options on the script tag:
     data-gwc-export="off"        hide the button on this page
     data-storage-keys="a,b"      localStorage keys to include in backup
*/
(function () {
  var tag = document.currentScript;
  var off = tag && tag.getAttribute("data-gwc-export") === "off";
  var keys = ((tag && tag.getAttribute("data-storage-keys")) || "").split(",").map(function (k) { return k.trim(); }).filter(Boolean);

  function slug(s) { return (s || "gwc").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60); }
  function today() { return new Date().toISOString().slice(0, 10); }
  function toast(msg) {
    var t = document.createElement("div");
    t.className = "gwc-toast";
    t.setAttribute("role", "status");
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function () { t.remove(); }, 3500);
  }
  function download(name, mime, text) {
    var blob = new Blob(["﻿", text], { type: mime });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  var OPTION_SEL = ".dq-opt, .chip, .option-btn, .radio-card, .scale-btn, .yn-btn, .q-item, .accom-item, .format-card, .opt";
  var SELECTED_RE = /(^|[\s-])(selected|sel|active|checked|done|on)([\s-]|$)|selected-|sel-/i;

  function buildWordHtml() {
    var title = (document.title || "GroundWork Collaborative").replace(/\s*[|—-]\s*GroundWork.*$/i, "");
    var orig = Array.prototype.slice.call(document.body.querySelectorAll("*"));
    var clone = document.body.cloneNode(true);
    var cl = Array.prototype.slice.call(clone.querySelectorAll("*"));
    var drop = [];
    orig.forEach(function (el, i) {
      var c = cl[i];
      if (!c) return;
      var tagName = el.tagName;
      if (/^(SCRIPT|STYLE|NOSCRIPT|SVG|CANVAS|IFRAME|NAV|TEMPLATE)$/i.test(tagName) || (el.closest && el.closest(".gwc-save, .gwc-toast"))) { drop.push(c); return; }
      var cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden" || el.hidden || el.getAttribute("aria-hidden") === "true") { drop.push(c); return; }
      if (tagName === "INPUT") {
        var type = (el.type || "text").toLowerCase();
        var rep;
        if (type === "checkbox" || type === "radio") rep = document.createTextNode(el.checked ? "[x] " : "[ ] ");
        else if (type === "button" || type === "submit" || type === "file" || type === "hidden") { drop.push(c); return; }
        else rep = document.createTextNode(el.value ? el.value : "(blank)");
        c.__rep = rep;
      } else if (tagName === "TEXTAREA") {
        var d = document.createElement("div");
        d.style.whiteSpace = "pre-wrap";
        d.textContent = el.value || "(blank)";
        c.__rep = d;
      } else if (tagName === "SELECT") {
        var o = el.options[el.selectedIndex];
        c.__rep = document.createTextNode(o ? o.textContent : "");
      } else if (el.matches && el.matches(OPTION_SEL)) {
        if (SELECTED_RE.test(el.className)) c.__mark = true; else { drop.push(c); return; }
      } else if (tagName === "BUTTON") {
        if (SELECTED_RE.test(el.className)) c.__mark = true; else { drop.push(c); return; }
      }
    });
    drop.forEach(function (n) { if (n.parentNode) n.parentNode.removeChild(n); });
    Array.prototype.slice.call(clone.querySelectorAll("*")).forEach(function (n) {
      if (n.__rep && n.parentNode) n.parentNode.replaceChild(n.__rep, n);
    });
    Array.prototype.slice.call(clone.querySelectorAll("*")).forEach(function (n) {
      if (n.__mark) n.insertBefore(document.createTextNode("✔ "), n.firstChild);
      n.removeAttribute("style"); n.removeAttribute("class"); n.removeAttribute("id");
      Array.prototype.slice.call(n.attributes).forEach(function (a) { if (/^(on|data-|aria-)/i.test(a.name)) n.removeAttribute(a.name); });
    });
    var css = "body{font-family:Calibri,Arial,sans-serif;color:#012A4E;font-size:11pt}" +
      "h1,h2,h3,h4{color:#03567A;font-family:Georgia,serif}" +
      "table{border-collapse:collapse}td,th{border:1px solid #B8D7D8;padding:4px 6px;vertical-align:top}" +
      "th{background:#E3F2F3}.meta{color:#03567A;font-size:9pt}";
    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8"><title>' + esc(title) + "</title><style>" + css + "</style></head><body>" +
      '<p class="meta">GroundWork Collaborative &middot; ' + esc(title) + " &middot; Saved " + today() + "</p>" +
      clone.innerHTML + "</body></html>";
  }

  function wordExport() {
    var name = slug(document.title.replace(/\s*[|—-]\s*GroundWork.*$/i, "")) + "-" + today() + ".doc";
    download(name, "application/msword", buildWordHtml());
    toast("Word file saved to your Downloads folder.");
  }

  function backup() {
    var data = {};
    keys.forEach(function (k) { try { var v = localStorage.getItem(k); if (v !== null) data[k] = v; } catch (e) {} });
    if (!Object.keys(data).length) { toast("Nothing saved yet to back up."); return; }
    download("gwc-backup-" + slug(document.title) + "-" + today() + ".json", "application/json", JSON.stringify({ gwcBackup: 1, saved: new Date().toISOString(), data: data }, null, 2));
    toast("Backup file saved to your Downloads folder.");
  }

  function restore(file) {
    var r = new FileReader();
    r.onload = function () {
      try {
        var obj = JSON.parse(String(r.result).replace(/^﻿/, ""));
        if (!obj || obj.gwcBackup !== 1 || !obj.data) throw new Error("bad file");
        var n = 0;
        Object.keys(obj.data).forEach(function (k) { if (keys.indexOf(k) !== -1) { localStorage.setItem(k, obj.data[k]); n++; } });
        if (!n) throw new Error("no match");
        toast("Backup restored. Reloading.");
        setTimeout(function () { location.reload(); }, 800);
      } catch (e) { toast("That file is not a backup for this tool."); }
    };
    r.readAsText(file);
  }

  function build() {
    var wrap = document.createElement("div");
    wrap.className = "gwc-save";
    var btn = document.createElement("button");
    btn.type = "button"; btn.className = "gwc-save-btn"; btn.textContent = "Save my work";
    btn.setAttribute("aria-haspopup", "true"); btn.setAttribute("aria-expanded", "false");
    var menu = document.createElement("div");
    menu.className = "gwc-save-menu"; menu.hidden = true;
    function item(label, fn) {
      var b = document.createElement("button");
      b.type = "button"; b.textContent = label;
      b.addEventListener("click", function () { menu.hidden = true; btn.setAttribute("aria-expanded", "false"); fn(); });
      menu.appendChild(b);
    }
    item("Download as Word", wordExport);
    item("Print or save as PDF", function () { window.print(); });
    if (keys.length) {
      item("Back up my data (file)", backup);
      var fi = document.createElement("input");
      fi.type = "file"; fi.accept = ".json,application/json"; fi.hidden = true;
      fi.addEventListener("change", function () { if (fi.files[0]) restore(fi.files[0]); fi.value = ""; });
      item("Restore from backup", function () { fi.click(); });
      wrap.appendChild(fi);
      var note = document.createElement("div");
      note.className = "gwc-save-note";
      note.textContent = "Your data stays on this device. Back it up to move it.";
      menu.appendChild(note);
    }
    btn.addEventListener("click", function () {
      var open = menu.hidden;
      menu.hidden = !open;
      btn.setAttribute("aria-expanded", open ? "true" : "false");
    });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !menu.hidden) { menu.hidden = true; btn.setAttribute("aria-expanded", "false"); btn.focus(); } });
    document.addEventListener("click", function (e) { if (!wrap.contains(e.target)) { menu.hidden = true; btn.setAttribute("aria-expanded", "false"); } });
    wrap.appendChild(menu); wrap.appendChild(btn);
    document.body.appendChild(wrap);
  }

  if (!off) {
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", build); else build();
  }
})();
