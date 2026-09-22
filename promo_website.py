#!/usr/bin/env python3
"""
promo_website.py

Regenerates the Active Promotions HTML page directly from the
"Active_Promotions" sheet produced by pharmacy_pipeline.py.

Usage:
    python promo_website.py <path_to_workbook.xlsx> [output.html]

If <path_to_workbook.xlsx> is the full multi-sheet dashboard that
pharmacy_pipeline.py outputs, this reads the "Active_Promotions" sheet
from it. If it's a workbook that only has one sheet (already just the
promotions export), it reads that sheet directly.

Drop this file anywhere in your pharmacy_pipeline.py project (e.g. next
to it in Y:/Reports_Export/) and either:
  - run it manually after each pipeline run, or
  - call it as the last step of pharmacy_pipeline.py itself, e.g.:
        subprocess.run(["python", "promo_website.py",
                         "Reports_Export/dashboard.xlsx",
                         "Reports_Export/index.html"])

The page supports both light and dark mode: it follows the visitor's
OS/browser preference by default, and a toggle button in the toolbar
lets them override it (remembered per-browser via localStorage).

Includes a camera barcode scanner (via the ZXing JS library, loaded
from jsDelivr) -- works in any modern mobile browser, including iOS
Safari, once this page is served over HTTPS (e.g. GitHub Pages). It
will NOT get camera access inside a sandboxed preview iframe (such as
a Claude.ai artifact preview) -- that's expected there, and normal on
a real hosted URL.
"""

import sys
import json
import pandas as pd


TEMPLATE_HEAD = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light dark">
<script>
(function(){
  try{
    var t = localStorage.getItem('promoTheme');
    if(t !== 'light' && t !== 'dark'){
      t = (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
    }
    document.documentElement.setAttribute('data-theme', t);
  }catch(e){}
})();
</script>
<title>Active Promotions — Al-Mohanna</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Kufi+Arabic:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  /* ---- light theme (default) ---- */
  --bg:#f3f4fb;
  --bg-deep:#ffffff;
  --card:rgba(23,25,51,0.035);
  --card-hover:rgba(23,25,51,0.065);
  --border:rgba(88,66,214,0.16);
  --border-strong:rgba(88,66,214,0.34);
  --ink:#181a33;
  --muted:#5b5e80;
  --muted-dim:#888cab;
  --violet:#5433d6;
  --blue:#0a6d8a;
  --pink:#b8215d;
  --gold:#9a6406;
  --grad1:rgba(109,79,255,0.14);
  --grad2:rgba(10,109,138,0.12);
  --grad3:rgba(184,33,93,0.09);
  --noise-dot:rgba(20,20,50,0.045);
  --sel-bg:rgba(84,51,214,0.22);
  --sel-ink:#14152c;
  --h1-start:var(--ink);
  --toolbar-bg:rgba(255,255,255,0.82);
  --panel-bg:#ffffff;
  --on-brand:#ffffff;
  --select-arrow:url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="6"><path d="M0 0l5 6 5-6z" fill="%235b5e80"/></svg>');
  --font-display:'Space Grotesk', system-ui, sans-serif;
  --font-sans:'Inter', system-ui, -apple-system, sans-serif;
  --font-mono:'IBM Plex Mono', ui-monospace, monospace;
  --font-ar:'Noto Kufi Arabic', system-ui, sans-serif;
  --radius-s:4px;
  --radius-m:10px;
  color-scheme:light;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0a0b1f;
    --bg-deep:#050614;
    --card:rgba(255,255,255,0.045);
    --card-hover:rgba(255,255,255,0.075);
    --border:rgba(151,140,255,0.18);
    --border-strong:rgba(151,140,255,0.4);
    --ink:#eceeff;
    --muted:#9297c2;
    --muted-dim:#6a6f96;
    --violet:#8b6bff;
    --blue:#4fd6ff;
    --pink:#ff5fa8;
    --gold:#ffcf6b;
    --grad1:rgba(139,107,255,0.32);
    --grad2:rgba(79,214,255,0.22);
    --grad3:rgba(255,95,168,0.16);
    --noise-dot:rgba(255,255,255,0.035);
    --sel-bg:rgba(139,107,255,0.45);
    --sel-ink:#ffffff;
    --h1-start:#ffffff;
    --toolbar-bg:rgba(10,11,31,0.82);
    --panel-bg:#0c0e24;
    --on-brand:#0a0b1f;
    --select-arrow:url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="6"><path d="M0 0l5 6 5-6z" fill="%239297c2"/></svg>');
    color-scheme:dark;
  }
}
:root[data-theme="dark"]{
  --bg:#0a0b1f;
  --bg-deep:#050614;
  --card:rgba(255,255,255,0.045);
  --card-hover:rgba(255,255,255,0.075);
  --border:rgba(151,140,255,0.18);
  --border-strong:rgba(151,140,255,0.4);
  --ink:#eceeff;
  --muted:#9297c2;
  --muted-dim:#6a6f96;
  --violet:#8b6bff;
  --blue:#4fd6ff;
  --pink:#ff5fa8;
  --gold:#ffcf6b;
  --grad1:rgba(139,107,255,0.32);
  --grad2:rgba(79,214,255,0.22);
  --grad3:rgba(255,95,168,0.16);
  --noise-dot:rgba(255,255,255,0.035);
  --sel-bg:rgba(139,107,255,0.45);
  --sel-ink:#ffffff;
  --h1-start:#ffffff;
  --toolbar-bg:rgba(10,11,31,0.82);
  --panel-bg:#0c0e24;
  --on-brand:#0a0b1f;
  --select-arrow:url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="6"><path d="M0 0l5 6 5-6z" fill="%239297c2"/></svg>');
  color-scheme:dark;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  background:var(--bg);
  color:var(--ink);
  font-family:var(--font-sans);
  -webkit-font-smoothing:antialiased;
  min-height:100vh;
  position:relative;
  transition:background .2s, color .2s;
}
body::before{
  content:"";
  position:fixed;
  inset:0;
  z-index:-2;
  background:
    radial-gradient(700px 480px at 12% -10%, var(--grad1), transparent 60%),
    radial-gradient(650px 500px at 88% 8%, var(--grad2), transparent 60%),
    radial-gradient(600px 460px at 50% 115%, var(--grad3), transparent 65%),
    var(--bg-deep);
}
body::after{
  content:"";
  position:fixed;
  inset:0;
  z-index:-1;
  opacity:0.5;
  background-image:radial-gradient(var(--noise-dot) 1px, transparent 1px);
  background-size:26px 26px;
}
::selection{background:var(--sel-bg);color:var(--sel-ink);}
a{color:inherit;}
button{font-family:inherit;}

/* ---------- Header ---------- */
.masthead{
  padding:40px 28px 26px;
  position:relative;
  border-bottom:1px solid var(--border);
}
.masthead-row{
  max-width:1320px;
  margin:0 auto;
}
.brandmark{
  display:inline-flex;align-items:center;gap:8px;
  font-family:var(--font-mono);
  font-size:12px;
  letter-spacing:0.14em;
  color:var(--blue);
  border:1px solid var(--border-strong);
  padding:5px 12px;
  border-radius:20px;
  text-transform:uppercase;
  background:rgba(79,214,255,0.06);
}
.brandmark .dot{
  width:6px;height:6px;border-radius:50%;
  background:var(--blue);
  box-shadow:0 0 8px 2px rgba(79,214,255,0.8);
}
h1{
  font-family:var(--font-display);
  font-size:clamp(30px,4.6vw,44px);
  font-weight:700;
  margin:16px 0 4px;
  letter-spacing:-0.01em;
  line-height:1.08;
  background:linear-gradient(100deg, var(--h1-start) 20%, var(--blue) 55%, var(--violet) 85%);
  -webkit-background-clip:text;
  background-clip:text;
  -webkit-text-fill-color:transparent;
}
.subtitle{
  font-family:var(--font-ar);
  font-size:15px;
  color:var(--muted);
  direction:rtl;
  margin:0;
}
.meta-line{
  font-size:13.5px;
  color:var(--muted);
  margin-top:10px;
}
.meta-line strong{color:var(--ink);font-weight:600;}

/* ---------- Toolbar ---------- */
.toolbar{
  position:sticky;top:0;z-index:20;
  background:var(--toolbar-bg);
  backdrop-filter:blur(14px);
  -webkit-backdrop-filter:blur(14px);
  border-bottom:1px solid var(--border);
  padding:14px 28px;
}
.toolbar-inner{
  max-width:1320px;margin:0 auto;
  display:flex;flex-direction:column;gap:10px;
}
.search-row{
  display:flex;gap:10px;align-items:center;
}
.scan-btn, .theme-btn{
  flex-shrink:0;
  width:44px;height:44px;
  display:flex;align-items:center;justify-content:center;
  border:1.5px solid var(--border-strong);
  border-radius:var(--radius-m);
  background:rgba(139,107,255,0.08);
  color:var(--blue);
  cursor:pointer;
  transition:background .15s, box-shadow .15s;
}
.scan-btn:hover, .theme-btn:hover{
  background:rgba(139,107,255,0.18);
  box-shadow:0 0 16px rgba(79,214,255,0.2);
}
.scan-btn svg{width:20px;height:20px;}
.theme-btn svg{width:20px;height:20px;display:none;}
.theme-btn .icon-moon{display:block;}
:root[data-theme="dark"] .theme-btn .icon-sun{display:block;}
:root[data-theme="dark"] .theme-btn .icon-moon{display:none;}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]) .theme-btn .icon-sun{display:block;}
  :root:not([data-theme="light"]) .theme-btn .icon-moon{display:none;}
}

/* ---------- Scanner modal ---------- */
.scan-overlay{
  position:fixed;inset:0;
  background:rgba(5,6,20,0.85);
  backdrop-filter:blur(6px);
  -webkit-backdrop-filter:blur(6px);
  display:none;
  align-items:center;justify-content:center;
  z-index:100;
  padding:20px;
}
.scan-overlay.open{display:flex;}
.scan-modal{
  width:100%;max-width:420px;
  background:var(--panel-bg);
  border:1.5px solid var(--border-strong);
  border-radius:14px;
  padding:18px;
  box-shadow:0 0 40px rgba(139,107,255,0.25);
  position:relative;
}
.scan-modal-head{
  display:flex;justify-content:space-between;align-items:center;
  margin-bottom:12px;
}
.scan-modal-head h2{
  font-family:var(--font-display);
  font-size:16px;margin:0;color:var(--ink);
}
.scan-close{
  background:none;border:none;color:var(--muted);
  font-size:20px;line-height:1;cursor:pointer;
  padding:4px 8px;
}
.scan-close:hover{color:var(--ink);}
.scan-video-wrap{
  position:relative;
  border-radius:10px;
  overflow:hidden;
  background:#000;
  aspect-ratio:4/3;
}
.scan-video-wrap video{
  width:100%;height:100%;object-fit:cover;display:block;
}
.scan-reticle{
  position:absolute;inset:18% 10%;
  border:2px solid var(--blue);
  border-radius:8px;
  box-shadow:0 0 0 999px rgba(0,0,0,0.35);
  pointer-events:none;
}
.scan-status{
  margin-top:12px;
  font-family:var(--font-mono);
  font-size:12.5px;
  color:var(--muted);
  text-align:center;
}
.search-box{
  flex:1;
  position:relative;
}
.search-box input{
  width:100%;
  font-family:var(--font-sans);
  font-size:15px;
  padding:11px 14px 11px 38px;
  border:1.5px solid var(--border);
  border-radius:var(--radius-m);
  background:var(--card);
  color:var(--ink);
  outline:none;
  transition:border-color .15s, box-shadow .15s;
}
.search-box input::placeholder{color:var(--muted-dim);}
.search-box input:focus{
  border-color:var(--violet);
  box-shadow:0 0 0 3px rgba(139,107,255,0.15);
}
.search-box svg{
  position:absolute;left:12px;top:50%;transform:translateY(-50%);
  width:16px;height:16px;color:var(--muted);pointer-events:none;
}
.result-count{
  font-family:var(--font-mono);
  font-size:12.5px;
  color:var(--blue);
  white-space:nowrap;
  min-width:112px;
  text-align:right;
}
.filters-row{
  display:flex;gap:8px;flex-wrap:wrap;align-items:center;
}
select{
  font-family:var(--font-sans);
  font-size:13.5px;
  padding:8px 30px 8px 12px;
  border:1.5px solid var(--border);
  border-radius:var(--radius-m);
  background:var(--card) var(--select-arrow) no-repeat right 12px center;
  color:var(--ink);
  appearance:none;
  cursor:pointer;
  max-width:180px;
}
select option{background:var(--panel-bg);color:var(--ink);}
select:focus{outline:none;border-color:var(--violet);}
.clear-btn{
  font-size:13px;
  color:var(--pink);
  background:none;
  border:1px solid transparent;
  padding:8px 10px;
  cursor:pointer;
  border-radius:var(--radius-m);
  font-weight:600;
}
.clear-btn:hover{background:rgba(255,95,168,0.1);}

/* ---------- Grid ---------- */
main{max-width:1320px;margin:0 auto;padding:26px 28px 60px;}
.grid{
  display:grid;
  grid-template-columns:repeat(auto-fill,minmax(272px,1fr));
  gap:14px;
}
.card{
  background:var(--card);
  border:1px solid var(--border);
  border-radius:var(--radius-m);
  padding:16px;
  display:flex;
  flex-direction:column;
  gap:9px;
  position:relative;
  backdrop-filter:blur(6px);
  transition:border-color .15s, background .15s, transform .15s;
}
.card:hover{
  border-color:var(--border-strong);
  background:var(--card-hover);
}
.card-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px;}
.brand-tag{
  font-family:var(--font-mono);
  font-size:11px;
  letter-spacing:0.05em;
  color:var(--blue);
  font-weight:600;
  text-transform:uppercase;
}
.subgroup-tag{
  font-size:11px;
  color:var(--muted);
  background:rgba(255,255,255,0.04);
  border:1px solid var(--border);
  padding:2px 8px;
  border-radius:20px;
  white-space:nowrap;
}
.item-name{
  font-family:var(--font-display);
  font-size:15px;
  font-weight:600;
  line-height:1.35;
  margin:0;
  color:var(--ink);
}
.item-name-ar{
  font-family:var(--font-ar);
  font-size:14px;
  color:var(--muted);
  direction:rtl;
  text-align:right;
  line-height:1.5;
  margin:0;
}
.tag-row{
  display:flex;
  gap:6px;
  flex-wrap:wrap;
}
.ending-tag{
  align-self:flex-start;
  font-size:11.5px;
  font-weight:600;
  color:var(--pink);
  background:rgba(255,95,168,0.1);
  border:1px solid rgba(255,95,168,0.3);
  padding:3px 9px;
  border-radius:20px;
}
.new-tag{
  align-self:flex-start;
  font-size:11.5px;
  font-weight:700;
  color:var(--blue);
  background:rgba(79,214,255,0.1);
  border:1px solid rgba(79,214,255,0.35);
  padding:3px 9px;
  border-radius:20px;
}
.price-row{
  display:flex;
  align-items:baseline;
  gap:8px;
  flex-wrap:wrap;
  margin-top:auto;
  padding-top:10px;
  border-top:1px dashed var(--border);
}
.offer-chip{
  margin-left:auto;
  font-family:var(--font-sans);
  font-size:11.5px;
  font-weight:600;
  color:var(--violet);
  background:rgba(139,107,255,0.1);
  border:1px solid rgba(139,107,255,0.35);
  padding:3px 10px;
  border-radius:20px;
  white-space:normal;
  text-align:right;
  line-height:1.3;
}
.price-now{
  font-family:var(--font-mono);
  font-size:17px;
  font-weight:600;
  color:var(--gold);
}
.price-was{
  font-family:var(--font-mono);
  font-size:12.5px;
  color:var(--muted-dim);
  text-decoration:line-through;
}
.card-footer{
  display:flex;
  justify-content:space-between;
  align-items:center;
  font-family:var(--font-mono);
  font-size:10.5px;
  color:var(--muted-dim);
}
.card-footer .code{color:var(--violet);}
.empty-state{
  grid-column:1/-1;
  text-align:center;
  padding:70px 20px;
  color:var(--muted);
}
.empty-state .glyph{font-size:34px;display:block;margin-bottom:10px;color:var(--violet);}
.empty-state button{
  margin-top:14px;
  background:linear-gradient(100deg, var(--violet), var(--blue));
  color:var(--on-brand);
  border:none;
  padding:9px 18px;
  border-radius:var(--radius-m);
  font-weight:700;
  cursor:pointer;
}

.load-more-wrap{
  display:flex;
  justify-content:center;
  margin-top:26px;
}
.load-more{
  background:rgba(255,255,255,0.03);
  border:1.5px solid var(--border-strong);
  color:var(--ink);
  font-weight:700;
  font-size:14px;
  padding:12px 30px;
  border-radius:var(--radius-m);
  cursor:pointer;
  transition:background .15s, box-shadow .15s;
}
.load-more:hover{
  background:linear-gradient(100deg, rgba(139,107,255,0.25), rgba(79,214,255,0.2));
  box-shadow:0 0 24px rgba(139,107,255,0.25);
}

@media(max-width:640px){
  .masthead{padding:28px 16px 20px;}
  .toolbar{padding:12px 16px;}
  main{padding:18px 16px 50px;}
  select{max-width:none;flex:1;}
}
</style>
</head>
<body>
"""

TEMPLATE_BODY = """<header class="masthead">
  <div class="masthead-row">
    <span class="brandmark"><span class="dot"></span>Al-Mohanna Pharmacies</span>
    <h1>Active Promotions</h1>
    <p class="subtitle">دليل العروض النشطة — تصفح وابحث عن المنتجات المخفضة</p>
    <p class="meta-line">Live directory of <strong id="totalCount">—</strong> promoted items across <strong id="brandCount">—</strong> brands.</p>
  </div>
</header>

<div class="toolbar">
  <div class="toolbar-inner">
    <div class="search-row">
      <div class="search-box">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
        <input id="searchInput" type="text" placeholder="Search by item name, brand, code or barcode…" autocomplete="off">
      </div>
      <button class="scan-btn" id="scanBtn" title="Scan barcode" type="button" aria-label="Scan barcode">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/><line x1="7" y1="8" x2="7" y2="16"/><line x1="10" y1="8" x2="10" y2="16"/><line x1="13" y1="8" x2="13" y2="16"/><line x1="17" y1="8" x2="17" y2="16"/></svg>
      </button>
      <button class="theme-btn" id="themeBtn" title="Toggle light / dark theme" type="button" aria-label="Toggle light or dark theme">
        <svg class="icon-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></svg>
        <svg class="icon-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3a7 7 0 0 0 9.79 9.79z"/></svg>
      </button>
      <div class="result-count" id="resultCount">— items</div>
    </div>
    <div class="filters-row">
      <select id="categoryFilter"><option value="">All categories</option></select>
      <select id="brandFilter"><option value="">All brands</option></select>
      <select id="typeFilter"><option value="">All promotion types</option></select>
      <select id="offerFilter"><option value="">All offers</option></select>
      <select id="sortSelect">
        <option value="recent">Sort: Recently added</option>
        <option value="name">Sort: Name (A–Z)</option>
        <option value="priceAsc">Sort: Price (low to high)</option>
        <option value="priceDesc">Sort: Price (high to low)</option>
        <option value="discount">Sort: Biggest discount</option>
        <option value="ending">Sort: Ending soonest</option>
        <option value="brand">Sort: Brand</option>
      </select>
      <button class="clear-btn" id="clearBtn">Clear filters</button>
    </div>
  </div>
</div>

<main>
  <div class="grid" id="grid"></div>
  <div class="load-more-wrap" id="loadMoreWrap">
    <button class="load-more" id="loadMoreBtn">Show more items</button>
  </div>
</main>

<div class="scan-overlay" id="scanOverlay">
  <div class="scan-modal">
    <div class="scan-modal-head">
      <h2>Scan barcode</h2>
      <button class="scan-close" id="scanCloseBtn" type="button">&times;</button>
    </div>
    <div class="scan-video-wrap">
      <video id="scanVideo" muted playsinline autoplay></video>
      <div class="scan-reticle"></div>
    </div>
    <p class="scan-status" id="scanStatus">Starting camera…</p>
  </div>
</div>

<script>
const DATA = __DATA__;
const PAGE_SIZE = 60;
let filtered = [];
let shown = 0;

const grid = document.getElementById('grid');
const searchInput = document.getElementById('searchInput');
const categoryFilter = document.getElementById('categoryFilter');
const brandFilter = document.getElementById('brandFilter');
const typeFilter = document.getElementById('typeFilter');
const offerFilter = document.getElementById('offerFilter');
const sortSelect = document.getElementById('sortSelect');
const resultCount = document.getElementById('resultCount');
const loadMoreWrap = document.getElementById('loadMoreWrap');
const loadMoreBtn = document.getElementById('loadMoreBtn');
const clearBtn = document.getElementById('clearBtn');

function uniqueSorted(key){
  const s = new Set();
  DATA.forEach(r => { if(r[key]) s.add(r[key]); });
  return Array.from(s).sort((a,b)=>a.localeCompare(b));
}

function populateSelect(sel, values){
  values.forEach(v=>{
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
}

function fmtDate(d){
  if(!d) return '';
  const dt = new Date(d+'T00:00:00');
  return dt.toLocaleDateString('en-GB', {day:'2-digit', month:'short', year:'numeric'});
}

function daysLeft(endDate){
  if(!endDate) return null;
  const end = new Date(endDate+'T00:00:00');
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.ceil((end - today) / 86400000);
}

function escapeHtml(s){
  if(s===null||s===undefined) return '';
  return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function computeDiscountPct(r){
  if(r.discPct) return r.discPct;
  if(r.priceTax && r.priceAfterDisc && r.priceTax > 0){
    return Math.round((1 - r.priceAfterDisc/r.priceTax) * 100);
  }
  return 0;
}

let MAX_ACT_DATE = null;

function cardHtml(r){
  const finalPrice = r.priceAfterDisc || r.priceTax;
  const showWas = r.priceAfterDisc && r.priceTax && r.priceAfterDisc < r.priceTax;
  const days = daysLeft(r.endDate);
  const endingSoon = days !== null && days <= 7 && days >= 0;
  const isNew = r.actDate && MAX_ACT_DATE && (MAX_ACT_DATE - new Date(r.actDate+'T00:00:00')) <= 3*86400000;
  return `<div class="card">
    <div class="card-top">
      <span class="brand-tag">${escapeHtml(r.brand||'')}</span>
      ${r.subGroup ? `<span class="subgroup-tag">${escapeHtml(r.subGroup)}</span>` : ''}
    </div>
    <p class="item-name">${escapeHtml(r.nameEn)}</p>
    ${r.nameAr ? `<p class="item-name-ar">${escapeHtml(r.nameAr)}</p>` : ''}
    <div class="tag-row">
      ${isNew ? `<span class="new-tag">New</span>` : ''}
      ${endingSoon ? `<span class="ending-tag">Ends in ${days}d</span>` : ''}
    </div>
    <div class="price-row">
      <span class="price-now">SAR ${finalPrice != null ? finalPrice.toFixed(2) : '—'}</span>
      ${showWas ? `<span class="price-was">SAR ${r.priceTax.toFixed(2)}</span>` : ''}
      ${r.offer ? `<span class="offer-chip">${escapeHtml(r.offer)}</span>` : ''}
    </div>
    <div class="card-footer">
      <span class="code">#${escapeHtml(r.code)}</span>
      <span>until ${fmtDate(r.endDate)}</span>
    </div>
  </div>`;
}

function applyFilters(){
  const q = searchInput.value.trim().toLowerCase();
  const cat = categoryFilter.value;
  const brand = brandFilter.value;
  const type = typeFilter.value;
  const offer = offerFilter.value;

  filtered = DATA.filter(r=>{
    if(cat && r.subGroup !== cat) return false;
    if(brand && r.brand !== brand) return false;
    if(type && r.type !== type) return false;
    if(offer && r.offer !== offer) return false;
    if(q){
      const hay = (r.nameEn+' '+r.nameAr+' '+r.brand+' '+r.code+' '+r.barcode).toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });

  const sort = sortSelect.value;
  filtered.sort((a,b)=>{
    if(sort === 'name') return (a.nameEn||'').localeCompare(b.nameEn||'');
    if(sort === 'brand') return (a.brand||'').localeCompare(b.brand||'');
    if(sort === 'priceAsc') return (a.priceAfterDisc||a.priceTax||0) - (b.priceAfterDisc||b.priceTax||0);
    if(sort === 'priceDesc') return (b.priceAfterDisc||b.priceTax||0) - (a.priceAfterDisc||a.priceTax||0);
    if(sort === 'discount') return computeDiscountPct(b) - computeDiscountPct(a);
    if(sort === 'ending') return new Date(a.endDate) - new Date(b.endDate);
    if(sort === 'recent') return new Date(b.actDate) - new Date(a.actDate);
    return 0;
  });

  shown = 0;
  renderPage(true);
}

function renderPage(reset){
  if(reset) grid.innerHTML = '';
  const next = filtered.slice(shown, shown + PAGE_SIZE);
  if(filtered.length === 0){
    grid.innerHTML = `<div class="empty-state"><span class="glyph">∅</span>No items match these filters.<br><button id="resetFromEmpty">Clear filters</button></div>`;
    document.getElementById('resetFromEmpty').onclick = clearFilters;
    loadMoreWrap.style.display = 'none';
  } else {
    grid.insertAdjacentHTML('beforeend', next.map(cardHtml).join(''));
    shown += next.length;
    loadMoreWrap.style.display = shown < filtered.length ? 'flex' : 'none';
  }
  resultCount.textContent = `${filtered.length} item${filtered.length===1?'':'s'}`;
}

function clearFilters(){
  searchInput.value = '';
  categoryFilter.value = '';
  brandFilter.value = '';
  typeFilter.value = '';
  offerFilter.value = '';
  sortSelect.value = 'recent';
  applyFilters();
}

loadMoreBtn.addEventListener('click', ()=>renderPage(false));
[searchInput].forEach(el=>el.addEventListener('input', applyFilters));
[categoryFilter, brandFilter, typeFilter, offerFilter, sortSelect].forEach(el=>el.addEventListener('change', applyFilters));
clearBtn.addEventListener('click', clearFilters);

// init
populateSelect(categoryFilter, uniqueSorted('subGroup'));
populateSelect(brandFilter, uniqueSorted('brand'));
populateSelect(typeFilter, uniqueSorted('type'));
populateSelect(offerFilter, uniqueSorted('offer'));

document.getElementById('totalCount').textContent = DATA.length;
document.getElementById('brandCount').textContent = new Set(DATA.map(r=>r.brand)).size;

const actDates = DATA.map(r=>r.actDate).filter(Boolean).map(d=>new Date(d+'T00:00:00'));
MAX_ACT_DATE = actDates.length ? new Date(Math.max(...actDates)) : null;

applyFilters();

// ---------- Theme toggle ----------
const themeBtn = document.getElementById('themeBtn');

function setTheme(theme){
  document.documentElement.setAttribute('data-theme', theme);
  try{ localStorage.setItem('promoTheme', theme); }catch(e){}
}

themeBtn.addEventListener('click', ()=>{
  const current = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  setTheme(current === 'dark' ? 'light' : 'dark');
});

if(window.matchMedia){
  const mq = window.matchMedia('(prefers-color-scheme: dark)');
  const handleSystemChange = (e)=>{
    let stored = null;
    try{ stored = localStorage.getItem('promoTheme'); }catch(err){}
    if(stored !== 'light' && stored !== 'dark'){
      document.documentElement.setAttribute('data-theme', e.matches ? 'dark' : 'light');
    }
  };
  if(mq.addEventListener) mq.addEventListener('change', handleSystemChange);
  else if(mq.addListener) mq.addListener(handleSystemChange);
}

// ---------- Barcode scanner ----------
const scanBtn = document.getElementById('scanBtn');
const scanOverlay = document.getElementById('scanOverlay');
const scanCloseBtn = document.getElementById('scanCloseBtn');
const scanVideo = document.getElementById('scanVideo');
const scanStatus = document.getElementById('scanStatus');

let zxingLoadPromise = null;
let codeReader = null;
let scanControls = null;

function ensureZXingLoaded(){
  if(window.ZXing) return Promise.resolve();
  if(zxingLoadPromise) return zxingLoadPromise;
  zxingLoadPromise = new Promise((resolve, reject)=>{
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/@zxing/library@0.20.0/umd/index.min.js';
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('Could not load the scanner library.'));
    document.head.appendChild(s);
  });
  return zxingLoadPromise;
}

async function openScanner(){
  scanOverlay.classList.add('open');
  scanStatus.textContent = 'Starting camera…';
  if(!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia){
    scanStatus.textContent = 'Camera access is not available in this browser.';
    return;
  }
  try{
    await ensureZXingLoaded();
    codeReader = new ZXing.BrowserMultiFormatReader();
    scanControls = await codeReader.decodeFromConstraints(
      { video: { facingMode: { ideal: 'environment' } } },
      scanVideo,
      (result, err, controls) => {
        if(result){
          searchInput.value = result.getText();
          applyFilters();
          closeScanner();
        }
      }
    );
    scanStatus.textContent = 'Point the camera at a barcode';
  } catch(e){
    const msg = (e && e.name === 'NotAllowedError')
      ? 'Camera permission was denied.'
      : (e && e.message) ? e.message : 'Camera unavailable in this view.';
    scanStatus.textContent = msg;
  }
}

function closeScanner(){
  scanOverlay.classList.remove('open');
  if(scanControls){
    try{ scanControls.stop(); } catch(e){}
    scanControls = null;
  }
}

scanBtn.addEventListener('click', openScanner);
scanCloseBtn.addEventListener('click', closeScanner);
scanOverlay.addEventListener('click', (e)=>{ if(e.target === scanOverlay) closeScanner(); });

</script>
</body>
</html>
"""


def load_promotions(xlsx_path: str) -> pd.DataFrame:
    xl = pd.ExcelFile(xlsx_path)
    sheet_name = "Active_Promotions" if "Active_Promotions" in xl.sheet_names else xl.sheet_names[0]
    return pd.read_excel(xl, sheet_name=sheet_name)


def clean(v):
    if pd.isna(v):
        return None
    return v


def parse_date(v):
    """Return v as an ISO 'YYYY-MM-DD' string.

    The pipeline's date columns aren't consistently typed -- Activation
    Date can come through as a 'DD/MM/YYYY' string while End Date is
    already a Timestamp -- so this handles either.
    """
    if pd.isna(v):
        return None
    if isinstance(v, str):
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
    else:
        dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.strftime("%Y-%m-%d")


def build_records(df: pd.DataFrame) -> list:
    records = []
    for _, row in df.iterrows():
        rec = {
            "q": clean(row.get("Quotation No")),
            "brand": clean(row.get("Brand")),
            "code": clean(row.get("Item Code")),
            # NOTE: in this sheet, "Foreign Name" holds the English name
            # and "Item Name" holds the Arabic name (columns are swapped
            # from what their labels suggest) -- confirmed against the
            # source data on 2026-09-17.
            "nameEn": clean(row.get("Foreign Name")),
            "nameAr": clean(row.get("Item Name")),
            "type": clean(row.get("Promotion Type")),
            "offer": clean(row.get("Offer")),
            "price": clean(row.get("Price")),
            "tax": clean(row.get("Tax%")),
            "discPct": clean(row.get("Discount Percentage")),
            "priceTax": clean(row.get("Price with Tax")),
            "priceAfterDisc": clean(row.get("Price After Disc")),
            "subGroup": clean(row.get("Sub Group Name")),
            "barcode": clean(row.get("Barcodes")),
            "vendor": clean(row.get("Main Vendor")),
            "actDate": parse_date(row.get("Activation Date")),
            "endDate": parse_date(row.get("End Date")),
        }
        records.append(rec)
    return records


def main():
    if len(sys.argv) < 2:
        print("Usage: python promo_website.py <workbook.xlsx> [output.html]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "index.html"

    df = load_promotions(xlsx_path)
    records = build_records(df)
    data_json = json.dumps(records, ensure_ascii=False)

    body = TEMPLATE_BODY.replace("__DATA__", data_json)
    full_html = TEMPLATE_HEAD + body

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(full_html)

    print(f"Wrote {len(records)} items to {out_path}")


if __name__ == "__main__":
    main()
