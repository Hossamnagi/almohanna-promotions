#!/usr/bin/env python3
"""
generate_promo_page.py

Regenerates the Active Promotions HTML page directly from the
"Active_Promotions" sheet produced by pharmacy_pipeline.py.

Usage:
    python generate_promo_page.py <path_to_workbook.xlsx> [output.html]

If <path_to_workbook.xlsx> is the full multi-sheet dashboard that
pharmacy_pipeline.py outputs, this reads the "Active_Promotions" sheet
from it. If it's a workbook that only has one sheet (already just the
promotions export), it reads that sheet directly.

Drop this file anywhere in your pharmacy_pipeline.py project (e.g. next
to it in Y:/Reports_Export/) and either:
  - run it manually after each pipeline run, or
  - call it as the last step of pharmacy_pipeline.py itself, e.g.:
        subprocess.run(["python", "generate_promo_page.py",
                         "Reports_Export/dashboard.xlsx",
                         "Reports_Export/active_promotions.html"])
"""

import sys
import json
import pandas as pd


TEMPLATE_HEAD = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Active Promotions — Al-Mohanna</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Kufi+Arabic:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
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
  --font-display:'Space Grotesk', system-ui, sans-serif;
  --font-sans:'Inter', system-ui, -apple-system, sans-serif;
  --font-mono:'IBM Plex Mono', ui-monospace, monospace;
  --font-ar:'Noto Kufi Arabic', system-ui, sans-serif;
  --radius-s:4px;
  --radius-m:10px;
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
}
body::before{
  content:"";
  position:fixed;
  inset:0;
  z-index:-2;
  background:
    radial-gradient(700px 480px at 12% -10%, rgba(139,107,255,0.32), transparent 60%),
    radial-gradient(650px 500px at 88% 8%, rgba(79,214,255,0.22), transparent 60%),
    radial-gradient(600px 460px at 50% 115%, rgba(255,95,168,0.16), transparent 65%),
    var(--bg-deep);
}
body::after{
  content:"";
  position:fixed;
  inset:0;
  z-index:-1;
  opacity:0.5;
  background-image:radial-gradient(rgba(255,255,255,0.035) 1px, transparent 1px);
  background-size:26px 26px;
}
::selection{background:rgba(139,107,255,0.45);color:#fff;}
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
  background:linear-gradient(100deg, #ffffff 20%, var(--blue) 55%, var(--violet) 85%);
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
  background:rgba(10,11,31,0.82);
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
  background:rgba(255,255,255,0.03);
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
  background:rgba(255,255,255,0.03) url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="6"><path d="M0 0l5 6 5-6z" fill="%239297c2"/></svg>') no-repeat right 12px center;
  color:var(--ink);
  appearance:none;
  cursor:pointer;
  max-width:180px;
}
select option{background:#12142f;color:var(--ink);}
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
  color:#0a0b1f;
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
  const today = new Date('2026-09-16T00:00:00');
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


def fmt_date(v):
    if pd.isna(v):
        return None
    return pd.to_datetime(v, dayfirst=True).strftime("%Y-%m-%d")


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
            "actDate": fmt_date(row.get("Activation Date")),
            "endDate": fmt_date(row.get("End Date")),
        }
        records.append(rec)
    return records


def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_promo_page.py <workbook.xlsx> [output.html]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "active_promotions.html"

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