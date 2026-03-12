#!/usr/bin/env python3
"""
Анализатор данных: читает Excel или HTML, строит граф + диаграмму вендоров.
Использование:
    python3 analyzer.py                    # интерактивный режим
    python3 analyzer.py myfile.xlsx        # прямой запуск
    python3 analyzer.py myfile.html
"""

import sys
import os
import json
import argparse
import webbrowser
import tempfile
from pathlib import Path

# ── зависимости ────────────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    print("Установите openpyxl:  pip install openpyxl")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("Установите beautifulsoup4:  pip install beautifulsoup4")
    sys.exit(1)

# ── константы ──────────────────────────────────────────────────────────────────
ROLE_COLORS = {
    0: "#4f8ef7",   # синий
    1: "#9b6dff",   # фиолетовый
    2: "#00c9b1",   # циан
    3: "#ffb547",   # янтарный
    4: "#ff4b6e",   # красный
    5: "#32d98b",   # зелёный
    6: "#ff6eb4",   # розовый
    7: "#f97316",   # оранжевый
}

RF_KEYWORDS = [
    "касперский", "kaspersky", "positive", "инфотекс", "крипто", "аладдин",
    "индид", "гарда", "solar", "солар", "bizone", "bi.zone", "ростелеком",
    "сбер", "мегафон", "eltex", "элтекс", "эшелон", "рбсофт", "softline",
    "доктор веб", "drweb", "vipnet", "випнет", "код безопасности",
    "usergate", "юзергейт", "infowatch", "zecurion", "group-ib", "группа иб",
    "контур", "диалог", "иртея", "кода", "sofline",
]


# ══════════════════════════════════════════════════════════════════════════════
# ЧТЕНИЕ ФАЙЛОВ
# ══════════════════════════════════════════════════════════════════════════════

def read_xlsx(path: str) -> tuple[list[str], list[dict]]:
    """Возвращает (columns, rows)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None)
    if not header:
        raise ValueError("Файл пустой")
    columns = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(header)]
    rows = []
    for row in rows_iter:
        obj = {}
        for i, val in enumerate(row):
            if i < len(columns):
                obj[columns[i]] = str(val).strip() if val is not None else ""
        if any(v for v in obj.values()):
            rows.append(obj)
    return columns, rows


def read_html(path: str) -> tuple[list[str], list[dict]]:
    """Находит самую большую таблицу в HTML."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
    tables = soup.find_all("table")
    if not tables:
        raise ValueError("Таблица не найдена в HTML")
    table = max(tables, key=lambda t: len(t.find_all("tr")))
    rows_tags = table.find_all("tr")
    if not rows_tags:
        raise ValueError("Таблица пустая")
    columns = [th.get_text(strip=True) for th in rows_tags[0].find_all(["th", "td"])]
    rows = []
    for tr in rows_tags[1:]:
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        if not any(cells):
            continue
        obj = {}
        for i, val in enumerate(cells):
            if i < len(columns):
                obj[columns[i]] = val
        if any(v for v in obj.values()):
            rows.append(obj)
    return columns, rows


# ══════════════════════════════════════════════════════════════════════════════
# НАСТРОЙКА СТОЛБЦОВ (интерактивно)
# ══════════════════════════════════════════════════════════════════════════════

ROLE_NAMES = [
    "Главная связь",
    "Процесс",
    "Наименование",
    "Сокращение",
    "Базовое решение",
    "Вендор",
    "Угрозы",
    "Дополнительно",
]


def auto_guess(role: str, col: str) -> bool:
    r, c = role.lower(), col.lower()
    if "связь" in r or "главн" in r:   return "ответ" in c or "владел" in c
    if "процесс" in r:                  return "процесс" in c
    if "наимен" in r:                   return "наимен" in c and "сокращ" not in c and "базов" not in c
    if "сокращ" in r:                   return "сокращ" in c
    if "базов" in r:                    return "базов" in c
    if "вендор" in r:                   return any(x in c for x in ["вендор", "произв", "поставщ", "vendor"])
    if "угроз" in r:                    return "угроз" in c
    return False


def ask_mapping(columns: list[str]) -> dict[str, str | None]:
    """Интерактивный CLI для назначения ролей столбцам."""
    print("\n" + "═" * 60)
    print("  НАСТРОЙКА СТОЛБЦОВ")
    print("═" * 60)
    print("Найдены столбцы:")
    for i, col in enumerate(columns):
        print(f"  [{i:2d}] {col}")
    print()

    mapping = {}
    for role in ROLE_NAMES:
        # Автоматическая подсказка
        guess = next((c for c in columns if auto_guess(role, c)), None)
        hint = f"  (авто: «{guess}»)" if guess else ""

        print(f"  Роль «{role}»{hint}")
        print(f"  Введите номер или часть названия столбца, Enter = пропустить: ", end="")
        raw = input().strip()

        if not raw:
            mapping[role] = guess  # принять авто или None
        else:
            # Попытка по номеру
            if raw.isdigit() and 0 <= int(raw) < len(columns):
                mapping[role] = columns[int(raw)]
            else:
                # Поиск по подстроке
                matches = [c for c in columns if raw.lower() in c.lower()]
                if len(matches) == 1:
                    mapping[role] = matches[0]
                elif len(matches) > 1:
                    print(f"  Несколько совпадений: {matches}")
                    print(f"  Уточните номер: ", end="")
                    idx = input().strip()
                    mapping[role] = columns[int(idx)] if idx.isdigit() and 0 <= int(idx) < len(columns) else None
                else:
                    print(f"  Не найдено, пропускаю.")
                    mapping[role] = None

        if mapping[role]:
            print(f"  ✓ {role} → «{mapping[role]}»")
        else:
            print(f"  — {role} → не используется")
        print()

    return mapping


def ask_rf_keywords() -> list[str]:
    print("═" * 60)
    print("  КЛЮЧЕВЫЕ СЛОВА ДЛЯ РФ-ВЕНДОРОВ")
    print("═" * 60)
    print(f"По умолчанию {len(RF_KEYWORDS)} слов (kaspersky, solar, positive...)")
    print("Добавить свои? Введите через запятую или Enter:")
    extra = input().strip()
    kws = list(RF_KEYWORDS)
    if extra:
        kws += [k.strip().lower() for k in extra.split(",") if k.strip()]
    return kws


# ══════════════════════════════════════════════════════════════════════════════
# ПОСТРОЕНИЕ ГРАФА
# ══════════════════════════════════════════════════════════════════════════════

def build_graph(rows: list[dict], mapping: dict, rf_keywords: list[str]):
    active_roles = [r for r in ROLE_NAMES if mapping.get(r)]
    nodes_map = {}
    links_set = set()
    links = []

    def ensure_node(role_idx, value):
        val = value.strip()
        if not val or val in ("-", "—", "None", "nan"):
            return None
        node_id = f"{role_idx}\u00A7{val}"
        if node_id not in nodes_map:
            role = ROLE_NAMES[role_idx]
            is_rf = False
            if role == "Вендор":
                vl = val.lower()
                is_rf = any(kw in vl for kw in rf_keywords)
            nodes_map[node_id] = {
                "id": node_id,
                "roleIdx": role_idx,
                "role": role,
                "label": val,
                "isRF": is_rf,
                "color": ROLE_COLORS.get(role_idx, "#888"),
            }
        return node_id

    for row in rows:
        prev_ids = []
        for role in active_roles:
            ri = ROLE_NAMES.index(role)
            col = mapping[role]
            raw = row.get(col, "")
            if not raw:
                continue
            parts = [p.strip() for p in raw.replace(";", ",").replace("\n", ",").split(",") if p.strip()]
            cur_ids = [nid for nid in (ensure_node(ri, p) for p in parts) if nid]
            for cid in cur_ids:
                for pid in prev_ids:
                    key = f"{pid}\u2192{cid}"
                    if key not in links_set:
                        links_set.add(key)
                        links.append({"source": pid, "target": cid})
            prev_ids = cur_ids

    nodes = list(nodes_map.values())
    return {"nodes": nodes, "links": links, "activeRoles": active_roles}


def compute_vendor_stats(graph_data: dict) -> dict:
    vendor_nodes = [n for n in graph_data["nodes"] if n["role"] == "Вендор"]
    if not vendor_nodes:
        return {"total": 0, "rf": 0, "foreign": 0, "pct": 0}
    rf = sum(1 for n in vendor_nodes if n["isRF"])
    total = len(vendor_nodes)
    return {
        "total": total,
        "rf": rf,
        "foreign": total - rf,
        "pct": round(rf / total * 100),
        "rf_list": [n["label"] for n in vendor_nodes if n["isRF"]],
        "foreign_list": [n["label"] for n in vendor_nodes if not n["isRF"]],
    }


# ══════════════════════════════════════════════════════════════════════════════
# ГЕНЕРАЦИЯ HTML
# ══════════════════════════════════════════════════════════════════════════════

def load_d3() -> str:
    """Ищет d3.min.js рядом со скриптом или в стандартных путях."""
    candidates = [
        Path(__file__).parent / "d3.min.js",
        Path("/home/claude/.npm-global/lib/node_modules/@mermaid-js/mermaid-cli/node_modules/d3/dist/d3.min.js"),
    ]
    for p in candidates:
        if p.exists():
            return p.read_text(encoding="utf-8")
    raise FileNotFoundError(
        "d3.min.js не найден. Скачайте его и положите рядом со скриптом:\n"
        "https://cdnjs.cloudflare.com/ajax/libs/d3/7.8.5/d3.min.js"
    )


def generate_html(graph_data: dict, vendor_stats: dict, title: str) -> str:
    d3_code = load_d3()
    nodes_json = json.dumps(graph_data["nodes"], ensure_ascii=False)
    links_json = json.dumps(graph_data["links"], ensure_ascii=False)
    roles_json = json.dumps(graph_data["activeRoles"], ensure_ascii=False)
    stats_json = json.dumps(vendor_stats, ensure_ascii=False)
    role_colors_json = json.dumps(ROLE_COLORS, ensure_ascii=False)
    role_names_json = json.dumps(ROLE_NAMES, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<script>{d3_code}</script>
<style>
:root {{
  --bg:#08090d; --s1:#0f1117; --s2:#161b26;
  --border:#1e2533; --border2:#2a3347;
  --text:#dce3f0; --muted:#4a5568; --muted2:#6b7898;
  --blue:#4f8ef7; --cyan:#00c9b1; --red:#ff4b6e;
  --amber:#ffb547; --purple:#9b6dff; --green:#32d98b; --pink:#ff6eb4;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:var(--bg); color:var(--text); font-family:'Consolas','Courier New',monospace;
       height:100vh; overflow:hidden; display:flex; flex-direction:column; }}

header {{ padding:0 20px; height:50px; border-bottom:1px solid var(--border);
          background:var(--s1); display:flex; align-items:center; gap:16px; flex-shrink:0; z-index:10; }}
.logo {{ font-family:'Segoe UI','Arial Black',sans-serif; font-weight:800; font-size:14px;
         color:var(--cyan); letter-spacing:1px; }}
.hdr-stat {{ font-size:10px; color:var(--muted2); padding:3px 10px;
             border:1px solid var(--border2); border-radius:20px; }}
.hdr-stat span {{ color:var(--cyan); }}
.hdr-title {{ font-size:11px; color:var(--muted); margin-left:8px; max-width:400px;
              white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}

.main {{ display:flex; flex:1; overflow:hidden; }}

/* sidebar */
.sidebar {{ width:290px; flex-shrink:0; background:var(--s1);
            border-right:1px solid var(--border); display:flex; flex-direction:column; overflow:hidden; }}
.stabs {{ display:flex; border-bottom:1px solid var(--border); }}
.stab {{ flex:1; padding:11px; font-size:10px; text-align:center; cursor:pointer;
         color:var(--muted2); border-bottom:2px solid transparent; transition:all .2s; letter-spacing:.5px; }}
.stab.active {{ color:var(--cyan); border-bottom-color:var(--cyan); }}
.spanel {{ display:none; flex:1; flex-direction:column; overflow:hidden; }}
.spanel.active {{ display:flex; }}

/* chart */
.chart-area {{ padding:18px; flex:1; overflow-y:auto; display:flex; flex-direction:column; gap:14px; }}
.chart-title {{ font-family:'Segoe UI','Arial Black',sans-serif; font-size:13px; font-weight:700; }}
.chart-sub {{ font-size:10px; color:var(--muted2); margin-top:-8px; }}
.donut-wrap {{ position:relative; width:170px; height:170px; margin:0 auto; }}
.donut-center {{ position:absolute; top:50%; left:50%; transform:translate(-50%,-50%);
                 text-align:center; pointer-events:none; }}
.donut-pct {{ font-family:'Segoe UI','Arial Black',sans-serif; font-size:30px; font-weight:800; line-height:1; }}
.donut-label {{ font-size:9px; color:var(--muted2); letter-spacing:1px; text-transform:uppercase; }}
.leg-row {{ display:flex; align-items:center; justify-content:space-between; font-size:11px; margin-bottom:6px; }}
.leg-left {{ display:flex; align-items:center; gap:8px; color:var(--muted2); }}
.leg-sw {{ width:10px; height:10px; border-radius:3px; flex-shrink:0; }}
.leg-cnt {{ font-size:10px; color:var(--muted); }}
.divider {{ height:1px; background:var(--border); margin:6px 0; }}
.bar-row {{ margin-bottom:9px; }}
.bar-label {{ display:flex; justify-content:space-between; font-size:10px; color:var(--muted2); margin-bottom:3px; }}
.bar-track {{ height:4px; background:var(--border2); border-radius:2px; }}
.bar-fill {{ height:100%; border-radius:2px; transition:width .6s; }}

/* legend panel */
.leg-area {{ padding:14px 18px; flex:1; overflow-y:auto; }}
.leg-sec {{ font-size:9px; color:var(--muted); letter-spacing:1.5px; text-transform:uppercase;
            margin-bottom:9px; margin-top:4px; }}
.lg-item {{ display:flex; align-items:center; gap:9px; margin-bottom:6px; font-size:11px; color:var(--muted2); }}
.lg-dot {{ width:10px; height:10px; border-radius:50%; flex-shrink:0; }}

/* graph */
.graph-main {{ flex:1; display:flex; flex-direction:column; overflow:hidden; position:relative; }}
.gtoolbar {{ padding:9px 14px; border-bottom:1px solid var(--border); display:flex;
             align-items:center; gap:8px; background:var(--s1); flex-shrink:0; }}
.lbtn {{ background:var(--s2); border:1px solid var(--border2); color:var(--muted2);
         padding:5px 12px; border-radius:6px; font-size:10px; font-family:inherit;
         cursor:pointer; transition:all .2s; letter-spacing:.5px; }}
.lbtn.active,.lbtn:hover {{ border-color:var(--cyan); color:var(--cyan); background:rgba(0,201,177,.06); }}
.tb-sep {{ width:1px; height:18px; background:var(--border2); }}
.ghint {{ font-size:10px; color:var(--muted); margin-left:auto; }}
.gcanvas {{ flex:1; position:relative; overflow:hidden; }}
svg.main {{ width:100%; height:100%; }}
.edge {{ fill:none; stroke:#1e2533; stroke-width:1.5; stroke-opacity:.65; }}
.node-g {{ cursor:pointer; }}
.tooltip {{ position:absolute; background:var(--s1); border:1px solid var(--border2);
            border-radius:8px; padding:9px 13px; font-size:11px; pointer-events:none;
            max-width:260px; z-index:100; display:none; line-height:1.7;
            box-shadow:0 12px 32px rgba(0,0,0,.7); }}
.tt-type {{ font-size:9px; letter-spacing:1.2px; text-transform:uppercase; margin-bottom:3px; }}
.tt-val {{ font-size:12px; font-weight:500; word-break:break-word; }}
.path-panel {{ position:absolute; bottom:14px; left:14px;
               background:rgba(15,17,23,.96); border:1px solid var(--border2);
               border-radius:10px; padding:14px 17px; max-width:230px; font-size:11px;
               line-height:1.8; backdrop-filter:blur(10px); display:none;
               box-shadow:0 8px 32px rgba(0,0,0,.6); }}
.zbtns {{ position:absolute; bottom:14px; right:14px; display:flex; flex-direction:column; gap:5px; }}
.zbtn {{ width:32px; height:32px; background:var(--s1); border:1px solid var(--border2);
         color:var(--text); border-radius:7px; cursor:pointer; font-size:15px;
         display:flex; align-items:center; justify-content:center; transition:border-color .2s; }}
.zbtn:hover {{ border-color:var(--cyan); }}
::-webkit-scrollbar {{ width:4px; height:4px; }}
::-webkit-scrollbar-thumb {{ background:var(--border2); border-radius:2px; }}
</style>
</head>
<body>

<header>
  <div class="logo">ANALYSER</div>
  <div class="hdr-stat">Узлов: <span id="hNodes">0</span></div>
  <div class="hdr-stat">Связей: <span id="hLinks">0</span></div>
  <div class="hdr-stat">Строк: <span id="hRows">0</span></div>
  <div class="hdr-title" id="hTitle"></div>
</header>

<div class="main">
  <!-- Sidebar -->
  <div class="sidebar">
    <div class="stabs">
      <div class="stab active" onclick="switchTab('chart',this)">📊 Диаграмма</div>
      <div class="stab" onclick="switchTab('legend',this)">🏷 Легенда</div>
    </div>
    <div class="spanel active" id="tabChart">
      <div class="chart-area">
        <div class="chart-title">Вендоры по происхождению</div>
        <div class="chart-sub" id="vendorSub"></div>
        <div class="donut-wrap">
          <svg id="donutSvg" viewBox="0 0 180 180"></svg>
          <div class="donut-center">
            <div class="donut-pct" id="donutPct">—</div>
            <div class="donut-label">РФ</div>
          </div>
        </div>
        <div id="donutLegend"></div>
        <div class="divider"></div>
        <div class="chart-title" style="font-size:12px">По уровням</div>
        <div id="barChart"></div>
      </div>
    </div>
    <div class="spanel" id="tabLegend">
      <div class="leg-area" id="legArea"></div>
    </div>
  </div>

  <!-- Graph -->
  <div class="graph-main">
    <div class="gtoolbar">
      <button class="lbtn active" data-layout="hierarchy">▤ Иерархия</button>
      <button class="lbtn" data-layout="radial">◎ Радиальный</button>
      <button class="lbtn" data-layout="force">⊛ Force</button>
      <div class="tb-sep"></div>
      <span class="ghint">Клик — путь узла · фон — сбросить</span>
    </div>
    <div class="gcanvas" id="gcanvas">
      <svg class="main" id="gsvg">
        <defs>
          <marker id="arr" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto">
            <path d="M0,1 L0,5 L6,3 z" fill="#2a3347"/>
          </marker>
        </defs>
        <g id="zg"></g>
      </svg>
      <div class="tooltip" id="tooltip"></div>
      <div class="path-panel" id="pathPanel"></div>
      <div class="zbtns">
        <button class="zbtn" id="zIn">+</button>
        <button class="zbtn" id="zOut">−</button>
        <button class="zbtn" id="zReset">⌂</button>
      </div>
    </div>
  </div>
</div>

<script>
// ── data injected by Python ──────────────────────────────────────────────────
const NODES        = {nodes_json};
const LINKS        = {links_json};
const ACTIVE_ROLES = {roles_json};
const VENDOR_STATS = {stats_json};
const ROLE_COLORS  = {role_colors_json};
const ROLE_NAMES   = {role_names_json};
const FILE_TITLE   = {json.dumps(title, ensure_ascii=False)};

// ── init ─────────────────────────────────────────────────────────────────────
document.getElementById('hNodes').textContent  = NODES.length;
document.getElementById('hLinks').textContent  = LINKS.length;
document.getElementById('hRows').textContent   = NODES.length;
document.getElementById('hTitle').textContent  = FILE_TITLE;

const NODE_R = 13;

function getColor(n) {{
  if (n.role === 'Вендор') return n.isRF ? '#ff4b6e' : '#4f8ef7';
  return ROLE_COLORS[n.roleIdx] || '#888';
}}

// ── tabs ─────────────────────────────────────────────────────────────────────
function switchTab(name, el) {{
  document.querySelectorAll('.stab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.spanel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab' + name.charAt(0).toUpperCase() + name.slice(1)).classList.add('active');
}}

// ── donut chart ──────────────────────────────────────────────────────────────
(function buildDonut() {{
  const s = VENDOR_STATS;
  if (!s.total) {{ document.getElementById('vendorSub').textContent = 'Столбец Вендор не назначен'; return; }}
  document.getElementById('vendorSub').textContent = 'По полю «Вендор»';
  const pctEl = document.getElementById('donutPct');
  pctEl.textContent = s.pct + '%';
  pctEl.style.color = s.pct >= 60 ? '#ff4b6e' : s.pct >= 30 ? '#ffb547' : '#32d98b';

  const svg = d3.select('#donutSvg');
  const arc = d3.arc().innerRadius(52).outerRadius(76).cornerRadius(3).padAngle(.03);
  const pie = d3.pie().value(d => d.v).sort(null);
  const data = [{{ label:'Российские', v:s.rf, c:'#ff4b6e' }}, {{ label:'Зарубежные', v:s.foreign, c:'#4f8ef7' }}];
  svg.append('g').attr('transform','translate(90,90)').selectAll('path').data(pie(data)).join('path')
    .attr('d', arc).attr('fill', d => d.data.c).attr('opacity',.85);

  document.getElementById('donutLegend').innerHTML = data.map(d =>
    `<div class="leg-row"><div class="leg-left"><div class="leg-sw" style="background:${{d.c}}"></div>${{d.label}}</div>
     <span class="leg-cnt">${{d.v}} (${{s.total ? Math.round(d.v/s.total*100) : 0}}%)</span></div>`
  ).join('') + `<div class="divider"></div>
    <div class="leg-row"><div class="leg-left" style="color:var(--muted2)">Всего</div>
    <span class="leg-cnt">${{s.total}}</span></div>`;
}})();

// ── bar chart ────────────────────────────────────────────────────────────────
(function buildBars() {{
  const counts = {{}};
  NODES.forEach(n => {{ counts[n.role] = (counts[n.role]||0)+1; }});
  const max = Math.max(...Object.values(counts), 1);
  document.getElementById('barChart').innerHTML = ACTIVE_ROLES.map(role => {{
    const ri = ROLE_NAMES.indexOf(role);
    const c = counts[role] || 0;
    const pct = Math.round(c/max*100);
    const col = ROLE_COLORS[ri] || '#888';
    return `<div class="bar-row">
      <div class="bar-label"><span>${{role}}</span><span style="color:${{col}}">${{c}}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${{pct}}%;background:${{col}}"></div></div>
    </div>`;
  }}).join('');
}})();

// ── legend ───────────────────────────────────────────────────────────────────
(function buildLegend() {{
  const el = document.getElementById('legArea');
  el.innerHTML = '<div class="leg-sec">Типы узлов</div>' +
    ACTIVE_ROLES.map(role => {{
      const ri = ROLE_NAMES.indexOf(role);
      return `<div class="lg-item"><div class="lg-dot" style="background:${{ROLE_COLORS[ri]}}"></div>${{role}}</div>`;
    }}).join('') +
    `<div class="divider" style="margin:12px 0"></div>
     <div class="leg-sec">Управление</div>
     <div class="lg-item" style="font-size:10px;line-height:1.9;color:var(--muted2)">
       🖱 Наведение — предпросмотр связей<br>
       🖱 Клик — полный путь узла<br>
       🖱 Повтор. клик / фон — сброс
     </div>`;
}})();

// ── layout ───────────────────────────────────────────────────────────────────
document.querySelector('.gtoolbar').addEventListener('click', e => {{
  const btn = e.target.closest('[data-layout]');
  if (!btn) return;
  document.querySelectorAll('.lbtn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  render(btn.dataset.layout);
}});

function render(type) {{
  const c = document.getElementById('gcanvas');
  const W = c.clientWidth, H = c.clientHeight;
  if (type === 'hierarchy') renderHierarchy(W, H);
  else if (type === 'radial') renderRadial(W, H);
  else renderForce(W, H);
}}

const orderedRoles = () => ACTIVE_ROLES.slice().sort((a,b) => ROLE_NAMES.indexOf(a)-ROLE_NAMES.indexOf(b));

function renderHierarchy(W, H) {{
  const roles = orderedRoles();
  const nodes = NODES.map(n => ({{...n}}));
  const colW = W / roles.length;
  roles.forEach((role, ci) => {{
    const rn = nodes.filter(n => n.role === role);
    rn.forEach((n,i) => {{
      n.x = colW*ci + colW/2;
      n.y = H*.08 + Math.min(H*.84/Math.max(rn.length,1),70)*i + Math.min(H*.84/Math.max(rn.length,1),70)/2;
    }});
  }});
  drawBase(nodes, roles, W, H, (container) => {{
    roles.forEach((role,ci) => {{
      const ri = ROLE_NAMES.indexOf(role);
      const cx = colW*ci+colW/2;
      container.append('line').attr('x1',cx).attr('x2',cx).attr('y1',0).attr('y2',H)
        .attr('stroke',ROLE_COLORS[ri]).attr('stroke-opacity',.08).attr('stroke-dasharray','4 7').attr('stroke-width',1);
      container.append('text').attr('x',cx).attr('y',20).attr('text-anchor','middle')
        .attr('fill',ROLE_COLORS[ri]).attr('font-size','9px').attr('opacity',.45).text(role.toUpperCase());
    }});
  }});
}}

function renderRadial(W, H) {{
  const roles = orderedRoles();
  const nodes = NODES.map(n => ({{...n}}));
  const cx=W/2, cy=H/2, maxR=Math.min(W,H)*.43, step=maxR/roles.length;
  roles.forEach((role,ci) => {{
    const r = step*(ci+.7);
    const rn = nodes.filter(n => n.role === role);
    rn.forEach((n,i) => {{
      const a = rn.length===1 ? -Math.PI/2 : i/rn.length*2*Math.PI - Math.PI/2;
      n.x = cx + r*Math.cos(a); n.y = cy + r*Math.sin(a);
    }});
  }});
  drawBase(nodes, roles, W, H, (container) => {{
    roles.forEach((role,ci) => {{
      const ri = ROLE_NAMES.indexOf(role);
      container.append('circle').attr('cx',cx).attr('cy',cy).attr('r',step*(ci+.7))
        .attr('fill','none').attr('stroke',ROLE_COLORS[ri]).attr('stroke-opacity',.08).attr('stroke-dasharray','4 7');
    }});
  }});
}}

function renderForce(W, H) {{
  const roles = orderedRoles();
  const nodes = NODES.map(n => ({{...n}}));
  nodes.forEach(n => {{
    const ci = roles.indexOf(n.role);
    n.x = W/roles.length*ci + W/roles.length/2;
    n.y = H/2 + (Math.random()-.5)*180;
  }});
  const nb = new Map(nodes.map(n=>[n.id,n]));
  const links = LINKS.map(l => ({{source:nb.get(l.source)||l.source, target:nb.get(l.target)||l.target}}));

  const svg = d3.select('#gsvg'), container = d3.select('#zg');
  container.selectAll('*').remove();

  const edgeEls = container.append('g').selectAll('path').data(links).join('path')
    .attr('class','edge').attr('marker-end','url(#arr)');
  const nodeEls = container.append('g').selectAll('g').data(nodes).join('g').attr('class','node-g')
    .call(d3.drag()
      .on('start',(e,d)=>{{if(!e.active)sim.alphaTarget(.3).restart();d.fx=d.x;d.fy=d.y;}})
      .on('drag', (e,d)=>{{d.fx=e.x;d.fy=e.y;}})
      .on('end',  (e,d)=>{{if(!e.active)sim.alphaTarget(0);d.fx=null;d.fy=null;}}));
  appendVisuals(nodeEls);

  const sim = d3.forceSimulation(nodes)
    .force('link', d3.forceLink(links).id(d=>d.id).distance(110).strength(.5))
    .force('charge', d3.forceManyBody().strength(-280))
    .force('x', d3.forceX(d=>W/roles.length*roles.indexOf(d.role)+W/roles.length/2).strength(.45))
    .force('y', d3.forceY(H/2).strength(.08))
    .force('collide', d3.forceCollide(NODE_R+24));
  sim.on('tick', () => {{
    edgeEls.attr('d', d=>ePath(d.source,d.target));
    nodeEls.attr('transform', d=>`translate(${{d.x}},${{d.y}})`);
  }});
  setupInteract(nodeEls, edgeEls, links, svg);
  setupZoom(svg);
}}

function drawBase(nodes, roles, W, H, drawBg) {{
  const nb = new Map(nodes.map(n=>[n.id,n]));
  const links = LINKS.map(l => ({{...l, source:nb.get(l.source)||nb.get(l.source?.id), target:nb.get(l.target)||nb.get(l.target?.id)}})).filter(l=>l.source&&l.target);
  const svg = d3.select('#gsvg'), container = d3.select('#zg');
  container.selectAll('*').remove();
  drawBg(container);
  const edgeEls = container.append('g').selectAll('path').data(links).join('path')
    .attr('class','edge').attr('marker-end','url(#arr)').attr('d',d=>ePath(d.source,d.target));
  const nodeEls = container.append('g').selectAll('g').data(nodes).join('g')
    .attr('class','node-g').attr('transform',d=>`translate(${{d.x}},${{d.y}})`);
  appendVisuals(nodeEls);
  setupInteract(nodeEls, edgeEls, links, svg);
  setupZoom(svg);
}}

function appendVisuals(nodeEls) {{
  nodeEls.append('circle').attr('r',NODE_R)
    .attr('fill',d=>getColor(d)+'22').attr('stroke',d=>getColor(d)).attr('stroke-width',1.5);
  nodeEls.filter(d=>d.role==='Вендор'&&d.isRF).append('text')
    .attr('x',8).attr('y',-8).attr('font-size','9px').attr('text-anchor','middle').text('🇷🇺');
  nodeEls.append('text').attr('y',NODE_R+12).attr('text-anchor','middle')
    .attr('fill',d=>getColor(d)).attr('font-size','10px')
    .text(d=>d.label.length>20?d.label.slice(0,19)+'…':d.label);
}}

function ePath(s,t) {{
  const dx=t.x-s.x, dy=t.y-s.y, dist=Math.sqrt(dx*dx+dy*dy)||1;
  const off=NODE_R+2, nx=dx/dist*off, ny=dy/dist*off;
  const bend=Math.min(35,dist*.12);
  const bx=(s.x+t.x)/2+-dy/dist*bend, by=(s.y+t.y)/2+dx/dist*bend;
  return `M${{s.x+nx}},${{s.y+ny}} Q${{bx}},${{by}} ${{t.x-nx*3}},${{t.y-ny*3}}`;
}}

// ── traverse ─────────────────────────────────────────────────────────────────
function traverse(id, links, dir) {{
  const vis=new Set(), q=[id];
  while(q.length){{
    const cur=q.shift(); if(vis.has(cur)) continue; vis.add(cur);
    links.forEach(l=>{{ const from=dir==='fwd'?l.source.id:l.target.id, to=dir==='fwd'?l.target.id:l.source.id;
      if(from===cur&&!vis.has(to)) q.push(to); }});
  }}
  return vis;
}}

// ── interactions ─────────────────────────────────────────────────────────────
function setupInteract(nodeEls, edgeEls, links, svg) {{
  const tt=document.getElementById('tooltip'), pp=document.getElementById('pathPanel');
  let pinned=null;

  function hl(id) {{
    if(!id) return rst();
    const anc=traverse(id,links,'bwd'), des=traverse(id,links,'fwd');
    const path=new Set([...anc,...des]);
    const pek=new Set(links.filter(l=>path.has(l.source.id)&&path.has(l.target.id)).map(l=>l.source.id+'>'+l.target.id));
    edgeEls.attr('stroke-opacity',l=>pek.has(l.source.id+'>'+l.target.id)?.95:.03)
      .attr('stroke',l=>pek.has(l.source.id+'>'+l.target.id)?getColor(l.source):'#1e2533')
      .attr('stroke-width',l=>pek.has(l.source.id+'>'+l.target.id)?2.5:1);
    nodeEls.attr('opacity',n=>path.has(n.id)?1:.07)
      .select('circle').attr('stroke-width',n=>n.id===id?3:1.5)
      .attr('filter',n=>n.id===id?`drop-shadow(0 0 7px ${{getColor(n)}})`:'none');
  }}

  function rst() {{
    edgeEls.attr('stroke-opacity',.65).attr('stroke','#1e2533').attr('stroke-width',1.5);
    nodeEls.attr('opacity',1).select('circle').attr('stroke-width',1.5).attr('filter','none');
  }}

  nodeEls.on('mouseover',(e,d)=>{{
    tt.style.display='block';
    tt.innerHTML=`<div class="tt-type" style="color:${{getColor(d)}}">${{d.role}}${{d.role==='Вендор'?(d.isRF?' · 🇷🇺 РФ':' · зарубежный'):''}}</div><div class="tt-val">${{d.label}}</div>`;
    if(!pinned) hl(d.id);
  }})
  .on('mousemove',e=>{{
    const r=document.getElementById('gcanvas').getBoundingClientRect();
    tt.style.left=(e.clientX-r.left+14)+'px'; tt.style.top=(e.clientY-r.top-12)+'px';
  }})
  .on('mouseout',()=>{{ tt.style.display='none'; if(!pinned) rst(); }})
  .on('click',(e,d)=>{{
    e.stopPropagation();
    if(pinned===d.id){{ pinned=null; rst(); pp.style.display='none'; return; }}
    pinned=d.id; hl(d.id);
    const anc=[...traverse(d.id,links,'bwd')].filter(x=>x!==d.id);
    const des=[...traverse(d.id,links,'fwd')].filter(x=>x!==d.id);
    const nm=new Map(); nodeEls.each(n=>nm.set(n.id,n));
    const fmt=ids=>ids.length?ids.map(id=>{{const n=nm.get(id);return n?`<span style="color:${{getColor(n)}}">${{n.label.length>24?n.label.slice(0,23)+'…':n.label}}</span>`:''}}).join('<br>'):'<span style="color:var(--muted)">—</span>';
    pp.style.display='block';
    pp.innerHTML=`<div style="font-size:9px;color:var(--muted);letter-spacing:1.5px;text-transform:uppercase;margin-bottom:7px">Путь</div>
      <div style="color:${{getColor(d)}};font-weight:600;font-size:12px;margin-bottom:9px">${{d.label}}</div>
      <div style="font-size:9px;color:var(--muted);margin-bottom:3px">↑ ПРЕДКИ (${{anc.length}})</div>
      <div style="margin-bottom:9px;font-size:11px;line-height:1.7">${{fmt(anc)}}</div>
      <div style="font-size:9px;color:var(--muted);margin-bottom:3px">↓ ПОТОМКИ (${{des.length}})</div>
      <div style="font-size:11px;line-height:1.7">${{fmt(des)}}</div>
      <div style="color:var(--muted);font-size:9px;margin-top:9px">Ещё раз — снять</div>`;
  }});
  svg.on('click',()=>{{ if(pinned){{ pinned=null; rst(); pp.style.display='none'; }} }});
}}

function setupZoom(svg) {{
  const zoom = d3.zoom().scaleExtent([.04,8]).on('zoom',e=>d3.select('#zg').attr('transform',e.transform));
  svg.call(zoom);
  document.getElementById('zIn').onclick    = ()=>svg.transition().call(zoom.scaleBy,1.4);
  document.getElementById('zOut').onclick   = ()=>svg.transition().call(zoom.scaleBy,.72);
  document.getElementById('zReset').onclick = ()=>svg.transition().call(zoom.transform,d3.zoomIdentity);
}}

// start
render('hierarchy');
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Анализатор данных → граф + диаграмма")
    parser.add_argument("file", nargs="?", help="Путь к .xlsx или .html файлу")
    parser.add_argument("--output", "-o", default="", help="Куда сохранить report.html (по умолчанию рядом с входным файлом)")
    parser.add_argument("--no-open", action="store_true", help="Не открывать браузер автоматически")
    args = parser.parse_args()

    # ── выбор файла ────────────────────────────────────────────────────────────
    filepath = args.file
    if not filepath:
        print("=" * 60)
        print("  ANALYSER  —  Excel / HTML → Граф + Диаграмма")
        print("=" * 60)
        print("Введите путь к файлу (.xlsx или .html):")
        filepath = input("  > ").strip().strip('"').strip("'")

    if not os.path.exists(filepath):
        print(f"Файл не найден: {filepath}")
        sys.exit(1)

    ext = Path(filepath).suffix.lower()

    # ── чтение ─────────────────────────────────────────────────────────────────
    print(f"\nЧитаю файл: {filepath}")
    try:
        if ext in (".xlsx", ".xls"):
            columns, rows = read_xlsx(filepath)
            print(f"  Excel: {len(rows)} строк, {len(columns)} столбцов")
        else:
            columns, rows = read_html(filepath)
            print(f"  HTML: {len(rows)} строк, {len(columns)} столбцов")
    except Exception as e:
        print(f"Ошибка чтения: {e}")
        sys.exit(1)

    # ── маппинг ────────────────────────────────────────────────────────────────
    mapping = ask_mapping(columns)
    active = [r for r in ROLE_NAMES if mapping.get(r)]
    if len(active) < 2:
        print("Нужно выбрать хотя бы 2 столбца.")
        sys.exit(1)

    rf_keywords = ask_rf_keywords()

    # ── построение ─────────────────────────────────────────────────────────────
    print("\nСтроим граф...")
    graph_data = build_graph(rows, mapping, rf_keywords)
    vendor_stats = compute_vendor_stats(graph_data)

    n_nodes = len(graph_data["nodes"])
    n_links = len(graph_data["links"])
    print(f"  Узлов: {n_nodes}, связей: {n_links}")

    if vendor_stats["total"]:
        print(f"  Вендоры РФ: {vendor_stats['rf']}/{vendor_stats['total']} ({vendor_stats['pct']}%)")

    # ── генерация HTML ─────────────────────────────────────────────────────────
    title = Path(filepath).stem
    html = generate_html(graph_data, vendor_stats, title)

    out_path = args.output
    if not out_path:
        out_path = str(Path(filepath).parent / (Path(filepath).stem + "_graph.html"))

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(out_path) // 1024
    print(f"\n✓ Сохранено: {out_path}  ({size_kb} KB)")

    if not args.no_open:
        print("  Открываю в браузере...")
        webbrowser.open(f"file://{os.path.abspath(out_path)}")

    print("\nГотово!")


if __name__ == "__main__":
    main()
