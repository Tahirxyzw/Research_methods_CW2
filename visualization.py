"""
NHS Hospital Admissions Parallel Coordinates Dashboard
Reserch Methods CourseWork 2 
"""

import os, re, glob, warnings, json
import pandas as pd
warnings.filterwarnings("ignore")

DATA_DIR = "primary_summary"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ICD-10 chapter mapping based on first 1-2 characters of diagnosis code
CHAPTER_MAP = {
    "A": "Infectious & Parasitic", "B": "Infectious & Parasitic",
    "C": "Neoplasms", "D0": "Neoplasms", "D1": "Neoplasms",
    "D2": "Neoplasms", "D3": "Neoplasms", "D4": "Neoplasms",
    "D5": "Blood Diseases", "D6": "Blood Diseases",
    "D7": "Blood Diseases", "D8": "Blood Diseases",
    "E": "Endocrine & Metabolic", "F": "Mental & Behavioural",
    "G": "Nervous System",
    "H0": "Eye Diseases", "H1": "Eye Diseases", "H2": "Eye Diseases",
    "H3": "Eye Diseases", "H4": "Eye Diseases", "H5": "Eye Diseases",
    "H6": "Ear Diseases", "H7": "Ear Diseases",
    "H8": "Ear Diseases", "H9": "Ear Diseases",
    "I": "Circulatory System", "J": "Respiratory System",
    "K": "Digestive System", "L": "Skin Diseases",
    "M": "Musculoskeletal", "N": "Genitourinary",
    "O": "Pregnancy & Childbirth", "P": "Perinatal Conditions",
    "Q": "Congenital Malformations", "R": "Symptoms & Signs",
    "S": "Injury & Trauma", "T": "Injury & Trauma",
    "U": "Special Codes", "Z": "Health Services Contact",
}


def get_chapter(code):
    """mapping ICD-10 code to its chapter name using first 1-2 characters"""
    cleaned = code.strip().lstrip("‡† ").strip()
    if not cleaned:
        return "Other"
    two_char = cleaned[:2]
    one_char = cleaned[0]
    return CHAPTER_MAP.get(two_char, CHAPTER_MAP.get(one_char, "Other"))


def extract_year(filename):
    """extracting year string from filename """
    base = os.path.basename(filename)
    match = re.search(r'(\d{4})-(\d{2,4})', base)
    if match:
        return f"{match.group(1)}-{match.group(2)[-2:]}"
    match = re.search(r'(\d{2})-(\d{2})', base)
    if match:
        prefix = "19" if int(match.group(1)) >= 90 else "20"
        return f"{prefix}{match.group(1)}-{match.group(2)}"
    return base


def safe_float(value):
    """converting value to float"""
    try:
        result = float(value)
        return result if result >= 0 else float("nan")
    except (ValueError, TypeError):
        return float("nan")


def find_column(headers, keyword_sets):
    """finding column index by trying multiple keyword sets against header names"""
    for keywords in keyword_sets:
        for i, header in enumerate(headers):
            if all(k.lower() in header.lower() for k in keywords):
                return i
    return None


def read_file(filepath):
    """reading one NHS HES file and returning a standardised DataFrame"""
    extension = os.path.splitext(filepath)[1].lower()
    year = extract_year(filepath)

    # reading the sheet name 
    try:
        if extension == ".xlsx":
            excel = pd.ExcelFile(filepath)
            known_names = [
                "Primary diagnosis - summary",
                "Primary Diagnosis Summary",
                "Primary diagnosis summary",
            ]
            sheet = next(
                (s for name in known_names for s in excel.sheet_names if s == name),
                None,
            )
            if not sheet:
                sheet = next(
                    (s for s in excel.sheet_names if "summary" in s.lower()),
                    excel.sheet_names[0],
                )
            df = pd.read_excel(filepath, sheet_name=sheet, header=None)
        else:
            excel = pd.ExcelFile(filepath)
            df = pd.read_excel(filepath, sheet_name=excel.sheet_names[0], header=None)
    except Exception as err:
        print(f"  ERROR: {filepath}: {err}")
        return pd.DataFrame()

    # finding the header row by looking for known column names
    header_row = None
    for i in range(min(30, len(df))):
        row_text = " ".join(str(v) for v in df.iloc[i] if pd.notna(v))
        if any(term in row_text for term in
               ["Admissions", "Finished Admission", "Finished Consultant", "FCE", "FAE"]):
            header_row = i
            break
    if header_row is None:
        print(f" WARNING: No header found in {filepath}")
        return pd.DataFrame()

    # cleaning headers and mapping column positions
    headers = [
        str(h).replace('\n', ' ').strip() if pd.notna(h) else ''
        for h in df.iloc[header_row]
    ]
    columns = {
        "fce": find_column(headers, [["Finished", "Consultant"], ["FCE"]]),
        "admissions": find_column(headers, [["Admission"], ["Admissions"]]),
        "male": find_column(headers, [["Male"]]),
        "emergency": find_column(headers, [["Emergency"]]),
        "mean_los": find_column(headers, [["Mean", "length"], ["Mean", "stay"]]),
        "mean_age": find_column(headers, [["Mean", "Age"]]),
        "bed_days": find_column(headers, [["Bed", "Day"], ["FCE", "bed"]]),
    }

    # detecting whether code and description are in one column or two
    first_row = header_row + 1
    while first_row < len(df) and pd.isna(df.iloc[first_row, 0]):
        first_row += 1
    test_val = df.iloc[first_row + 1, 1] if first_row + 1 < len(df) else None
    try:
        float(test_val)
        combined_format = True
    except (ValueError, TypeError):
        combined_format = not (
            isinstance(test_val, str)
            and not test_val.replace(",", "").replace(".", "").isdigit()
        )

    # parsing each data row
    records = []
    stop_words = ["Copyright", "Source:", "Responsible", "Contact",
                  "enquiries", "4-char", "Note:"]
    code_pattern = re.compile(
        r'^[‡†\s]*((?:[A-Z]\d{2}(?:\.\d)?(?:-[A-Z]\d{2}(?:\.\d)?)?(?:,\s*)?)+)\s+(.*)'
    )

    for i in range(first_row, len(df)):
        cell = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""
        if not cell or cell == "nan":
            continue
        if any(s in cell for s in stop_words):
            break
        if cell.startswith("Total"):
            continue

        # parsing code and description
        if combined_format:
            match = code_pattern.match(cell)
            if not match:
                continue
            code = match.group(1).strip()
            description = match.group(2).strip()
        else:
            code = cell.lstrip("‡† ").strip()
            description = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
            if not re.match(r'^[A-Z]', code):
                continue

        # extracting numeric values
        def get_value(key):
            idx = columns.get(key)
            return safe_float(df.iloc[i, idx]) if idx is not None else float("nan")

        admissions = get_value("admissions")
        emergency = get_value("emergency")
        if pd.isna(admissions) or admissions == 0:
            continue

        start_year_match = re.match(r'(\d{4})', year)
        start_year = int(start_year_match.group(1)) if start_year_match else 0

        emergency_pct = (
            round(emergency / admissions * 100, 1)
            if admissions > 0 and not pd.isna(emergency) else float("nan")
        )

        records.append({
            "year": year,
            "start_year": start_year,
            "code": code,
            "description": description[:60],
            "chapter": get_chapter(code),
            "fce": get_value("fce"),
            "admissions": admissions,
            "emergency": emergency,
            "mean_los": get_value("mean_los"),
            "mean_age": get_value("mean_age"),
            "bed_days": get_value("bed_days"),
            "emergency_pct": emergency_pct,
        })

    return pd.DataFrame(records)


def load_all_data():
    """loading and merging all HES files from the data directory"""
    files = sorted(
        glob.glob(os.path.join(DATA_DIR, "*.xls"))
        + glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    )
    if not files:
        print(f"No files found in {DATA_DIR}")
        return pd.DataFrame()

    print(f"Found {len(files)} files\n")
    dataframes = []
    for filepath in files:
        print(f"  {os.path.basename(filepath)} → {extract_year(filepath)}", end="")
        result = read_file(filepath)
        if not result.empty:
            print(f"  ✓ {len(result)}")
            dataframes.append(result)
        else:
            print("  ✗")

    if not dataframes:
        return pd.DataFrame()
    merged = pd.concat(dataframes, ignore_index=True)
    return merged.sort_values(["start_year", "code"]).reset_index(drop=True)


def build_dashboard(df):
    """generating an interactive HTML dashboard from the processed data"""
    required_cols = ["admissions", "emergency_pct", "mean_age", "mean_los", "bed_days"]
    df = df.dropna(subset=required_cols).copy()

    # preparing individual level data
    display_cols = [
        "year", "start_year", "code", "description", "chapter",
        "admissions", "emergency", "emergency_pct",
        "mean_age", "mean_los", "bed_days", "fce",
    ]
    individual_data = df[display_cols].to_dict(orient="records")
    years = sorted(df["year"].unique().tolist())
    chapters = sorted(df["chapter"].unique().tolist())

    # preparing chapter level aggregated data
    chapter_agg = df.groupby(["year", "start_year", "chapter"]).agg(
        admissions=("admissions", "sum"),
        emergency=("emergency", "sum"),
        bed_days=("bed_days", "sum"),
        fce=("fce", "sum"),
        mean_los=("mean_los", "mean"),
        mean_age=("mean_age", "mean"),
    ).reset_index()
    chapter_agg["emergency_pct"] = (
        chapter_agg["emergency"] / chapter_agg["admissions"] * 100
    ).round(1)
    chapter_agg = chapter_agg.dropna(subset=required_cols)
    chapter_data = chapter_agg.to_dict(orient="records")

    # building HTML components
    year_checkboxes = "\n".join(
        f'            <label class="dropdown-item">'
        f'<input type="checkbox" class="year-checkbox" value="{y}" '
        f'checked onchange="onYearChange()">{y}</label>'
        for y in years
    )
    chapter_options = "\n".join(
        f'          <option value="{c}">{c}</option>' for c in chapters
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>NHS Hospital Admissions — Parallel Coordinates Explorer</title>
  <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}

    body {{
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: #f0f2f5;
      color: #1a1a2e;
    }}

    /* Header */
    .header {{
      background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
      color: #fff;
      padding: 20px 30px;
    }}
    .header h1 {{ font-size: 22px; font-weight: 600; }}
    .header p {{ font-size: 12px; color: #8a9cc5; margin-top: 4px; }}

    /* Layout */
    .layout {{ display: flex; min-height: calc(100vh - 80px); }}

    /* Sidebar */
    .sidebar {{
      width: 280px;
      background: #fff;
      padding: 20px;
      border-right: 1px solid #e0e4ea;
      flex-shrink: 0;
      overflow-y: auto;
    }}
    .sidebar h3 {{
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 1.5px;
      color: #666;
      margin-bottom: 10px;
      font-weight: 600;
    }}

    /* Filter groups */
    .filter-group {{ margin-bottom: 18px; }}
    .filter-group label {{
      display: block;
      font-size: 13px;
      font-weight: 500;
      color: #333;
      margin-bottom: 5px;
    }}
    .filter-group select,
    .filter-group input {{
      width: 100%;
      padding: 8px 10px;
      border: 1px solid #d0d5dd;
      border-radius: 6px;
      font-size: 13px;
      background: #fafbfc;
      color: #333;
    }}

    /* Stats cards */
    .stats-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-top: 12px;
    }}
    .stat-card {{
      background: #f8f9fb;
      border-radius: 8px;
      padding: 10px;
      border: 1px solid #e8ebf0;
    }}
    .stat-card .stat-label {{
      font-size: 10px;
      color: #888;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .stat-card .stat-value {{
      font-size: 18px;
      font-weight: 700;
      color: #1a1a2e;
      margin-top: 2px;
    }}

    /* Main content */
    .main-content {{ flex: 1; padding: 20px; overflow: hidden; }}
    .chart-card {{
      background: #fff;
      border-radius: 10px;
      padding: 20px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06);
      border: 1px solid #e8ebf0;
      height: calc(100vh - 120px);
    }}

    /* Controls */
    .controls-row {{
      display: flex;
      gap: 16px;
      margin-bottom: 16px;
      align-items: center;
      flex-wrap: wrap;
    }}
    .view-buttons {{ display: flex; gap: 4px; }}
    .view-btn {{
      padding: 7px 18px;
      font-size: 12px;
      border: 1px solid #d0d5dd;
      border-radius: 6px;
      cursor: pointer;
      background: #fff;
      color: #555;
      font-weight: 500;
    }}
    .view-btn.active {{
      background: #1a1a2e;
      color: #fff;
      border-color: #1a1a2e;
    }}
    .view-btn:hover:not(.active) {{ background: #f0f2f5; }}

    .colour-picker {{ display: flex; align-items: center; gap: 6px; }}
    .colour-picker label {{
      font-size: 12px;
      color: #666;
      white-space: nowrap;
      font-weight: 500;
    }}
    .colour-picker select {{
      padding: 7px 10px;
      border: 1px solid #d0d5dd;
      border-radius: 6px;
      font-size: 12px;
    }}

    /* Chart titles */
    .chart-title {{ font-size: 16px; font-weight: 600; color: #1a1a2e; margin-bottom: 2px; }}
    .chart-subtitle {{ font-size: 12px; color: #888; margin-bottom: 12px; }}

    /* Custom dropdown for year multi-select */
    .multi-dropdown {{ position: relative; }}
    .multi-dropdown-trigger {{
      width: 100%;
      padding: 8px 10px;
      border: 1px solid #d0d5dd;
      border-radius: 6px;
      font-size: 13px;
      background: #fafbfc;
      color: #333;
      cursor: pointer;
    }}
    .multi-dropdown-trigger:hover {{ border-color: #0f3460; }}
    .multi-dropdown-menu {{
      display: none;
      position: absolute;
      top: 100%;
      left: 0;
      right: 0;
      background: #fff;
      border: 1px solid #d0d5dd;
      border-radius: 6px;
      max-height: 250px;
      overflow-y: auto;
      z-index: 100;
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
      margin-top: 2px;
    }}
    .multi-dropdown-menu.open {{ display: block; }}
    .dropdown-item {{
      display: block;
      padding: 5px 10px;
      font-size: 12px;
      cursor: pointer;
      color: #333;
    }}
    .dropdown-item:hover {{ background: #f0f2f5; }}
    .dropdown-item input {{ margin-right: 6px; }}

    /* Axis guide */
    .axis-guide {{
      margin-top: 14px;
      padding: 12px;
      background: #f9f9f9;
      border-radius: 8px;
      border: 1px solid #eee;
    }}
    .axis-guide h4 {{ font-size: 11px; color: #444; margin-bottom: 6px; font-weight: 600; }}
    .axis-guide p {{ font-size: 10px; color: #666; margin-bottom: 3px; }}
    .axis-guide strong {{ color: #333; }}
  </style>
</head>
<body>

  <!-- Header -->
  <div class="header">
    <h1>NHS Hospital Admissions Parallel Coordinates Explorer</h1>
    <p>Hospital Episode Statistics · England · {years[0]} to {years[-1]} · Overview of Primary Diagnosis Summary</p>
  </div>

  <div class="layout">
    <!-- Sidebar -->
    <div class="sidebar">
      <h3>Filters</h3>

      <!-- Year multi-select dropdown -->
      <div class="filter-group">
        <label>Years</label>
        <div class="multi-dropdown" id="year-dropdown">
          <div class="multi-dropdown-trigger" onclick="toggleDropdown('year-dropdown')">
            <span id="year-label">All Years</span>
            <span style="float:right">▾</span>
          </div>
          <div class="multi-dropdown-menu">
            <div style="padding:4px 8px; display:flex; gap:6px; border-bottom:1px solid #eee; margin-bottom:4px">
              <button onclick="selectAllYears(true)"
                style="font-size:10px; padding:2px 8px; border:1px solid #d0d5dd;
                       border-radius:4px; background:#fff; cursor:pointer; color:#555">All</button>
              <button onclick="selectAllYears(false)"
                style="font-size:10px; padding:2px 8px; border:1px solid #d0d5dd;
                       border-radius:4px; background:#fff; cursor:pointer; color:#555">None</button>
            </div>
{year_checkboxes}
          </div>
        </div>
      </div>

      <!-- Chapter filter -->
      <div class="filter-group">
        <label>ICD-10 Chapter</label>
        <select id="chapter-filter">
          <option value="all">All Chapters</option>
{chapter_options}
        </select>
      </div>

      <!-- Admissions threshold slider -->
      <div class="filter-group">
        <label>Min Admissions</label>
        <input type="range" id="min-admissions" min="0" max="100000" step="1000" value="0">
        <div style="font-size:11px; color:#888; margin-top:4px">
          ≥ <strong id="min-adm-label">0</strong>
        </div>
      </div>

      <!-- Summary stats -->
      <div>
        <h3 style="margin-top:8px">Summary</h3>
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-label">Lines</div>
            <div class="stat-value" id="stat-lines">-</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">Admissions</div>
            <div class="stat-value" id="stat-admissions">-</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">Emerg %</div>
            <div class="stat-value" id="stat-emergency">-</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">Avg Stay</div>
            <div class="stat-value" id="stat-stay">-</div>
          </div>
        </div>
      </div>

      <!-- Axis reference -->
      <div class="axis-guide">
        <h4>Axes</h4>
        <p><strong>Chapter:</strong> ICD-10 body system</p>
        <p><strong>Year:</strong> NHS financial year</p>
        <p><strong>Admissions:</strong> Finished Admission Episodes</p>
        <p><strong>Emergency %:</strong> % unplanned admissions</p>
        <p><strong>Mean Age:</strong> Average patient age</p>
        <p><strong>Mean Stay:</strong> Avg days per patient</p>
        <p><strong>Bed Days:</strong> Admissions × Stay</p>
      </div>
    </div>

    <!-- Main chart area -->
    <div class="main-content">
      <div class="controls-row">
        <div class="view-buttons">
          <button class="view-btn active" onclick="switchView('chapters')">
            Chapter Overview (Grouped)
          </button>
          <button class="view-btn" onclick="switchView('categories')">
            Individual Diagnoses
          </button>
        </div>
        <div class="colour-picker">
          <label>Colour lines by:</label>
          <select id="colour-by" onchange="updateChart()">
            <option value="start_year">Year</option>
            <option value="emergency_pct">Emergency %</option>
            <option value="mean_age">Mean Age</option>
            <option value="mean_los">Mean Length of Stay</option>
            <option value="admissions">Admissions Volume</option>
          </select>
        </div>
      </div>

      <div class="chart-card">
        <div class="chart-title" id="chart-title">Chapter Overview</div>
        <div class="chart-subtitle" id="chart-subtitle"></div>
        <div id="chart" style="width:100%; height:calc(100% - 50px)"></div>
      </div>
    </div>
  </div>

  <script>
    // Data injected from Python
    var individualData = {json.dumps(individual_data, default=str)};
    var chapterData = {json.dumps(chapter_data, default=str)};
    var allChapters = {json.dumps(chapters)};
    var currentView = 'chapters';

    function formatNumber(n) {{
      if (n >= 1e6) return (n / 1e6).toFixed(1) + 'M';
      if (n >= 1e3) return (n / 1e3).toFixed(0) + 'K';
      return n.toFixed(0);
    }}

    // Dropdown toggle
    function toggleDropdown(id) {{
      document.querySelector('#' + id + ' .multi-dropdown-menu').classList.toggle('open');
    }}
    document.addEventListener('click', function(e) {{
      if (!e.target.closest('.multi-dropdown')) {{
        document.querySelectorAll('.multi-dropdown-menu').forEach(function(m) {{
          m.classList.remove('open');
        }});
      }}
    }});

    // Year selection
    function selectAllYears(checked) {{
      document.querySelectorAll('.year-checkbox').forEach(function(cb) {{
        cb.checked = checked;
      }});
      onYearChange();
    }}

    function onYearChange() {{
      var selected = getSelectedYears();
      var total = document.querySelectorAll('.year-checkbox').length;
      var label = document.getElementById('year-label');

      if (selected.length === 0) label.textContent = 'None';
      else if (selected.length === total) label.textContent = 'All Years (' + total + ')';
      else if (selected.length <= 3) label.textContent = selected.join(', ');
      else label.textContent = selected.length + ' years';

      updateChart();
    }}

    function getSelectedYears() {{
      var checkboxes = document.querySelectorAll('.year-checkbox:checked');
      return Array.from(checkboxes).map(function(cb) {{ return cb.value; }});
    }}

    // View switching
    function switchView(view) {{
      currentView = view;
      document.querySelectorAll('.view-btn').forEach(function(btn) {{
        btn.classList.remove('active');
      }});
      event.target.classList.add('active');
      updateChart();
    }}

    // Filtering
    function getFilteredData() {{
      var selectedYears = getSelectedYears();
      var chapter = document.getElementById('chapter-filter').value;
      var minAdm = parseInt(document.getElementById('min-admissions').value);
      var source = (currentView === 'chapters') ? chapterData : individualData;

      return source.filter(function(row) {{
        if (selectedYears.length > 0 && selectedYears.indexOf(row.year) === -1) return false;
        if (chapter !== 'all' && row.chapter !== chapter) return false;
        if (row.admissions < minAdm) return false;
        return true;
      }});
    }}

    // Updating summary stats
    function updateStats(data) {{
      var totalAdm = data.reduce(function(sum, r) {{ return sum + r.admissions; }}, 0);
      var avgEmerg = data.length
        ? (data.reduce(function(s, r) {{ return s + (r.emergency_pct || 0); }}, 0) / data.length).toFixed(1)
        : '0';
      var avgStay = data.length
        ? (data.reduce(function(s, r) {{ return s + (r.mean_los || 0); }}, 0) / data.length).toFixed(1)
        : '0';

      document.getElementById('stat-lines').textContent = data.length;
      document.getElementById('stat-admissions').textContent = formatNumber(totalAdm);
      document.getElementById('stat-emergency').textContent = avgEmerg + '%';
      document.getElementById('stat-stay').textContent = avgStay + 'd';
    }}

    // Main chart update
    function updateChart() {{
      var data = getFilteredData();
      var colourBy = document.getElementById('colour-by').value;
      updateStats(data);

      if (data.length === 0) {{
        Plotly.purge('chart');
        document.getElementById('chart-title').textContent = 'No data matches filters';
        return;
      }}

      // Building axes
      var activeChapters = Array.from(new Set(data.map(function(r) {{ return r.chapter; }}))).sort();
      var activeYears = Array.from(new Set(data.map(function(r) {{ return r.start_year; }}))).sort();

      var dimensions = [
        {{
          label: 'ICD-10 Chapter',
          values: data.map(function(r) {{ return activeChapters.indexOf(r.chapter); }}),
          tickvals: activeChapters.map(function(_, i) {{ return i; }}),
          ticktext: activeChapters
        }},
        {{
          label: 'Year',
          values: data.map(function(r) {{ return r.start_year; }}),
          tickvals: activeYears,
          ticktext: activeYears.map(String)
        }},
        {{ label: 'Admissions (FAE)', values: data.map(function(r) {{ return r.admissions; }}) }},
        {{ label: 'Emergency %', values: data.map(function(r) {{ return r.emergency_pct; }}) }},
        {{ label: 'Mean Age (Yrs)', values: data.map(function(r) {{ return r.mean_age; }}) }},
        {{ label: 'Mean Stay (Days)', values: data.map(function(r) {{ return r.mean_los; }}) }},
        {{ label: 'Bed Days', values: data.map(function(r) {{ return r.bed_days; }}) }}
      ];

      // Colour scales
      var colourScales = {{
        start_year: [[0, '#2166ac'], [0.5, '#67a9cf'], [1, '#d6604d']],
        emergency_pct: 'YlOrRd',
        mean_age: 'Blues',
        mean_los: 'Purples',
        admissions: 'Viridis'
      }};
      var colourLabels = {{
        start_year: 'Year',
        emergency_pct: 'Emergency %',
        mean_age: 'Mean Age',
        mean_los: 'Mean Stay',
        admissions: 'Admissions'
      }};

      // Build trace
      var trace = {{
        type: 'parcoords',
        line: {{
          color: data.map(function(r) {{ return r[colourBy]; }}),
          colorscale: colourScales[colourBy],
          showscale: true,
          colorbar: {{
            title: {{ text: colourLabels[colourBy], font: {{ size: 12 }} }},
            thickness: 20,
            len: 0.5,
            tickfont: {{ size: 10 }}
          }}
        }},
        dimensions: dimensions,
        labelfont: {{ size: 11, color: '#1a1a2e' }},
        tickfont: {{ size: 9, color: '#666' }},
        rangefont: {{ size: 9, color: '#1a1a2e' }}
      }};

      var layout = {{
        font: {{ family: 'Segoe UI, system-ui, sans-serif', size: 12, color: '#333' }},
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        margin: {{ l: 100, r: 60, t: 40, b: 20 }}
      }};

      Plotly.react('chart', [trace], layout, {{ responsive: true }});

      // Update titles
      if (currentView === 'chapters') {{
        document.getElementById('chart-title').textContent =
          'Chapter Overview — ' + activeChapters.length + ' groups';
      }} else {{
        var uniqueCodes = new Set(data.map(function(r) {{ return r.code; }})).size;
        document.getElementById('chart-title').textContent =
          'Individual Diagnoses — ' + uniqueCodes + ' categories';
      }}
      document.getElementById('chart-subtitle').textContent =
        data.length + ' lines · Drag any axis to filter';
    }}

    // Event listeners
    document.getElementById('chapter-filter').addEventListener('change', updateChart);
    document.getElementById('min-admissions').addEventListener('input', function() {{
      document.getElementById('min-adm-label').textContent =
        formatNumber(parseInt(this.value));
      updateChart();
    }});

    // Initial render
    updateChart();
  </script>
</body>
</html>"""

    output_path = os.path.join(OUTPUT_DIR, "dashboard.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Saved: {output_path}")


#main execution
if __name__ == "__main__":
    print("Loading NHS Hospital Admissions data...\n")
    df = load_all_data()
    if df.empty:
        print("No data loaded.")
        exit(1)

    print(f"\n {len(df):,} rows | {df['year'].nunique()} years | {df['code'].nunique()} categories\n")
    build_dashboard(df)
    print("Done")