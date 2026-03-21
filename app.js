const SHEET_ID = '1_eutPpUeEWZG_3F8bPjPyqxy8wufu_mM';
const SHEET_GID = '2023701082';
const GVIZ_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&gid=${SHEET_GID}`;
const XLSX_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx&gid=${SHEET_GID}`;

const translations = {
  bg: {
    title: 'БЪЛГАРСКИ РЕКОРДИ ПО СТРЕЛБА С ЛЪК',
    season: 'Сезон',
    category: 'Категория',
    division: 'Дивизия',
    recordType: 'Вид запис',
    archer: 'Стрелец',
    club: 'Клуб',
    points: 'Точки',
    date: 'Дата',
    allSeasons: 'Всички сезони',
    allCategories: 'Всички категории',
    allDivisions: 'Всички дивизии',
    allTypes: 'Всички видове',
    allYears: 'Всички години',
    searchArcher: 'Търси стрелец...',
    searchClub: 'Търси клуб...',
    noResults: 'Няма намерени резултати',
    loading: 'Зареждане...',
    error: 'Грешка при зареждане на данните',
    dataFrom: 'Данни от: Google Sheets',
    updated: 'Актуализирано',
    records: 'записа',
    shown: 'показани',
  },
  en: {
    title: 'BULGARIAN ARCHERY RECORDS',
    season: 'Season',
    category: 'Category',
    division: 'Division',
    recordType: 'Record Type',
    archer: 'Archer',
    club: 'Club',
    points: 'Points',
    date: 'Date',
    allSeasons: 'All Seasons',
    allCategories: 'All Categories',
    allDivisions: 'All Divisions',
    allTypes: 'All Types',
    allYears: 'All Years',
    searchArcher: 'Search archer...',
    searchClub: 'Search club...',
    noResults: 'No results found',
    loading: 'Loading...',
    error: 'Error loading data',
    dataFrom: 'Data from: Google Sheets',
    updated: 'Updated',
    records: 'records',
    shown: 'shown',
  }
};

// parse XLSX — reads merged-cell division labels via SheetJS
async function fetchRecordsFromXLSX() {
  const resp = await fetch(XLSX_URL, { redirect: 'follow' });
  if (!resp.ok) throw new Error(`XLSX HTTP ${resp.status}`);
  const ab = await resp.arrayBuffer();
  const wb = XLSX.read(ab, { type: 'array', cellDates: true });

  // find the sheet by scanning for the "Актуали към:" metadata marker
  let ws = null;
  for (const name of wb.SheetNames) {
    const s = wb.Sheets[name];
    outer: for (let r = 0; r < 5; r++) {
      for (let c = 0; c < 5; c++) {
        const cell = s[XLSX.utils.encode_cell({ r, c })];
        if (cell && typeof cell.v === 'string' && cell.v.includes('Актуал')) {
          ws = s; break outer;
        }
      }
    }
    if (ws) break;
  }
  if (!ws) ws = wb.Sheets[wb.SheetNames[0]];

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const merges = ws['!merges'] || [];

  // division header rows are wide merges (≥3 cols); map each row index to its label
  const divisionByRow = {};
  for (const m of merges) {
    const colSpan = m.e.c - m.s.c + 1;
    if (colSpan < 3) continue;
    const cell = ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })];
    if (!cell || typeof cell.v !== 'string' || !cell.v.trim()) continue;
    const text = cell.v.trim();
    if (text === 'Тип рекорд' || text === 'Стрелец' || text.startsWith('Актуал')) continue;
    for (let r = m.s.r; r <= m.e.r; r++) divisionByRow[r] = text;
  }

  const parsed = [];
  let lastUpdated = '';
  let currentDivision = '';

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cell = c => ws[XLSX.utils.encode_cell({ r, c })];
    const strVal = c => { const x = cell(c); return x ? String(x.v ?? '').trim() : ''; };

    if (divisionByRow[r] !== undefined) {
      currentDivision = divisionByRow[r];
      continue;
    }

    const bVal = strVal(1);
    if (bVal.includes('Актуал')) { lastUpdated = strVal(2); continue; }
    if (bVal === 'Тип рекорд') continue;

    const aCell = cell(0);
    if (!aCell || aCell.v === null || aCell.v === undefined) continue;
    if (typeof aCell.v !== 'number') continue;

    const recordType = bVal;
    const archer = strVal(2);
    const club = strVal(3);
    const eCell = cell(4);
    const points = eCell ? Number(eCell.v) : null;

    const fCell = cell(5);
    let date = '';
    if (fCell) {
      if (fCell.t === 'd' && fCell.v instanceof Date) {
        const d = fCell.v;
        date = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
      } else if (fCell.w) {
        date = fCell.w;
      } else {
        date = String(fCell.v ?? '');
      }
    }

    if (!archer || archer === 'Стрелец') continue;

    const dp = currentDivision.split('/').map(s => s.trim());
    parsed.push({ division: currentDivision, season: dp[0] || '', category: dp[1] || '', divisionPart: dp[2] || '', recordType, archer, club, points, date });
  }

  return { records: parsed, lastUpdated };
}

// parse gviz JSON fallback (no division data)
async function fetchRecordsFromGviz() {
  const resp = await fetch(GVIZ_URL);
  const text = await resp.text();
  const json = JSON.parse(text.replace(/^[^{]*/, '').replace(/[^}]*$/, ''));

  const rows = json.table.rows;
  const parsed = [];
  let lastUpdated = '';

  for (const row of rows) {
    const cells = row.c || [];
    const colA = cells[0];

    if (colA === null || colA === undefined || colA.v === null || colA.v === undefined) {
      const colB = cells[1];
      if (colB && String(colB.v ?? '').includes('Актуал')) {
        const colC = cells[2];
        if (colC && colC.v) lastUpdated = String(colC.v).trim();
      }
      continue;
    }

    const str = idx => { const c = cells[idx]; return (c && c.v != null) ? String(c.v).trim() : ''; };
    const fmt = idx => { const c = cells[idx]; if (!c) return ''; return c.f ?? (c.v != null ? String(c.v) : ''); };

    const archer = str(2);
    if (!archer || archer === 'Стрелец') continue;

    parsed.push({
      division: '', season: '', category: '', divisionPart: '',
      recordType: str(1),
      archer,
      club: str(3),
      points: cells[4]?.v != null ? Number(cells[4].v) : null,
      date: fmt(5),
    });
  }

  return { records: parsed, lastUpdated };
}

function archerApp() {
  return {
    lang: 'bg',
    records: [],
    hasDivisions: false,
    loading: true,
    error: null,
    lastUpdated: null,

    f: {
      season: '',
      category: '',
      divisionPart: '',
      recordType: '',
      archer: '',
      club: '',
      date: '',
    },

    suggest: { field: null, items: [], idx: -1 },

    expanded: {},
    toggleRow(i) { this.expanded[i] = !this.expanded[i]; },

    sortCol: null,
    sortDir: 'asc',

    get t() { return translations[this.lang]; },

    get filtered() {
      let rows = this.records
        .filter(r => !this.f.season       || r.season === this.f.season)
        .filter(r => !this.f.category     || r.category === this.f.category)
        .filter(r => !this.f.divisionPart || r.divisionPart === this.f.divisionPart)
        .filter(r => !this.f.recordType   || r.recordType === this.f.recordType)
        .filter(r => !this.f.archer       || r.archer.toLowerCase().includes(this.f.archer.toLowerCase()))
        .filter(r => !this.f.club         || r.club.toLowerCase().includes(this.f.club.toLowerCase()))
        .filter(r => !this.f.date         || r.date.includes(this.f.date));

      if (this.sortCol) {
        rows = [...rows].sort((a, b) => {
          let va = a[this.sortCol], vb = b[this.sortCol];
          if (this.sortCol === 'points') {
            va = Number(va) || 0; vb = Number(vb) || 0;
          } else if (this.sortCol === 'date') {
            const iso = d => { if (!d) return ''; const p = d.split('/'); return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : d; };
            va = iso(va); vb = iso(vb);
          } else {
            va = (va || '').toLowerCase(); vb = (vb || '').toLowerCase();
          }
          if (va < vb) return this.sortDir === 'asc' ? -1 : 1;
          if (va > vb) return this.sortDir === 'asc' ? 1 : -1;
          return 0;
        });
      }
      return rows;
    },

    withoutFilter(key) {
      return this.records.filter(r =>
        Object.keys(this.f).every(k => {
          if (k === key || !this.f[k]) return true;
          if (k === 'archer') return r.archer.toLowerCase().includes(this.f[k].toLowerCase());
          if (k === 'club')   return r.club.toLowerCase().includes(this.f[k].toLowerCase());
          if (k === 'date')   return r.date.includes(this.f[k]);
          return r[k] === this.f[k];
        })
      );
    },

    get seasonOptions() {
      return [...new Set(this.withoutFilter('season').map(r => r.season).filter(Boolean))].sort();
    },
    get categoryOptions() {
      return [...new Set(this.withoutFilter('category').map(r => r.category).filter(Boolean))].sort();
    },
    get divisionPartOptions() {
      return [...new Set(this.withoutFilter('divisionPart').map(r => r.divisionPart).filter(Boolean))].sort();
    },
    get recordTypeOptions() {
      return [...new Set(this.withoutFilter('recordType').map(r => r.recordType).filter(Boolean))].sort();
    },
    get yearOptions() {
      const years = this.withoutFilter('date').map(r => {
        const p = (r.date || '').split('/'); return p.length === 3 ? p[2] : null;
      }).filter(Boolean);
      return [...new Set(years)].sort().reverse();
    },

    setSort(col) {
      if (this.sortCol === col) this.sortDir = this.sortDir === 'asc' ? 'desc' : 'asc';
      else { this.sortCol = col; this.sortDir = 'asc'; }
    },

    suggestFor(field) {
      const val = (this.f[field] || '').toLowerCase();
      if (!val) return [];
      const seen = new Set();
      const out = [];
      for (const r of this.withoutFilter(field)) {
        const v = r[field];
        if (v && v.toLowerCase().includes(val) && !seen.has(v)) { seen.add(v); out.push(v); }
      }
      return out.sort().slice(0, 12);
    },

    onSuggestInput(field) {
      this.suggest.field = field;
      this.suggest.items = this.suggestFor(field);
      this.suggest.idx = -1;
    },

    onSuggestKeydown(e, field) {
      const { items } = this.suggest;
      if (e.key === 'ArrowDown') {
        if (!items.length) return;
        e.preventDefault();
        this.suggest.idx = Math.min(this.suggest.idx + 1, items.length - 1);
      } else if (e.key === 'ArrowUp') {
        if (!items.length) return;
        e.preventDefault();
        this.suggest.idx = Math.max(this.suggest.idx - 1, -1);
      } else if (e.key === 'Enter' && this.suggest.idx >= 0) {
        e.preventDefault();
        this.f[field] = items[this.suggest.idx];
        this.suggest = { field: null, items: [], idx: -1 };
      } else if (e.key === 'Escape') {
        this.suggest = { field: null, items: [], idx: -1 };
      }
    },

    pickSuggest(field, val) {
      this.f[field] = val;
      this.suggest = { field: null, items: [], idx: -1 };
    },

    closeSuggest() {
      this.suggest = { field: null, items: [], idx: -1 };
    },

    clearFilter(key) { this.f[key] = ''; this.closeSuggest(); },
    get hasActiveFilters() { return Object.values(this.f).some(v => v !== ''); },
    clearAll() { Object.keys(this.f).forEach(k => this.f[k] = ''); this.closeSuggest(); this.expanded = {}; },

    async init() {
      try {
        if (typeof XLSX !== 'undefined') {
          try {
            const { records, lastUpdated } = await fetchRecordsFromXLSX();
            this.records = records;
            this.lastUpdated = lastUpdated || new Date().toLocaleDateString('bg-BG');
            this.hasDivisions = records.some(r => r.division);
            return;
          } catch (xlsxErr) {
            console.warn('XLSX fetch failed, falling back to gviz:', xlsxErr.message);
          }
        }
        const { records, lastUpdated } = await fetchRecordsFromGviz();
        this.records = records;
        this.lastUpdated = lastUpdated || new Date().toLocaleDateString('bg-BG');
        this.hasDivisions = false;
      } catch (e) {
        this.error = e.message;
      } finally {
        this.loading = false;
      }
    },
  };
}
