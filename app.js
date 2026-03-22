const SHEET_ID = '1_eutPpUeEWZG_3F8bPjPyqxy8wufu_mM';
const SHEET_GID = '1898570785';
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
    year: 'Година',
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
    from: 'от',
    shown: 'показани',
    onlyCurrentRecord: 'Само текущ рекорд',
    onlyCurrentDiscipline: 'Само актуални дисциплини',
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
    year: 'Year',
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
    from: 'from',
    shown: 'shown',
    onlyCurrentRecord: 'Current records only',
    onlyCurrentDiscipline: 'Active disciplines only',
  }
};

// builds the category label from Възрастова Група + Индивидуално/Отборно
function buildCategory(ageGroup, indivTeam) {
  const it = indivTeam.toLowerCase();
  if (it === 'индивидуално' || !it) return ageGroup;
  // "Смесен отбор" and ageGroup already encodes it → no suffix
  if (it === 'смесен отбор' && ageGroup.toLowerCase().includes('смесен отбор')) return ageGroup;
  return ageGroup ? `${ageGroup} Отборно` : 'Отборно';
}

// parse XLSX — flat table, row 0 is header
async function fetchRecordsFromXLSX() {
  const resp = await fetch(XLSX_URL, { redirect: 'follow' });
  if (!resp.ok) throw new Error(`XLSX HTTP ${resp.status}`);
  const ab = await resp.arrayBuffer();
  const wb = XLSX.read(ab, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const parsed = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cell = c => ws[XLSX.utils.encode_cell({ r, c })];
    const strVal = c => { const x = cell(c); return x ? String(x.v ?? '').trim() : ''; };

    const recordType = strVal(0);
    if (!recordType || recordType === 'Тип рекорд') continue;  // header or blank

    const style     = strVal(1);
    const discip    = strVal(2);
    const ageGroup  = strVal(3);
    const indivTeam = strVal(4);
    const pCell     = cell(5);
    const points    = pCell ? Number(pCell.v) : null;
    const archer    = strVal(9);
    const club      = strVal(10);
    if (!archer) continue;

    const dCell = cell(11);
    let date = '';
    if (dCell) {
      if (dCell.t === 'd' && dCell.v instanceof Date) {
        const d = dCell.v;
        date = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
      } else if (dCell.w) {
        date = dCell.w;
      } else {
        date = String(dCell.v ?? '');
      }
    }

    const activeDisc = strVal(7);
    const currentRec = strVal(8);
    const category = buildCategory(ageGroup, indivTeam);
    const divAgeGroup = (indivTeam && indivTeam !== 'Индивидуално' && !(indivTeam.toLowerCase() === 'смесен отбор' && ageGroup.toLowerCase().includes('смесен отбор'))) ? `${ageGroup} ${indivTeam}` : ageGroup;
    const division = [style, discip, divAgeGroup].filter(Boolean).join(' / ');
    parsed.push({ division, season: discip, category, divisionPart: style, recordType, archer, club, points, date, activeDisc, currentRec });
  }

  return { records: parsed, lastUpdated: '' };
}

// parse gviz JSON fallback
async function fetchRecordsFromGviz() {
  const resp = await fetch(GVIZ_URL);
  const text = await resp.text();
  const json = JSON.parse(text.replace(/^[^{]*/, '').replace(/[^}]*$/, ''));
  const rows = json.table.rows;
  const parsed = [];

  for (const row of rows) {
    const cells = row.c || [];
    const str = idx => { const c = cells[idx]; return (c && c.v != null) ? String(c.v).trim() : ''; };
    const fmt = idx => { const c = cells[idx]; if (!c) return ''; return c.f ?? (c.v != null ? String(c.v) : ''); };

    const recordType = str(0);
    if (!recordType || recordType === 'Тип рекорд') continue;

    const style     = str(1);
    const discip    = str(2);
    const ageGroup  = str(3);
    const indivTeam = str(4);
    const points    = cells[5]?.v != null ? Number(cells[5].v) : null;
    const archer    = str(9);
    const club      = str(10);
    const date      = fmt(11);

    if (!archer) continue;

    const activeDisc = str(7);
    const currentRec = str(8);
    const category = buildCategory(ageGroup, indivTeam);
    const divAgeGroup = (indivTeam && indivTeam !== 'Индивидуално' && !(indivTeam.toLowerCase() === 'смесен отбор' && ageGroup.toLowerCase().includes('смесен отбор'))) ? `${ageGroup} ${indivTeam}` : ageGroup;
    const division = [style, discip, divAgeGroup].filter(Boolean).join(' / ');
    parsed.push({ division, season: discip, category, divisionPart: style, recordType, archer, club, points, date, activeDisc, currentRec });
  }

  return { records: parsed, lastUpdated: '' };
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
      onlyCurrentRecord: false,
      onlyCurrentDiscipline: false,
    },

    suggest: { field: null, items: [], idx: -1 },

    expanded: {},
    toggleRow(i) { this.expanded[i] = !this.expanded[i]; },

    sortCol: null,
    sortDir: 'asc',

    get t() { return translations[this.lang]; },

    get filtered() {
      let rows = this.records
        .filter(r => !this.f.season                 || r.season === this.f.season)
        .filter(r => !this.f.category               || r.category === this.f.category)
        .filter(r => !this.f.divisionPart           || r.divisionPart === this.f.divisionPart)
        .filter(r => !this.f.recordType             || r.recordType === this.f.recordType)
        .filter(r => !this.f.archer                 || r.archer.toLowerCase().includes(this.f.archer.toLowerCase()))
        .filter(r => !this.f.club                   || r.club.toLowerCase().includes(this.f.club.toLowerCase()))
        .filter(r => !this.f.date                   || r.date.includes(this.f.date))
        .filter(r => !this.f.onlyCurrentRecord      || r.currentRec === 'Текущ Рекорд')
        .filter(r => !this.f.onlyCurrentDiscipline  || r.activeDisc === 'Актуална Дисциплина');

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
          if (k === 'archer')                return r.archer.toLowerCase().includes(this.f[k].toLowerCase());
          if (k === 'club')                  return r.club.toLowerCase().includes(this.f[k].toLowerCase());
          if (k === 'date')                  return r.date.includes(this.f[k]);
          if (k === 'onlyCurrentRecord')     return r.currentRec === 'Текущ Рекорд';
          if (k === 'onlyCurrentDiscipline') return r.activeDisc === 'Актуална Дисциплина';
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

    clearFilter(key) { this.f[key] = typeof this.f[key] === 'boolean' ? false : ''; this.closeSuggest(); },
    get hasActiveFilters() { return Object.values(this.f).some(v => v !== '' && v !== false); },
    clearAll() { Object.keys(this.f).forEach(k => this.f[k] = typeof this.f[k] === 'boolean' ? false : ''); this.closeSuggest(); this.expanded = {}; },

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
        this.hasDivisions = records.some(r => r.division);
      } catch (e) {
        this.error = e.message;
      } finally {
        this.loading = false;
      }
    },
  };
}
