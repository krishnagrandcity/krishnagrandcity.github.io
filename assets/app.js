/* ───── part 1 ───── */
/* Global toast helper — showToast(message, type) where type = 'ok' | 'err' | '' */
window.showToast = function (message, type) {
  try {
    var wrap = document.getElementById('toast-wrap');
    if (!wrap) { wrap = document.createElement('div'); wrap.id = 'toast-wrap'; document.body.appendChild(wrap); }
    var t = document.createElement('div');
    t.className = 'toast' + (type ? ' ' + type : '');
    var mark = type === 'ok' ? '✓' : type === 'err' ? '!' : 'i';
    var tick = document.createElement('span'); tick.className = 'tick'; tick.textContent = mark;
    var text = (window.__kgcTranslate ? window.__kgcTranslate(message) : message);
    var span = document.createElement('span'); span.textContent = text;
    t.appendChild(tick); t.appendChild(span);
    wrap.appendChild(t);
    var life = setTimeout(close, 3600);
    function close() {
      clearTimeout(life);
      t.classList.add('out');
      t.addEventListener('animationend', function () { t.remove(); }, { once: true });
    }
    t.addEventListener('click', close);
  } catch (e) { /* toast is non-critical */ }
};

/* ───── part 2 ───── */
/* ──────────────────────────────────────────────────
   MULTI-USER LOGIN ACCOUNTS
   Each account has: email, password, name, role, initials
   Role controls which sidebar options are visible
   ────────────────────────────────────────────────── */
/* ──────────────────────────────────────────────────
   ACCESS TIERS
   admin     — President, Secretary, Treasurer, Maintenance Officer.
               Full editing rights: members, analytics, notices, complaints.
   committee — Joint Secretary, other committee members, Security.
               Can raise complaints and book Club Hall / Gym only.
   resident  — One login per household (generated below from `members`).
               Can view only their own profile/metrics, book Club Hall / Gym,
               raise complaints, and read the Notice Board.
   ────────────────────────────────────────────────── */
/* ──────────────────────────────────────────────────
   PERIOD CONSTANTS — must be defined before member/analytics
   data is seeded below, since seeding runs immediately.
   ────────────────────────────────────────────────── */
/* ──────────────────────────────────────────────────
   BACKEND CONNECTION
   All real login/data now goes through the Node.js backend
   (kgc-backend), which talks to Supabase on your behalf.
   Change BACKEND_URL if your Render URL is ever different.
   ────────────────────────────────────────────────── */
const BACKEND_URL = 'https://kgc-backend.onrender.com';

// Wraps fetch() to automatically attach the logged-in user's token,
// and to consistently surface errors. Use this for every future
// call to the backend instead of raw fetch().
async function apiRequest(path, options = {}) {
  const token = sessionStorage.getItem('kgc_token');
  const headers = { 'Content-Type': 'application/json', ...(options.headers || {}) };
  if (token) headers['Authorization'] = `Bearer ${token}`;

  const res = await fetch(`${BACKEND_URL}${path}`, { ...options, headers });
  let data = null;
  try { data = await res.json(); } catch (_) { /* no JSON body */ }

  if (!res.ok) {
    const message = (data && data.error) ? data.error : `Request failed (${res.status})`;
    throw new Error(message);
  }
  return data;
}

const AVAILABLE_YEARS = Array.from({length: 100}, (_, i) => 2015 + i); // FY2015–16 through FY2114–15

/* ════════════════════════════════════════════════════
   IMPORT MEMBERS FROM EXCEL
   ════════════════════════════════════════════════════ */
let importParsedRows = []; // parsed + mapped rows, ready to send to the backend

function openImportModal() {
  document.getElementById('import-modal').classList.remove('hidden');
  document.getElementById('import-step-upload').classList.remove('hidden');
  document.getElementById('import-step-preview').classList.add('hidden');
  document.getElementById('import-step-result').classList.add('hidden');
  document.getElementById('import-file-input').value = '';
  document.getElementById('import-parse-error').classList.add('hidden');
  importParsedRows = [];
}
function closeImportModal() { document.getElementById('import-modal').classList.add('hidden'); }

function handleImportFile(event) {
  const file = event.target.files[0];
  const errBox = document.getElementById('import-parse-error');
  errBox.classList.add('hidden');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });

      // ── Unit type detection ──────────────────────────────────
      // Matches the actual house number formats in the Excel file:
      //   "Antilia - 01", "Royale - 14", "Empress - 24"  → rowhouse
      //   "Regal 1", "Regal 2"                           → 2bhk
      //   "A1 - 1", "A1 - 2"  (Block A-1)               → 2bhk
      //   "A2 - 33"            (Block A-2)               → 2bhk
      //   "Kingston - 01"                                → 3bhk
      //   "Shop - 12"                                    → shop
      // Falls back to the sheet name if the prefix is not recognised.
      function deriveUnitType(houseNumber, sheetName) {
        const hn = houseNumber.trim();
        const sn = (sheetName || '').toLowerCase();

        // By house number prefix (primary method)
        if (/^(antilia|antillia|empress|royale)\b/i.test(hn)) return 'rowhouse';
        if (/^regal\b/i.test(hn))                             return '2bhk';
        if (/^sunshine\b/i.test(hn))                          return '2bhk';
        if (/^(a1|a-1)\b/i.test(hn))                         return '2bhk';  // legacy, pre-rename
        if (/^(a2|a-2)\b/i.test(hn))                         return '2bhk';  // legacy, pre-rename
        if (/^kingston\b/i.test(hn))                          return '3bhk';
        if (/^shop\b/i.test(hn))                              return 'shop';

        // Fallback: derive from sheet name
        if (sn.includes('row house'))                           return 'rowhouse';
        if (sn.includes('2 bhk') || sn.includes('a - 1') ||
            sn.includes('a - 2') || sn.includes('a-1') ||
            sn.includes('a-2'))                                 return '2bhk';
        if (sn.includes('3 bhk'))                              return '3bhk';
        if (sn.includes('shop'))                               return 'shop';

        return ''; // unknown — will show as error in preview
      }

      // Sheets to skip (none now — Shop units are imported as unit type "shop")
      const SKIP_SHEETS = [];

      // Read ALL sheets and merge into one flat array
      let allRows = [];
      for (const sheetName of workbook.SheetNames) {
        if (SKIP_SHEETS.includes(sheetName.toLowerCase().trim())) continue;

        const sheet = workbook.Sheets[sheetName];
        const rows  = XLSX.utils.sheet_to_json(sheet, { defval: '' });

        rows.forEach(r => {
          const houseNumber = String(r.house_number || '').trim();
          const fullName    = String(r.full_name    || '').trim();
          if (!fullName && !houseNumber) return; // skip blank rows

          allRows.push({
            fullName,
            houseNumber,
            unitType:              deriveUnitType(houseNumber, sheetName),
            // Strip hyphens/spaces from phone numbers (e.g. "73549-20530" → "7354920530")
            phonePrimary:          String(r.phone_primary           || '').trim().replace(/[-\s]/g, ''),
            phoneSecondary:        String(r.phone_secondary         || '').trim().replace(/[-\s]/g, '') || null,
            email:                 String(r.email                   || '').trim() || null,
            vehicle1Reg:           String(r.vehicle_1_reg           || '').trim() || null,
            vehicle2Reg:           String(r.vehicle_2_reg           || '').trim() || null,
            emergencyContactName:  String(r.emergency_contact_name  || '').trim() || null,
            emergencyContactPhone: String(r.emergency_contact_phone || '').trim() || null,
            _sheet: sheetName, // shown in import preview
          });
        });
      }

      if (allRows.length === 0) {
        errBox.textContent = 'No data rows found across all sheets.';
        errBox.classList.remove('hidden');
        return;
      }

      importParsedRows = allRows;

      renderImportPreview();
      document.getElementById('import-step-upload').classList.add('hidden');
      document.getElementById('import-step-preview').classList.remove('hidden');
    } catch (err) {
      console.error('Import parse error:', err);
      errBox.textContent = 'Could not read that file. Make sure it is a valid .xlsx file.';
      errBox.classList.remove('hidden');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderImportPreview() {
  document.getElementById('import-row-count').textContent = importParsedRows.length;
  document.getElementById('import-preview-body').innerHTML = importParsedRows.map(r => {
    const missing = !r.fullName || !r.houseNumber || !r.unitType || !r.phonePrimary;
    const rowStyle = missing ? 'background:#fdecea;' : '';
    return `<tr style="${rowStyle}">
      <td style="padding:6px 8px;font-size:11px;color:var(--text2);">${r._sheet || ''}</td>
      <td style="padding:6px 8px;">${r.houseNumber || '<span style="color:var(--red)">missing</span>'}</td>
      <td style="padding:6px 8px;">${r.fullName || '<span style="color:var(--red)">missing</span>'}</td>
      <td style="padding:6px 8px;">${r.unitType || '<span style="color:var(--red)">unknown</span>'}</td>
      <td style="padding:6px 8px;">${r.phonePrimary || '<span style="color:var(--red)">missing</span>'}</td>
    </tr>`;
  }).join('');
}

async function confirmImport() {
  const btn     = document.getElementById('import-confirm-btn');
  const BATCH   = 50; // send 50 rows at a time to avoid timeouts on Render free tier
  btn.disabled  = true;

  const totalRows   = importParsedRows.length;
  let succeeded     = 0;
  let allFailed     = [];
  let batchCount    = Math.ceil(totalRows / BATCH);

  try {
    for (let i = 0; i < totalRows; i += BATCH) {
      const batch     = importParsedRows.slice(i, i + BATCH);
      const batchNum  = Math.floor(i / BATCH) + 1;
      btn.textContent = `Importing… (${Math.min(i + BATCH, totalRows)} / ${totalRows})`;

      const result = await apiRequest('/api/members/import', {
        method: 'POST',
        body: JSON.stringify({ members: batch }),
      });

      succeeded   += result.successCount;
      allFailed    = allFailed.concat(result.failed || []);
    }

    // Show result
    document.getElementById('import-step-preview').classList.add('hidden');
    document.getElementById('import-step-result').classList.remove('hidden');
    document.getElementById('import-result-summary').innerHTML =
      `✅ Imported <strong>${succeeded}</strong> of ${totalRows} rows successfully.` +
      (allFailed.length > 0
        ? ` ⚠️ <strong>${allFailed.length}</strong> row(s) failed — see details below.`
        : ' 🎉 All rows imported successfully.');

    document.getElementById('import-result-failures').innerHTML =
      allFailed.length === 0 ? '' :
      '<table style="width:100%;margin-top:12px;"><thead><tr>' +
      '<th style="text-align:left;padding:4px;">House</th>' +
      '<th style="text-align:left;padding:4px;">Reason</th>' +
      '</tr></thead><tbody>' +
      allFailed.map(f =>
        `<tr><td style="padding:4px;">${f.row}</td><td style="padding:4px;color:var(--red);">${f.error}</td></tr>`
      ).join('') +
      '</tbody></table>';

    if (typeof loadMembers === 'function') loadMembers();
  } catch (e) {
    alert('Import failed: ' + (e.message || 'Unknown error. Check your internet connection and try again.'));
  } finally {
    btn.disabled    = false;
    btn.textContent = 'Confirm & Import';
  }
}

const QUARTER_LABELS  = { Q1:'Q1 (Apr–Jun)', Q2:'Q2 (Jul–Sep)', Q3:'Q3 (Oct–Dec)', Q4:'Q4 (Jan–Mar)' };
const QUARTERS = ['Q1','Q2','Q3','Q4'];

// NOTE: This array is no longer used for real authentication — login now
// goes through the real backend (see doLogin() and BACKEND_URL above).
// Kept here only as a reference for which staff roles originally existed
// in this demo; safe to delete once real committee accounts exist in the database.
const ACCOUNTS = [
  { email: 'president@kgc.in',    pass: 'kgc@2025',   name: 'Rajiv Sharma',   role: 'President',        initials: 'RS', tier: 'admin' },
  { email: 'secretary@kgc.in',    pass: 'sec@2025',   name: 'Anita Mehra',    role: 'Secretary',        initials: 'AM', tier: 'admin' },
  { email: 'treasurer@kgc.in',    pass: 'tre@2025',   name: 'Deepak Nair',    role: 'Treasurer',        initials: 'DN', tier: 'admin' },
  { email: 'maintenance@kgc.in',  pass: 'mnt@2025',   name: 'Suresh Iyer',    role: 'Maintenance Officer',initials: 'SI', tier: 'admin' },
  { email: 'jointsec@kgc.in',     pass: 'jts@2025',   name: 'Priya Kapoor',   role: 'Joint Secretary',  initials: 'PK', tier: 'committee' },
  { email: 'committee@kgc.in',    pass: 'com@2025',   name: 'Neha Bhatt',     role: 'Committee Member',initials: 'NB', tier: 'committee' },
  { email: 'security@kgc.in',     pass: 'sec@guard',  name: 'Guard Account',  role: 'Security',         initials: 'GD', tier: 'committee' },
];

/* ──────────────────────────────────────────────────
   MEMBER DATA STORE
   Populated from the real backend — see loadMembers() below.
   Payment/dues fields intentionally do not exist here yet; that
   data belongs to the Maintenance Bills module, not yet wired.
   ────────────────────────────────────────────────── */
let members = [];

async function loadMembers() {
  const tbody = document.getElementById('members-tbody');
  if (tbody) tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:32px;color:var(--text2);">Loading…</td></tr>';
  try {
    const rows = await apiRequest('/api/members');
    members = rows.map(r => ({
      id: r.id,
      name: r.full_name,
      house: r.house_number,
      type: r.unit_type,
      phone: r.phone_primary,
      email: r.email || '—',
      isActive: r.is_active,
      photo: '',
    }));
    // Other pages (Maintenance, Finances ...) load members only for the house-number lookup
    if (document.getElementById('m-total')) renderMemberMetrics();
    if (tbody) renderMembersTable();
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:32px;color:var(--red);">Could not load members: ${e.message}</td></tr>`;
  }
}


/* ──────────────────────────────────────────────────
   RESIDENT LOGIN ACCOUNTS
   One login is auto-generated per household from `members`, so every
   resident can sign in and see only their own profile.
   Email  = house number, lowercased, no hyphen, e.g. A-101 → a101@kgc.in
   Password = "kgc" + last 4 digits of the registered phone number
   ────────────────────────────────────────────────── */
function buildResidentAccounts() {
  return members.map(m => {
    const digits = m.phone.replace(/\D/g, '');
    return {
      email:  m.house.toLowerCase().replace(/-/g, '') + '@kgc.in',
      pass:   'kgc' + digits.slice(-4),
      name:   m.name,
      role:   'Resident',
      initials: m.name.split(' ').map(n=>n[0]).join('').slice(0,2),
      tier:   'resident',
      memberId: m.id,
    };
  });
}
let RESIDENT_ACCOUNTS = buildResidentAccounts();

/* ──────────────────────────────────────────────────
   MEMBER FINANCE RECORDS
   Builds m.payments[year][quarter] = { due, paid, paidDate } for every
   member, across AVAILABLE_YEARS. Seeded from the member's current
   status/defaulter flag so the table has realistic-looking history;
   from here on, real payments are recorded via the calendar picker on
   the member's profile page and overwrite these values.
   ────────────────────────────────────────────────── */
// FY semantics: "year" X = the financial year starting Apr X and ending Mar (X+1).
// Q1 = Apr-Jun X · Q2 = Jul-Sep X · Q3 = Oct-Dec X · Q4 = Jan-Mar (X+1)
// "Today" for this demo portal is treated as sitting inside FY2026 Q2 (Jul-Sep 2026).
const CURRENT_FY = 2026;
const CURRENT_FY_Q = 'Q2';

// Flat chronological list of every {year,quarter} across AVAILABLE_YEARS, oldest → newest
const CHRONO_QUARTERS = AVAILABLE_YEARS.flatMap(y => QUARTERS.map(q => ({ year: y, quarter: q })));
const CURRENT_CHRONO_INDEX = CHRONO_QUARTERS.findIndex(c => c.year === CURRENT_FY && c.quarter === CURRENT_FY_Q);
function chronoIndex(year, q) { return CHRONO_QUARTERS.findIndex(c => c.year === year && c.quarter === q); }

function seedMemberPayments(m) {
  m.payments = {};
  const due = m.maintenance * 3; // one quarter = 3 months of maintenance
  AVAILABLE_YEARS.forEach(year => {
    m.payments[year] = {};
    QUARTERS.forEach(q => {
      const idx = chronoIndex(year, q);
      const future = idx > CURRENT_CHRONO_INDEX; // hasn't started yet — nothing can be paid
      let paidFlag;
      if (future) {
        paidFlag = false;
      } else if (m.defaulter) {
        // Defaulter: the current quarter and the one immediately before it are unpaid
        paidFlag = idx !== CURRENT_CHRONO_INDEX && idx !== CURRENT_CHRONO_INDEX - 1;
      } else if (m.status === 'pending') {
        // Pending: only the current (in-progress) quarter is unpaid
        paidFlag = idx !== CURRENT_CHRONO_INDEX;
      } else {
        paidFlag = true;
      }
      m.payments[year][q] = { due, paid: paidFlag ? due : 0, paidDate: paidFlag ? quarterPaidDate(year, q) : '', future };
    });
  });
}
// Picks a plausible payment date inside a given financial quarter
function quarterPaidDate(year, q) {
  const startMonth = { Q1: 3, Q2: 6, Q3: 9, Q4: 0 }[q];
  const y = q === 'Q4' ? year + 1 : year;
  const d = new Date(y, startMonth, 8);
  return d.toISOString().slice(0, 10);
}
function seedAllMemberPayments() { members.forEach(seedMemberPayments); }
seedAllMemberPayments();

// Formats an ISO date (YYYY-MM-DD) as "8 Apr 2025"; returns '—' if empty
function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
}
// Which FY quarter does a given ISO date fall in? Returns {year, quarter}
function quarterOf(iso) {
  const d = new Date(iso);
  const m = d.getMonth(); // 0-11
  if (m >= 3 && m <= 5)  return { year: d.getFullYear(),     quarter: 'Q1' };
  if (m >= 6 && m <= 8)  return { year: d.getFullYear(),     quarter: 'Q2' };
  if (m >= 9 && m <= 11) return { year: d.getFullYear(),     quarter: 'Q3' };
  return                        { year: d.getFullYear() - 1, quarter: 'Q4' };
}
// Is a given member currently a "defaulter" for a given year? (2+ unpaid, already-due quarters)
// Real members loaded from the backend don't have a `.payments` object yet
// (that's the not-yet-built Maintenance Bills data) — guard against that
// so Analytics/Charts don't crash on real data, they just show zero for now.
function isDefaulterFor(m, year) {
  if (!m.payments) return false;
  const yr = m.payments[year];
  if (!yr) return false;
  return QUARTERS.filter(q => !yr[q].future && yr[q].paid < yr[q].due).length >= 2;
}

// Returns a member's due/paid/status for a given year + period ('Q1'..'Q4'|'yearly'|'defaulters').
// Used by the Analytics charts. Real members from the backend don't have
// `.payments` yet (that data belongs to the not-yet-built Maintenance
// Bills module) — this safely returns zeros for them rather than crashing.
function figuresFor(m, year, period) {
  if (!m.payments || !m.payments[year]) {
    return { due: 0, paid: 0, defaulting: false, future: false };
  }
  if (period === 'yearly' || period === 'defaulters') {
    const due  = QUARTERS.reduce((s, q) => s + m.payments[year][q].due, 0);
    const paid = QUARTERS.reduce((s, q) => s + m.payments[year][q].paid, 0);
    return { due, paid, defaulting: isDefaulterFor(m, year), future: false };
  }
  const rec = m.payments[year][period];
  return { due: rec.due, paid: rec.paid, defaulting: !rec.future && rec.paid < rec.due, future: rec.future };
}

/* ──────────────────────────────────────────────────
   NOTICE BOARD DATA STORE
   Populated from the real backend — see loadNotices() below.
   ────────────────────────────────────────────────── */
let notices = [];
let editingNoticeId = null;

/* ──────────────────────────────────────────────────
   COMPLAINTS DATA STORE
   Populated from the real backend — see loadComplaints() below.
   ────────────────────────────────────────────────── */
let complaints = [];

/* ──────────────────────────────────────────────────
   GYM & HALL BOOKING DATA STORES
   ────────────────────────────────────────────────── */
let gymMemberships = [];
let hallBookings   = [];

/* ──────────────────────────────────────────────────
   APP STATE VARIABLES
   ────────────────────────────────────────────────── */
let currentFilter    = 'all';      // Active member filter tab
let currentSearch    = '';         // Member search query string
let loggedIn         = false;      // Whether a user is currently logged in
let chartsInitialized = false;     // Prevent duplicate chart initialization
let currentUser      = null;       // The logged-in user account object
let noticeFilter     = 'all';      // Active notice category filter
let complaintFilter  = 'all';      // Active complaint status filter
let selectedPlanEl   = null;       // Currently selected gym plan card element
let selectedGymPlan  = '';         // Plan identifier string
let memberPeriod     = 'Q1';       // 'Q1'|'Q2'|'Q3'|'Q4'|'yearly'|'defaulters' — active Members-table period tab
let memberYear        = 2026;      // Active year for the Members table
let currentDetailMemberId = null;  // Member currently shown on the profile page
let currentDetailMember = null;    // Full profile object of that member (for editing)
let editingMemberId = null;        // Set while the Add/Edit modal is in "edit" mode
let detailYear        = 2026;      // Active year selector inside the member profile page
let analyticsPeriod  = 'yearly';   // 'Q1'|'Q2'|'Q3'|'Q4'|'yearly' — active Analytics period tab
let analyticsYear    = 2026;       // Active year for Analytics

/* ──────────────────────────────────────────────────
   Chart.js instance references
   Stored so they can be destroyed and rebuilt when data changes
   ────────────────────────────────────────────────── */
let chartInstances = {};

/* ──────────────────────────────────────────────────
   ANALYTICS DATA STORE — keyed by year, then by quarter (Q1–Q4)
   Defaults are generated so every year/quarter has something to show.
   Replaced automatically when a spreadsheet is uploaded, or edited via
   the Manual Entry (Advanced) panel. "Yearly" figures are always the
   sum of that year's four quarters — never entered directly.
   ────────────────────────────────────────────────── */
const BASE_MONTHLY   = [5.8, 6.1, 6.2, 5.9, 6.3, 6.7, 6.5, 6.8, 6.4, 6.9, 7.1, 6.5];
const BASE_EXP_CATS  = [14.2, 12.8, 8.4, 9.1, 5.3, 4.2, 7.5]; // Security, Maintenance, Landscaping, Electricity, Water, Admin, Reserves
const EXPENSE_CATEGORY_LABELS = ['Security','Maintenance','Landscaping','Electricity','Water','Admin','Reserves'];

function genQuarterData(yearIdx, qIdx) {
  const scale = 1 + yearIdx * 0.08; // slightly larger figures in later years
  const monthly = BASE_MONTHLY.slice(qIdx * 3, qIdx * 3 + 3).map(v => +(v * scale).toFixed(1));
  const expenses_cat = BASE_EXP_CATS.map(v => +((v / 4) * scale).toFixed(1));
  return {
    collected: +monthly.reduce((a, b) => a + b, 0).toFixed(1),
    pending:   +((4.8 / 4) * scale).toFixed(1),
    expenses:  +expenses_cat.reduce((a, b) => a + b, 0).toFixed(1),
    corpus:    +((32.1 / 4) * scale).toFixed(1),
    monthly, expenses_cat,
  };
}
function buildAnalyticsStore() {
  const store = {};
  AVAILABLE_YEARS.forEach((y, yi) => {
    store[y] = {};
    QUARTERS.forEach((q, qi) => { store[y][q] = genQuarterData(yi, qi); });
  });
  return store;
}
let analyticsStore = buildAnalyticsStore();

// Sums the four quarters of a year into one "Yearly (Cumulative)" dataset
function getYearlyAnalytics(year) {
  const yr = analyticsStore[year];
  if (!yr) return null;
  const expenses_cat = EXPENSE_CATEGORY_LABELS.map((_, i) => +QUARTERS.reduce((s, q) => s + yr[q].expenses_cat[i], 0).toFixed(1));
  return {
    collected: +QUARTERS.reduce((s, q) => s + yr[q].collected, 0).toFixed(1),
    pending:   +QUARTERS.reduce((s, q) => s + yr[q].pending, 0).toFixed(1),
    expenses:  +QUARTERS.reduce((s, q) => s + yr[q].expenses, 0).toFixed(1),
    corpus:    +Math.max(...QUARTERS.map(q => yr[q].corpus)).toFixed(1),
    monthly:   QUARTERS.flatMap(q => yr[q].monthly),
    expenses_cat,
  };
}
// Returns whichever dataset the current period-bar selection points to
function getCurrentAnalyticsData() {
  if (analyticsPeriod === 'yearly') return getYearlyAnalytics(analyticsYear);
  return analyticsStore[analyticsYear][analyticsPeriod];
}


/* ════════════════════════════════════════════════════
   PAGE ROUTING — shows/hides top-level pages
   ════════════════════════════════════════════════════ */
/* ──────────────────────────────────────────────────
   PAGE ROUTING — works in two modes
   • Single-file mode: every page / dashboard section is in this document,
     so showPage() / showDash() just show or hide it.
   • Multi-page mode (the separate-pages build): each page lives in its own
     .html file; when the requested page / section is not in the current
     document, showPage() / showDash() open its file instead. The same
     buttons therefore work as links in both modes.
   ────────────────────────────────────────────────── */
const SITE_PAGE_FILES = { home: 'index.html', about: 'about.html', login: 'login.html' };
const DASH_PAGE_FILES = {
  'members': 'members.html', 'detail': 'member.html', 'finances': 'finances.html',
  'bills': 'maintenance.html', 'pay-upi': 'pay-upi.html', 'notices': 'notices.html',
  'complaints': 'complaints.html', 'facilities': 'facilities.html',
  'resident-home': 'profile.html', 'res-bills': 'my-bills.html', 'res-complaints': 'my-complaints.html',
};
function goToFile(file) { window.location.href = file; }

// The dashboard section each kind of user lands on after signing in
function landingView() {
  if (!currentUser) return 'members';
  if (currentUser.tier === 'resident')  return 'resident-home';
  if (currentUser.tier === 'committee') return 'notices';
  return 'members';          // administrator tiers and manager
}

function showPage(page) {
  const target = document.getElementById('page-' + page);
  if (!target) {                                   // multi-page mode: open the page's own file
    if (page === 'dashboard') goToFile(loggedIn ? DASH_PAGE_FILES[landingView()] : SITE_PAGE_FILES.login);
    else if (SITE_PAGE_FILES[page]) goToFile(SITE_PAGE_FILES[page]);
    return;
  }
  ['home','about','login','dashboard'].forEach(p => {
    const el = document.getElementById('page-' + p);
    if (el) el.classList.toggle('hidden', p !== page);
  });
  // Update active state on nav buttons
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const btns = document.querySelectorAll('.nav-btn');
  if (page === 'home'  && btns[0]) btns[0].classList.add('active');
  if (page === 'about' && btns[1]) btns[1].classList.add('active');
  // Guard against accessing dashboard without login
  if (page === 'dashboard') {
    if (!loggedIn) { showPage('login'); return; }
    if (currentUser.tier === 'resident') {
      showDash('resident-home');
    } else if (currentUser.tier === 'committee') {
      showDash('notices');       // Committee lands on notices
    } else if (currentUser.tier === 'manager') {
      showDash('members');       // Manager lands on members
    } else {
      loadMembers();             // Admin tiers land on members
    }
  }
  window.scrollTo(0, 0);
}


/* ──────────────────────────────────────────────────
   CHANGE PASSWORD (self-service, for the logged-in user)
   ────────────────────────────────────────────────── */
function openChangePasswordModal() {
  document.getElementById('cp-current').value = '';
  document.getElementById('cp-new').value = '';
  document.getElementById('cp-error').classList.add('hidden');
  document.getElementById('cp-success').classList.add('hidden');
  document.getElementById('change-password-modal').classList.remove('hidden');
}

function closeChangePasswordModal() {
  document.getElementById('change-password-modal').classList.add('hidden');
}

async function submitChangePassword() {
  const currentPassword = document.getElementById('cp-current').value;
  const newPassword     = document.getElementById('cp-new').value;
  const errBox     = document.getElementById('cp-error');
  const successBox = document.getElementById('cp-success');
  const btn        = document.getElementById('cp-submit-btn');

  errBox.classList.add('hidden');
  successBox.classList.add('hidden');

  if (!currentPassword || !newPassword) {
    errBox.textContent = 'Please fill in both fields.';
    errBox.classList.remove('hidden');
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Updating…';

  try {
    await apiRequest('/api/auth/change-password', {
      method: 'POST',
      body: JSON.stringify({ currentPassword, newPassword }),
    });
    successBox.textContent = 'Password updated. Use it next time you log in.';
    successBox.classList.remove('hidden');
    document.getElementById('cp-current').value = '';
    document.getElementById('cp-new').value = '';
  } catch (e) {
    errBox.textContent = e.message || 'Could not update password.';
    errBox.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Update Password';
  }
}

/* ──────────────────────────────────────────────────
   AUTHENTICATION — login, logout, and role-based access control
   ════════════════════════════════════════════════════ */
// Applies a successful login response to the app's UI state.
// Shared by doLogin() (fresh login) and restoreSession() (page reload).
// Show / hide a password field (login page eye button)
function togglePasswordVisibility(inputId, btn) {
  const input = document.getElementById(inputId);
  if (!input) return;
  const show = input.type === 'password';
  input.type = show ? 'text' : 'password';
  btn.setAttribute('aria-pressed', show ? 'true' : 'false');
  const label = show ? 'Hide password' : 'Show password';
  btn.setAttribute('aria-label', label);
  btn.title = label;
}

function applySession(userData) {
  // tier controls what the frontend shows/hides.
  // 'admin'     → full access + deletions (super_admin, president, secretary, treasurer)
  // 'manager'   → add/edit, no deletions
  // 'committee' → view + notices/complaints only
  // 'resident'  → own data only
  const tierMap = {
    super_admin: 'admin',
    president:   'admin',
    secretary:   'admin',
    treasurer:   'admin',
    manager:     'manager',
    committee:   'committee',
    resident:    'resident',
  };
  const initials = userData.displayName.split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();

  loggedIn = true;
  currentUser = {
    name: userData.displayName,
    role: {
      super_admin: 'Administrator',
      president:   'President',
      secretary:   'Secretary',
      treasurer:   'Treasurer',
      manager:     'Manager',
      committee:   'Committee Member',
      resident:    'Resident',
    }[userData.role] || userData.role,
    tier: tierMap[userData.role] || 'resident',
    initials,
    houseNumber: userData.houseNumber,
    memberId: userData.memberId,
  };

  const setText = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  setText('sb-user-name', currentUser.name);
  setText('sb-user-role', currentUser.role);
  setText('sb-avatar-initials', currentUser.initials);
  const loginBtn = document.getElementById('nav-login-btn');
  if (loginBtn) {
    if (document.body.dataset.kind === 'public') {
      // Multi-page mode, on Home / About: the button becomes a way back to the dashboard
      loginBtn.textContent = 'Dashboard';
      loginBtn.setAttribute('onclick', "showPage('dashboard')");
      loginBtn.style.display = '';
    } else {
      loginBtn.style.display = 'none';
    }
  }
  const logoutBtn = document.getElementById('nav-logout-btn');
  if (logoutBtn) logoutBtn.style.display = '';
  applyPermissions();
}

// Runs once when the page first loads. If a login token was saved from
// before the reload, ask the backend "who does this belong to?" and, if
// it's still valid, silently restore the logged-in view instead of
// dumping the person back on the public home page every refresh.
async function restoreSession() {
  // Multi-page mode: <body data-kind="public|dash" data-view="..."> says what this file is.
  const kind = document.body.dataset.kind || '';
  const token = sessionStorage.getItem('kgc_token');

  if (!token) {
    if (kind === 'dash') goToFile(SITE_PAGE_FILES.login);   // dashboard pages need a login
    return;                                                  // otherwise nobody was logged in
  }

  try {
    const data = await apiRequest('/api/auth/me');
    applySession(data.user);
    if (kind === 'dash') {
      initDashPage();
    } else if (kind === 'public') {
      if (document.body.dataset.page === 'login') showPage('dashboard');   // already signed in
    } else {
      showPage('dashboard');                                // single-file mode
    }
  } catch (e) {
    // Token expired, account deactivated, or backend unreachable —
    // fall back to a clean logged-out state rather than a broken one.
    sessionStorage.removeItem('kgc_token');
    if (kind === 'dash') goToFile(SITE_PAGE_FILES.login);
  }
}

// Multi-page mode: opens this file's own dashboard section (with the role guard)
function initDashPage() {
  document.body.classList.add('auth-ok');   // reveal the page (it stays hidden until the login is confirmed)
  const view = document.body.dataset.view || 'members';
  if (view === 'detail') {
    // Only administrators and managers open individual member profiles
    if (!isAdmin() && !isManager()) { showDash('members'); return; }
    const id = new URLSearchParams(window.location.search).get('id');
    if (id) { showMemberDetail(id); return; }
    showDash('members'); return;
  }
  if (view === 'pay-upi') {
    const saved = sessionStorage.getItem('kgc_pay_upi_summary');
    const el = document.getElementById('pay-upi-bill-summary');
    if (saved && el) el.textContent = saved;
  }
  showDash(view);
}
window.addEventListener('DOMContentLoaded', restoreSession);

async function doLogin() {
  const identifier = document.getElementById('login-email').value.trim();
  const pass       = document.getElementById('login-password').value.trim();
  const err        = document.getElementById('login-error');
  const btn        = document.getElementById('login-submit-btn');

  if (!identifier || !pass) {
    err.textContent = 'Please enter both a login ID and password.';
    err.classList.remove('hidden');
    return;
  }

  err.classList.add('hidden');
  btn.disabled = true;
  btn.textContent = 'Signing in…';

  try {
    const data = await apiRequest('/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ identifier, password: pass }),
    });

    // Save the session token so future requests to the backend are authenticated,
    // AND so restoreSession() can bring you back in after a page reload.
    // sessionStorage clears when the tab closes — a reasonable default for a
    // society admin portal. The token itself also expires server-side.
    sessionStorage.setItem('kgc_token', data.token);
    applySession(data.user);
    showPage('dashboard');
  } catch (e) {
    err.textContent = e.message || 'Incorrect login ID or password. Please try again.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Sign In →';
  }
}

function logout() {
  loggedIn          = false;
  currentUser       = null;
  chartsInitialized = false;
  sessionStorage.removeItem('kgc_token'); // end the backend session too
  { const pw = document.getElementById('login-password'), tg = document.getElementById('login-pw-toggle'); if (pw) { pw.type = 'password'; pw.value = ''; } if (tg) { tg.setAttribute('aria-pressed','false'); tg.setAttribute('aria-label','Show password'); tg.title = 'Show password'; } }
  const loginBtn = document.getElementById('nav-login-btn');
  if (loginBtn) { loginBtn.textContent = 'Admin Login'; loginBtn.setAttribute('onclick', "showPage('login')"); loginBtn.style.display = ''; }
  const logoutBtn = document.getElementById('nav-logout-btn');
  if (logoutBtn) logoutBtn.style.display = 'none';
  showPage('home');
}

/* ──────────────────────────────────────────────────
   ROLE-BASED ACCESS CONTROL
   admin     → sees everything, all edit controls
   committee → Notice Board (read), Complaints (raise), Facilities
               (book Gym/Hall). No Members list, no Analytics.
   resident  → My Profile (own record only), Notice Board (read),
               Complaints (raise), Facilities (book Gym/Hall).
   ────────────────────────────────────────────────── */
function isAdmin()     { return currentUser && currentUser.tier === 'admin'; }
function isManager()   { return currentUser && currentUser.tier === 'manager'; }
function isCommittee() { return currentUser && currentUser.tier === 'committee'; }
function isResident()  { return currentUser && currentUser.tier === 'resident'; }
// Can write = admin or manager
function canWrite()    { return isAdmin() || isManager(); }
// Can delete = admin only
function canDelete()   { return isAdmin(); }

// Show / hide an element by id when it exists on this page (not every page has every element)
function setPermHidden(id, hide) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('perm-hidden', hide);
}

function applyPermissions() {
  // Tier on <body> drives CSS rules that also cover rows rendered later (ledger/expense Edit/Del columns)
  document.body.dataset.tier = (currentUser && currentUser.tier) || '';
  const admin     = isAdmin();
  const manager   = isManager();
  const committee = isCommittee();
  const resident  = isResident();
  const staff     = admin || manager || committee; // anyone non-resident

  // ── Resident-only sidebar items ──
  setPermHidden('sb-myprofile',      !resident);
  setPermHidden('sb-res-bills',      !resident);
  setPermHidden('sb-res-complaints', !resident);

  // ── Staff sidebar items (admin + manager + committee) ──
  setPermHidden('sb-members', !staff);
  setPermHidden('sb-complaints', !staff);
  setPermHidden('sb-facilities', !staff);

  // ── Notices — visible to everyone including residents ──
  setPermHidden('sb-notices', false);

  // ── Admin + manager only (not committee) ──
  // (committee gets read-only views of both; write controls are .write-only)
  setPermHidden('sb-bills', !staff);
  setPermHidden('sb-finances', !staff);

  // ── .admin-only elements: show for admin tier only (deletions, destructive actions) ──
  document.querySelectorAll('.admin-only').forEach(el =>
    el.classList.toggle('perm-hidden', !admin)
  );

  // ── .write-only elements: show for admin + manager (add/edit buttons) ──
  document.querySelectorAll('.write-only').forEach(el =>
    el.classList.toggle('perm-hidden', !(admin || manager))
  );

  // ── .admin-mgr-only: dues grid etc. — admin, manager (not committee/resident) ──
  document.querySelectorAll('.admin-mgr-only').forEach(el =>
    el.classList.toggle('perm-hidden', !(admin || manager))
  );

  // ── btn-add-member: admin + manager ──
  const addMemberBtn = document.getElementById('btn-add-member');
  if (addMemberBtn) addMemberBtn.classList.toggle('perm-hidden', !(admin || manager));
}


/* ════════════════════════════════════════════════════
   DASHBOARD SUB-VIEW ROUTING
   Controls which panel is shown in the main content area
   ════════════════════════════════════════════════════ */
function showDash(view) {
  // Route guard — committee/resident tiers cannot open Members or Analytics; committee also can't open Notice Board
  // Committee can only access: notices, complaints, facilities
  // (they cannot see members list, bills, or finances)
  const committeeOnly = ['members', 'finances', 'bills', 'notices', 'complaints', 'facilities'];
  if (isCommittee() && !committeeOnly.includes(view)) view = 'notices';

  // Residents go to their own scoped views
  if (isResident() && !['resident-home','res-bills','res-complaints','pay-upi','facilities','notices'].includes(view)) {
    view = 'resident-home';
  }
  // Staff have no use for the resident-only pages
  if (currentUser && !isResident() && ['resident-home','res-bills','res-complaints'].includes(view)) {
    view = landingView();
  }
  // Multi-page mode: the section lives in its own file — open it
  if (!document.getElementById('dash-' + view)) {
    goToFile(DASH_PAGE_FILES[view] || SITE_PAGE_FILES.home);
    return;
  }
  // All dashboard sub-views (only the ones present in this document)
  const views = [
    'members','detail','finances','notices','complaints','bills','pay-upi','facilities',
    'resident-home','res-bills','res-complaints',
  ];
  views.forEach(v => {
    const el = document.getElementById('dash-'+v);
    if (el) el.classList.toggle('hidden', v !== view);
  });

  // Clear active state on all sidebar buttons
  document.querySelectorAll('.sb-btn').forEach(b => b.classList.remove('active'));

  // Set active sidebar button
  const sbMap = {
    'members':       'sb-members',
    'finances':      'sb-finances',
    'notices':       'sb-notices',
    'complaints':    'sb-complaints',
    'bills':         'sb-bills',
    'facilities':    'sb-facilities',
    'resident-home': 'sb-myprofile',
    'res-bills':     'sb-res-bills',
    'res-complaints':'sb-res-complaints',
  };
  if (sbMap[view]) document.getElementById(sbMap[view])?.classList.add('active');

  // ── Data loading hooks ──
  if (view === 'finances') loadFinances();
  if (view === 'members')        loadMembers();
  if (view === 'notices')        loadNotices();
  if (view === 'complaints')     loadComplaints();
  if (view === 'bills')          { initLedgerYearDropdowns(); loadLedger(); loadLedgerSummary(); }
  if (view === 'facilities')     loadHallBookings();
  if (view === 'resident-home')  loadResidentHome();
  if (view === 'res-bills')      loadResBills();
  if (view === 'res-complaints') loadResComplaints();
}

// Resident-only entry point — jumps straight into their own profile
function showMyProfile() {
  showDash('resident-home');
}


/* ════════════════════════════════════════════════════
   MEMBER UTILITIES
   ════════════════════════════════════════════════════ */

// Returns human-readable unit type label
function getTypeLabel(t) {
  return {
    '2bhk':     '2 BHK (Regal / Sunshine)',
    '3bhk':     '3 BHK (Kingston)',
    'rowhouse': 'Row House (Antillia / Empress / Royale)',
    'shop':     'Shop',
  }[t] || t;
}

// Shorter label used in tight spaces like table badges
function getTypeShortLabel(t) {
  return { '2bhk': '2 BHK', '3bhk': '3 BHK', 'rowhouse': 'Row House', 'shop': 'Shop' }[t] || t;
}
// Returns one of 8 avatar background colours based on member id.
// Member ids are UUIDs (strings) from the real database, so we hash
// the string rather than assuming it's a number.
function avatarBg(id) {
  const colors = ['#1a3a6b','#c9a84c','#2d7a4f','#8b4f9e','#c0392b','#1e6b6b','#6b4f1e','#2c5f8a'];
  const str = String(id);
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = (hash * 31 + str.charCodeAt(i)) % colors.length;
  return colors[Math.abs(hash) % colors.length];
}

// Generates avatar HTML — photo if available, else initials circle
function avatarHtml(m) {
  if (m.photo) {
    return `<div class="avatar"><img src="${m.photo}" alt="${m.name}" /></div>`;
  }
  const initials = m.name.split(' ').map(n=>n[0]).join('').slice(0,2);
  return `<div class="avatar" style="background:${avatarBg(m.id)}">${initials}</div>`;
}


/* ════════════════════════════════════════════════════
   MEMBER METRICS — summary counts above the table
   ════════════════════════════════════════════════════ */
function renderMemberMetrics() {
  document.getElementById('m-total').textContent    = members.length;
  document.getElementById('m-2bhk').textContent     = members.filter(m => m.type === '2bhk').length;
  document.getElementById('m-3bhk').textContent     = members.filter(m => m.type === '3bhk').length;
  document.getElementById('m-rh').textContent       = members.filter(m => m.type === 'rowhouse').length;
  document.getElementById('m-shop').textContent     = members.filter(m => m.type === 'shop').length;
  document.getElementById('m-inactive').textContent = members.filter(m => !m.isActive).length;
}


/* ════════════════════════════════════════════════════
   MEMBERS TABLE RENDERING
   Applies active filter and search query, then renders rows.
   Payment/dues columns intentionally excluded — see the note under
   the table; that data belongs to the not-yet-wired Bills module.
   ════════════════════════════════════════════════════ */
// Natural-order comparator for house numbers: sorts by the block name
// (text prefix) first, then by the unit number as a true integer — so
// "Sunshine - 2" comes before "Sunshine - 10", and blocks group together.
// Used everywhere members are listed so new entries always slot into place.
function compareHouseNumbers(a, b) {
  const parse = s => {
    const str = String(s || '').trim();
    const m = str.match(/^(.*?)(\d+)\s*$/);      // split trailing number off the prefix
    const prefix = (m ? m[1] : str).replace(/[\s\-.]+$/,'').toLowerCase().trim();
    const num    = m ? parseInt(m[2], 10) : 0;
    return { prefix, num };
  };
  const pa = parse(a), pb = parse(b);
  if (pa.prefix !== pb.prefix) return pa.prefix.localeCompare(pb.prefix);
  return pa.num - pb.num;
}

function renderMembersTable() {
  const filtered = members.filter(m => {
    const matchType = currentFilter === 'all' || m.type === currentFilter;
    const q = currentSearch.toLowerCase();
    const matchSearch = !q || m.name.toLowerCase().includes(q) || m.house.toLowerCase().includes(q) || m.phone.includes(q) || m.email.toLowerCase().includes(q);
    return matchType && matchSearch;
  }).sort((a, b) => compareHouseNumbers(a.house, b.house));
  document.getElementById('member-count-lbl').textContent = filtered.length + ' of ' + members.length + ' members';
  const tbody = document.getElementById('members-tbody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:32px;color:var(--text2);">No members found.</td></tr>';
    return;
  }
  tbody.innerHTML = filtered.map(m => `
    <tr onclick="showMemberDetail('${m.id}')">
      <td>
        <div class="member-cell">
          ${avatarHtml(m)}
          <div>
            <div class="member-name">${m.name}</div>
            <div class="member-unit">${m.email}</div>
          </div>
        </div>
      </td>
      <td><span class="badge badge-navy">${getTypeShortLabel(m.type)}</span></td>
      <td style="font-weight:500">${m.house}</td>
      <td style="color:var(--text2);font-size:13px">${m.phone}</td>
      <td><span class="badge ${m.isActive ? 'badge-green' : 'badge-navy'}">${m.isActive ? 'Active' : 'Inactive'}</span></td>
    </tr>
  `).join('');
}

// Applies a filter tab selection
function filterMembers(filter, btn) {
  currentFilter = filter;
  document.querySelectorAll('.filter-tabs .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderMembersTable();
}

// Updates search state and re-renders table
function searchMembers(q) {
  currentSearch = q;
  renderMembersTable();
}


/* ════════════════════════════════════════════════════
   MEMBER DETAIL PAGE
   Shows full profile for a single member.
   If tenant, displays owner name in an amber banner.
   ════════════════════════════════════════════════════ */
async function showMemberDetail(id) {
  if (isCommittee()) return; // committee sees the members list only, not individual profiles
  if (!document.getElementById('dash-detail')) {            // multi-page mode: profile has its own page
    goToFile('member.html?id=' + encodeURIComponent(id));
    return;
  }
  currentDetailMemberId = id;
  const container = document.getElementById('detail-content');
  container.innerHTML = '<div style="padding:48px;text-align:center;color:var(--text2);">Loading…</div>';
  showDash('detail');
  window.scrollTo(0, 0);

  let m;
  try {
    m = await apiRequest(`/api/members/${id}`);
  } catch (e) {
    container.innerHTML = `<div style="padding:48px;text-align:center;color:var(--red);">Could not load this member: ${e.message}</div>`;
    return;
  }
  currentDetailMember = m; // stashed for openEditModal()

  const initials = m.full_name.split(' ').map(n=>n[0]).join('').slice(0,2);
  const photoHtml = `<div class="detail-photo" style="font-size:72px;opacity:0.25;">${initials}</div>`;
  const vehicles = [m.vehicle_1_reg, m.vehicle_2_reg].filter(Boolean).join(', ') || '—';
  const emergency = m.emergency_contact_name
    ? `${m.emergency_contact_name}${m.emergency_contact_phone ? ' – ' + m.emergency_contact_phone : ''}`
    : '—';
  const actionButtons = (isAdmin() || isManager()) ? `
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:16px;">
      <button class="btn-sm btn-sm-outline" onclick="openEditModal()">✏️ Edit Profile</button>
      <button class="btn-sm btn-sm-outline" onclick="toggleMemberActive('${m.id}', ${!m.is_active})">${m.is_active ? '🚫 Mark Inactive' : '✅ Reactivate'}</button>
      ${isAdmin() && m.is_active ? `<button class="btn-sm btn-sm-outline" onclick="openTransferModal()">🏠 Transfer Ownership</button>` : ''}
      ${isAdmin() ? `<button class="btn-sm btn-sm-red" onclick="deleteMember('${m.id}')">🗑️ Delete</button>` : ''}
    </div>` : '';

  container.innerHTML = `
    <div class="detail-card">
      ${photoHtml}
      <div class="detail-info">
        <h2 class="detail-name">${m.full_name}</h2>
        <p class="detail-type">${getTypeLabel(m.unit_type)} &nbsp;·&nbsp; ${m.house_number}</p>
        <div class="detail-rows">
          <div class="detail-row"><span class="detail-key">Phone</span><span class="detail-val">${m.phone_primary}</span></div>
          ${m.phone_secondary ? `<div class="detail-row"><span class="detail-key">Alternate Phone</span><span class="detail-val">${m.phone_secondary}</span></div>` : ''}
          <div class="detail-row"><span class="detail-key">Email</span><span class="detail-val" style="font-size:13px">${m.email || '—'}</span></div>
          <div class="detail-row"><span class="detail-key">Vehicle(s)</span><span class="detail-val">${vehicles}</span></div>
          <div class="detail-row"><span class="detail-key">Emergency Contact</span><span class="detail-val" style="font-size:12px;max-width:160px;text-align:right">${emergency}</span></div>
          <div class="detail-row"><span class="detail-key">Status</span>
            <span class="badge ${m.is_active ? 'badge-green' : 'badge-navy'}">${m.is_active ? 'Active' : 'Inactive'}</span>
          </div>
        </div>
        ${actionButtons}
      </div>
    </div>
    <div class="detail-right">
      <div class="info-section" id="member-bills-section">
        <h3>🧾 Maintenance & Payments</h3>
        <div id="member-bills-content" style="text-align:center;padding:20px;color:var(--text2);">Loading billing history…</div>
      </div>
      ${(isAdmin() || isManager()) ? `<div class="info-section" id="house-history-section">
        <h3>🏠 House Maintenance History</h3>
        <div id="house-history-content" style="padding:12px 0;color:var(--text2);font-size:13px;">Loading…</div>
      </div>` : ''}
    </div>
  `;
  applyPermissions();
  loadMemberBillsSection(m.id);
  if (isAdmin() || isManager()) loadHouseHistory(m.id);
}


/* ════════════════════════════════════════════════════
   ADD MEMBER MODAL — open / close / submit
   ════════════════════════════════════════════════════ */
function openAddModal() {
  editingMemberId = null;
  document.getElementById('add-modal-title').textContent = 'Add New Member';
  document.getElementById('add-member-submit-btn').textContent = 'Add Member';
  document.getElementById('add-modal').classList.remove('hidden');
  document.getElementById('modal-error').classList.add('hidden');
  // Reset all form fields
  ['f-name','f-house','f-phone','f-email','f-vehicle','f-emergency-name','f-emergency-phone'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('f-type').value   = '';
  // Reset photo preview
  document.getElementById('f-photo-preview').style.display = 'none';
  document.getElementById('f-upload-icon').style.display   = '';
  document.getElementById('f-photo-input').value           = '';
}

// Opens the same modal pre-filled with the currently-viewed member's data,
// switched into "edit" mode (submits a PUT instead of a POST).
function openEditModal() {
  const m = currentDetailMember;
  if (!m) return;
  editingMemberId = m.id;
  document.getElementById('add-modal-title').textContent = 'Edit Member Profile';
  document.getElementById('add-member-submit-btn').textContent = 'Save Changes';
  document.getElementById('add-modal').classList.remove('hidden');
  document.getElementById('modal-error').classList.add('hidden');

  document.getElementById('f-name').value   = m.full_name || '';
  document.getElementById('f-house').value  = m.house_number || '';
  document.getElementById('f-type').value   = m.unit_type || '';
  document.getElementById('f-phone').value  = m.phone_primary || '';
  document.getElementById('f-email').value  = m.email || '';
  document.getElementById('f-vehicle').value = m.vehicle_1_reg || '';
  document.getElementById('f-emergency-name').value  = m.emergency_contact_name || '';
  document.getElementById('f-emergency-phone').value = m.emergency_contact_phone || '';
  // Photo editing isn't wired to storage yet — keep the preview blank
  document.getElementById('f-photo-preview').style.display = 'none';
  document.getElementById('f-upload-icon').style.display   = '';
  document.getElementById('f-photo-input').value           = '';
}

function closeAddModal() {
  document.getElementById('add-modal').classList.add('hidden');
  editingMemberId = null;
}

// Handles photo file selection — preview only for now; not yet sent to the
// backend, since real photo storage (Supabase Storage) isn't wired up yet.
function previewPhoto(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const preview = document.getElementById('f-photo-preview');
    preview.src = e.target.result;
    preview.style.display = 'block';
    document.getElementById('f-upload-icon').style.display = 'none';
  };
  reader.readAsDataURL(file);
}

// Validates and sends a new (or edited) member to the real backend.
// Behaves as POST (add) or PUT (edit) depending on editingMemberId.
async function addMember() {
  const name  = document.getElementById('f-name').value.trim();
  const house = document.getElementById('f-house').value.trim();
  const type  = document.getElementById('f-type').value;
  const phone = document.getElementById('f-phone').value.trim();
  const err   = document.getElementById('modal-error');

  if (!name || !house || !type || !phone) {
    err.textContent = 'Please fill in all required fields (marked *).';
    err.classList.remove('hidden');
    return;
  }
  err.classList.add('hidden');

  const btn = document.getElementById('add-member-submit-btn');
  const isEditing = !!editingMemberId;
  if (btn) { btn.disabled = true; btn.textContent = isEditing ? 'Saving…' : 'Adding…'; }

  const payload = {
    fullName: name,
    houseNumber: house,
    unitType: type,
    phonePrimary: phone,
    email: document.getElementById('f-email').value.trim() || null,
    vehicle1Reg: document.getElementById('f-vehicle').value.trim() || null,
    emergencyContactName: document.getElementById('f-emergency-name').value.trim() || null,
    emergencyContactPhone: document.getElementById('f-emergency-phone').value.trim() || null,
  };

  try {
    if (isEditing) {
      await apiRequest(`/api/members/${editingMemberId}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      await apiRequest('/api/members', { method: 'POST', body: JSON.stringify(payload) });
    }
    const wasEditingId = editingMemberId;
    closeAddModal();
    await loadMembers();
    if (wasEditingId) await showMemberDetail(wasEditingId); // refresh the detail view with saved changes
    showToast(isEditing ? 'Member updated.' : 'Member added.', 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not save member.';
    err.classList.remove('hidden');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = isEditing ? 'Save Changes' : 'Add Member'; }
  }
}

// Marks a member active/inactive — the recommended way to "remove" someone
// who has existing history (complaints, bills, etc.) that a hard delete
// would be blocked from touching anyway.
async function toggleMemberActive(id, makeActive) {
  const confirmMsg = makeActive
    ? 'Reactivate this member?'
    : 'Mark this member as inactive? They\'ll be hidden from active views but their history is preserved.';
  if (!confirm(confirmMsg)) return;
  try {
    await apiRequest(`/api/members/${id}/active`, { method: 'PATCH', body: JSON.stringify({ isActive: makeActive }) });
    await loadMembers();
    await showMemberDetail(id);
    showToast(makeActive ? 'Member reactivated.' : 'Member deactivated.', 'ok');
  } catch (e) {
    alert('Could not update status: ' + e.message);
  }
}

// Permanently deletes a member. The database itself blocks this if the
// member has any existing complaints, bills, bookings, etc. — in that
// case the backend returns a clear message suggesting deactivation instead.
async function deleteMember(id) {
  if (!confirm('Permanently delete this member? This cannot be undone.')) return;
  try {
    await apiRequest(`/api/members/${id}`, { method: 'DELETE' });
    await loadMembers();
    showDash('members');
    showToast('Member deleted.', 'ok');
  } catch (e) {
    // If the block is specifically because of existing history, offer the
    // separate, more dangerous "force delete" path — which needs a much
    // stronger confirmation since it destroys that history permanently.
    if (e.message.includes('existing records')) {
      offerForceDelete(id);
    } else {
      alert(e.message);
    }
  }
}

// Shown only when a normal delete was blocked by existing history.
// Requires typing the member's exact name — a deliberately higher bar
// than a single "OK" click, since this permanently destroys complaints,
// bills, bookings, etc. along with the member.
async function offerForceDelete(id) {
  const m = currentDetailMember;
  const wantsForce = confirm(
    `${m.full_name} has existing records (complaints, bills, bookings, etc.) that a normal delete won't touch.\n\n` +
    `Click OK to permanently delete ${m.full_name} AND all of their related records forever.\n` +
    `Click Cancel to keep everything and mark them Inactive instead (recommended).`
  );
  if (!wantsForce) return;

  const typed = prompt(`To confirm, type this member's exact name: ${m.full_name}`);
  if (typed !== m.full_name) {
    alert('Name did not match exactly — nothing was deleted.');
    return;
  }

  try {
    const result = await apiRequest(`/api/members/${id}/force`, { method: 'DELETE' });
    const counts = result.deletedCounts;
    const summary = Object.entries(counts).filter(([,v]) => v > 0).map(([k,v]) => `${v} ${k}`).join(', ') || 'no related records';
    showToast(`Member deleted — also removed ${summary}.`, 'ok');
    await loadMembers();
    showDash('members');
  } catch (e) {
    alert('Force delete failed: ' + e.message);
  }
}


/* ════════════════════════════════════════════════════
   MAINTENANCE BILLS
   ════════════════════════════════════════════════════ */
let bills = [];
let billsFilter = 'all';
let currentPaymentBill = null; // bill currently open in the Record Payment modal

const billStatusColors = { unpaid: 'badge-red', partial: 'badge-amber', paid: 'badge-green', waived: 'badge-navy' };

async function loadBills(memberId) {
  const tbody = document.getElementById('bills-tbody');
  if (tbody) tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--text2);">Loading…</td></tr>';
  try {
    const query = memberId ? `?memberId=${memberId}` : '';
    bills = await apiRequest(`/api/bills${query}`);
    renderBills();
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--red);">Could not load bills: ${e.message}</td></tr>`;
  }
}

function renderBills() {
  const tbody = document.getElementById('bills-tbody');
  if (!tbody) return; // not on the bills page (e.g. called from member detail context elsewhere)
  const filtered = bills.filter(b => billsFilter === 'all' || b.status === billsFilter);
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--text2);">No bills found.</td></tr>';
    return;
  }
  tbody.innerHTML = filtered.map(b => {
    const monthLabel = new Date(b.billing_month).toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
    const canPay = (isAdmin() || isCommittee()) && b.status !== 'paid' && b.status !== 'waived';
    return `
      <tr>
        <td style="font-weight:500">${b.house_number}</td>
        <td>${b.member_name}</td>
        <td>${monthLabel}</td>
        <td>₹${parseFloat(b.total_amount_due).toLocaleString('en-IN')}</td>
        <td style="color:var(--green)">₹${parseFloat(b.amount_paid).toLocaleString('en-IN')}</td>
        <td style="color:${parseFloat(b.balance_due) > 0 ? 'var(--red)' : 'var(--text2)'}">₹${parseFloat(b.balance_due).toLocaleString('en-IN')}</td>
        <td><span class="badge ${billStatusColors[b.status]}">${b.status.charAt(0).toUpperCase()+b.status.slice(1)}</span></td>
        <td class="admin-only">
          <div style="display:flex;gap:6px;flex-wrap:wrap;">
            ${canPay ? `<button class="btn-sm btn-sm-outline" onclick="openPaymentModal('${b.id}')">Record Payment</button>` : ''}
            ${parseFloat(b.amount_paid) > 0 ? `<button class="btn-sm btn-sm-outline" onclick="viewBillPayments('${b.id}')">🧾 Payments</button>` : ''}
            ${canDelete() ? `<button class="btn-sm btn-sm-outline" onclick="openEditBillModal('${b.id}','${b.house_number}','${b.member_name}','${b.billing_month}',${b.total_amount_due},${b.amount_paid},'${b.status}')">✏️ Edit</button>` : ''}
          </div>
        </td>
      </tr>`;
  }).join('');
}

function filterBills(status, btn) {
  billsFilter = status;
  document.querySelectorAll('#dash-bills .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderBills();
}

/* ── View Bill Payments modal (committee/admin download access from the main Bills page) ── */
async function viewBillPayments(billId) {
  document.getElementById('bill-payments-modal').classList.remove('hidden');
  const box = document.getElementById('bill-payments-content');
  box.innerHTML = '<div style="text-align:center;padding:24px;color:var(--text2);">Loading…</div>';
  try {
    const bill = await apiRequest(`/api/bills/${billId}`);
    memberBillsDetailCache[billId] = bill; // reuse the same cache downloadReceipt() reads from
    if (!bill.payments || bill.payments.length === 0) {
      box.innerHTML = '<p style="font-size:13px;color:var(--text2);">No payments recorded yet for this bill.</p>';
      return;
    }
    box.innerHTML = bill.payments.map(p => {
      const payDate = new Date(p.payment_date).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
      return `<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);font-size:13px;">
        <div>
          <div style="font-weight:500">₹${parseFloat(p.amount_paid).toLocaleString('en-IN')} — ${p.payment_mode.replace('_',' ')}</div>
          <div style="color:var(--text2);font-size:11px;">${payDate} · Receipt: ${p.receipt_number_auto}${p.receipt_number_manual ? ' · Bill No: ' + p.receipt_number_manual : ''}${p.bill_book_number ? ' · Book: ' + p.bill_book_number : ''}</div>
        </div>
        <button class="btn-sm btn-sm-outline" onclick="downloadReceipt('${billId}','${p.id}')">⬇ Receipt</button>
      </div>`;
    }).join('');
  } catch (e) {
    box.innerHTML = `<p style="font-size:13px;color:var(--red);">Could not load payments: ${e.message}</p>`;
  }
}
function closeBillPaymentsModal() { document.getElementById('bill-payments-modal').classList.add('hidden'); }

/* ── Generate Bills modal ── */
function openGenerateBillsModal() {
  document.getElementById('generate-bills-modal').classList.remove('hidden');
  document.getElementById('gb-error').classList.add('hidden');
  document.getElementById('gb-result').classList.add('hidden');
  const now = new Date();
  document.getElementById('gb-month').value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  populateFYDropdown('gb-fy');
}

// Populate a financial year <select> with 100 years:
// 50 years back and 50 years forward from the current FY.
// Indian financial year runs Apr–Mar, so FY for a given calendar
// year Y is "Y-YY" (e.g. 2025-26).
// The current FY is selected by default.
function populateFYDropdown(selectId) {
  const sel = document.getElementById(selectId);
  if (!sel) return;

  const now       = new Date();
  // Current FY start year: if we're in Jan–Mar, FY started last calendar year
  const curFYStart = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;

  const startYear = 2018;             // Society established 2018–19
  const endYear   = curFYStart + 50;  // 50 years forward

  sel.innerHTML = '';
  for (let y = endYear; y >= startYear; y--) {
    const value   = `${y}-${String(y + 1).slice(-2)}`;   // e.g. "2025-26"
    const label   = `${y}–${y + 1}`;                      // e.g. "2025–2026"
    const opt     = document.createElement('option');
    opt.value     = value;
    opt.textContent = label;
    if (y === curFYStart) opt.selected = true;
    sel.appendChild(opt);
  }
}
function closeGenerateBillsModal() { document.getElementById('generate-bills-modal').classList.add('hidden'); }

async function submitGenerateBills() {
  const monthVal = document.getElementById('gb-month').value; // "YYYY-MM"
  const fy = document.getElementById('gb-fy').value;
  const err = document.getElementById('gb-error');
  const resultBox = document.getElementById('gb-result');
  if (!monthVal) { err.textContent = 'Please choose a billing month.'; err.classList.remove('hidden'); return; }
  err.classList.add('hidden');
  resultBox.classList.add('hidden');

  const btn = document.getElementById('gb-submit-btn');
  btn.disabled = true; btn.textContent = 'Generating…';

  try {
    const result = await apiRequest('/api/bills/generate', {
      method: 'POST',
      body: JSON.stringify({ billingMonth: `${monthVal}-01`, financialYear: fy }),
    });
    resultBox.innerHTML = `✅ Generated <strong>${result.generated}</strong> new bill(s). ` +
      (result.skippedExisting > 0 ? `${result.skippedExisting} already existed. ` : '') +
      (result.skippedNoRate > 0 ? `⚠️ ${result.skippedNoRate} member(s) skipped — no rate set for their unit type/FY.` : '');
    resultBox.classList.remove('hidden');
    await loadBills();
  } catch (e) {
    err.textContent = e.message || 'Could not generate bills.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Generate';
  }
}

/* ── Record Payment modal ── */
function openPaymentModal(billId) {
  currentPaymentBill = bills.find(b => b.id === billId);
  if (!currentPaymentBill) return;
  const b = currentPaymentBill;
  document.getElementById('pay-bill-summary').textContent =
    `${b.member_name} (${b.house_number}) — Balance due: ₹${parseFloat(b.balance_due).toLocaleString('en-IN')}`;
  document.getElementById('pay-amount').value = b.balance_due;
  document.getElementById('pay-date').value = new Date().toISOString().split('T')[0];
  document.getElementById('pay-mode').value = 'cash';
  document.getElementById('pay-remarks').value = '';
  document.getElementById('pay-receipt-manual').value = '';
  document.getElementById('pay-bill-book').value = '';
  ['pay-cheque-number','pay-cheque-date','pay-bank-name','pay-upi-txn','pay-upi-app',
   'pay-bank-ref','pay-transfer-bank','pay-dd-number','pay-dd-date','pay-dd-bank'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  onPaymentModeChange();
  document.getElementById('pay-error').classList.add('hidden');
  document.getElementById('payment-modal').classList.remove('hidden');
}
function closePaymentModal() { document.getElementById('payment-modal').classList.add('hidden'); currentPaymentBill = null; }

// Shows/hides the payment-mode-specific field groups
function onPaymentModeChange() {
  const mode = document.getElementById('pay-mode').value;
  document.getElementById('pay-fields-cheque').classList.toggle('hidden', mode !== 'cheque');
  document.getElementById('pay-fields-upi').classList.toggle('hidden', mode !== 'online_upi');
  document.getElementById('pay-fields-transfer').classList.toggle('hidden', !['online_neft','online_imps','online_rtgs'].includes(mode));
  document.getElementById('pay-fields-dd').classList.toggle('hidden', mode !== 'demand_draft');
}

async function submitPayment() {
  if (!currentPaymentBill) return;
  const err = document.getElementById('pay-error');
  err.classList.add('hidden');

  const amountPaid = parseFloat(document.getElementById('pay-amount').value);
  const paymentMode = document.getElementById('pay-mode').value;
  if (!amountPaid || amountPaid <= 0) { err.textContent = 'Please enter a valid amount.'; err.classList.remove('hidden'); return; }

  const payload = {
    amountPaid, paymentMode,
    paymentDate: document.getElementById('pay-date').value || undefined,
    remarks: document.getElementById('pay-remarks').value.trim() || undefined,
    receiptNumberManual: document.getElementById('pay-receipt-manual').value.trim() || undefined,
    billBookNumber: document.getElementById('pay-bill-book').value.trim() || undefined,
  };
  if (paymentMode === 'cheque') {
    payload.chequeNumber = document.getElementById('pay-cheque-number').value.trim();
    payload.chequeDate = document.getElementById('pay-cheque-date').value || undefined;
    payload.bankName = document.getElementById('pay-bank-name').value.trim() || undefined;
  } else if (paymentMode === 'online_upi') {
    payload.upiTransactionId = document.getElementById('pay-upi-txn').value.trim();
    payload.upiApp = document.getElementById('pay-upi-app').value.trim() || undefined;
  } else if (['online_neft','online_imps','online_rtgs'].includes(paymentMode)) {
    payload.bankReferenceNumber = document.getElementById('pay-bank-ref').value.trim();
    payload.transferBankName = document.getElementById('pay-transfer-bank').value.trim() || undefined;
  } else if (paymentMode === 'demand_draft') {
    payload.demandDraftNumber = document.getElementById('pay-dd-number').value.trim() || undefined;
    payload.ddDate = document.getElementById('pay-dd-date').value || undefined;
    payload.ddBankName = document.getElementById('pay-dd-bank').value.trim() || undefined;
  }

  const btn = document.getElementById('pay-submit-btn');
  btn.disabled = true; btn.textContent = 'Saving…';
  try {
    const result = await apiRequest(`/api/bills/${currentPaymentBill.id}/payments`, {
      method: 'POST', body: JSON.stringify(payload),
    });
    closePaymentModal();
    await loadBills();
    showToast(`Payment recorded — receipt ${result.receiptNumber}. Bill is now ${result.newStatus}.`, 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not record payment.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Record Payment';
  }
}

// Fetches and renders a specific member's bill history — used on the
// member detail / My Profile page (not the main Bills list page).
// Cache of full bill+payments objects, keyed by bill id — populated by
// loadMemberBillsSection() and read by downloadReceipt() so the PDF
// generator has every field it needs without re-fetching per click.
let memberBillsDetailCache = {};

async function loadMemberBillsSection(memberId) {
  const box = document.getElementById('member-bills-content');
  if (!box) return;
  try {
    const memberBills = await apiRequest(`/api/bills?memberId=${memberId}`);
    if (memberBills.length === 0) {
      box.innerHTML = '<p style="font-size:13px;color:var(--text2);">No bills generated yet for this member.</p>';
      return;
    }
    // Fetch full detail (including each payment's receipt numbers, mode-specific
    // fields, etc.) for every bill, in parallel — needed for the receipt PDFs.
    const detailedBills = await Promise.all(memberBills.map(b => apiRequest(`/api/bills/${b.id}`)));
    detailedBills.forEach(b => { memberBillsDetailCache[b.id] = b; });

    const totalDue = memberBills.reduce((s,b) => s + parseFloat(b.balance_due), 0);
    const canPay = isAdmin() || isCommittee();
    const showUpiPayButton = isResident();
    box.innerHTML = `
      <div class="detail-row"><span class="detail-key">Outstanding Balance</span>
        <span class="detail-val" style="color:${totalDue > 0 ? 'var(--red)' : 'var(--green)'}">₹${totalDue.toLocaleString('en-IN')}</span>
      </div>
      <div style="max-height:340px;overflow-y:auto;margin-top:10px;">
        ${detailedBills.map(b => {
          const monthLabel = new Date(b.billing_month).toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
          const paymentRows = (b.payments || []).map(p => {
            const payDate = new Date(p.payment_date).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
            return `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0 6px 14px;font-size:12px;color:var(--text2);border-top:1px dashed var(--border);">
              <div>₹${parseFloat(p.amount_paid).toLocaleString('en-IN')} paid ${payDate} · ${p.receipt_number_auto}</div>
              <button class="btn-sm btn-sm-outline" onclick="downloadReceipt('${b.id}','${p.id}')">⬇ Receipt</button>
            </div>`;
          }).join('');
          return `<div style="padding:8px 0;border-bottom:1px solid var(--border);font-size:13px;">
            <div style="display:flex;justify-content:space-between;align-items:center;">
              <div>
                <div style="font-weight:500">${monthLabel}</div>
                <div style="color:var(--text2);font-size:11px;">₹${parseFloat(b.total_amount_due).toLocaleString('en-IN')} due</div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;">
                <span class="badge ${billStatusColors[b.status]}">${b.status}</span>
                ${canPay && b.status !== 'paid' && b.status !== 'waived' ? `<button class="btn-sm btn-sm-outline" onclick="jumpToPayBill('${b.id}','${memberId}')">Record Payment</button>` : ''}
                ${showUpiPayButton && b.status !== 'paid' && b.status !== 'waived' ? `<button class="btn-sm" style="background:var(--gold-bg,#fdf3d8);color:var(--amber,#8a6d1a);" onclick="showUpiPaymentPage('${b.id}', '${monthLabel}', '${parseFloat(b.balance_due).toLocaleString('en-IN')}')">Pay</button>` : ''}
              </div>
            </div>
            ${paymentRows}
          </div>`;
        }).join('')}
      </div>
    `;
  } catch (e) {
    box.innerHTML = `<p style="font-size:13px;color:var(--red);">Could not load billing history: ${e.message}</p>`;
  }
}

// Builds a real, downloadable PDF receipt for one specific payment,
// showing the auto-generated receipt number, the manual bill no., and
// the bill book no. side by side — matching the physical receipt book
// so both records stay traceable to each other.
function downloadReceipt(billId, paymentId) {
  const bill = memberBillsDetailCache[billId];
  if (!bill) { alert('Bill details not loaded yet — please reopen this profile and try again.'); return; }
  const p = bill.payments.find(x => x.id === paymentId);
  if (!p) { alert('Payment not found.'); return; }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a5' });
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = 40;

  doc.setFont('helvetica', 'bold'); doc.setFontSize(16);
  doc.text('Krishna Grand City', pageWidth/2, y, { align: 'center' }); y += 20;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(11);
  doc.text('Maintenance Payment Receipt', pageWidth/2, y, { align: 'center' }); y += 24;
  doc.setDrawColor(180); doc.line(30, y, pageWidth-30, y); y += 22;

  const row = (label, value) => {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.text(label, 34, y);
    doc.setFont('helvetica', 'normal'); doc.text(String(value ?? '—'), 220, y);
    y += 18;
  };

  row('Receipt No. (auto):', p.receipt_number_auto);
  row('Bill No.:', p.receipt_number_manual || '—');
  row('Bill Book No.:', p.bill_book_number || '—');
  y += 6; doc.line(30, y, pageWidth-30, y); y += 20;

  row('Resident:', `${bill.member_name} (${bill.house_number})`);
  row('Billing Month:', new Date(bill.billing_month).toLocaleDateString('en-IN', {month:'long', year:'numeric'}));
  row('Financial Year:', bill.financial_year);
  y += 6; doc.line(30, y, pageWidth-30, y); y += 20;

  row('Payment Date:', new Date(p.payment_date).toLocaleDateString('en-IN', {day:'numeric', month:'long', year:'numeric'}));
  row('Payment Mode:', p.payment_mode.replace('_',' ').replace(/\b\w/g, c => c.toUpperCase()));

  // Mode-specific reference detail
  if (p.payment_mode === 'cheque') row('Cheque No. / Date:', `${p.cheque_number || '—'} / ${p.cheque_date ? new Date(p.cheque_date).toLocaleDateString('en-IN') : '—'}`);
  if (p.payment_mode === 'online_upi') row('UPI Transaction ID:', p.upi_transaction_id || '—');
  if (['online_neft','online_imps','online_rtgs'].includes(p.payment_mode)) row('Bank Reference (UTR):', p.bank_reference_number || '—');
  if (p.payment_mode === 'demand_draft') row('DD No. / Date:', `${p.demand_draft_number || '—'} / ${p.dd_date ? new Date(p.dd_date).toLocaleDateString('en-IN') : '—'}`);

  y += 6; doc.line(30, y, pageWidth-30, y); y += 20;
  // Note: using "Rs." rather than the ₹ symbol here specifically — jsPDF's
  // built-in fonts don't include the Rupee glyph, so ₹ renders as a broken
  // character in the PDF even though it displays fine everywhere else in the app.
  row('Amount Paid:', `Rs. ${parseFloat(p.amount_paid).toLocaleString('en-IN')}`);
  row('Balance After This Payment:', `Rs. ${parseFloat(p.balance_outstanding).toLocaleString('en-IN')}`);
  if (p.remarks) row('Remarks:', p.remarks);
  row('Collected By:', p.collected_by_name || '—');

  y += 20;
  doc.setFontSize(9); doc.setTextColor(130);
  doc.text('This is a system-generated receipt from the Krishna Grand City resident portal.', pageWidth/2, y, { align: 'center' });

  doc.save(`Receipt_${p.receipt_number_auto}.pdf`);
}

// Jumps from a member's profile page to the Bills page's payment modal
// for a specific bill — loads that member's bills first so the bill
// object is available for openPaymentModal to find.
async function jumpToPayBill(billId, memberId) {
  await loadBills(memberId);
  openPaymentModal(billId);
}

// Takes a resident to the (placeholder) online payment page for a specific
// bill. Real UPI integration isn't built yet — this just shows where the
// QR code / UPI ID will go once that's wired up.
function showUpiPaymentPage(billId, monthLabel, balanceDue) {
  const text = `${monthLabel} — ₹${balanceDue} due`;
  const el = document.getElementById('pay-upi-bill-summary');
  if (el) el.textContent = text;
  else sessionStorage.setItem('kgc_pay_upi_summary', text);   // multi-page mode: handed to pay-upi.html
  showDash('pay-upi');
}


/* ════════════════════════════════════════════════════
   ANALYTICS — chart initialization and data refresh
   Collected / Pending / Payment Status / 3-year Trend now pull from
   REAL bill data. Expenses / Corpus Fund / expense breakdown remain
   manual entry or file-upload — no expense-tracking backend exists yet.
   ════════════════════════════════════════════════════ */

// Converts a bill's billing_month (calendar date) into its position (0-11)
// within the Indian financial-year month order: Apr=0, May=1, ..., Mar=11.
function fyMonthIndex(billingMonthISO) {
  const calendarMonth = new Date(billingMonthISO).getMonth(); // 0=Jan..11=Dec
  return (calendarMonth - 3 + 12) % 12;
}

// Fetches real bills for one financial year and reduces them into
// everything the charts need: totals scoped to the selected quarter/year,
// a full 12-month collected breakdown, and status counts for the donut.
// Fails safe (all zeros) rather than crashing the page if the fetch fails.
async function loadBillingAnalyticsData(year, period) {
  const fy = `${year}-${String(year + 1).slice(2)}`;
  const zeroResult = {
    collectedL: 0, pendingL: 0,
    monthlyL: new Array(period === 'yearly' ? 12 : 3).fill(0),
    statusCounts: { paid: 0, partial: 0, unpaid: 0, waived: 0 },
  };
  try {
    const rows = await apiRequest(`/api/bills?financialYear=${fy}`);
    const monthlyRupees = new Array(12).fill(0);
    const statusCounts = { paid: 0, partial: 0, unpaid: 0, waived: 0 };
    const qIdx = period === 'yearly' ? -1 : QUARTERS.indexOf(period);
    let scopedCollected = 0, scopedPending = 0;

    rows.forEach(b => {
      const idx = fyMonthIndex(b.billing_month);
      const paid = parseFloat(b.amount_paid) || 0;
      const pending = parseFloat(b.balance_due) || 0;
      monthlyRupees[idx] += paid;

      const inScope = period === 'yearly' || Math.floor(idx / 3) === qIdx;
      if (inScope) {
        scopedCollected += paid;
        scopedPending += pending;
        if (statusCounts[b.status] !== undefined) statusCounts[b.status]++;
      }
    });

    const monthlyL = period === 'yearly'
      ? monthlyRupees.map(v => v / 100000)
      : monthlyRupees.slice(qIdx * 3, qIdx * 3 + 3).map(v => v / 100000);

    return { collectedL: scopedCollected / 100000, pendingL: scopedPending / 100000, monthlyL, statusCounts };
  } catch (e) {
    console.error('Billing analytics fetch failed:', e);
    return zeroResult;
  }
}

// Real 3-year collections trend — fetches each of the last 3 years' yearly totals.
async function getRealTrendData() {
  const years = AVAILABLE_YEARS.slice(-3);
  const results = await Promise.all(years.map(y => loadBillingAnalyticsData(y, 'yearly')));
  return results.map(r => r.collectedL);
}
function getTrendLabels() {
  return AVAILABLE_YEARS.slice(-3).map(y => y + '–' + String(y+1).slice(2));
}

async function initCharts() {
  // Destroy existing instances before recreating (prevents duplicates)
  Object.values(chartInstances).forEach(c => c.destroy());
  chartInstances = {};
  const manualData = getCurrentAnalyticsData(); // expenses/corpus/expense-category breakdown only — still manual
  const real = await loadBillingAnalyticsData(analyticsYear, analyticsPeriod);
  const monthsAll = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
  const qIdx = QUARTERS.indexOf(analyticsPeriod);
  const months = analyticsPeriod === 'yearly' ? monthsAll : monthsAll.slice(qIdx*3, qIdx*3+3);
  const expColors = ['#0f1f3d','#1a3060','#c9a84c','#2d7a4f','#7b5ea7','#b8870f','#5a5a6e'];

  // Update KPI cards — Collected/Pending from real bills, Expenses/Corpus still manual
  document.getElementById('akpi-collected').textContent = '₹'+real.collectedL.toFixed(1)+'L';
  document.getElementById('akpi-pending').textContent   = '₹'+real.pendingL.toFixed(1)+'L';
  document.getElementById('akpi-expenses').textContent  = '₹'+manualData.expenses+'L';
  document.getElementById('akpi-corpus').textContent    = '₹'+manualData.corpus+'L';

  // Collections bar chart — now real, from actual bill payments
  chartInstances.collections = new Chart(document.getElementById('chart-collections'), {
    type: 'bar',
    data: { labels: months, datasets: [{ label: 'Collection (₹L)', data: real.monthlyL, backgroundColor: '#0f1f3d', borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { callback: v => '₹'+v+'L' } } } }
  });

  // Unit type donut chart — computed from members data (not period-dependent)
  chartInstances.units = new Chart(document.getElementById('chart-units'), {
    type: 'doughnut',
    data: { labels: ['2 BHK','3 BHK','Row House','Shop'], datasets: [{ data: [members.filter(m=>m.type==='2bhk').length, members.filter(m=>m.type==='3bhk').length, members.filter(m=>m.type==='rowhouse').length, members.filter(m=>m.type==='shop').length], backgroundColor: ['#0f1f3d','#c9a84c','#2d7a4f','#8b4f9e'], borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, cutout: '65%' }
  });

  // Expense breakdown horizontal bar chart — still manual entry / upload
  chartInstances.expense = new Chart(document.getElementById('chart-expense'), {
    type: 'bar',
    data: { labels: EXPENSE_CATEGORY_LABELS, datasets: [{ label: '₹ Lakhs', data: [...manualData.expenses_cat], backgroundColor: expColors, borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { callback: v => '₹'+v+'L' } } } }
  });

  // Payment status donut — now real, from actual bill statuses
  chartInstances.payment = await buildPaymentChart();

  // 3-year trend line chart — now real, yearly totals for the most recent 3 years
  chartInstances.trend = new Chart(document.getElementById('chart-trend'), {
    type: 'line',
    data: { labels: getTrendLabels(), datasets: [{ label: 'Collections (₹L)', data: await getRealTrendData(), borderColor: '#c9a84c', backgroundColor: 'rgba(201,168,76,0.08)', tension: 0.4, fill: true, pointBackgroundColor: '#c9a84c', pointRadius: 6, pointHoverRadius: 8 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: false, ticks: { callback: v => '₹'+v+'L' } } } }
  });
}

// Builds the payment status donut from real bill status counts
// (Paid / Partial / Unpaid / Waived — the actual bill_status values).
async function buildPaymentChart() {
  const real = await loadBillingAnalyticsData(analyticsYear, analyticsPeriod);
  const c = real.statusCounts;
  return new Chart(document.getElementById('chart-payment'), {
    type: 'doughnut',
    data: { labels: ['Paid','Partial','Unpaid','Waived'], datasets: [{ data: [c.paid, c.partial, c.unpaid, c.waived], backgroundColor: ['#2d7a4f','#b8870f','#c0392b','#5a5a6e'], borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, cutout: '65%' }
  });
}

// Refreshes only the payment status donut chart (e.g. after a payment is recorded)
async function refreshPaymentChart() {
  if (!chartInstances.payment) return;
  chartInstances.payment.destroy();
  chartInstances.payment = await buildPaymentChart();
}

/* ════════════════════════════════════════════════════
   ANALYTICS PERIOD BAR — Q1–Q4 / Yearly tabs + year dropdown
   ════════════════════════════════════════════════════ */
function populateAnalyticsYearSelect() {
  const sel = document.getElementById('analytics-year-select');
  sel.innerHTML = AVAILABLE_YEARS.map(y => `<option value="${y}" ${y === analyticsYear ? 'selected' : ''}>FY ${y}–${String(y+1).slice(2)}</option>`).join('');
}
function setAnalyticsPeriod(period, btn) {
  analyticsPeriod = period;
  document.querySelectorAll('#analytics-period-tabs .ptab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  initCharts();
}
function onAnalyticsYearChange(val) {
  analyticsYear = parseInt(val);
  initCharts();
}

// Shows/hides the file-upload panel (primary data-entry path)
function toggleAnalyticsUpload() {
  document.getElementById('analytics-upload-panel').classList.toggle('hidden');
  document.getElementById('analytics-input-panel').classList.add('hidden');
}
// Switches from the upload panel to the manual-entry fallback panel
function toggleAnalyticsManual() {
  document.getElementById('analytics-upload-panel').classList.add('hidden');
  document.getElementById('analytics-input-panel').classList.remove('hidden');
  const monthsAll = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
  const note = document.getElementById('manual-entry-scope-note');
  const grid = document.getElementById('monthly-grid-container');
  const title = document.getElementById('monthly-grid-title');
  if (analyticsPeriod === 'yearly') {
    note.innerHTML = '⚠ "Yearly" is a computed total of Q1–Q4 and can\'t be edited directly. Pick a quarter tab above, then come back here to enter that quarter\'s figures.';
    grid.innerHTML = '';
    title.textContent = 'Monthly Collections (₹ Lakhs)';
    return;
  }
  note.textContent = `Enter figures for ${QUARTER_LABELS[analyticsPeriod]}, FY ${analyticsYear}–${String(analyticsYear+1).slice(2)}. Charts update automatically when you click "Apply".`;
  const qIdx = QUARTERS.indexOf(analyticsPeriod);
  const monthNames = monthsAll.slice(qIdx*3, qIdx*3+3);
  title.textContent = 'Monthly Collections (₹ Lakhs) — ' + monthNames.join(', ');
  grid.innerHTML = monthNames.map((mn,i) => `<div class="form-group"><label class="form-label">${mn}</label><input class="form-input" id="mc-${i}" type="number" step="0.1" /></div>`).join('');
  // Pre-fill inputs with the currently active quarter's data so edits are incremental
  const data = getCurrentAnalyticsData();
  const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
  setVal('ai-collected', data.collected); setVal('ai-pending', data.pending);
  setVal('ai-expenses', data.expenses);   setVal('ai-corpus', data.corpus);
  data.monthly.forEach((v,i) => setVal('mc-'+i, v));
  data.expenses_cat.forEach((v,i) => setVal('ex-'+i, v));
}
// Legacy alias kept so nothing else in the app breaks
function toggleAnalyticsInput() { toggleAnalyticsManual(); }

/* ════════════════════════════════════════════════════
   ANALYTICS — spreadsheet (.xlsx / .csv) upload & parsing
   Expected columns: Year | Quarter (Q1-Q4, blank = whole year) | Metric | Value
   Metric may be a KPI name, a month name, or an expense-category name.
   ════════════════════════════════════════════════════ */
function handleAnalyticsFile(input) {
  const file = input.files[0];
  const status = document.getElementById('analytics-upload-status');
  if (!file) return;
  status.style.color = 'var(--text2)';
  status.textContent = 'Reading ' + file.name + '…';
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'binary' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const result = applyAnalyticsRows(rows);
      status.style.color = 'var(--green)';
      status.textContent = `✓ Imported ${result.count} rows across ${result.years.size} year(s). Charts updated below.`;
      initCharts();
    } catch (err) {
      status.style.color = 'var(--red)';
      status.textContent = '⚠ Could not read that file — please check it matches the expected column format.';
    }
  };
  reader.readAsBinaryString(file);
}

// Parses uploaded rows into analyticsStore, keyed by Year → Quarter
function applyAnalyticsRows(rows) {
  const monthIndex = { Apr:0, May:1, Jun:2, Jul:3, Aug:4, Sep:5, Oct:6, Nov:7, Dec:8, Jan:9, Feb:10, Mar:11 };
  const monthToQuarter = { Apr:'Q1',May:'Q1',Jun:'Q1', Jul:'Q2',Aug:'Q2',Sep:'Q2', Oct:'Q3',Nov:'Q3',Dec:'Q3', Jan:'Q4',Feb:'Q4',Mar:'Q4' };
  const kpiKeys = { collected:'collected', pending:'pending', expenses:'expenses', corpus:'corpus' };
  const years = new Set();
  let count = 0;
  rows.forEach(row => {
    const year = parseInt(row.Year || row.year);
    let quarter = String(row.Quarter || row.quarter || '').trim().toUpperCase();
    const metric = String(row.Metric || row.metric || '').trim();
    const value  = parseFloat(row.Value || row.value);
    if (!year || !metric || isNaN(value)) return;
    if (!analyticsStore[year]) { analyticsStore[year] = {}; QUARTERS.forEach((q,qi)=>{ analyticsStore[year][q] = genQuarterData(AVAILABLE_YEARS.indexOf(year) >= 0 ? AVAILABLE_YEARS.indexOf(year) : 0, qi); }); }
    // A month name implies its quarter directly
    if (monthIndex.hasOwnProperty(metric)) {
      quarter = quarter || monthToQuarter[metric];
      const qData = analyticsStore[year][quarter];
      const posInQuarter = monthIndex[metric] % 3;
      qData.monthly[posInQuarter] = value;
      qData.collected = +qData.monthly.reduce((a,b)=>a+b,0).toFixed(1);
      years.add(year); count++;
      return;
    }
    const catIdx = EXPENSE_CATEGORY_LABELS.findIndex(c => c.toLowerCase() === metric.toLowerCase());
    if (catIdx >= 0 && quarter) {
      const qData = analyticsStore[year][quarter];
      qData.expenses_cat[catIdx] = value;
      qData.expenses = +qData.expenses_cat.reduce((a,b)=>a+b,0).toFixed(1);
      years.add(year); count++;
      return;
    }
    const kpi = kpiKeys[metric.toLowerCase()];
    if (kpi && quarter) {
      analyticsStore[year][quarter][kpi] = value;
      years.add(year); count++;
    }
  });
  if (years.size && !AVAILABLE_YEARS.includes(analyticsYear)) analyticsYear = [...years][0];
  return { count, years };
}

// Reads all manual-entry input fields, writes them into the active year/quarter, rebuilds charts
function applyAnalyticsData() {
  const n = (id, def) => parseFloat(document.getElementById(id).value) || def;
  const target = analyticsPeriod === 'yearly'
    ? null // yearly is always derived — manual edits apply to the selected quarter instead
    : analyticsStore[analyticsYear][analyticsPeriod];
  if (!target) {
    alert('Manual entry edits individual quarters. Switch to a Q1–Q4 tab to enter figures, then Yearly will total them automatically.');
    return;
  }
  target.collected    = n('ai-collected', target.collected);
  target.pending      = n('ai-pending',   target.pending);
  target.expenses     = n('ai-expenses',  target.expenses);
  target.corpus       = n('ai-corpus',    target.corpus);
  target.monthly      = Array.from({length:3}, (_,i) => n('mc-'+i, target.monthly[i]));
  target.expenses_cat = Array.from({length:7}, (_,i) => n('ex-'+i, target.expenses_cat[i]));

  initCharts();
  toggleAnalyticsManual();
  document.getElementById('analytics-input-panel').classList.add('hidden');
}


/* ════════════════════════════════════════════════════
   NOTICE BOARD — render, filter, post, edit, delete
   ════════════════════════════════════════════════════ */
const categoryColors = {
  general:     'badge-navy',
  maintenance: 'badge-amber',
  event:       'badge-green',
  emergency:   'badge-red',
  finance:     'badge-purple',
  legal:       'badge-navy',
};
const priorityColors = { normal: 'badge-navy', high: 'badge-amber', urgent: 'badge-red' };

async function loadNotices() {
  const list = document.getElementById('notices-list');
  list.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">Loading…</div>';
  try {
    notices = await apiRequest('/api/notices');
    renderNotices();
  } catch (e) {
    list.innerHTML = `<div style="text-align:center;padding:48px;color:var(--red);">Could not load notices: ${e.message}</div>`;
  }
}

function renderNotices() {
  const filtered = notices.filter(n => noticeFilter === 'all' || n.category === noticeFilter);
  const list = document.getElementById('notices-list');
  if (filtered.length === 0) {
    list.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">No notices found in this category.</div>';
    return;
  }
  list.innerHTML = filtered.map(n => {
    const dateStr = new Date(n.created_at).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
    const expiryStr = n.expires_at ? ' · Valid till ' + new Date(n.expires_at).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'}) : '';
    return `
    <div class="notice-card">
      <div class="notice-header">
        <div>
          <div class="notice-title">${n.is_pinned ? '📌 ' : ''}${n.title}</div>
          <div class="notice-meta">Posted by ${n.posted_by_name || 'Committee'} · ${dateStr}${expiryStr}${n.visible_to === 'admin_only' ? ' · 🔒 Committee/Admin Only' : ''}</div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          <span class="badge ${categoryColors[n.category] || 'badge-navy'}">${n.category.charAt(0).toUpperCase()+n.category.slice(1)}</span>
          <span class="badge ${priorityColors[n.priority] || 'badge-navy'}">${n.priority.charAt(0).toUpperCase()+n.priority.slice(1)}</span>
        </div>
      </div>
      <div class="notice-body">${n.body}</div>
      ${canWrite() ? `<div class="notice-actions">
        <button class="btn-sm btn-sm-outline" onclick="openEditNoticeModal('${n.id}')">Edit</button>
        ${canDelete() ? `<button class="btn-sm btn-sm-red" onclick="deleteNotice('${n.id}')">Delete</button>` : ''}
      </div>` : ''}
    </div>
  `; }).join('');
}

function filterNotices(cat, btn) {
  noticeFilter = cat;
  document.querySelectorAll('#dash-notices .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderNotices();
}

function openNoticeModal() {
  editingNoticeId = null;
  document.getElementById('notice-modal-title').textContent = 'Post New Notice';
  document.getElementById('notice-submit-btn').textContent = 'Post Notice';
  document.getElementById('notice-modal').classList.remove('hidden');
  document.getElementById('notice-error').classList.add('hidden');
  ['n-title','n-body','n-expiry'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('n-category').value = 'general';
  document.getElementById('n-priority').value = 'normal';
  document.getElementById('n-visible-to').value = 'all';
  document.getElementById('n-pinned').checked = false;
}

// Opens the same modal pre-filled with an existing notice's data,
// switched into "edit" mode (submits a PUT instead of a POST).
function openEditNoticeModal(id) {
  const n = notices.find(x => x.id === id);
  if (!n) return;
  editingNoticeId = id;
  document.getElementById('notice-modal-title').textContent = 'Edit Notice';
  document.getElementById('notice-submit-btn').textContent = 'Save Changes';
  document.getElementById('notice-modal').classList.remove('hidden');
  document.getElementById('notice-error').classList.add('hidden');
  document.getElementById('n-title').value = n.title;
  document.getElementById('n-body').value = n.body;
  document.getElementById('n-category').value = n.category;
  document.getElementById('n-priority').value = n.priority;
  document.getElementById('n-visible-to').value = n.visible_to;
  document.getElementById('n-pinned').checked = n.is_pinned;
  document.getElementById('n-expiry').value = n.expires_at ? n.expires_at.split('T')[0] : '';
}

function closeNoticeModal() {
  document.getElementById('notice-modal').classList.add('hidden');
  editingNoticeId = null;
}

async function postNotice() {
  const title = document.getElementById('n-title').value.trim();
  const body  = document.getElementById('n-body').value.trim();
  const err   = document.getElementById('notice-error');
  if (!title || !body) { err.textContent = 'Please fill in the Title and Body.'; err.classList.remove('hidden'); return; }
  err.classList.add('hidden');

  const btn = document.getElementById('notice-submit-btn');
  const isEditing = !!editingNoticeId;
  btn.disabled = true; btn.textContent = isEditing ? 'Saving…' : 'Posting…';

  const payload = {
    title, body,
    category: document.getElementById('n-category').value,
    priority: document.getElementById('n-priority').value,
    visibleTo: document.getElementById('n-visible-to').value,
    isPinned: document.getElementById('n-pinned').checked,
    expiresAt: document.getElementById('n-expiry').value || null,
  };

  try {
    if (isEditing) {
      await apiRequest(`/api/notices/${editingNoticeId}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      await apiRequest('/api/notices', { method: 'POST', body: JSON.stringify(payload) });
    }
    closeNoticeModal();
    await loadNotices();
    showToast(isEditing ? 'Notice updated.' : 'Notice posted.', 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not save notice.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = isEditing ? 'Save Changes' : 'Post Notice';
  }
}

async function deleteNotice(id) {
  if (!confirm('Delete this notice?')) return;
  try {
    await apiRequest(`/api/notices/${id}`, { method: 'DELETE' });
    await loadNotices();
    showToast('Notice deleted.', 'ok');
  } catch (e) {
    alert('Could not delete notice: ' + e.message);
  }
}


/* ════════════════════════════════════════════════════
   COMPLAINTS — render, filter, log, status update, feedback
   ════════════════════════════════════════════════════ */
const complaintStatusColors = {
  open: 'badge-red', acknowledged: 'badge-amber', in_progress: 'badge-amber',
  resolved: 'badge-green', closed: 'badge-navy', rejected: 'badge-navy',
};
const complaintStatusLabels = {
  open: 'Open', acknowledged: 'Acknowledged', in_progress: 'In Progress',
  resolved: 'Resolved', closed: 'Closed', rejected: 'Rejected',
};
const priorityBadge = { low: 'badge-navy', medium: 'badge-amber', high: 'badge-red', critical: 'badge-red' };

// Pulls the real, current complaint list for this logged-in user from the backend.
// Residents automatically see only their own (enforced server-side); staff see all.
async function loadComplaints() {
  const list = document.getElementById('complaints-list');
  list.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">Loading…</div>';
  try {
    complaints = await apiRequest('/api/complaints');
    renderComplaints();
  } catch (e) {
    list.innerHTML = `<div style="text-align:center;padding:48px;color:var(--red);">Could not load complaints: ${e.message}</div>`;
  }
}

function renderComplaints() {
  const filtered = complaints.filter(c => complaintFilter === 'all' || c.status === complaintFilter);
  const list = document.getElementById('complaints-list');
  if (filtered.length === 0) {
    list.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">No complaints found.</div>';
    return;
  }
  list.innerHTML = filtered.map(c => {
    const dateStr = new Date(c.created_at).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
    const canRate = isResident() && (c.status === 'resolved' || c.status === 'closed') && !c.resident_rating;
    return `
    <div class="complaint-card">
      <div class="complaint-top">
        <div>
          <div class="complaint-title">${c.title}</div>
          <div class="complaint-meta">${c.complaint_number} · ${c.member_name} · ${c.house_number} · ${dateStr}</div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          <span class="badge badge-navy">${c.category.replace('_',' ').replace(/\b\w/g, ch => ch.toUpperCase())}</span>
          <span class="badge ${priorityBadge[c.priority]}">${c.priority.charAt(0).toUpperCase()+c.priority.slice(1)}</span>
          <span class="badge ${complaintStatusColors[c.status]}">${complaintStatusLabels[c.status]}</span>
        </div>
      </div>
      <div class="complaint-body">${c.description}</div>
      ${c.resolution_notes ? `<div class="complaint-body" style="color:var(--green);font-size:13px;"><strong>Resolution:</strong> ${c.resolution_notes}</div>` : ''}
      ${c.resident_rating ? `<div class="complaint-body" style="font-size:13px;">⭐ ${c.resident_rating}/5${c.resident_feedback ? ' — ' + c.resident_feedback : ''}</div>` : ''}
      ${(isAdmin() || isCommittee()) ? `<div class="complaint-footer">
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          ${c.status === 'open' ? `<button class="btn-sm btn-sm-outline" onclick="updateComplaintStatus('${c.id}','acknowledged')">Acknowledge</button>` : ''}
          ${(c.status === 'open' || c.status === 'acknowledged') ? `<button class="btn-sm btn-sm-outline" onclick="updateComplaintStatus('${c.id}','in_progress')">Mark In Progress</button>` : ''}
          ${c.status !== 'resolved' && c.status !== 'closed' && c.status !== 'rejected' ? `<button class="btn-sm" style="background:var(--green-bg);color:var(--green);" onclick="promptResolve('${c.id}')">Mark Resolved</button>` : ''}
          ${c.status === 'resolved' ? `<button class="btn-sm btn-sm-outline" onclick="updateComplaintStatus('${c.id}','closed')">Close</button>` : ''}
          ${c.status !== 'rejected' && c.status !== 'resolved' && c.status !== 'closed' ? `<button class="btn-sm btn-sm-outline" onclick="updateComplaintStatus('${c.id}','rejected')">Reject</button>` : ''}
        </div>
      </div>` : ''}
      ${canRate ? `<div class="complaint-footer">
        <button class="btn-sm btn-sm-outline" onclick="promptRating('${c.id}')">⭐ Rate this repair</button>
      </div>` : ''}
    </div>
  `; }).join('');
}

function filterComplaints(status, btn) {
  complaintFilter = status;
  document.querySelectorAll('#dash-complaints .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderComplaints();
}

async function updateComplaintStatus(id, newStatus, extra = {}) {
  try {
    await apiRequest(`/api/complaints/${id}/status`, {
      method: 'PATCH',
      body: JSON.stringify({ newStatus, ...extra }),
    });
    await loadComplaints();
    showToast('Complaint marked ' + String(newStatus).replace(/_/g, ' ') + '.', 'ok');
  } catch (e) {
    alert('Could not update status: ' + e.message);
  }
}

function promptResolve(id) {
  const notes = prompt('How was this resolved? (shown to the resident)');
  if (notes === null) return; // cancelled
  updateComplaintStatus(id, 'resolved', { resolutionNotes: notes });
}

async function promptRating(id) {
  const rating = prompt('Rate this repair from 1 (poor) to 5 (excellent):');
  if (rating === null) return;
  const num = parseInt(rating, 10);
  if (!num || num < 1 || num > 5) { alert('Please enter a number from 1 to 5.'); return; }
  const feedback = prompt('Any additional feedback? (optional)') || '';
  try {
    await apiRequest(`/api/complaints/${id}/feedback`, {
      method: 'POST',
      body: JSON.stringify({ rating: num, feedback }),
    });
    await loadComplaints();
  } catch (e) {
    alert('Could not submit rating: ' + e.message);
  }
}

function openComplaintModal() {
  document.getElementById('complaint-modal').classList.remove('hidden');
  document.getElementById('complaint-error').classList.add('hidden');
  // Only staff need to specify whose house this complaint is for —
  // a resident's own complaint is always tied to their own account automatically.
  document.getElementById('c-house-group').style.display = (isAdmin() || isCommittee()) ? '' : 'none';
}
function closeComplaintModal() { document.getElementById('complaint-modal').classList.add('hidden'); }

async function logComplaint() {
  const house   = document.getElementById('c-house').value.trim();
  const cat     = document.getElementById('c-category').value;
  const subject = document.getElementById('c-subject').value.trim();
  const desc    = document.getElementById('c-description').value.trim();
  const err     = document.getElementById('complaint-error');
  const btn     = document.getElementById('complaint-submit-btn');

  if (!cat || !subject || !desc) { err.textContent = 'Please fill in all required fields.'; err.classList.remove('hidden'); return; }

  btn.disabled = true;
  btn.textContent = 'Submitting…';
  err.classList.add('hidden');

  try {
    await apiRequest('/api/complaints', {
      method: 'POST',
      body: JSON.stringify({
        title: subject,
        description: desc,
        category: cat,
        priority: document.getElementById('c-priority').value,
        houseNumber: house || undefined,
      }),
    });
    closeComplaintModal();
    await loadComplaints();
    ['c-house','c-subject','c-description'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('c-category').value = '';
    document.getElementById('c-priority').value = 'medium';
  } catch (e) {
    err.textContent = e.message || 'Could not submit complaint.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Submit Complaint';
  }
}


/* ════════════════════════════════════════════════════
   FACILITIES — sub-tab switching
   ════════════════════════════════════════════════════ */
function showFacilityTab(tab, btn) {
  ['overview','gym','hall'].forEach(t => {
    document.getElementById('fac-'+t).classList.toggle('hidden', t !== tab);
  });
  document.querySelectorAll('.facility-tabs .ftab').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  // Load live bookings from backend whenever the hall tab becomes visible
  if (tab === 'hall') loadHallBookings();
}

/* ════════════════════════════════════════════════════
   GYM MEMBERSHIP — plan selection and enrolment
   ════════════════════════════════════════════════════ */
function selectPlan(el, plan) {
  if (selectedPlanEl) selectedPlanEl.classList.remove('selected');
  el.classList.add('selected');
  selectedPlanEl  = el;
  selectedGymPlan = plan;
  // Sync the dropdown in the form below
  const planMap = { monthly: 'Monthly — ₹800', quarterly: 'Quarterly — ₹2,100', annual: 'Annual — ₹7,500', family: 'Family Annual — ₹12,000' };
  const sel = document.getElementById('gym-plan');
  for (let i = 0; i < sel.options.length; i++) {
    if (sel.options[i].text.startsWith(plan.charAt(0).toUpperCase()+plan.slice(1))) {
      sel.selectedIndex = i; break;
    }
  }
}

// Calculates membership valid-till date based on plan type
function calcValidTill(start, plan) {
  const d = new Date(start);
  if (plan === 'monthly')   d.setMonth(d.getMonth() + 1);
  if (plan === 'quarterly') d.setMonth(d.getMonth() + 3);
  if (plan === 'annual' || plan === 'family') d.setFullYear(d.getFullYear() + 1);
  return d.toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
}

function enrollGym() {
  const name  = document.getElementById('gym-name').value.trim();
  const house = document.getElementById('gym-house').value.trim();
  const phone = document.getElementById('gym-phone').value.trim();
  const plan  = document.getElementById('gym-plan').value;
  const start = document.getElementById('gym-start').value;
  const err   = document.getElementById('gym-error');
  const suc   = document.getElementById('gym-success');
  if (!name || !house || !phone || !plan || !start) { err.classList.remove('hidden'); suc.classList.add('hidden'); return; }
  err.classList.add('hidden');
  gymMemberships.push({ name, house, phone, plan, start: new Date(start).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'}), validTill: calcValidTill(start, plan) });
  suc.classList.remove('hidden');
  renderGymTable();
  ['gym-name','gym-house','gym-phone','gym-start'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('gym-plan').value = '';
}

function renderGymTable() {
  const planLabels = { monthly:'Monthly', quarterly:'Quarterly', annual:'Annual', family:'Family Annual' };
  const tbody = document.getElementById('gym-members-tbody');
  if (gymMemberships.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--text2);">No memberships enrolled yet.</td></tr>';
    return;
  }
  tbody.innerHTML = gymMemberships.map(m => `
    <tr>
      <td style="font-weight:500">${m.name}</td>
      <td>${m.house}</td>
      <td><span class="badge badge-navy">${planLabels[m.plan]||m.plan}</span></td>
      <td>${m.start}</td>
      <td>${m.validTill}</td>
      <td><span class="badge badge-green">Active</span></td>
    </tr>
  `).join('');
}


/* ════════════════════════════════════════════════════
   COMMUNITY HALL BOOKING — fully wired to /api/hall-bookings
   ════════════════════════════════════════════════════ */

// Submit a new booking request to the backend
async function submitHallBooking() {
  const name      = document.getElementById('hall-name').value.trim();
  const house     = document.getElementById('hall-house').value.trim();
  const phone     = document.getElementById('hall-phone').value.trim();
  const eventType = document.getElementById('hall-event-type').value;
  const date      = document.getElementById('hall-date').value;
  const guests    = document.getElementById('hall-guests').value;
  const startT    = document.getElementById('hall-start-time').value;
  const endT      = document.getElementById('hall-end-time').value;
  const notes     = document.getElementById('hall-notes').value.trim();
  const errEl     = document.getElementById('hall-error');
  const sucEl     = document.getElementById('hall-success');

  // Client-side required-field check
  if (!name || !house || !phone || !eventType || !date || !guests || !startT || !endT) {
    errEl.textContent = 'Please fill all required fields.';
    errEl.classList.remove('hidden');
    sucEl.classList.add('hidden');
    return;
  }
  errEl.classList.add('hidden');
  sucEl.classList.add('hidden');

  try {
    // Residents book for themselves (backend picks up memberId from JWT).
    // Admin/committee pass houseNumber and backend resolves the member.
    const body = {
      eventName:      eventType,   // Using event type as the event name for simplicity
      eventType,
      eventDate:      date,
      startTime:      startT,
      endTime:        endT,
      expectedGuests: guests,
      phone,
      specialRequests: notes,
    };
    // Staff booking on behalf of a resident — pass house number so backend can resolve
    if (!isResident()) body.houseNumber = house;

    await apiRequest('/api/hall-bookings', {
      method: 'POST',
      body: JSON.stringify(body),
    });

    sucEl.classList.remove('hidden');

    // Clear form
    ['hall-name','hall-house','hall-phone','hall-guests','hall-notes'].forEach(id => {
      document.getElementById(id).value = '';
    });
    document.getElementById('hall-event-type').value  = '';
    document.getElementById('hall-date').value        = '';
    document.getElementById('hall-start-time').value  = '';
    document.getElementById('hall-end-time').value    = '';

    // Reload the table so the new pending booking appears immediately
    await loadHallBookings();
  } catch (err) {
    errEl.textContent = err.message || 'Could not submit booking. Please try again.';
    errEl.classList.remove('hidden');
  }
}

// Fetch all bookings from backend and re-render the table
async function loadHallBookings() {
  // Residents cannot see the bookings table — only admin/committee can.
  if (isResident()) return;
  const tbody = document.getElementById('hall-bookings-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px;color:var(--text2);">Loading…</td></tr>';

  try {
    hallBookings = await apiRequest('/api/hall-bookings');
    renderHallTable();
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:24px;color:var(--red);">${err.message || 'Could not load bookings.'}</td></tr>`;
  }
}

// Update a booking's status (admin/committee only)
async function updateHallBookingStatus(id, newStatus) {
  let rejectionReason = null;
  if (newStatus === 'rejected') {
    rejectionReason = prompt('Enter reason for rejection (required):');
    if (!rejectionReason || !rejectionReason.trim()) return; // cancelled the prompt
  }
  if (newStatus === 'cancelled') {
    if (!confirm('Are you sure you want to cancel this booking?')) return;
  }

  try {
    const body = { newStatus };
    if (rejectionReason) body.rejectionReason = rejectionReason.trim();
    await apiRequest(`/api/hall-bookings/${id}/status`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    await loadHallBookings();
    showToast('Booking ' + newStatus + '.', 'ok');
  } catch (err) {
    alert(err.message || 'Could not update booking status.');
  }
}

// Render the bookings table from the hallBookings array
function renderHallTable() {
  const tbody = document.getElementById('hall-bookings-tbody');
  if (!tbody) return;

  if (!hallBookings || hallBookings.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px;color:var(--text2);">No bookings yet.</td></tr>';
    return;
  }

  const canManage = isAdmin() || isCommittee();

  // Show or hide the Actions column header to match whether action buttons appear
  const actionsTh = document.getElementById('hall-actions-th');
  if (actionsTh) actionsTh.style.display = canManage ? '' : 'none';

  // Status badge colour
  const statusBadge = {
    pending:   'badge-amber',
    confirmed: 'badge-green',
    rejected:  'badge-red',
    cancelled:'badge-red',
    completed: 'badge-navy',
  };
  // Status display label
  const statusLabel = {
    pending:   'Pending Approval',
    confirmed: 'Confirmed',
    rejected:  'Rejected',
    cancelled: 'Cancelled',
    completed: 'Completed',
  };

  tbody.innerHTML = hallBookings.map(b => {
    const dateStr = b.event_date
      ? new Date(b.event_date).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'})
      : '—';
    const timeStr = (b.start_time && b.end_time) ? `${b.start_time} – ${b.end_time}` : '—';
    const badge   = statusBadge[b.status] || 'badge-navy';
    const label   = statusLabel[b.status] || b.status;

    // Action buttons — only shown to admin/committee, and only for actionable statuses
    let actions = '';
    if (canManage) {
      if (b.status === 'pending') {
        actions = `
          <button class="btn-sm btn-sm-green" onclick="updateHallBookingStatus('${b.id}','confirmed')" title="Approve">✓ Approve</button>
          <button class="btn-sm btn-sm-red"   onclick="updateHallBookingStatus('${b.id}','rejected')"  title="Reject">✕ Reject</button>`;
      } else if (b.status === 'confirmed') {
        actions = `
          <button class="btn-sm btn-sm-outline" onclick="updateHallBookingStatus('${b.id}','completed')"  title="Mark completed">✓ Done</button>
          <button class="btn-sm btn-sm-red"     onclick="updateHallBookingStatus('${b.id}','cancelled')"  title="Cancel">Cancel</button>`;
      }
    }

    return `
      <tr>
        <td style="font-weight:500">${b.member_name || '—'}</td>
        <td>${b.house_number || '—'}</td>
        <td>${b.event_type || b.event_name || '—'}</td>
        <td>${dateStr}</td>
        <td style="font-size:12px">${timeStr}</td>
        <td>${b.expected_guests}</td>
        <td><span class="badge ${badge}">${label}</span></td>
        ${canManage ? `<td style="white-space:nowrap">${actions}</td>` : ''}
      </tr>`;
  }).join('');
}

/* ════════════════════════════════════════════════════════
   RESIDENT PORTAL — data stores, filters, and all functions
   The backend already scopes /api/bills, /api/complaints,
   /api/hall-bookings to the resident's own records when the
   JWT role is 'resident' — no extra filtering needed on POST.
   ════════════════════════════════════════════════════════ */

let expenseRows   = []; // cached from last loadExpenses() call — used by edit modal
let resBills      = [];
let resComplaints = [];
let resBillsFilter      = 'all';
let resComplaintsFilter = 'all';

// ── Resident Home ────────────────────────────────────────
// Loads profile, summary KPIs and a 3-bill preview all in parallel.
async function loadResidentHome() {
  if (!currentUser) return;

  // Greeting — always shown even if member record is not linked yet
  const hour = new Date().getHours();
  const greeting = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
  document.getElementById('res-home-greeting').textContent = `${greeting}, ${currentUser.name.split(' ')[0]}`;
  document.getElementById('res-home-house').textContent = `House ${currentUser.houseNumber || ''} · Krishna Grand City Society`;

  // If no member record is linked, show a helpful message but still show the page
  if (!currentUser.memberId) {
    document.getElementById('res-profile-content').innerHTML =
      '<span style="color:var(--amber);">Your account is not yet linked to a member record. Please contact the society office.</span>';
    document.getElementById('res-home-bills').innerHTML =
      '<span style="color:var(--text2);">No billing data available.</span>';
    document.getElementById('res-kpi-balance').textContent   = '—';
    document.getElementById('res-kpi-complaints').textContent = '—';
    return;
  }

  // Load all three in parallel
  const [memberData, billData, complaintData] = await Promise.allSettled([
    apiRequest(`/api/members/${currentUser.memberId}`),
    apiRequest('/api/bills'),
    apiRequest('/api/complaints'),
  ]);

  // ── Profile card ──
  const profBox = document.getElementById('res-profile-content');
  if (memberData.status === 'fulfilled') {
    const m = memberData.value;
    const vehicles = [m.vehicle_1_reg, m.vehicle_2_reg].filter(Boolean).join(', ') || '—';
    profBox.innerHTML = `
      <div class="detail-rows">
        <div class="detail-row"><span class="detail-key">House Number</span><span class="detail-val">${m.house_number}</span></div>
        <div class="detail-row"><span class="detail-key">Unit Type</span><span class="detail-val">${getTypeLabel(m.unit_type)}</span></div>
        <div class="detail-row"><span class="detail-key">Phone</span><span class="detail-val">${m.phone_primary}</span></div>
        <div class="detail-row"><span class="detail-key">Email</span><span class="detail-val">${m.email || '—'}</span></div>
        <div class="detail-row"><span class="detail-key">Vehicle(s)</span><span class="detail-val">${vehicles}</span></div>
        ${m.emergency_contact_name ? `<div class="detail-row"><span class="detail-key">Emergency Contact</span><span class="detail-val">${m.emergency_contact_name}${m.emergency_contact_phone ? ' – ' + m.emergency_contact_phone : ''}</span></div>` : ''}
      </div>`;
  } else {
    profBox.innerHTML = '<span style="color:var(--red);">Could not load profile.</span>';
  }

  // ── Bills KPI + preview ──
  if (billData.status === 'fulfilled') {
    resBills = billData.value;
    const totalDue = resBills.reduce((s,b) => s + parseFloat(b.balance_due || 0), 0);
    document.getElementById('res-kpi-balance').textContent =
      '₹' + totalDue.toLocaleString('en-IN');
    document.getElementById('res-kpi-balance').style.color =
      totalDue > 0 ? 'var(--red)' : 'var(--green)';

    const recentBills = [...resBills]
      .sort((a,b) => new Date(b.billing_month) - new Date(a.billing_month))
      .slice(0, 4);
    const homeBillsBox = document.getElementById('res-home-bills');
    if (recentBills.length === 0) {
      homeBillsBox.innerHTML = '<p>No bills generated yet.</p>';
    } else {
      homeBillsBox.innerHTML = recentBills.map(b => {
        const monthLabel = new Date(b.billing_month).toLocaleDateString('en-IN', {month:'short', year:'numeric'});
        const due = parseFloat(b.balance_due);
        return `<div class="detail-row">
          <span class="detail-key">${monthLabel}</span>
          <div style="display:flex;align-items:center;gap:10px;">
            <span class="badge ${billStatusColors[b.status]}">${b.status.charAt(0).toUpperCase()+b.status.slice(1)}</span>
            ${due > 0 ? `<span style="font-weight:500;color:var(--red);">₹${due.toLocaleString('en-IN')} due</span>` : '<span style="color:var(--green);">Paid</span>'}
            ${due > 0 ? `<button class="btn-sm" style="background:var(--gold-bg,#fdf3d8);color:var(--amber,#8a6d1a);" onclick="showUpiPaymentPage('${b.id}','${monthLabel}','${due.toLocaleString('en-IN')}')">Pay</button>` : ''}
          </div>
        </div>`;
      }).join('');
    }
  }

  // ── Complaints KPI ──
  if (complaintData.status === 'fulfilled') {
    resComplaints = complaintData.value;
    const openCount = resComplaints.filter(c => c.status === 'open' || c.status === 'acknowledged' || c.status === 'in_progress').length;
    document.getElementById('res-kpi-complaints').textContent = openCount;
  }
}

// ── Resident Bills ───────────────────────────────────────
async function loadResBills() {
  const box = document.getElementById('res-bills-list');
  box.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">Loading...</div>';
  try {
    resBills = await apiRequest('/api/bills');
    renderResBills();
  } catch (e) {
    box.innerHTML = `<div style="text-align:center;padding:48px;color:var(--red);">Could not load bills: ${e.message}</div>`;
  }
  loadResDues(); // quarterly dues panel
}

// Resident's own quarterly dues & penalties
async function loadResDues() {
  const list  = document.getElementById('res-dues-list');
  const total = document.getElementById('res-dues-total');
  if (!list) return;
  const fmt = n => '₹' + Math.round(n).toLocaleString('en-IN');
  try {
    const d = await apiRequest('/api/dues/me');
    total.textContent = fmt(d.totalOwed);
    total.style.color = d.totalOwed > 0 ? 'var(--red)' : 'var(--green)';

    // Group breakdown by financial year, newest first
    const byFy = {};
    d.breakdown.forEach(b => { (byFy[b.fyStart] = byFy[b.fyStart] || []).push(b); });
    const fys = Object.keys(byFy).map(Number).sort((a,b) => b - a);

    if (fys.length === 0) { list.innerHTML = '<span style="color:var(--green);">No dues on record.</span>'; return; }

    const qNames = {1:'Q1 Apr–Jun', 2:'Q2 Jul–Sep', 3:'Q3 Oct–Dec', 4:'Q4 Jan–Mar'};
    list.innerHTML = fys.map(fy => {
      const rows = byFy[fy].sort((a,b) => a.quarter - b.quarter).map(b => {
        let right;
        if (b.isPaid) {
          right = '<span style="color:var(--green);font-weight:600;">Paid</span>';
        } else if (!b.ended) {
          right = `<span style="color:var(--amber);">Due ${fmt(b.charge)} (current quarter)</span>`;
        } else {
          right = `<span style="color:var(--red);font-weight:600;">${fmt(b.owed)}</span>` +
                  (b.penalty ? ` <span style="font-size:11px;color:var(--red);">(incl. ₹700 penalty)</span>` : '');
        }
        return `<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);">
          <span>${qNames[b.quarter]}</span>${right}</div>`;
      }).join('');
      return `<div style="margin-bottom:14px;">
        <div style="font-weight:600;color:var(--navy);margin-bottom:4px;">FY ${fy}-${String(fy+1).slice(-2)}</div>
        ${rows}</div>`;
    }).join('');
  } catch (e) {
    list.innerHTML = `<span style="color:var(--red);">Could not load dues: ${e.message}</span>`;
    if (total) total.textContent = '—';
  }
}

function filterResBills(status, btn) {
  resBillsFilter = status;
  document.querySelectorAll('#dash-res-bills .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderResBills();
}

function renderResBills() {
  const box = document.getElementById('res-bills-list');
  const banner = document.getElementById('res-bills-banner');
  if (!box) return;

  const filtered = resBills.filter(b => resBillsFilter === 'all' || b.status === resBillsFilter);

  // Show outstanding banner
  const totalDue = resBills.reduce((s,b) => s + parseFloat(b.balance_due || 0), 0);
  if (totalDue > 0 && banner) {
    banner.style.display = '';
    document.getElementById('res-bills-total').textContent = '₹' + totalDue.toLocaleString('en-IN');
  } else if (banner) {
    banner.style.display = 'none';
  }

  if (filtered.length === 0) {
    box.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">No bills found.</div>';
    return;
  }

  box.innerHTML = filtered.map(b => {
    const monthLabel = new Date(b.billing_month).toLocaleDateString('en-IN', {month:'long', year:'numeric'});
    const due = parseFloat(b.balance_due || 0);
    const paid = parseFloat(b.amount_paid || 0);
    return `
      <div class="complaint-card" style="margin-bottom:12px;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px;">
          <div>
            <div style="font-weight:600;font-size:15px;color:var(--navy);">${monthLabel}</div>
            <div style="font-size:12px;color:var(--text2);margin-top:3px;">
              Total due: ₹${parseFloat(b.total_amount_due).toLocaleString('en-IN')}
              ${paid > 0 ? ` · Paid: ₹${paid.toLocaleString('en-IN')}` : ''}
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <span class="badge ${billStatusColors[b.status]}">${b.status.charAt(0).toUpperCase()+b.status.slice(1)}</span>
            ${due > 0 ? `<span style="font-weight:600;color:var(--red);">₹${due.toLocaleString('en-IN')} remaining</span>` : ''}
            ${due > 0 ? `<button class="btn-sm" style="background:var(--gold-bg,#fdf3d8);color:var(--amber,#8a6d1a);" onclick="showUpiPaymentPage('${b.id}','${monthLabel}','${due.toLocaleString('en-IN')}')">Pay Now</button>` : ''}
          </div>
        </div>
      </div>`;
  }).join('');
}

// ── Resident Complaints ──────────────────────────────────
async function loadResComplaints() {
  const box = document.getElementById('res-complaints-list');
  box.innerHTML = '<div style="text-align:center;padding:48px;color:var(--text2);">Loading...</div>';
  try {
    resComplaints = await apiRequest('/api/complaints');
    renderResComplaints();
  } catch (e) {
    box.innerHTML = `<div style="text-align:center;padding:48px;color:var(--red);">Could not load complaints: ${e.message}</div>`;
  }
}

function filterResComplaints(status, btn) {
  resComplaintsFilter = status;
  document.querySelectorAll('#dash-res-complaints .ftab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderResComplaints();
}

function renderResComplaints() {
  const box = document.getElementById('res-complaints-list');
  if (!box) return;

  const filtered = resComplaints.filter(c =>
    resComplaintsFilter === 'all' || c.status === resComplaintsFilter
  );

  if (filtered.length === 0) {
    box.innerHTML = `
      <div style="text-align:center;padding:48px;color:var(--text2);">
        <div style="font-size:32px;margin-bottom:12px;">✅</div>
        <div>No complaints ${resComplaintsFilter === 'all' ? 'yet' : 'with this status'}.</div>
        ${resComplaintsFilter === 'all' ? '<div style="margin-top:8px;font-size:13px;">Click "+ Raise Complaint" to report an issue.</div>' : ''}
      </div>`;
    return;
  }

  box.innerHTML = filtered.map(c => {
    const dateStr = new Date(c.created_at).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
    const canRate = (c.status === 'resolved' || c.status === 'closed') && !c.resident_rating;
    return `
      <div class="complaint-card" style="margin-bottom:12px;">
        <div class="complaint-top">
          <div>
            <div class="complaint-title">${c.title}</div>
            <div class="complaint-meta">${c.complaint_number} · ${dateStr}${c.assigned_to_name ? ' · Assigned to: ' + c.assigned_to_name : ''}</div>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;">
            <span class="badge badge-navy">${c.category.replace('_',' ').replace(/\b\w/g, ch => ch.toUpperCase())}</span>
            <span class="badge ${priorityBadge[c.priority]}">${c.priority.charAt(0).toUpperCase()+c.priority.slice(1)}</span>
            <span class="badge ${complaintStatusColors[c.status]}">${complaintStatusLabels[c.status]}</span>
          </div>
        </div>
        <div class="complaint-body">${c.description}</div>
        ${c.resolution_notes ? `<div class="complaint-body" style="color:var(--green);font-size:13px;margin-top:6px;"><strong>Resolution:</strong> ${c.resolution_notes}</div>` : ''}
        ${c.resident_rating ? `<div style="font-size:13px;margin-top:6px;">Your rating: ${'⭐'.repeat(c.resident_rating)} ${c.resident_feedback ? '— ' + c.resident_feedback : ''}</div>` : ''}
        ${canRate ? `<div class="complaint-footer"><button class="btn-sm btn-sm-outline" onclick="promptResRating('${c.id}')">⭐ Rate this repair</button></div>` : ''}
      </div>`;
  }).join('');
}

function toggleResComplaintForm() {
  const form = document.getElementById('res-complaint-form');
  form.classList.toggle('hidden');
  if (!form.classList.contains('hidden')) {
    document.getElementById('rc-subject').focus();
  }
}

async function submitResComplaint() {
  const cat  = document.getElementById('rc-category').value;
  const subj = document.getElementById('rc-subject').value.trim();
  const desc = document.getElementById('rc-description').value.trim();
  const pri  = document.getElementById('rc-priority').value;
  const err  = document.getElementById('rc-error');
  const btn  = document.getElementById('rc-submit-btn');

  if (!cat || !subj || !desc) {
    err.textContent = 'Please fill in category, subject, and details.';
    err.classList.remove('hidden');
    return;
  }
  err.classList.add('hidden');
  btn.disabled = true;
  btn.textContent = 'Submitting...';

  try {
    await apiRequest('/api/complaints', {
      method: 'POST',
      body: JSON.stringify({ title: subj, description: desc, category: cat, priority: pri }),
    });
    // Reset form and hide it
    ['rc-subject','rc-description'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('rc-category').value = '';
    document.getElementById('rc-priority').value = 'medium';
    document.getElementById('res-complaint-form').classList.add('hidden');
    // Reload list
    await loadResComplaints();
  } catch (e) {
    err.textContent = e.message || 'Could not submit complaint. Please try again.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Submit Complaint';
  }
}

async function promptResRating(id) {
  const rating = prompt('Rate this repair from 1 (poor) to 5 (excellent):');
  if (rating === null) return;
  const num = parseInt(rating, 10);
  if (!num || num < 1 || num > 5) { alert('Please enter a number from 1 to 5.'); return; }
  const feedback = prompt('Any additional feedback? (optional — press OK to skip)') || '';
  try {
    await apiRequest(`/api/complaints/${id}/feedback`, {
      method: 'POST',
      body: JSON.stringify({ rating: num, feedback }),
    });
    await loadResComplaints();
  } catch (e) {
    alert('Could not submit rating: ' + e.message);
  }
}

/* ════════════════════════════════════════════════════════
   CSV EXPORT — Expenses and Maintenance Bills
   Always fetches ALL records (no filter applied) so the
   exported file is a complete historical record.
   ════════════════════════════════════════════════════════ */

// Generic CSV download helper
function downloadCSV(filename, headers, rows) {
  const escape = val => {
    const str = (val === null || val === undefined) ? '' : String(val);
    // Wrap in quotes if contains comma, quote, or newline
    return str.includes(',') || str.includes('"') || str.includes('\n')
      ? `"${str.replace(/"/g, '""')}"` : str;
  };
  const csv = [
    headers.map(escape).join(','),
    ...rows.map(row => row.map(escape).join(',')),
  ].join('\r\n');

  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' }); // BOM for Excel
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// Export ALL expenses (ignores current month/year/category filter)
async function exportExpensesCSV() {
  try {
    const btn = event.target;
    btn.textContent = 'Exporting...';
    btn.disabled = true;

    const rows = await apiRequest('/api/expenses'); // no query params = all records

    const modeLabel = {
      cash:'Cash', cheque:'Cheque', online_upi:'UPI',
      online_neft:'NEFT', online_rtgs:'RTGS', demand_draft:'Demand Draft',
    };

    const headers = [
      'Voucher No.', 'Date', 'Category', 'Description',
      'Payee / Vendor', 'GST Number', 'Bill Number',
      'Amount (₹)', 'GST Amount (₹)', 'Payment Mode',
    ];

    const csvRows = rows.map(e => [
      e.voucher_number || '',
      e.expense_date ? new Date(e.expense_date).toLocaleDateString('en-IN') : '',
      e.category || '',
      e.description || '',
      e.vendor_name || '',
      e.vendor_gst_number || '',
      e.invoice_number || '',
      parseFloat(e.amount || 0).toFixed(2),
      parseFloat(e.gst_amount || 0).toFixed(2),
      modeLabel[e.payment_mode] || e.payment_mode || '',
    ]);

    const today    = new Date().toISOString().split('T')[0];
    downloadCSV(`KGC_Expenses_All_${today}.csv`, headers, csvRows);

    btn.textContent = '⬇ Export CSV';
    btn.disabled = false;
  } catch (err) {
    alert('Export failed: ' + err.message);
    event.target.textContent = '⬇ Export CSV';
    event.target.disabled = false;
  }
}

// Export ALL maintenance bills with payment details
async function exportBillsCSV() {
  try {
    const btn = event.target;
    btn.textContent = 'Exporting...';
    btn.disabled = true;

    const rows = await apiRequest('/api/bills'); // all bills, no filter

    const headers = [
      'House No.', 'Member Name', 'Billing Month',
      'Total Due (₹)', 'Amount Paid (₹)', 'Balance Due (₹)',
      'Status', 'Financial Year',
    ];

    const csvRows = rows.map(b => [
      b.house_number || '',
      b.member_name  || '',
      b.billing_month ? new Date(b.billing_month).toLocaleDateString('en-IN', {month:'long', year:'numeric'}) : '',
      parseFloat(b.total_amount_due || 0).toFixed(2),
      parseFloat(b.amount_paid      || 0).toFixed(2),
      parseFloat(b.balance_due      || 0).toFixed(2),
      b.status ? b.status.charAt(0).toUpperCase() + b.status.slice(1) : '',
      b.financial_year || '',
    ]);

    const today = new Date().toISOString().split('T')[0];
    downloadCSV(`KGC_MaintenanceBills_All_${today}.csv`, headers, csvRows);

    btn.textContent = '⬇ Export CSV';
    btn.disabled = false;
  } catch (err) {
    alert('Export failed: ' + err.message);
    event.target.textContent = '⬇ Export CSV';
    event.target.disabled = false;
  }
}

/* ════════════════════════════════════════════════════════
   EDIT EXPENSE MODAL
   ════════════════════════════════════════════════════════ */

let currentEditExpenseId = null;

function openEditExpenseModal(id) {
  const e = expenseRows.find(r => r.id === id);
  if (!e) { alert('Could not find expense data. Please reload the page.'); return; }

  currentEditExpenseId = id;

  // Voucher label
  document.getElementById('ee-voucher-label').textContent =
    `Voucher: ${e.voucher_number || '—'}`;

  // Populate category dropdown with all known categories
  const allCats = [...DEFAULT_CATEGORIES, ...getExtraCategories()];
  const catSel  = document.getElementById('ee-category');
  catSel.innerHTML = [...new Set(allCats)].map(c =>
    `<option value="${c}" ${c === e.category ? 'selected' : ''}>${c}</option>`
  ).join('');
  // If expense has a category not in the list, add it
  if (e.category && !allCats.includes(e.category)) {
    catSel.innerHTML += `<option value="${e.category}" selected>${e.category}</option>`;
  }

  // Pre-fill fields
  const dateVal = e.expense_date ? e.expense_date.split('T')[0] : '';
  document.getElementById('ee-date').value         = dateVal;
  document.getElementById('ee-description').value  = e.description  || '';
  document.getElementById('ee-payee').value         = e.vendor_name  || '';
  document.getElementById('ee-gst-number').value    = e.vendor_gst_number || '';
  document.getElementById('ee-amount').value        = parseFloat(e.amount || 0);
  document.getElementById('ee-gst-amount').value    = parseFloat(e.gst_amount || 0);
  document.getElementById('ee-bill-number').value   = e.invoice_number || '';
  document.getElementById('ee-payment-mode').value  = e.payment_mode || 'cash';

  document.getElementById('ee-error').classList.add('hidden');
  document.getElementById('edit-expense-modal').classList.remove('hidden');
}

function closeEditExpenseModal() {
  document.getElementById('edit-expense-modal').classList.add('hidden');
  currentEditExpenseId = null;
}

async function submitEditExpense() {
  const btn = document.getElementById('ee-submit-btn');
  const err = document.getElementById('ee-error');
  err.classList.add('hidden');

  const expenseDate   = document.getElementById('ee-date').value;
  const category      = document.getElementById('ee-category').value;
  const description   = document.getElementById('ee-description').value.trim();
  const vendorName    = document.getElementById('ee-payee').value.trim();
  const amount        = parseFloat(document.getElementById('ee-amount').value);
  const paymentMode   = document.getElementById('ee-payment-mode').value;

  if (!expenseDate || !category || !description || !vendorName || isNaN(amount) || !paymentMode) {
    err.textContent = 'Please fill in all required fields.';
    err.classList.remove('hidden');
    return;
  }

  btn.disabled = true; btn.textContent = 'Saving...';

  try {
    await apiRequest(`/api/expenses/${currentEditExpenseId}`, {
      method: 'PATCH',
      body: JSON.stringify({
        expenseDate,
        category,
        description,
        vendorName,
        vendorGstNumber:  document.getElementById('ee-gst-number').value.trim()  || null,
        amount,
        gstAmount:        parseFloat(document.getElementById('ee-gst-amount').value) || 0,
        paymentMode,
        invoiceNumber:    document.getElementById('ee-bill-number').value.trim() || null,
      }),
    });

    closeEditExpenseModal();
    await Promise.all([loadExpenses(), loadFinanceSummary()]);
    showToast('Expense updated.', 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not update expense.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Save Changes';
  }
}

/* ════════════════════════════════════════════════════════
   BILL IMPORT — historical maintenance bills from XLSX
   ════════════════════════════════════════════════════════ */

function toggleBillImportPanel() {
  document.getElementById('bill-import-panel').classList.toggle('hidden');
}

function handleBillImportFile(input) {
  if (!input.files || !input.files[0]) return;
  const statusEl = document.getElementById('bill-import-status');
  statusEl.textContent = 'Reading file...';

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const sheet    = workbook.Sheets[workbook.SheetNames[0]];
      const rows     = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      if (rows.length === 0) {
        statusEl.innerHTML = '<span style="color:var(--red);">No data rows found.</span>';
        return;
      }

      // Parse date — supports DDMMYYYY (8 digits), DD.MM.YYYY, DD/MM/YYYY, YYYY-MM-DD, JS Date
      function parseBillDate(val, asBillingMonth) {
        if (!val) return null;
        if (typeof val === 'object' && val instanceof Date) {
          const iso = val.toISOString().split('T')[0];
          return asBillingMonth ? iso.slice(0, 8) + '01' : iso;
        }
        const s = String(val).trim().replace(/\s/g, '');

        // DDMMYYYY — exactly 8 digits
        if (/^\d{8}$/.test(s)) {
          const d = s.slice(0,2), m = s.slice(2,4), y = s.slice(4,8);
          const iso = `${y}-${m}-${d}`;
          return asBillingMonth ? `${y}-${m}-01` : iso;
        }
        // DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
        const dmy = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
        if (dmy) {
          const iso = `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}`;
          return asBillingMonth ? `${dmy[3]}-${dmy[2].padStart(2,'0')}-01` : iso;
        }
        // YYYY-MM-DD already
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
          return asBillingMonth ? s.slice(0,8) + '01' : s;
        }
        // YYYYMM — 6 digits (month only, e.g. 202508)
        if (/^\d{6}$/.test(s)) {
          return `${s.slice(0,4)}-${s.slice(4,6)}-01`;
        }
        return null;
      }

      // Normalise financial year — accept "2025-2026", "2025-26", "25-26"
      function normFY(val) {
        const s = String(val || '').trim();
        // Already in short format like "2025-26"
        if (/^\d{4}-\d{2}$/.test(s)) return s;
        // Long format like "2025-2026"
        const m = s.match(/^(\d{4})[- \/](\d{4})$/);
        if (m) return `${m[1]}-${m[2].slice(-2)}`;
        return s;
      }

      function normMode(val) {
        const v = String(val || '').trim().toLowerCase().replace(/[\s_]+/g,'_');
        const map = {
          cash:'cash', cheque:'cheque', check:'cheque',
          upi:'online_upi', online_upi:'online_upi',
          neft:'online_neft', online_neft:'online_neft',
          rtgs:'online_rtgs', online_rtgs:'online_rtgs',
          dd:'demand_draft', demand_draft:'demand_draft', imps:'online_neft',
        };
        return map[v] || 'cash';
      }

      const bills = rows
        .filter(r => r.house_number)
        .map(r => ({
          houseNumber:    String(r.house_number    || '').trim(),
          billingMonth:   parseBillDate(r.billing_month, true),   // always 1st of month
          financialYear:  normFY(r.financial_year),
          amountDue:      parseFloat(r.amount_due)  || 0,          // default 0
          amountPaid:     parseFloat(r.amount_paid) || 0,
          paymentDate:    parseBillDate(r.payment_date, false),
          paymentMode:    normMode(r.payment_mode),
          billBookNumber: String(r.bill_book_number || '').trim() || null,
        }))
        .filter(r => r.billingMonth); // skip rows with unparseable billing month

      if (bills.length === 0) {
        statusEl.innerHTML = '<span style="color:var(--red);">No valid rows found. Check that billing_month column is present and formatted correctly.</span>';
        return;
      }

      statusEl.textContent = `Found ${bills.length} rows. Importing...`;

      let succeeded = 0, allFailed = [];
      const BATCH = 50;
      for (let i = 0; i < bills.length; i += BATCH) {
        const batch  = bills.slice(i, i + BATCH);
        const result = await apiRequest('/api/bills/import', {
          method: 'POST',
          body: JSON.stringify({ bills: batch }),
        });
        succeeded += result.successCount;
        allFailed  = allFailed.concat(result.failed || []);
        statusEl.textContent = `Importing... ${Math.min(i + BATCH, bills.length)} / ${bills.length}`;
      }

      statusEl.innerHTML =
        `<span style="color:var(--green);">✅ Imported ${succeeded} of ${bills.length} bills.</span>` +
        (allFailed.length
          ? `<br><span style="color:var(--red);">⚠️ ${allFailed.length} failed:</span><br>` +
            allFailed.map(f => `<span style="font-size:12px;color:var(--red);">• ${f.row}: ${f.error}</span>`).join('<br>')
          : '');

      input.value = '';
      await loadBills();
      await loadQuarterlySummary();
    } catch (err) {
      statusEl.innerHTML = `<span style="color:var(--red);">Error: ${err.message}</span>`;
    }
  };
  reader.readAsArrayBuffer(input.files[0]);
}

/* ════════════════════════════════════════════════════════
   EDIT BILL MODAL — admin/president/secretary/treasurer
   ════════════════════════════════════════════════════════ */

let currentEditBillId = null;

function openEditBillModal(id, house, name, billingMonth, amountDue, amountPaid, status) {
  currentEditBillId = id;
  const monthLabel = new Date(billingMonth).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
  document.getElementById('edit-bill-info').textContent =
    `${house} — ${name} — ${monthLabel}`;
  document.getElementById('eb-amount-due').value  = amountDue;
  document.getElementById('eb-amount-paid').value = amountPaid;
  document.getElementById('eb-status').value      = status;
  document.getElementById('eb-error').classList.add('hidden');
  document.getElementById('edit-bill-modal').classList.remove('hidden');
}

function closeEditBillModal() {
  document.getElementById('edit-bill-modal').classList.add('hidden');
  currentEditBillId = null;
}

async function submitEditBill() {
  const btn = document.getElementById('eb-submit-btn');
  const err = document.getElementById('eb-error');
  err.classList.add('hidden');

  const amountDue  = parseFloat(document.getElementById('eb-amount-due').value);
  const amountPaid = parseFloat(document.getElementById('eb-amount-paid').value) || 0;
  const status     = document.getElementById('eb-status').value;

  if (isNaN(amountDue) || amountDue < 0) {
    err.textContent = 'Please enter a valid amount due.';
    err.classList.remove('hidden');
    return;
  }

  btn.disabled = true; btn.textContent = 'Saving...';

  try {
    await apiRequest(`/api/bills/${currentEditBillId}`, {
      method: 'PATCH',
      body: JSON.stringify({ totalAmountDue: amountDue, amountPaid, status }),
    });
    closeEditBillModal();
    await loadBills();
    await loadQuarterlySummary();
    showToast('Bill updated.', 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not update bill.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Save Changes';
  }
}

/* ════════════════════════════════════════════════════════
   PAYMENT LEDGER — flat payment log (transition phase)
   ════════════════════════════════════════════════════════ */

let ledgerRows = []; // cached for edit lookup

const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const LEDGER_MODE_LABELS = {
  cash:'Cash', cheque:'Cheque', online:'Online',
  online_upi:'UPI', online_neft:'NEFT', online_rtgs:'RTGS', demand_draft:'DD',
};

/* ════════════════════════════════════════════════════════
   MAINTENANCE DUES & PENALTIES
   ₹4,000 per quarter; ₹700 penalty once an unpaid quarter ends;
   unpaid balances carry forward. Admin toggles paid/unpaid per
   quarter; totals are computed by the backend.
   ════════════════════════════════════════════════════════ */
let duesData = null; // last payload from /api/dues

async function loadDues() {
  const body = document.getElementById('dues-grid-body');
  if (!body) return;
  body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:28px;color:var(--text2);">Loading…</td></tr>';
  try {
    duesData = await apiRequest('/api/dues');
    populateDuesFyDropdown();
    renderDuesKpis();
    renderDuesGrid();
  } catch (e) {
    body.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:28px;color:var(--red);">${e.message}</td></tr>`;
  }
}

function populateDuesFyDropdown() {
  const sel = document.getElementById('dues-fy-select');
  if (!sel || !duesData) return;
  const cur = sel.value;
  let html = '';
  for (let fy = duesData.currentFyStart; fy >= duesData.firstFyStart; fy--) {
    html += `<option value="${fy}">FY ${fy}-${String(fy+1).slice(-2)}</option>`;
  }
  sel.innerHTML = html;
  if (cur) sel.value = cur; else sel.value = duesData.currentFyStart;
}

function renderDuesKpis() {
  if (!duesData) return;
  const fmt = n => '₹' + Math.round(n).toLocaleString('en-IN');
  let withDues = 0, maint = 0, pen = 0;
  duesData.members.forEach(m => {
    if (m.totalOwed > 0) withDues++;
    maint += m.totalDue;
    pen   += m.totalPenalty;
  });
  document.getElementById('dues-kpi-count').textContent   = withDues;
  document.getElementById('dues-kpi-maint').textContent   = fmt(maint);
  document.getElementById('dues-kpi-penalty').textContent = fmt(pen);
  document.getElementById('dues-kpi-total').textContent   = fmt(maint + pen);
}

function renderDuesGrid() {
  const body = document.getElementById('dues-grid-body');
  if (!body || !duesData) return;
  const fy = parseInt(document.getElementById('dues-fy-select').value) || duesData.currentFyStart;
  const q  = (document.getElementById('dues-search').value || '').toLowerCase().trim();
  const fmt = n => '₹' + Math.round(n).toLocaleString('en-IN');

  const list = duesData.members
    .filter(m => !q || m.name.toLowerCase().includes(q) || (m.houseNumber||'').toLowerCase().includes(q))
    .sort((a,b) => compareHouseNumbers(a.houseNumber, b.houseNumber));

  if (list.length === 0) {
    body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:28px;color:var(--text2);">No residents found.</td></tr>';
    return;
  }

  body.innerHTML = list.map(m => {
    // Build a quarter cell for the selected FY
    const cell = qNum => {
      const b = m.breakdown.find(x => x.fyStart === fy && x.quarter === qNum);
      if (!b) {
        // Quarter hasn't started yet this FY
        return `<td style="text-align:center"><span class="q-cell q-future">—</span></td>`;
      }
      const cls = b.isPaid ? 'q-paid' : 'q-unpaid';
      const label = b.isPaid ? 'Paid' : 'Unpaid';
      const sub = b.isPaid ? '' : `<small>${fmt(b.owed)}${b.penalty ? ' (+₹700)' : ''}</small>`;
      return `<td style="text-align:center">
        <span class="q-cell ${cls}" onclick="toggleDue('${m.memberId}',${fy},${qNum},${b.isPaid})">
          ${label}${sub}
        </span></td>`;
    };
    return `<tr>
      <td style="font-weight:500">${m.houseNumber || '—'}</td>
      <td>${m.name}</td>
      ${cell(1)}${cell(2)}${cell(3)}${cell(4)}
      <td style="text-align:right;font-weight:600;color:${m.totalOwed > 0 ? 'var(--red)' : 'var(--green)'}">
        ${m.totalOwed > 0 ? fmt(m.totalOwed) : '₹0'}
      </td>
    </tr>`;
  }).join('');
}

async function toggleDue(memberId, fyStart, quarter, currentlyPaid) {
  try {
    await apiRequest('/api/dues', {
      method: 'PATCH',
      body: JSON.stringify({ memberId, fyStart, quarter, isPaid: !currentlyPaid }),
    });
    await loadDues(); // reload so totals + penalties recompute
    showToast(!currentlyPaid ? `${quarter} marked paid.` : `${quarter} marked unpaid.`, 'ok');
  } catch (e) {
    alert('Could not update: ' + e.message);
  }
}

function initLedgerYearDropdowns() {
  const cur = new Date().getFullYear();
  // FY summary dropdown
  const fySel = document.getElementById('ledger-fy-year');
  if (fySel && fySel.options.length === 0) {
    for (let y = cur; y >= 2018; y--) {
      const label = `FY ${y}-${String(y+1).slice(-2)}`;
      const opt = document.createElement('option');
      opt.value = y; opt.textContent = label;
      // Default to current FY (if before April, FY started last year)
      const fyStart = (new Date().getMonth() >= 3) ? cur : cur - 1;
      if (y === fyStart) opt.selected = true;
      fySel.appendChild(opt);
    }
  }
  // Table filter year dropdown
  const yrSel = document.getElementById('ledger-filter-year');
  if (yrSel && yrSel.options.length <= 1) {
    yrSel.innerHTML = '<option value="">All Years</option>';
    for (let y = cur; y >= 2018; y--) {
      const opt = document.createElement('option');
      opt.value = y; opt.textContent = y;
      if (y === cur) opt.selected = true;
      yrSel.appendChild(opt);
    }
  }
}

// Load monthly collection summary cards for the selected FY
async function loadLedgerSummary() {
  const grid = document.getElementById('ledger-month-cards');
  const totalEl = document.getElementById('ledger-fy-total');
  if (!grid) return;
  const year = document.getElementById('ledger-fy-year')?.value || new Date().getFullYear();

  try {
    const data = await apiRequest(`/api/ledger/summary?year=${year}`);
    const fmt = n => 'Rs. ' + parseFloat(n).toLocaleString('en-IN', {maximumFractionDigits:0});

    // FY runs Apr(4) .. Mar(3). Build 12 month slots in that order.
    const order = [4,5,6,7,8,9,10,11,12,1,2,3];
    const byMonth = {};
    data.months.forEach(m => { byMonth[m.month] = m; });

    grid.innerHTML = order.map(mNum => {
      const m = byMonth[mNum];
      const total = m ? parseFloat(m.total) : 0;
      const count = m ? m.count : 0;
      const faded = total === 0 ? 'opacity:0.45;' : '';
      // The calendar year for this month within the FY
      const calYear = mNum >= 4 ? data.fyStart : data.fyStart + 1;
      return `
        <div class="metric-card" style="${faded}">
          <div class="metric-lbl">${MONTH_NAMES[mNum-1]} ${String(calYear).slice(-2)}</div>
          <div class="metric-val" style="color:var(--green);font-size:20px;">${fmt(total)}</div>
          <div class="metric-delta" style="color:var(--text2);">${count} payment${count === 1 ? '' : 's'}</div>
        </div>`;
    }).join('');

    totalEl.innerHTML = `FY ${data.fyLabel} total: <span style="color:var(--green);">${fmt(data.fyTotal)}</span> across ${data.fyCount} payments`;
  } catch (err) {
    grid.innerHTML = `<div style="color:var(--red);padding:16px;">${err.message}</div>`;
  }
}

// Load the payment table (filtered by month/year)
async function loadLedger() {
  const tbody = document.getElementById('ledger-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--text2);">Loading...</td></tr>';

  try {
    const month = document.getElementById('ledger-filter-month')?.value || '';
    const year  = document.getElementById('ledger-filter-year')?.value  || '';
    let qs = [];
    if (month) qs.push(`month=${month}`);
    if (year)  qs.push(`year=${year}`);
    const query = qs.length ? '?' + qs.join('&') : '';

    ledgerRows = await apiRequest('/api/ledger' + query);

    if (ledgerRows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--text2);">No payments found for this period.</td></tr>';
      document.getElementById('ledger-total-row').textContent = '';
      return;
    }

    tbody.innerHTML = ledgerRows.map(p => {
      const dateStr = new Date(p.payment_date).toLocaleDateString('en-IN', {day:'2-digit', month:'short', year:'numeric'});
      const amt = parseFloat(p.amount).toLocaleString('en-IN', {maximumFractionDigits:0});
      return `<tr class="row-click" tabindex="0" role="button" aria-label="Open payment details" onclick="openRowDetail('payment','${p.id}')" onkeydown="rowKey(event,'payment','${p.id}')">
        <td style="white-space:nowrap;">${dateStr}</td>
        <td style="font-weight:500;">${p.payer_name}</td>
        <td>${p.house_number || '-'}</td>
        <td style="color:var(--text2);font-size:13px;">${p.description || 'Maintenance'}</td>
        <td style="font-weight:600;">Rs. ${amt}</td>
        <td><span class="badge badge-navy">${LEDGER_MODE_LABELS[p.payment_mode] || p.payment_mode}</span></td>
        <td>${p.bill_book_number || '-'}</td>
        <td>${p.receipt_number || '-'}</td>
        <td style="white-space:nowrap;">${paymentCoversLabel(p)}</td>
        <td>${p.remark ? `<span class="badge badge-amber">${escapeHtmlText(p.remark)}</span>` : '<span style="color:var(--text3);">-</span>'}</td>
      </tr>`;
    }).join('');

    const total = ledgerRows.reduce((s,p) => s + parseFloat(p.amount), 0);
    document.getElementById('ledger-total-row').innerHTML =
      `Total for selected period: <span style="color:var(--green);">Rs. ${total.toLocaleString('en-IN', {maximumFractionDigits:0})}</span> (${ledgerRows.length} payments)`;
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--red);">${err.message}</td></tr>`;
  }
}

// ── Add / Edit payment modal ──
let currentEditLedgerId = null;

// Canonical key for a house number, so "Kingston - 05", "K-5", "k 05" and
// "KINGSTON5" all compare equal. Case, spaces, hyphens/dots and leading zeros
// are ignored, and block names collapse to their abbreviations
// (Kingston=K, Regal=R, Row House=RH).
function houseKey(str) {
  return String(str || '').toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .replace(/^KINGSTON/, 'K')
    .replace(/^REGAL/, 'R')
    .replace(/^ROWHOUSE/, 'RH')
    .replace(/([A-Z])0+(\d)/g, '$1$2');
}

/* ── Payment duration (which quarters a payment covers) ──
   Default = 1 quarter, matched to the financial quarter of the payment date.
   The recorder can pick 2 or 3 quarters, or the complete financial year.
   The remark is chosen / typed by the recorder (nothing is added automatically);
   the dialog only hints when the amount is above what the duration allows. */
let peStartTouched = false;
const PAY_Q_NAMES = { 1: 'Q1 (Apr–Jun)', 2: 'Q2 (Jul–Sep)', 3: 'Q3 (Oct–Dec)', 4: 'Q4 (Jan–Mar)' };

function localISODate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function payQuarterOfDate(iso) {
  const m = parseInt(String(iso).slice(5, 7), 10);
  return m >= 4 && m <= 6 ? 1 : m >= 7 && m <= 9 ? 2 : m >= 10 && m <= 12 ? 3 : 4;
}
function payFyLabel(iso) {
  const y = parseInt(String(iso).slice(0, 4), 10), m = parseInt(String(iso).slice(5, 7), 10);
  const fy = m >= 4 ? y : y - 1;
  return `${fy}-${String(fy + 1).slice(-2)}`;
}
function payQuarterFee() {
  const found = findMemberByHouse(document.getElementById('pe-house').value.trim());
  return (found.member && found.member.type === 'shop') ? EXPORT_QUARTERLY_FEE.shop : EXPORT_QUARTERLY_FEE.home;
}
// Short text such as "Q2", "Q2–Q3" or "Full year"
function payCoversText(n, s) {
  return n === 4 ? 'Full year' : (n === 1 ? `Q${s}` : `Q${s}–Q${s + n - 1}`);
}
// Table cell: which quarters a recorded payment covers
function paymentCoversLabel(p) {
  const n = p.quarters_covered, st = p.start_quarter;
  if (!n || !st) return '<span style="color:var(--text3);">-</span>';
  return `<span style="font-weight:500;">${payCoversText(n, st)}</span><div style="font-size:11px;color:var(--text2);">FY ${p.financial_year}</div>`;
}

// Escapes text typed by users before it is placed into table HTML
function escapeHtmlText(v) {
  return String(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ── Remark field: dropdown (none / Payment of previous dues / Other) + typed text ──
const REMARK_PRESETS = ['Payment of previous dues'];
function onRemarkChange() {
  const sel = document.getElementById('pe-remark'), txt = document.getElementById('pe-remark-text');
  const other = sel.value === '__other';
  txt.classList.toggle('hidden', !other);
  if (other) txt.focus();
}
function setRemarkFields(remark) {
  const sel = document.getElementById('pe-remark'), txt = document.getElementById('pe-remark-text');
  const r = (remark || '').trim();
  if (!r) { sel.value = ''; txt.value = ''; }
  else if (REMARK_PRESETS.includes(r)) { sel.value = r; txt.value = ''; }
  else { sel.value = '__other'; txt.value = r; }
  onRemarkChange();
}
function getRemarkValue() {
  const v = document.getElementById('pe-remark').value;
  if (v === '__other') return document.getElementById('pe-remark-text').value.trim() || null;
  return v || null;
}

// Keeps the Starting-quarter list, the hint and the warning in step with the form.
function updatePaymentDuration() {
  const dSel = document.getElementById('pe-duration'), sSel = document.getElementById('pe-start-q');
  const wrap = document.getElementById('pe-start-wrap'), hint = document.getElementById('pe-duration-hint');
  if (!dSel || !sSel || !hint) return;

  if (dSel.value === 'legacy') {            // older entry with no duration on record
    wrap.classList.add('hidden');
    hint.style.color = 'var(--text2)';
    hint.textContent = 'This older entry has no duration recorded. Choose a duration to record one, or leave it as it is.';
    return;
  }
  const n = parseInt(dSel.value, 10) || 1;
  const date = document.getElementById('pe-date').value;
  const dateQ = date ? payQuarterOfDate(date) : 1;
  const maxStart = 5 - n;                   // the quarters must fit inside one financial year
  const prev = parseInt(sSel.value, 10) || 0;
  sSel.innerHTML = '';
  for (let q = 1; q <= maxStart; q++) {
    const o = document.createElement('option');
    o.value = q; o.textContent = PAY_Q_NAMES[q];
    sSel.appendChild(o);
  }
  const start = (peStartTouched && prev >= 1 && prev <= maxStart) ? prev : Math.min(dateQ, maxStart);
  sSel.value = start;
  wrap.classList.toggle('hidden', n === 4);

  const fy = date ? payFyLabel(date) : '';
  const covers = n === 4 ? `Complete financial year ${fy}`
               : n === 1 ? `${PAY_Q_NAMES[start]} of FY ${fy}`
               : `Q${start}–Q${start + n - 1} of FY ${fy}`;
  const amt = parseFloat(document.getElementById('pe-amount').value) || 0;
  const limit = payQuarterFee() * n;
  const span = n === 1 ? '1 quarter' : n === 4 ? 'the complete financial year' : `${n} quarters`;
  if (amt > limit) {
    hint.style.color = 'var(--amber)';
    hint.textContent = `Covers ${covers}. Rs. ${amt.toLocaleString('en-IN')} is more than the Rs. ${limit.toLocaleString('en-IN')} due for ${span}, so you may want to choose "Payment of previous dues" in Remark. If it pays for more quarters, increase the duration.`;
  } else {
    hint.style.color = 'var(--text2)';
    hint.textContent = `Covers ${covers}.`;
  }
}
// Shows or hides the "not recorded" choice used when editing older entries.
function setLegacyDurationOption(show) {
  const dSel = document.getElementById('pe-duration');
  const existing = dSel.querySelector('option[value="legacy"]');
  if (show && !existing) {
    const o = document.createElement('option');
    o.value = 'legacy'; o.textContent = 'Not recorded (older entry)';
    dSel.insertBefore(o, dSel.firstChild);
  } else if (!show && existing) {
    existing.remove();
  }
}

// Split a house number into its block letters and unit number: "Kingston - 05" -> {block:'KINGSTON', num:5}
function splitHouse(str) {
  const t = String(str || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const m = t.match(/^([A-Z]*?)0*(\d+)$/);
  return m ? { block: m[1], num: parseInt(m[2], 10) } : null;
}

// Find the member for a typed house number. Tries the exact canonical key first,
// then falls back to "same unit number and the block names start alike"
// (K ↔ Kingston, R ↔ Regal, RH ↔ Row House ...), accepted only if exactly one member fits.
function findMemberByHouse(typed) {
  const key = houseKey(typed);
  const exact = members.find(x => houseKey(x.house) === key);
  if (exact) return { member: exact };
  const t = splitHouse(typed);
  if (!t || !t.block) return { member: null, candidates: [] };
  const aliases = { RH: 'ROWHOUSE', K: 'KINGSTON', R: 'REGAL' };
  const tb = aliases[t.block] || t.block;
  const hits = members.filter(x => {
    const h = splitHouse(x.house);
    if (!h || h.num !== t.num) return false;
    const hb = aliases[h.block] || h.block;
    return hb.startsWith(tb) || tb.startsWith(hb);
  });
  if (hits.length === 1) return { member: hits[0] };
  return { member: null, candidates: hits };
}

// Fill the house-number suggestions from the members list.
function populateHouseSuggestions() {
  const dl = document.getElementById('pe-house-list');
  if (!dl) return;
  dl.innerHTML = [...members].sort((a, b) => compareHouseNumbers(a.house, b.house))
    .map(m => `<option value="${String(m.house).replace(/"/g, '&quot;')}">${String(m.name || '').replace(/</g, '&lt;')}</option>`).join('');
}

// Map a house number to its owner from the loaded members list.
function fillOwnerFromHouse() {
  const house  = document.getElementById('pe-house').value.trim();
  const nameEl = document.getElementById('pe-name');
  const hint   = document.getElementById('pe-owner-hint');
  if (!house) { nameEl.value = ''; if (hint) hint.textContent = ''; return; }
  const found = findMemberByHouse(house);
  const m = found.member;
  if (m) {
    nameEl.value = m.name;
    nameEl.dataset.house = m.house;   // canonical house number, sent on save
    if (hint) { hint.textContent = '✓ ' + m.house + ' · ' + m.name + (m.type ? ' · ' + getTypeShortLabel(m.type) : ''); hint.style.color = 'var(--green)'; }
  } else {
    nameEl.value = '';
    delete nameEl.dataset.house;
    if (hint) {
      const c = found.candidates || [];
      hint.textContent = c.length > 1
        ? 'More than one match — pick one: ' + c.slice(0, 4).map(x => x.house).join(', ')
        : (members.length ? 'No member found for this house number. Pick from the suggestions.' : 'Members are still loading — try again in a moment.');
      hint.style.color = 'var(--red)';
    }
  }
  updatePaymentDuration();
}

async function openAddPaymentModal() {
  if (!members || members.length === 0) { try { await loadMembers(); } catch (e) {} }
  populateHouseSuggestions();
  currentEditLedgerId = null;
  document.getElementById('pe-modal-title').textContent = 'Add Payment';
  document.getElementById('pe-date').value = localISODate();
  document.getElementById('pe-amount').value = '';
  document.getElementById('pe-name').value = '';
  document.getElementById('pe-house').value = '';
  document.getElementById('pe-description').value = 'Maintenance';
  document.getElementById('pe-mode').value = 'cash';
  document.getElementById('pe-billbook').value = '';
  document.getElementById('pe-receipt').value = '';
  setRemarkFields('');
  const hint = document.getElementById('pe-owner-hint'); if (hint) hint.textContent = '';
  setLegacyDurationOption(false);
  document.getElementById('pe-duration').value = '1';     // default: 1 quarter, matching the date
  peStartTouched = false;
  updatePaymentDuration();
  document.getElementById('pe-error').classList.add('hidden');
  document.getElementById('payment-entry-modal').classList.remove('hidden');
}

async function openEditPaymentModal(id) {
  if (!members || members.length === 0) { try { await loadMembers(); } catch (e) {} }
  populateHouseSuggestions();
  const p = ledgerRows.find(r => r.id === id);
  if (!p) { alert('Could not find entry. Reload the page.'); return; }
  currentEditLedgerId = id;
  document.getElementById('pe-modal-title').textContent = 'Edit Payment';
  document.getElementById('pe-date').value = p.payment_date ? p.payment_date.split('T')[0] : '';
  document.getElementById('pe-amount').value = parseFloat(p.amount);
  document.getElementById('pe-house').value = p.house_number || '';
  document.getElementById('pe-name').value = p.payer_name || '';
  document.getElementById('pe-description').value = p.description || 'Maintenance';
  document.getElementById('pe-mode').value = p.payment_mode || 'cash';
  document.getElementById('pe-billbook').value = p.bill_book_number || '';
  document.getElementById('pe-receipt').value = p.receipt_number || '';
  setRemarkFields(p.remark || '');
  fillOwnerFromHouse();  // re-map name from current members list
  if (p.quarters_covered) {
    setLegacyDurationOption(false);
    document.getElementById('pe-duration').value = String(p.quarters_covered);
    peStartTouched = false;
    updatePaymentDuration();                               // rebuild the quarter list
    document.getElementById('pe-start-q').value = String(p.start_quarter || 1);
    peStartTouched = true;
  } else {
    setLegacyDurationOption(true);
    document.getElementById('pe-duration').value = 'legacy';
    peStartTouched = false;
  }
  updatePaymentDuration();
  document.getElementById('pe-error').classList.add('hidden');
  document.getElementById('payment-entry-modal').classList.remove('hidden');
}

function closePaymentModal() {
  document.getElementById('payment-entry-modal').classList.add('hidden');
  currentEditLedgerId = null;
}

async function submitPaymentEntry() {
  const btn = document.getElementById('pe-submit-btn');
  const err = document.getElementById('pe-error');
  err.classList.add('hidden');

  const body = {
    paymentDate:    document.getElementById('pe-date').value,
    payerName:      document.getElementById('pe-name').value.trim(),
    houseNumber:    document.getElementById('pe-name').dataset.house || document.getElementById('pe-house').value.trim(),
    description:    document.getElementById('pe-description').value.trim() || 'Maintenance',
    amount:         parseFloat(document.getElementById('pe-amount').value),
    paymentMode:    document.getElementById('pe-mode').value,
    billBookNumber: document.getElementById('pe-billbook').value.trim() || null,
    receiptNumber:  document.getElementById('pe-receipt').value.trim() || null,
    remark:         getRemarkValue(),
  };
  const durVal = document.getElementById('pe-duration').value;
  if (durVal === 'legacy') {
    body.quartersCovered = null; body.startQuarter = null;       // keep "not recorded"
  } else {
    body.quartersCovered = parseInt(durVal, 10);
    body.startQuarter = body.quartersCovered === 4 ? 1 : parseInt(document.getElementById('pe-start-q').value, 10);
  }

  if (!body.paymentDate || !body.houseNumber || isNaN(body.amount)) {
    err.textContent = 'House number, date, and amount are required.';
    err.classList.remove('hidden');
    return;
  }
  if (!body.payerName) {
    err.textContent = 'No member matches that house number — check the house number or add the member first.';
    err.classList.remove('hidden');
    return;
  }

  btn.disabled = true; btn.textContent = 'Saving...';
  try {
    if (currentEditLedgerId) {
      await apiRequest(`/api/ledger/${currentEditLedgerId}`, { method:'PATCH', body: JSON.stringify(body) });
    } else {
      await apiRequest('/api/ledger', { method:'POST', body: JSON.stringify(body) });
    }
    const wasLedgerEdit = !!currentEditLedgerId;
    closePaymentModal();
    await Promise.all([loadLedger(), loadLedgerSummary()]);
    showToast(wasLedgerEdit ? 'Payment entry updated.' : 'Payment recorded.', 'ok');
  } catch (e) {
    err.textContent = e.message || 'Could not save.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Save';
  }
}

async function deleteLedgerEntry(id) {
  if (!confirm('Delete this payment entry? This cannot be undone.')) return;
  try {
    await apiRequest(`/api/ledger/${id}`, { method:'DELETE' });
    await Promise.all([loadLedger(), loadLedgerSummary()]);
    showToast('Payment entry deleted.', 'ok');
  } catch (e) {
    alert('Could not delete: ' + e.message);
  }
}

function toggleLedgerImportPanel() {
  document.getElementById('ledger-import-panel').classList.toggle('hidden');
}

// ── Import from Excel (reads ALL sheets, skips CANCELLED) ──
function handleLedgerImportFile(input) {
  if (!input.files || !input.files[0]) return;
  const statusEl = document.getElementById('ledger-import-status');
  statusEl.textContent = 'Reading file...';

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });

      function parseDate(val) {
        if (!val) return null;
        if (typeof val === 'object' && val instanceof Date) return val.toISOString().split('T')[0];
        const s = String(val).trim().replace(/\s/g, '');
        if (/^\d{8}$/.test(s)) return `${s.slice(4,8)}-${s.slice(2,4)}-${s.slice(0,2)}`;
        const dmy = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
        if (dmy) return `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}`;
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
        const d = new Date(s); return isNaN(d) ? null : d.toISOString().split('T')[0];
      }

      function normMode(val) {
        const v = String(val || '').trim().toLowerCase().replace(/\s+/g,'_');
        const map = { cash:'cash', online:'online', cheque:'cheque', check:'cheque',
          upi:'online_upi', neft:'online_neft', rtgs:'online_rtgs', dd:'demand_draft' };
        return map[v] || 'cash';
      }

      // Read every sheet, merge rows
      let all = [];
      for (const sheetName of workbook.SheetNames) {
        if (sheetName.trim().toLowerCase() === 'shop') continue;
        const sheet = workbook.Sheets[sheetName];
        const rows  = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        rows.forEach(r => {
          const name = String(r['Name'] || '').trim();
          // Skip CANCELLED rows and rows without a name or amount
          if (!name || name.toUpperCase() === 'CANCELLED') return;
          const amount = parseFloat(r['Register Entry Amount']);
          if (!amount || isNaN(amount)) return;
          const date = parseDate(r['Date']);
          if (!date) return;
          all.push({
            paymentDate:    date,
            payerName:      name,
            houseNumber:    String(r['House No'] || '').trim() || null,
            description:    String(r['Description'] || 'Maintenance').trim() || 'Maintenance',
            amount:         amount,
            paymentMode:    normMode(r['Payment Mode']),
            billBookNumber: String(r['Bill Book1 No.'] || '').trim() || null,
            receiptNumber:  String(r['BB1 Receipt No.'] || '').trim() || null,
          });
        });
      }

      if (all.length === 0) {
        statusEl.innerHTML = '<span style="color:var(--red);">No valid payment rows found.</span>';
        return;
      }

      statusEl.textContent = `Found ${all.length} payments. Importing...`;

      let succeeded = 0, failed = [];
      const BATCH = 100;
      for (let i = 0; i < all.length; i += BATCH) {
        const batch  = all.slice(i, i + BATCH);
        const result = await apiRequest('/api/ledger/import', {
          method: 'POST', body: JSON.stringify({ payments: batch }),
        });
        succeeded += result.successCount;
        failed = failed.concat(result.failed || []);
        statusEl.textContent = `Importing... ${Math.min(i+BATCH, all.length)} / ${all.length}`;
      }

      statusEl.innerHTML =
        `<span style="color:var(--green);">Imported ${succeeded} of ${all.length} payments.</span>` +
        (failed.length ? `<br><span style="color:var(--red);">${failed.length} failed.</span>` : '');

      input.value = '';
      await Promise.all([loadLedger(), loadLedgerSummary()]);
    } catch (err) {
      statusEl.innerHTML = `<span style="color:var(--red);">Error: ${err.message}</span>`;
    }
  };
  reader.readAsArrayBuffer(input.files[0]);
}

// Export all ledger entries to CSV
/* ════════════════════════════════════════════════════════
   LEDGER CSV EXPORT — choose what, and for which period
   Period: complete financial year, Quarter 1-4 of a FY, or a custom date range.
   What:   the payment entries, or a paid / unpaid list by house.
   ════════════════════════════════════════════════════════ */
const EXPORT_QUARTERLY_FEE = { home: 4000, shop: 750 };   // Rs. per quarter
const EXPORT_PERIOD_NAMES  = { fy: 'Full year', q1: 'Q1 (Apr–Jun)', q2: 'Q2 (Jul–Sep)', q3: 'Q3 (Oct–Dec)', q4: 'Q4 (Jan–Mar)' };

function exportDateLabel(iso) {
  if (!iso) return '';
  return new Date(iso + 'T00:00:00').toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

// Opens the export dialog (the "Export CSV" button calls this).
async function exportLedgerCSV() {
  if (!members || members.length === 0) { try { await loadMembers(); } catch (e) {} }
  const fySel = document.getElementById('ex-fy');
  const now = new Date();
  const curFy = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  if (fySel.options.length === 0) {
    for (let y = curFy; y >= 2018; y--) {
      const o = document.createElement('option');
      o.value = y; o.textContent = `FY ${y}-${String(y + 1).slice(-2)}`;
      fySel.appendChild(o);
    }
  }
  fySel.value = curFy;
  document.getElementById('ex-kind').value   = 'payments';
  document.getElementById('ex-period').value = 'fy';
  document.getElementById('ex-status').value = 'all';
  const today = now.toISOString().split('T')[0];
  document.getElementById('ex-from').value = `${curFy}-04-01`;
  document.getElementById('ex-to').value   = today;
  document.getElementById('ex-error').classList.add('hidden');
  updateExportDialog();
  document.getElementById('export-modal').classList.remove('hidden');
}

function closeExportModal() {
  document.getElementById('export-modal').classList.add('hidden');
}

// Works out the date range for the current dialog choices.
function exportRange() {
  const period = document.getElementById('ex-period').value;
  if (period === 'custom') {
    const start = document.getElementById('ex-from').value;
    const end   = document.getElementById('ex-to').value;
    if (!start || !end) return { error: 'Choose both a From and a To date.' };
    if (start > end)    return { error: 'The From date must be on or before the To date.' };
    return { period, start, end, fyStart: null, n: null, label: `${exportDateLabel(start)} – ${exportDateLabel(end)}`, fileTag: `${start}_to_${end}` };
  }
  const fy = parseInt(document.getElementById('ex-fy').value);
  const spans = {
    fy: [`${fy}-04-01`,     `${fy + 1}-03-31`, 4],
    q1: [`${fy}-04-01`,     `${fy}-06-30`,     1],
    q2: [`${fy}-07-01`,     `${fy}-09-30`,     2],
    q3: [`${fy}-10-01`,     `${fy}-12-31`,     3],
    q4: [`${fy + 1}-01-01`, `${fy + 1}-03-31`, 4],
  };
  const [start, end, n] = spans[period];
  const fyLabel = `FY ${fy}-${String(fy + 1).slice(-2)}`;
  return {
    period, start, end, fyStart: fy, n,
    label: `${fyLabel} · ${EXPORT_PERIOD_NAMES[period]}`,
    fileTag: `FY${fy}-${String(fy + 1).slice(-2)}` + (period === 'fy' ? '' : '_' + period.toUpperCase()),
  };
}

// Shows/hides the right controls and explains how the list is worked out.
function updateExportDialog() {
  const kind   = document.getElementById('ex-kind').value;
  const period = document.getElementById('ex-period').value;
  const custom = period === 'custom';
  document.getElementById('ex-fy-wrap').classList.toggle('hidden', custom);
  document.getElementById('ex-custom-wrap').classList.toggle('hidden', !custom);
  document.getElementById('ex-status-wrap').classList.toggle('hidden', kind !== 'status');

  const r = exportRange();
  document.getElementById('ex-preview').textContent = r.error ? '' : `${exportDateLabel(r.start)} – ${exportDateLabel(r.end)}`;

  const hint = document.getElementById('ex-hint');
  if (kind === 'payments') {
    hint.textContent = 'One row per payment received in this period, with the owner name taken from the members list.';
  } else if (custom) {
    hint.textContent = 'A house is Paid if it has at least one payment in this date range, otherwise Unpaid.';
  } else {
    hint.textContent = 'Each house is compared with what it owes up to the end of this period (homes Rs. 4,000 and shops Rs. 750 per quarter), counting everything paid since 1 April of the financial year. Paid = fully covered, Partly paid = some paid, Unpaid = nothing paid.';
  }
}

async function runLedgerExport() {
  const err = document.getElementById('ex-error');
  const btn = document.getElementById('ex-submit-btn');
  err.classList.add('hidden');
  const r = exportRange();
  if (r.error) { err.textContent = r.error; err.classList.remove('hidden'); return; }
  const kind = document.getElementById('ex-kind').value;

  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    if (kind === 'payments') {
      const rows = await apiRequest(`/api/ledger?from=${r.start}&to=${r.end}`);
      if (rows.length === 0) { err.textContent = 'No payments were recorded in this period.'; err.classList.remove('hidden'); return; }
      const headers = ['Date','Owner Name','House No.','Description','Amount','Payment Mode','Bill Book No.','Receipt No.','Covers','Remark'];
      const csvRows = rows.map(p => [
        p.payment_date ? new Date(p.payment_date).toLocaleDateString('en-IN') : '',
        p.payer_name || '',
        p.house_number || '',
        p.description || '',
        parseFloat(p.amount || 0).toFixed(2),
        LEDGER_MODE_LABELS[p.payment_mode] || p.payment_mode || '',
        p.bill_book_number || '',
        p.receipt_number || '',
        (p.quarters_covered && p.start_quarter) ? `${payCoversText(p.quarters_covered, p.start_quarter)} (FY ${p.financial_year})` : '',
        p.remark || '',
      ]);
      downloadCSV(`KGC_Payments_${r.fileTag}.csv`, headers, csvRows);
      closeExportModal();
      showToast(`Exported ${rows.length} payment${rows.length === 1 ? '' : 's'}.`, 'ok');
      return;
    }

    // ── Paid & unpaid list, one row per active member ──
    // For a quarter/FY we need every payment from 1 April of the FY up to the period end.
    const fetchFrom = r.fyStart !== null ? `${r.fyStart}-04-01` : r.start;
    const rows = await apiRequest(`/api/ledger?from=${fetchFrom}&to=${r.end}`);
    const by = new Map();   // member id -> totals
    let unmatched = 0;
    rows.forEach(p => {
      const d = String(p.payment_date || '').slice(0, 10);
      if (p.member_id === null || p.member_id === undefined) {
        if (d >= r.start && d <= r.end) unmatched++;
        return;
      }
      const k = String(p.member_id);
      const t = by.get(k) || { toDate: 0, inPeriod: 0, count: 0, last: '' };
      const amt = parseFloat(p.amount || 0);
      t.toDate += amt;
      if (d >= r.start && d <= r.end) { t.inPeriod += amt; t.count++; }
      if (d && d > t.last) t.last = d;
      by.set(k, t);
    });

    const statusFilter = document.getElementById('ex-status').value;
    const fmt = n => n.toFixed(2);
    const list = members.filter(m => m.isActive !== false)
      .sort((a, b) => compareHouseNumbers(a.house, b.house))
      .map(m => {
        const t = by.get(String(m.id)) || { toDate: 0, inPeriod: 0, count: 0, last: '' };
        let status, cells;
        if (r.fyStart !== null) {
          const fee = m.type === 'shop' ? EXPORT_QUARTERLY_FEE.shop : EXPORT_QUARTERLY_FEE.home;
          const due = fee * r.n;
          status = t.toDate >= due ? 'Paid' : (t.toDate > 0 ? 'Partly paid' : 'Unpaid');
          cells = [m.house, m.name, getTypeShortLabel(m.type), EXPORT_PERIOD_NAMES[r.period],
                   fmt(due), fmt(t.inPeriod), fmt(t.toDate), fmt(Math.max(due - t.toDate, 0)),
                   status, t.last ? exportDateLabel(t.last) : ''];
        } else {
          status = t.inPeriod > 0 ? 'Paid' : 'Unpaid';
          cells = [m.house, m.name, getTypeShortLabel(m.type), fmt(t.inPeriod), t.count,
                   status, t.last ? exportDateLabel(t.last) : ''];
        }
        return { status, cells };
      })
      .filter(x => statusFilter === 'all' || (statusFilter === 'paid' ? x.status === 'Paid' : x.status !== 'Paid'));

    if (list.length === 0) { err.textContent = 'No houses match this selection.'; err.classList.remove('hidden'); return; }
    const headers = r.fyStart !== null
      ? ['House No.','Owner Name','Unit Type','Period','Due till period end (Rs.)','Paid in period (Rs.)','Paid since 1 April (Rs.)','Balance (Rs.)','Status','Last payment']
      : ['House No.','Owner Name','Unit Type','Paid in range (Rs.)','No. of payments','Status','Last payment'];
    const suffix = statusFilter === 'all' ? '' : (statusFilter === 'paid' ? '_PaidOnly' : '_UnpaidOnly');
    downloadCSV(`KGC_PaidUnpaid_${r.fileTag}${suffix}.csv`, headers, list.map(x => x.cells));
    closeExportModal();
    showToast(`Exported ${list.length} house${list.length === 1 ? '' : 's'}.`, 'ok');
    if (unmatched > 0) {
      setTimeout(() => showToast(`${unmatched} payment${unmatched === 1 ? '' : 's'} in this period could not be matched to a member (check the house number) and ${unmatched === 1 ? 'is' : 'are'} not counted.`, 'err'), 400);
    }
  } catch (e) {
    err.textContent = 'Export failed: ' + (e.message || 'please try again.');
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = 'Download CSV';
  }
}

/* ════════════════════════════════════════════════════════
   QUARTERLY BILLING SUMMARY
   Loads maintenance collected + expenses per quarter
   for the selected calendar year and renders the 4 cards.
   ════════════════════════════════════════════════════════ */

function initBillsYearDropdown() {
  const sel = document.getElementById('bills-fy-year');
  if (!sel || sel.options.length > 1) return;
  const cur = new Date().getFullYear();
  sel.innerHTML = '';
  for (let y = cur; y >= 2018; y--) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    if (y === cur) opt.selected = true;
    sel.appendChild(opt);
  }
}

async function loadQuarterlySummary() {
  initBillsYearDropdown();
  const year = document.getElementById('bills-fy-year')?.value || new Date().getFullYear();
  const grid = document.getElementById('quarterly-cards-grid');
  if (!grid) return;

  try {
    const data = await apiRequest(`/api/bills/quarterly-summary?year=${year}`);
    const fmt  = n => '₹' + Math.abs(parseFloat(n)).toLocaleString('en-IN', { maximumFractionDigits: 0 });
    const curQ = Math.ceil((new Date().getMonth() + 1) / 3);
    const curY = new Date().getFullYear();

    grid.innerHTML = data.quarters.map(q => {
      const surplus    = parseFloat(q.surplus);
      const isPositive = surplus >= 0;
      const isCurrent  = parseInt(year) === curY && q.quarter === curQ;
      const isFuture   = parseInt(year) > curY || (parseInt(year) === curY && q.quarter > curQ);

      return `
        <div class="metric-card" style="${isCurrent ? 'border:2px solid var(--gold,#c9a227);' : ''}${isFuture ? 'opacity:0.5;' : ''}">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;">
            <div class="metric-lbl">${q.label}</div>
            ${isCurrent ? '<span style="font-size:10px;background:var(--gold-bg,#fdf3d8);color:var(--amber,#8a6d1a);padding:2px 7px;border-radius:10px;font-weight:600;">Current</span>' : ''}
          </div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:8px;">
            <div style="display:flex;justify-content:space-between;margin-bottom:3px;">
              <span>Collected</span>
              <strong style="color:var(--green);">${fmt(q.collected)}</strong>
            </div>
            <div style="display:flex;justify-content:space-between;">
              <span>Expenses</span>
              <strong style="color:var(--red);">${fmt(q.expenses)}</strong>
            </div>
          </div>
          <div style="border-top:1px solid var(--border);padding-top:8px;display:flex;justify-content:space-between;align-items:center;">
            <span style="font-size:12px;color:var(--text2);">${isPositive ? 'Surplus' : 'Deficit'}</span>
            <span class="metric-val" style="font-size:18px;color:${isPositive ? 'var(--green)' : 'var(--red)'};">
              ${isPositive ? '+' : '−'}${fmt(surplus)}
            </span>
          </div>
        </div>`;
    }).join('');
  } catch (err) {
    document.getElementById('quarterly-cards-grid').innerHTML =
      `<div style="color:var(--red);font-size:13px;padding:16px;">Could not load quarterly summary: ${err.message}</div>`;
  }
}

/* ════════════════════════════════════════════════════════
   RESIDENT SELF-EDIT
   Residents can update their own contact details only.
   House number and unit type are displayed read-only.
   ════════════════════════════════════════════════════════ */

// Store the current resident's member data for pre-filling the edit form
let residentMemberData = null;

async function openResEditModal() {
  // Fetch fresh member data if we don't have it yet
  if (!residentMemberData && currentUser.memberId) {
    try {
      residentMemberData = await apiRequest(`/api/members/${currentUser.memberId}`);
    } catch (e) {
      alert('Could not load your profile data. Please try again.');
      return;
    }
  }

  const m = residentMemberData;
  if (!m) return;

  // Show read-only reference fields
  document.getElementById('re-house-display').textContent = m.house_number || '—';
  document.getElementById('re-type-display').textContent  = getTypeShortLabel(m.unit_type);

  // Pre-fill editable fields
  document.getElementById('re-phone-primary').value   = m.phone_primary           || '';
  document.getElementById('re-phone-secondary').value = m.phone_secondary         || '';
  document.getElementById('re-email').value           = m.email                   || '';
  document.getElementById('re-vehicle-1').value       = m.vehicle_1_reg           || '';
  document.getElementById('re-vehicle-2').value       = m.vehicle_2_reg           || '';
  document.getElementById('re-emergency-name').value  = m.emergency_contact_name  || '';
  document.getElementById('re-emergency-phone').value = m.emergency_contact_phone || '';

  document.getElementById('re-error').classList.add('hidden');
  document.getElementById('re-success').classList.add('hidden');

  document.getElementById('res-edit-modal').classList.remove('hidden');
}

function closeResEditModal() {
  document.getElementById('res-edit-modal').classList.add('hidden');
}

async function submitResEdit() {
  const btn   = document.getElementById('re-submit-btn');
  const err   = document.getElementById('re-error');
  const suc   = document.getElementById('re-success');
  err.classList.add('hidden');
  suc.classList.add('hidden');

  const phonePrimary = document.getElementById('re-phone-primary').value.trim();
  if (!phonePrimary) {
    err.textContent = 'Primary phone number is required.';
    err.classList.remove('hidden');
    return;
  }

  btn.disabled    = true;
  btn.textContent = 'Saving...';

  try {
    await apiRequest('/api/members/me', {
      method: 'PATCH',
      body: JSON.stringify({
        phonePrimary,
        phoneSecondary:        document.getElementById('re-phone-secondary').value.trim() || null,
        email:                 document.getElementById('re-email').value.trim()           || null,
        vehicle1Reg:           document.getElementById('re-vehicle-1').value.trim()       || null,
        vehicle2Reg:           document.getElementById('re-vehicle-2').value.trim()       || null,
        emergencyContactName:  document.getElementById('re-emergency-name').value.trim()  || null,
        emergencyContactPhone: document.getElementById('re-emergency-phone').value.trim() || null,
      }),
    });

    suc.textContent = 'Your details have been updated successfully.';
    suc.classList.remove('hidden');

    // Clear cached member data so the profile card refreshes on next load
    residentMemberData = null;

    // Refresh the profile card in the background
    setTimeout(() => {
      closeResEditModal();
      loadResidentHome();
    }, 1200);
  } catch (e) {
    err.textContent = e.message || 'Could not save changes. Please try again.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled    = false;
    btn.textContent = 'Save Changes';
  }
}

/* ════════════════════════════════════════════════════════
   FINANCES PAGE
   Handles expense form, XLSX import, KPI summary cards,
   and the expenses table with month/year/category filters.
   ════════════════════════════════════════════════════════ */

// Default expense categories (shown in the dropdown).
// Admin can add more via the + button — stored in localStorage
// so they persist across sessions on the same browser.
const DEFAULT_CATEGORIES = [
  'Lift Maintenance','Labour Salary','Security Salary','Housekeeping Salary',
  'Electricity','Water','Generator Maintenance','Plumbing Repairs',
  'Electrical Repairs','Pest Control','Garden / Landscaping',
  'CCTV / Security System','Painting','Stationery','Bank Charges',
  'Audit Fees','Festival / Events','Others',
];

// Load extra categories from localStorage (admin-added ones)
function getExtraCategories() {
  try { return JSON.parse(localStorage.getItem('kgc_extra_categories') || '[]'); }
  catch { return []; }
}
function saveExtraCategory(name) {
  const existing = getExtraCategories();
  if (!existing.includes(name)) {
    existing.push(name);
    localStorage.setItem('kgc_extra_categories', JSON.stringify(existing));
  }
}

// Rebuild the category <select> options with defaults + admin-added ones
function refreshCategoryDropdown(selectId) {
  const sel   = document.getElementById(selectId);
  if (!sel) return;
  const all   = [...DEFAULT_CATEGORIES, ...getExtraCategories()];
  const cur   = sel.value;
  sel.innerHTML = '<option value="">Select category</option>' +
    [...new Set(all)].map(c => `<option value="${c}"${c === cur ? ' selected' : ''}>${c}</option>`).join('');
}

// Admin clicks + to add a new category to the dropdown
function addFinanceCategory() {
  const name = prompt('Enter new category name:');
  if (!name || !name.trim()) return;
  const trimmed = name.trim();
  saveExtraCategory(trimmed);
  refreshCategoryDropdown('fin-category');
  refreshCategoryDropdown('fin-filter-category');
  document.getElementById('fin-category').value = trimmed;
}

// Populate the year filter in the expenses table header
function populateFinanceYearFilter() {
  const sel  = document.getElementById('fin-filter-year');
  if (!sel) return;
  const cur  = new Date().getFullYear();
  sel.innerHTML = '<option value="">All Years</option>' +
    [cur, cur-1, cur-2, cur-3].map(y =>
      `<option value="${y}" ${y === cur ? 'selected' : ''}>${y}</option>`
    ).join('');

  // Set default month filter to current month
  const mSel = document.getElementById('fin-filter-month');
  if (mSel && !mSel.value) {
    mSel.value = String(new Date().getMonth() + 1).padStart(2, '0');
  }
}

// Main load function — called when the Finances view is opened
async function loadFinances() {
  populateFinanceYearFilter();
  refreshCategoryDropdown('fin-category');
  refreshCategoryDropdown('fin-filter-category');

  // Set default date to today
  const dateInput = document.getElementById('fin-date');
  if (dateInput && !dateInput.value) {
    dateInput.value = new Date().toISOString().split('T')[0];
  }

  populateReceivableFilters();
  await Promise.all([loadFinanceSummary(), loadExpenses(), loadReceivables()]);
}

// Load the KPI summary cards from the backend
async function loadFinanceSummary() {
  try {
    const s = await apiRequest('/api/expenses/summary');
    const fmt = n => '₹' + parseFloat(n).toLocaleString('en-IN', {maximumFractionDigits:0});

    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun',
                        'Jul','Aug','Sep','Oct','Nov','Dec'];
    const mName = monthNames[s.month - 1];

    document.getElementById('fin-kpi-collected').textContent    = fmt(s.totalCollected);
    document.getElementById('fin-kpi-collected-sub').textContent =
      `FY ${s.fyLabel}: ${fmt(s.yearCollected)} (as in Monthly Collections)`;
    document.getElementById('fin-kpi-receivables').textContent       = fmt(s.yearReceivables);
    document.getElementById('fin-kpi-receivables-label').textContent = `FY ${s.fyLabel} · ${fmt(s.monthReceivables)} this month`;
    document.getElementById('fin-kpi-month-exp').textContent    = fmt(s.monthExpenses);
    document.getElementById('fin-kpi-month-label').textContent  = `${mName} ${s.year}`;
    document.getElementById('fin-kpi-year-exp').textContent     = fmt(s.yearExpenses);
    document.getElementById('fin-kpi-year-label').textContent   = `FY ${s.fyLabel} (Apr – Mar)`;

    const surplusEl = document.getElementById('fin-kpi-surplus');
    const surplus   = parseFloat(s.netSurplus);
    surplusEl.textContent  = fmt(Math.abs(surplus));
    surplusEl.style.color  = surplus >= 0 ? 'var(--green)' : 'var(--red)';
    // Prefix with + or -
    surplusEl.textContent  = (surplus >= 0 ? '+' : '−') + fmt(Math.abs(surplus));
  } catch (e) {
    console.error('loadFinanceSummary error:', e);
  }
}

// Load and render the expenses table with active filters
async function loadExpenses() {
  const tbody = document.getElementById('expenses-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--text2);">Loading...</td></tr>';

  try {
    const month    = document.getElementById('fin-filter-month')?.value  || '';
    const year     = document.getElementById('fin-filter-year')?.value   || '';
    const category = document.getElementById('fin-filter-category')?.value || '';

    let qs = [];
    if (month)    qs.push(`month=${month}`);
    if (year)     qs.push(`year=${year}`);
    if (category) qs.push(`category=${encodeURIComponent(category)}`);
    const query = qs.length ? '?' + qs.join('&') : '';

    const rows = await apiRequest('/api/expenses' + query);
    expenseRows = rows; // store for edit modal lookup

    if (rows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--text2);">No expenses found for the selected period.</td></tr>';
      document.getElementById('expenses-total-row').textContent = '';
      return;
    }

    const modeLabel = {
      cash:'Cash', cheque:'Cheque', online_upi:'UPI',
      online_neft:'NEFT', online_rtgs:'RTGS', demand_draft:'DD',
    };

    tbody.innerHTML = rows.map(e => {
      const dateStr = new Date(e.expense_date).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
      const amt     = parseFloat(e.amount).toLocaleString('en-IN', {maximumFractionDigits:0});
      return `<tr class="row-click" tabindex="0" role="button" aria-label="Open expense details" onclick="openRowDetail('expense','${e.id}')" onkeydown="rowKey(event,'expense','${e.id}')">
        <td style="font-size:12px;font-weight:500;color:var(--navy);">${e.voucher_number || '—'}</td>
        <td>${dateStr}</td>
        <td><span class="badge badge-navy">${e.category}</span></td>
        <td>${e.description}</td>
        <td style="color:var(--text2);font-size:13px;">${e.vendor_name}</td>
        <td style="font-weight:600;">₹${amt}</td>
        <td><span class="badge badge-navy">${modeLabel[e.payment_mode] || e.payment_mode}</span></td>
      </tr>`;
    }).join('');

    // Total row
    const total = rows.reduce((s, e) => s + parseFloat(e.amount), 0);
    document.getElementById('expenses-total-row').innerHTML =
      `Total for selected period: <strong>₹${total.toLocaleString('en-IN', {maximumFractionDigits:0})}</strong>`;
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--red);">${err.message}</td></tr>`;
  }
}

// Submit a single expense from the form
async function submitExpense() {
  const btn = document.getElementById('fin-submit-btn');
  const err = document.getElementById('fin-form-error');
  err.classList.add('hidden');

  const expenseDate  = document.getElementById('fin-date').value;
  const category     = document.getElementById('fin-category').value;
  const description  = document.getElementById('fin-description').value.trim();
  const vendorName   = document.getElementById('fin-payee').value.trim();
  const amount       = document.getElementById('fin-amount').value;
  const paymentMode  = document.getElementById('fin-payment-mode').value;

  if (!expenseDate || !category || !description || !vendorName || !amount || !paymentMode) {
    err.textContent = 'Please fill in all required fields (marked with *).';
    err.classList.remove('hidden');
    return;
  }

  btn.disabled    = true;
  btn.textContent = 'Saving...';

  try {
    const result = await apiRequest('/api/expenses', {
      method: 'POST',
      body: JSON.stringify({
        expenseDate,
        category,
        description,
        vendorName,
        vendorGstNumber:   document.getElementById('fin-gst-number').value.trim()  || null,
        amount:            parseFloat(amount),
        gstAmount:         parseFloat(document.getElementById('fin-gst-amount').value) || 0,
        amountExcludingGst: parseFloat(amount) - (parseFloat(document.getElementById('fin-gst-amount').value) || 0),
        paymentMode,
        invoiceNumber:     document.getElementById('fin-bill-number').value.trim() || null,
      }),
    });

    closeExpenseModal();
    resetFinanceForm();
    showToast(`Expense saved. Voucher No: ${result.voucherNumber}`, 'ok');

    // Refresh table and summary, and show the expenses list
    setFinanceView('expenses');
    await Promise.all([loadExpenses(), loadFinanceSummary()]);
  } catch (e) {
    err.textContent = e.message || 'Could not save expense. Please try again.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled    = false;
    btn.textContent = 'Save Expense';
  }
}

// Clear the expense form
function resetFinanceForm() {
  ['fin-description','fin-payee','fin-gst-number','fin-amount',
   'fin-gst-amount','fin-bill-number'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('fin-category').value     = '';
  document.getElementById('fin-payment-mode').value = '';
  document.getElementById('fin-date').value =
    new Date().toISOString().split('T')[0];  // reset to today
  document.getElementById('fin-form-error').classList.add('hidden');
}

// Delete an expense record (super admin only)
async function deleteExpense(id, voucherNo) {
  if (!confirm(`Delete expense ${voucherNo}? This cannot be undone.`)) return;
  try {
    await apiRequest(`/api/expenses/${id}`, { method: 'DELETE' });
    await Promise.all([loadExpenses(), loadFinanceSummary()]);
    showToast('Expense deleted.', 'ok');
  } catch (e) {
    alert('Could not delete: ' + e.message);
  }
}


/* ════════════════════════════════════════════════════════
   FINANCES — NEW RECORD MENU, EXPENSE DIALOG, RECEIVABLES
   ════════════════════════════════════════════════════════ */

let receivableRows = [];            // cached from last loadReceivables()
let currentEditReceivableId = null; // null = adding a new receivable

const RECEIVABLE_CATEGORIES = ['Hall / Club House Booking','Facility Charges','Penalty / Late Fee','Donation',
  'Bank Interest','Advertisement / Hoarding','Rent / Lease','Event / Sponsorship','Refund Received','Others'];
const PAY_MODE_LABEL = { cash:'Cash', cheque:'Cheque', online_upi:'UPI', online_neft:'NEFT', online_rtgs:'RTGS', demand_draft:'DD' };

function toggleNewRecordMenu(ev) {
  if (ev) ev.stopPropagation();
  const menu = document.getElementById('newrec-menu');
  const open = menu.classList.contains('hidden');
  menu.classList.toggle('hidden', !open);
  document.getElementById('newrec-btn').setAttribute('aria-expanded', open ? 'true' : 'false');
  if (open) menu.querySelector('button')?.focus();
}
function closeNewRecordMenu() {
  const menu = document.getElementById('newrec-menu');
  if (menu) menu.classList.add('hidden');
  const b = document.getElementById('newrec-btn');
  if (b) b.setAttribute('aria-expanded', 'false');
}
document.addEventListener('click', e => { if (!e.target.closest('#newrec-wrap')) closeNewRecordMenu(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeNewRecordMenu(); });

// Show either the Expenses or the Receivables table
function setFinanceView(view) {
  const rec = view === 'receivables';
  document.getElementById('expenses-wrap').classList.toggle('hidden', rec);
  document.getElementById('receivables-wrap').classList.toggle('hidden', !rec);
  document.getElementById('fin-tab-expenses').classList.toggle('active', !rec);
  document.getElementById('fin-tab-receivables').classList.toggle('active', rec);
}

// ── New Expense dialog (the existing form, now opened on demand) ──
function openExpenseModal() {
  closeNewRecordMenu();
  refreshCategoryDropdown('fin-category');
  const d = document.getElementById('fin-date');
  if (d && !d.value) d.value = new Date().toISOString().split('T')[0];
  document.getElementById('fin-form-error').classList.add('hidden');
  document.getElementById('expense-modal').classList.remove('hidden');
}
function closeExpenseModal() {
  document.getElementById('expense-modal').classList.add('hidden');
}

// ── Receivable dialog (add + edit) ──
async function openReceivableModal(id) {
  closeNewRecordMenu();
  if (!members || members.length === 0) { try { await loadMembers(); } catch (e) {} }
  const editing = id ? receivableRows.find(r => r.id === id) : null;
  currentEditReceivableId = editing ? editing.id : null;
  document.getElementById('rv-error').classList.add('hidden');
  document.getElementById('rv-title').textContent = editing ? 'Edit Receivable' : 'Record New Receivable';
  document.getElementById('rv-submit-btn').textContent = editing ? 'Save Changes' : 'Save Receivable';
  populateHouseSuggestions();

  const catSel = document.getElementById('rv-category');
  if (editing && editing.category && ![...catSel.options].some(o => o.value === editing.category)) {
    catSel.add(new Option(editing.category, editing.category));
  }
  const set = (k, v) => { document.getElementById(k).value = v; };
  if (editing) {
    set('rv-date', String(editing.receipt_date).split('T')[0]);
    set('rv-category', editing.category || '');
    set('rv-description', editing.description || '');
    set('rv-house', editing.house_number || '');
    set('rv-payer', editing.payer_name || '');
    set('rv-amount', parseFloat(editing.amount));
    set('rv-mode', editing.payment_mode || '');
    set('rv-reference', editing.reference_number || '');
    document.getElementById('rv-receipt-preview').textContent = editing.receipt_number;
  } else {
    set('rv-date', new Date().toISOString().split('T')[0]);
    ['rv-category','rv-description','rv-house','rv-payer','rv-amount','rv-mode','rv-reference'].forEach(k => set(k, ''));
    document.getElementById('rv-receipt-preview').textContent = 'Will be auto-generated on save (e.g. RCV-' + new Date().getFullYear() + '-0001)';
  }
  document.getElementById('rv-payer').dataset.auto = '';
  document.getElementById('receivable-modal').classList.remove('hidden');
}
function closeReceivableModal() {
  document.getElementById('receivable-modal').classList.add('hidden');
  currentEditReceivableId = null;
}

// Typing a house number fills "Received From" with the owner (still editable)
function fillReceivableFromHouse() {
  const houseEl = document.getElementById('rv-house');
  const payerEl = document.getElementById('rv-payer');
  const house = houseEl.value.trim();
  if (!house) { if (payerEl.dataset.auto === '1') { payerEl.value = ''; payerEl.dataset.auto = ''; } return; }
  const m = findMemberByHouse(house).member;
  if (m && (!payerEl.value.trim() || payerEl.dataset.auto === '1')) {
    payerEl.value = m.name;
    payerEl.dataset.auto = '1';
  }
}

async function submitReceivable() {
  const btn = document.getElementById('rv-submit-btn');
  const err = document.getElementById('rv-error');
  err.classList.add('hidden');

  const v = k => document.getElementById(k).value.trim();
  const body = {
    receiptDate: v('rv-date'), category: v('rv-category'), description: v('rv-description'),
    payerName: v('rv-payer'), houseNumber: v('rv-house') || null,
    amount: parseFloat(v('rv-amount')), paymentMode: v('rv-mode'), referenceNumber: v('rv-reference') || null,
  };
  if (!body.receiptDate || !body.category || !body.description || !body.payerName || !(body.amount > 0) || !body.paymentMode) {
    err.textContent = 'Please fill in all required fields (marked with *).';
    err.classList.remove('hidden');
    return;
  }

  const editing = currentEditReceivableId;
  btn.disabled = true; btn.textContent = 'Saving...';
  try {
    const result = await apiRequest(editing ? `/api/receivables/${editing}` : '/api/receivables', {
      method: editing ? 'PATCH' : 'POST',
      body: JSON.stringify(body),
    });
    closeReceivableModal();
    showToast(editing ? 'Receivable updated.' : `Receivable saved. Receipt No: ${result.receiptNumber}`, 'ok');
    setFinanceView('receivables');
    await Promise.all([loadReceivables(), loadFinanceSummary()]);
  } catch (e) {
    err.textContent = e.message || 'Could not save receivable. Please try again.';
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = currentEditReceivableId ? 'Save Changes' : 'Save Receivable';
  }
}

function populateReceivableFilters() {
  const ySel = document.getElementById('rv-filter-year');
  if (ySel && !ySel.options.length) {
    const cur = new Date().getFullYear();
    ySel.innerHTML = '<option value="">All Years</option>' +
      [cur, cur-1, cur-2, cur-3].map(y => `<option value="${y}" ${y === cur ? 'selected' : ''}>${y}</option>`).join('');
    document.getElementById('rv-filter-month').value = String(new Date().getMonth() + 1).padStart(2, '0');
  }
  const cSel = document.getElementById('rv-filter-category');
  if (cSel && cSel.options.length <= 1) {
    cSel.innerHTML = '<option value="">All Categories</option>' +
      RECEIVABLE_CATEGORIES.map(c => `<option value="${c}">${c}</option>`).join('');
  }
}

async function loadReceivables() {
  const tbody = document.getElementById('receivables-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--text2);">Loading...</td></tr>';
  try {
    const qs = [];
    const month = document.getElementById('rv-filter-month')?.value;
    const year  = document.getElementById('rv-filter-year')?.value;
    const cat   = document.getElementById('rv-filter-category')?.value;
    if (month) qs.push('month=' + month);
    if (year)  qs.push('year=' + year);
    if (cat)   qs.push('category=' + encodeURIComponent(cat));
    const rows = await apiRequest('/api/receivables' + (qs.length ? '?' + qs.join('&') : ''));
    receivableRows = rows;

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--text2);">No receivables found for the selected period.</td></tr>';
      document.getElementById('receivables-total-row').textContent = '';
      return;
    }
    tbody.innerHTML = rows.map(r => {
      const dateStr = new Date(r.receipt_date).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'});
      const amt = parseFloat(r.amount).toLocaleString('en-IN', {maximumFractionDigits:0});
      return `<tr class="row-click" tabindex="0" role="button" aria-label="Open receivable details" onclick="openRowDetail('receivable','${r.id}')" onkeydown="rowKey(event,'receivable','${r.id}')">
        <td style="font-size:12px;font-weight:500;color:var(--navy);">${r.receipt_number || '—'}</td>
        <td>${dateStr}</td>
        <td><span class="badge badge-navy">${r.category}</span></td>
        <td>${r.description}</td>
        <td style="color:var(--text2);font-size:13px;">${r.payer_name}${r.house_number ? ' · ' + r.house_number : ''}</td>
        <td style="font-weight:600;">₹${amt}</td>
        <td><span class="badge badge-navy">${PAY_MODE_LABEL[r.payment_mode] || r.payment_mode}</span></td>
      </tr>`;
    }).join('');
    const total = rows.reduce((t, r) => t + parseFloat(r.amount), 0);
    document.getElementById('receivables-total-row').innerHTML =
      `Total for selected period: <strong>₹${total.toLocaleString('en-IN', {maximumFractionDigits:0})}</strong>`;
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--red);">${e.message}</td></tr>`;
  }
}

async function deleteReceivable(id, receiptNo) {
  if (!confirm(`Delete receivable ${receiptNo}? This cannot be undone.`)) return;
  try {
    await apiRequest(`/api/receivables/${id}`, { method: 'DELETE' });
    await Promise.all([loadReceivables(), loadFinanceSummary()]);
    showToast('Receivable deleted.', 'ok');
  } catch (e) {
    alert('Could not delete: ' + e.message);
  }
}

async function exportReceivablesCSV(btn) {
  const label = btn.textContent;
  try {
    btn.textContent = 'Exporting...'; btn.disabled = true;
    const rows = await apiRequest('/api/receivables');
    const headers = ['Receipt No.','Date','Category','Description','Received From','House No.','Amount (₹)','Payment Mode','Reference No.'];
    const csv = rows.map(r => [
      r.receipt_number || '',
      r.receipt_date ? new Date(r.receipt_date).toLocaleDateString('en-IN') : '',
      r.category || '', r.description || '', r.payer_name || '', r.house_number || '',
      parseFloat(r.amount || 0).toFixed(2),
      PAY_MODE_LABEL[r.payment_mode] || r.payment_mode || '', r.reference_number || '',
    ]);
    downloadCSV(`KGC_Receivables_${new Date().toISOString().split('T')[0]}.csv`, headers, csv);
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    btn.textContent = label; btn.disabled = false;
  }
}


/* ════════════════════════════════════════════════════════
   ROW DETAIL POPUP — click a row in Maintenance or Finance.
   Everyone allowed to open the page sees the details;
   Edit / Delete appear only for administrator, president,
   secretary and treasurer.
   ════════════════════════════════════════════════════════ */
function rowKey(ev, kind, id) {
  if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); openRowDetail(kind, id); }
}

function openRowDetail(kind, id) {
  const money = n => 'Rs. ' + parseFloat(n || 0).toLocaleString('en-IN', {maximumFractionDigits: 2});
  const dt = v => v ? new Date(v).toLocaleDateString('en-IN', {day:'numeric', month:'short', year:'numeric'}) : '—';
  const mode = v => PAY_MODE_LABEL[v] || (v === 'online' ? 'Online' : (v || '—'));
  let title, sub, rows, edit, del;

  if (kind === 'payment') {
    const p = ledgerRows.find(r => r.id === id); if (!p) return;
    title = p.payer_name;
    sub = money(p.amount) + ' · ' + dt(p.payment_date);
    rows = [['House No.', p.house_number], ['Description', p.description], ['Mode', mode(p.payment_mode)],
            ['Covers', paymentCoversLabel(p).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().replace(/^-$/, '')],
            ['Bill book no.', p.bill_book_number], ['Receipt no.', p.receipt_number], ['Remark', p.remark]];
    edit = () => openEditPaymentModal(id);
    del  = () => deleteLedgerEntry(id);
  } else if (kind === 'expense') {
    const e = expenseRows.find(r => r.id === id); if (!e) return;
    title = e.vendor_name;
    sub = money(e.amount) + ' · ' + dt(e.expense_date);
    rows = [['Voucher no.', e.voucher_number], ['Category', e.category], ['Description', e.description],
            ['Mode', mode(e.payment_mode)], ['Bill no.', e.invoice_number],
            ['GST amount', parseFloat(e.gst_amount || 0) ? money(e.gst_amount) : ''], ['GST no.', e.vendor_gst_number]];
    edit = () => openEditExpenseModal(id);
    del  = () => deleteExpense(id, e.voucher_number);
  } else {
    const r = receivableRows.find(x => x.id === id); if (!r) return;
    title = r.payer_name;
    sub = money(r.amount) + ' · ' + dt(r.receipt_date);
    rows = [['Receipt no.', r.receipt_number], ['Category', r.category], ['Description', r.description],
            ['House / Unit', r.house_number], ['Mode', mode(r.payment_mode)], ['Reference no.', r.reference_number]];
    edit = () => openReceivableModal(id);
    del  = () => deleteReceivable(id, r.receipt_number);
  }

  document.getElementById('rd-title').textContent = title || '—';
  document.getElementById('rd-sub').textContent = sub;
  document.getElementById('rd-list').innerHTML = rows
    .filter(([, v]) => v !== undefined && v !== null && String(v).trim() !== '')
    .map(([k, v]) => `<dt>${escapeHtmlText(k)}</dt><dd>${escapeHtmlText(v)}</dd>`).join('');

  const actions = document.getElementById('rd-actions');
  actions.innerHTML = '';
  const mk = (label, cls, fn) => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = cls; b.textContent = label;
    b.onclick = fn; actions.appendChild(b); return b;
  };
  mk('Close', 'btn-cancel', closeRowDetail);
  if (canDelete()) {
    mk('Delete', 'btn-sm btn-sm-red', () => { closeRowDetail(); del(); }).style.padding = '10px 18px';
    mk('Edit', 'btn-add btn-gold', () => { closeRowDetail(); edit(); });
  }
  document.getElementById('row-detail-modal').classList.remove('hidden');
}

/* ── House ownership transfer + house maintenance history ── */
function openTransferModal() {
  const m = currentDetailMember;
  if (!m) return;
  document.getElementById('tf-sub').textContent = m.house_number + ' · currently ' + m.full_name;
  ['tf-name','tf-phone','tf-email','tf-note'].forEach(id => { document.getElementById(id).value = ''; });
  document.getElementById('tf-error').classList.add('hidden');
  document.getElementById('transfer-modal').classList.remove('hidden');
  document.getElementById('tf-name').focus();
}
function closeTransferModal() { document.getElementById('transfer-modal').classList.add('hidden'); }
async function submitTransfer() {
  const m = currentDetailMember, err = document.getElementById('tf-error'), btn = document.getElementById('tf-submit');
  const fullName = document.getElementById('tf-name').value.trim(), phone = document.getElementById('tf-phone').value.trim();
  if (!fullName || !phone) { err.textContent = 'Enter the new owner\'s name and phone.'; err.classList.remove('hidden'); return; }
  if (!confirm('Transfer ' + m.house_number + ' from ' + m.full_name + ' to ' + fullName + '?')) return;
  btn.disabled = true;
  try {
    const r = await apiRequest('/api/transfers', { method: 'POST', body: JSON.stringify({
      fromMemberId: m.id,
      newOwner: { fullName, phonePrimary: phone, email: document.getElementById('tf-email').value.trim() || null },
      note: document.getElementById('tf-note').value.trim() || null }) });
    closeTransferModal();
    await loadMembers();
    showToast('Ownership transferred.', 'ok');
    await showMemberDetail(r.newMemberId);
  } catch (e) {
    err.textContent = e.message; err.classList.remove('hidden');
  } finally { btn.disabled = false; }
}
async function loadHouseHistory(memberId) {
  const box = document.getElementById('house-history-content');
  if (!box) return;
  try {
    const h = await apiRequest('/api/transfers/member/' + memberId);
    const fmtD = d => new Date(d).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    const inr = n => '₹' + Number(n).toLocaleString('en-IN');
    const mode = v => String(v || '').replace(/_/g, ' ');
    box.innerHTML = h.owners.map((o, i) => {
      const rows = o.payments.length ? o.payments.map(p => `
        <div style="display:flex;justify-content:space-between;gap:10px;padding:5px 0;border-top:1px dashed var(--border);font-size:12px;color:var(--text2);">
          <span>${fmtD(p.payment_date)}${p.receipt_number ? ' · #' + escapeHtmlText(p.receipt_number) : ''} · ${escapeHtmlText(mode(p.payment_mode))}</span>
          <span>${inr(p.amount)}</span></div>`).join('')
        : '<div style="font-size:12px;color:var(--text2);padding:5px 0;">No payments recorded.</div>';
      const marker = i > 0 ? `<div style="margin:12px 0;padding:6px 10px;background:var(--gold-bg,#fdf3d8);border-radius:6px;font-size:12px;text-align:center;">
          ⇅ Ownership transferred${o.acquiredAt ? ' · recorded ' + fmtD(o.acquiredAt) : ''}${o.note ? ' · ' + escapeHtmlText(o.note) : ''}</div>` : '';
      return marker + `<div style="font-size:13px;">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
            <strong>${escapeHtmlText(o.name)}</strong>
            <span class="badge ${o.isActive ? 'badge-green' : 'badge-navy'}">${o.isActive ? 'Current owner' : 'Previous owner'}</span></div>
          <div style="font-size:11px;color:var(--text2);margin:2px 0 4px;">${o.payments.length} payment${o.payments.length === 1 ? '' : 's'} · ${inr(o.total)}</div>
          ${rows}</div>`;
    }).join('');
  } catch (e) {
    box.textContent = 'Could not load the house history: ' + e.message;
  }
}

function closeRowDetail() {
  document.getElementById('row-detail-modal').classList.add('hidden');
}
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !document.getElementById('row-detail-modal').classList.contains('hidden')) closeRowDetail();
  if (e.key === 'Escape' && !document.getElementById('transfer-modal').classList.contains('hidden')) closeTransferModal();
});

// Toggle the XLSX upload panel
function toggleFinanceUpload() {
  document.getElementById('finance-upload-panel').classList.toggle('hidden');
}

// Handle XLSX upload — parse and bulk-import
function handleFinanceFile(input) {
  if (!input.files || !input.files[0]) return;
  const statusEl = document.getElementById('finance-upload-status');
  statusEl.textContent = 'Reading file...';

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const sheet    = workbook.Sheets[workbook.SheetNames[0]];
      const rows     = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      if (rows.length === 0) {
        statusEl.innerHTML = '<span style="color:var(--red);">No data rows found in the file.</span>';
        return;
      }

      // ── Date parser — handles multiple formats ──────────────
      // Supports: JS Date object, YYYY-MM-DD, DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
      function parseDate(val) {
        if (!val) return null;
        if (typeof val === 'object' && val instanceof Date) {
          return val.toISOString().split('T')[0];
        }
        const s = String(val).trim();
        // DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
        const dmy = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
        if (dmy) return `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}`;
        // YYYY-MM-DD already
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
        // Fallback — let Date parse it
        const d = new Date(s);
        return isNaN(d) ? null : d.toISOString().split('T')[0];
      }

      // ── Payment mode normaliser ──────────────────────────────
      // Maps user-friendly values to the DB enum values
      function normaliseMode(val) {
        const raw = String(val || '').trim().toLowerCase().replace(/\s+/g, '_');
        const map = {
          'cash':         'cash',
          'cheque':       'cheque',
          'check':        'cheque',
          'upi':          'online_upi',
          'online_upi':   'online_upi',
          'neft':         'online_neft',
          'online_neft':  'online_neft',
          'rtgs':         'online_rtgs',
          'online_rtgs':  'online_rtgs',
          'dd':           'demand_draft',
          'demand_draft': 'demand_draft',
          'imps':         'online_neft',
        };
        return map[raw] || 'cash';
      }

      // Clean a cell value — treat 'None', 'null', '' as empty
      function clean(val) {
        if (val === null || val === undefined) return '';
        const s = String(val).trim();
        return (s === 'None' || s === 'null' || s === 'undefined') ? '' : s;
      }

      // Map rows to the format the backend expects
      const expenses = rows
        .filter(r => r.date && r.amount)
        .map(r => ({
          date:         parseDate(r.date),
          category:     clean(r.category)     || 'Others',
          description:  clean(r.description)  || '',
          payee_name:   clean(r.payee_name)   || '',
          amount:       parseFloat(r.amount)   || 0,
          gst_amount:   parseFloat(r.gst_amount || 0),
          payment_mode: normaliseMode(r.payment_mode),
          bill_number:  clean(r.bill_number)  || null,
          gst_number:   clean(r.gst_number)   || null,
        }))
        .filter(r => r.date); // skip rows where date couldn't be parsed

      statusEl.textContent = `Importing ${expenses.length} rows...`;

      // Send in batches of 50
      let succeeded = 0, failed = [];
      for (let i = 0; i < expenses.length; i += 50) {
        const batch  = expenses.slice(i, i + 50);
        const result = await apiRequest('/api/expenses/bulk', {
          method: 'POST',
          body: JSON.stringify({ expenses: batch }),
        });
        succeeded += result.successCount;
        failed     = failed.concat(result.failed || []);
        statusEl.textContent = `Importing... ${Math.min(i + 50, expenses.length)} / ${expenses.length}`;
      }

      statusEl.innerHTML =
        `<span style="color:var(--green);">✅ Imported ${succeeded} of ${expenses.length} expenses.</span>` +
        (failed.length ? `<br><span style="color:var(--red);">⚠️ ${failed.length} failed:</span><br>` +
          failed.map(f => `<span style="color:var(--red);font-size:12px;">• ${f.row}: ${f.error}</span>`).join('<br>') : '');

      input.value = '';
      await Promise.all([loadExpenses(), loadFinanceSummary()]);
    } catch (err) {
      statusEl.innerHTML = `<span style="color:var(--red);">Error: ${err.message}</span>`;
    }
  };
  reader.readAsArrayBuffer(input.files[0]);
}

/* ───── part 3 ───── */
/* ══════════════════════════════════════════════════════════════
   MOBILE NAVIGATION — burger drawer
   Builds its contents from the current context each time it opens:
   the dashboard sections the user can access (read live from the
   sidebar, so permissions are respected) plus the site links.
   ══════════════════════════════════════════════════════════════ */
function buildMobileMenu() {
  const menu = document.getElementById('mobile-menu');
  let html = '<button class="mm-close no-translate" onclick="closeMobileMenu()" aria-label="Close">&times;</button>';

  const isLoggedIn = (typeof loggedIn !== 'undefined' && loggedIn);
  const dash = document.getElementById('page-dashboard');
  const onDash = isLoggedIn && dash && !dash.classList.contains('hidden');

  // Dashboard sections — mirror the visible sidebar buttons
  if (onDash) {
    html += '<div class="mm-label">Menu</div>';
    document.querySelectorAll('.sidebar .sb-btn').forEach(btn => {
      if (btn.classList.contains('perm-hidden')) return;
      const label   = btn.textContent.trim();
      const onclick = (btn.getAttribute('onclick') || '').replace(/"/g, '&quot;');
      const active  = btn.classList.contains('active') ? ' active' : '';
      html += `<button class="mm-btn${active}" onclick="closeMobileMenu(); ${onclick}">${label}</button>`;
    });
    html += `<button class="mm-btn" onclick="closeMobileMenu(); openChangePasswordModal()">🔒 Change Password</button>`;
  }

  // Site links
  html += '<div class="mm-label">Site</div>';
  html += `<button class="mm-btn" onclick="closeMobileMenu(); showPage('home')">Home</button>`;
  html += `<button class="mm-btn" onclick="closeMobileMenu(); showPage('about')">About</button>`;
  if (isLoggedIn && !onDash && document.body.dataset.kind === 'public') {
    html += `<button class="mm-btn" onclick="closeMobileMenu(); showPage('dashboard')">Dashboard</button>`;
  }
  if (isLoggedIn) {
    html += `<button class="mm-btn gold" onclick="closeMobileMenu(); logout()">Logout</button>`;
  } else {
    html += `<button class="mm-btn gold" onclick="closeMobileMenu(); showPage('login')">Admin Login</button>`;
  }

  // Install the app (only offered when the browser allows it and it is not installed yet)
  if (typeof kgcCanInstall === 'function' && kgcCanInstall()) {
    html += '<div class="mm-label">App</div>';
    html += `<button class="mm-btn" onclick="closeMobileMenu(); kgcInstallApp()">📲 Install app</button>`;
  }

  // Language switch
  const curLang = localStorage.getItem('kgc_lang') || 'en';
  html += '<div class="mm-label">Language</div>';
  html += `<button class="mm-btn no-translate" onclick="closeMobileMenu(); toggleLanguage()">${curLang === 'hi' ? 'English' : 'हिंदी'}</button>`;

  menu.innerHTML = html;
}

function toggleMobileMenu() {
  if (document.body.classList.contains('mobile-nav-open')) {
    closeMobileMenu();
  } else {
    buildMobileMenu();
    document.body.classList.add('mobile-nav-open');
  }
}
function closeMobileMenu() {
  document.body.classList.remove('mobile-nav-open');
}
// Close the drawer automatically if the viewport grows to desktop size
window.addEventListener('resize', () => {
  if (window.innerWidth > 768) closeMobileMenu();
});

/* ───── part 4 ───── */
/* ══════════════════════════════════════════════════════════════
   BILINGUAL SUPPORT — English / Hindi (Devanagari)
   Text-node translation with reversal + MutationObserver so
   dynamically rendered content (tables, modals, badges) is
   translated automatically. Nothing else about the layout changes.
   ══════════════════════════════════════════════════════════════ */
(function () {
  // Exact English string  →  Hindi (Devanagari)
  const DICT = {
    // ── Home / about / login subtexts ──
    'A modern administrative platform for Krishna Grand City Society — transparent, efficient, and always at your service.':'कृष्णा ग्रैंड सिटी सोसायटी के लिए एक आधुनिक प्रशासनिक मंच — पारदर्शी, कुशल और सदैव आपकी सेवा में।',
    'From maintenance tracking to member records — our admin portal gives committee members full visibility and control.':'मेंटेनेंस ट्रैकिंग से लेकर सदस्य रिकॉर्ड तक — हमारा एडमिन पोर्टल समिति सदस्यों को पूर्ण दृश्यता और नियंत्रण देता है।',
    'Detailed profiles for all residents — categorized by unit type with contact details and payment history.':'सभी निवासियों की विस्तृत प्रोफ़ाइल — इकाई प्रकार के अनुसार वर्गीकृत, संपर्क विवरण और भुगतान इतिहास सहित।',
    'Monitor payment status in real-time. Identify defaulters instantly and send reminders directly from the portal.':'भुगतान स्थिति की वास्तविक समय में निगरानी करें। बकायादारों को तुरंत पहचानें और पोर्टल से सीधे अनुस्मारक भेजें।',
    'Role-based login for committee members. Your data stays protected with modern authentication.':'समिति सदस्यों के लिए भूमिका-आधारित लॉगिन। आपका डेटा आधुनिक प्रमाणीकरण से सुरक्षित रहता है।',
    'Forgot your password? Contact your society office — a committee member or admin can reset it for you.':'पासवर्ड भूल गए? अपने सोसायटी कार्यालय से संपर्क करें — कोई समिति सदस्य या एडमिन इसे आपके लिए रीसेट कर सकता है।',
    // ── Dues section subtext ──
    '₹4,000 per quarter · ₹700 penalty once a quarter ends unpaid · carries forward across years':'₹4,000 प्रति तिमाही · तिमाही अवैतनिक समाप्त होने पर ₹700 जुर्माना · वर्षों में आगे बढ़ता है',
    'Mark quarters for FY':'वित्तीय वर्ष के लिए तिमाही चिह्नित करें',
    'Will be auto-generated on save (e.g. VCH-2026-0001)':'सहेजने पर स्वतः बन जाएगा (जैसे VCH-2026-0001)',
    // ── Facilities subtexts ──
    'Light refreshments and beverages available. A comfortable space to socialise with neighbours.':'हल्के जलपान और पेय उपलब्ध। पड़ोसियों के साथ मेलजोल के लिए एक आरामदायक स्थान।',
    'Fully air-conditioned':'पूर्णतः वातानुकूलित', 'Up to 250 guests':'250 अतिथियों तक',
    '48 hours minimum':'न्यूनतम 48 घंटे', '30 members':'30 सदस्य', '45+ machines':'45+ मशीनें',
    'First floor of the Club House · Capacity: 250 persons':'क्लब हाउस की पहली मंज़िल · क्षमता: 250 व्यक्ति',
    // ── Resident bills subtext ──
    'To pay, visit the society office or use the UPI option on each unpaid bill. Contact the treasurer for queries.':'भुगतान के लिए सोसायटी कार्यालय जाएँ या प्रत्येक अवैतनिक बिल पर यूपीआई विकल्प का उपयोग करें। प्रश्नों के लिए कोषाध्यक्ष से संपर्क करें।',
    // ── Member/photo + import subtexts ──
    'Photo preview only for now — saving photos requires file storage setup, coming in a future step.':'अभी केवल फ़ोटो पूर्वावलोकन — फ़ोटो सहेजने के लिए फ़ाइल स्टोरेज सेटअप आवश्यक है, जो आगे के चरण में आएगा।',
    'Upload an .xlsx file with the columns below.':'नीचे दिए गए कॉलम के साथ एक .xlsx फ़ाइल अपलोड करें।',
    'unit_type is auto-detected from the house number':'इकाई प्रकार मकान नंबर से स्वतः पहचाना जाता है',
    '— you do not need a unit_type column.':'— आपको unit_type कॉलम की आवश्यकता नहीं है।',
    'Upload an .xlsx file with columns:':'इन कॉलम के साथ एक .xlsx फ़ाइल अपलोड करें:',
    // ── Dues grid quarter headers (split nodes) ──
    'Apr–Jun':'अप्रैल–जून', 'Jul–Sep':'जुलाई–सित', 'Oct–Dec':'अक्टू–दिस', 'Jan–Mar':'जन–मार्च',

    // ── Nav / shell ──
    'Home':'होम', 'About':'हमारे बारे में', 'Admin Login':'एडमिन लॉगिन',
    'Logout':'लॉगआउट', 'Admin Portal':'एडमिन पोर्टल', 'Admin Portal →':'एडमिन पोर्टल →',
    'Learn More':'और जानें',
    'Est. 2018 · Sheetala Colony, Kohka, Bhilai':'स्थापित 2018 · शीतला कॉलोनी, कोहका, भिलाई',
    'Your Community,':'आपका समुदाय,', 'Managed with Care':'देखभाल के साथ प्रबंधित',
    'Why KGC-Admin':'केजीसी-एडमिन क्यों',
    'Everything your society needs,':'आपकी सोसायटी की हर ज़रूरत,', 'in one place.':'एक ही जगह।',
    'Member Directory':'सदस्य निर्देशिका', 'Maintenance Tracking':'मेंटेनेंस ट्रैकिंग',
    'Analytics Dashboard':'एनालिटिक्स डैशबोर्ड',
    'Visual reports on collections, occupancy, expense categories and year-over-year trends.':'संग्रह, अधिभोग, व्यय श्रेणियों और वर्ष-दर-वर्ष रुझानों पर दृश्य रिपोर्ट।',
    'Secure Access':'सुरक्षित पहुँच', 'Notice Board':'सूचना पट्ट',
    'Publish announcements and circulars to residents from a single admin interface.':'एक ही एडमिन इंटरफ़ेस से निवासियों के लिए घोषणाएँ और परिपत्र प्रकाशित करें।',
    'Facility Management':'सुविधा प्रबंधन',
    'Track maintenance requests, vendor contracts and common area bookings with ease.':'मेंटेनेंस अनुरोध, विक्रेता अनुबंध और सामान्य क्षेत्र बुकिंग आसानी से ट्रैक करें।',
    'Society · All rights reserved':'सोसायटी · सर्वाधिकार सुरक्षित', '· All rights reserved':'· सर्वाधिकार सुरक्षित',

    // ── About ──
    'Our Society':'हमारी सोसायटी', 'A Thriving Community Since 2018':'2018 से एक समृद्ध समुदाय',
    'Committee Members 2025–26':'समिति सदस्य 2025–26',
    'Our Facilities':'हमारी सुविधाएँ', 'Life at Krishna Grand City':'कृष्णा ग्रैंड सिटी में जीवन',
    'We take pride in maintaining world-class facilities for our residents and their families.':'हमें अपने निवासियों और उनके परिवारों के लिए विश्व-स्तरीय सुविधाएँ बनाए रखने पर गर्व है।',
    'Gymnasium':'जिम',
    'Fully equipped modern gym with certified trainers available Monday to Saturday.':'सोमवार से शनिवार तक प्रमाणित प्रशिक्षकों के साथ पूर्ण सुसज्जित आधुनिक जिम।',
    'Community Hall':'सामुदायिक हॉल',
    'Spacious hall for events, meetings, and celebrations — bookable by residents.':'आयोजनों, बैठकों और समारोहों के लिए विशाल हॉल — निवासियों द्वारा बुक किया जा सकता है।',
    'Landscaped Gardens':'भूदृश्य उद्यान',
    "3 acres of manicured green spaces, jogging tracks and children's play areas.":'3 एकड़ सुव्यवस्थित हरित स्थान, जॉगिंग ट्रैक और बच्चों के खेल क्षेत्र।',
    '24/7 Security':'24/7 सुरक्षा',
    'CCTV coverage, biometric entry and trained security personnel round-the-clock.':'चौबीसों घंटे सीसीटीवी कवरेज, बायोमेट्रिक प्रवेश और प्रशिक्षित सुरक्षाकर्मी।',
    'Club House':'क्लब हाउस',
    'Exclusive clubhouse with indoor games, library, and social spaces for residents.':'निवासियों के लिए इनडोर खेल, पुस्तकालय और सामाजिक स्थानों वाला विशेष क्लब हाउस।',

    // ── Login ──
    'Sign in with your email, phone number, or house number':'अपने ईमेल, फ़ोन नंबर या मकान नंबर से साइन इन करें',
    '(e.g.':'(जैसे', ') and your password.':') और अपना पासवर्ड।',
    'Incorrect login ID or password. Please try again.':'गलत लॉगिन आईडी या पासवर्ड। कृपया पुनः प्रयास करें।',
    'Email, phone, or house number':'ईमेल, फ़ोन या मकान नंबर', 'Password':'पासवर्ड',
    'Sign In →':'साइन इन →', 'Signing in…':'साइन इन हो रहा है…',
    'Please enter both a login ID and password.':'कृपया लॉगिन आईडी और पासवर्ड दोनों दर्ज करें।',

    // ── Sidebar ──
    'Main':'मुख्य', 'Members':'सदस्य', 'My Profile':'मेरी प्रोफ़ाइल', 'My Bills':'मेरे बिल',
    'My Complaints':'मेरी शिकायतें', 'Finances':'वित्त', 'Management':'प्रबंधन',
    'Complaints':'शिकायतें', 'Maintenance Bills':'मेंटेनेंस बिल', 'Facilities':'सुविधाएँ',
    'Change Password':'पासवर्ड बदलें', '🔒 Change Password':'🔒 पासवर्ड बदलें',
    'Current password':'वर्तमान पासवर्ड', 'New password (min. 8 characters)':'नया पासवर्ड (कम से कम 8 अक्षर)',
    'Update Password':'पासवर्ड अपडेट करें', 'Updating…':'अपडेट हो रहा है…',
    'Password updated. Use it next time you log in.':'पासवर्ड अपडेट हो गया। अगली बार लॉगिन करते समय इसका उपयोग करें।',
    'Please fill in both fields.':'कृपया दोनों फ़ील्ड भरें।',

    // ── Common buttons / words ──
    'Cancel':'रद्द करें', 'Save':'सहेजें', 'Save Changes':'बदलाव सहेजें', 'Clear':'साफ़ करें',
    'Close':'बंद करें', 'Apply':'लागू करें', 'Done':'हो गया', 'Actions':'कार्रवाई',
    'Status':'स्थिति', 'Loading...':'लोड हो रहा है...', 'Loading…':'लोड हो रहा है…',
    'Saving...':'सहेजा जा रहा है...', 'Saving…':'सहेजा जा रहा है…',
    'Submitting...':'सबमिट हो रहा है...', 'Submitting…':'सबमिट हो रहा है…',
    'Exporting...':'एक्सपोर्ट हो रहा है...', 'Adding…':'जोड़ा जा रहा है…',
    '⬇ Export CSV':'⬇ CSV एक्सपोर्ट करें', 'Export failed: ':'एक्सपोर्ट विफल: ',
    'Please fill in all required fields.':'कृपया सभी आवश्यक फ़ील्ड भरें।',
    'Please fill all required fields.':'कृपया सभी आवश्यक फ़ील्ड भरें।',
    'Please fill in all required fields (marked *).':'कृपया सभी आवश्यक फ़ील्ड भरें (* से चिह्नित)।',
    'Please fill in all required fields (marked with *).':'कृपया सभी आवश्यक फ़ील्ड भरें (* से चिह्नित)।',

    // ── Members page ──
    'Society Members':'सोसायटी सदस्य', 'Manage and view all resident records':'सभी निवासी रिकॉर्ड प्रबंधित करें और देखें',
    '📤 Import from Excel':'📤 एक्सेल से इंपोर्ट करें', '+ Add New Member':'+ नया सदस्य जोड़ें',
    'Total Members':'कुल सदस्य', '2 BHK Units':'2 बीएचके इकाइयाँ', '3 BHK Units':'3 बीएचके इकाइयाँ',
    'Shops':'दुकानें', 'Shop':'दुकान', 'e.g. Kingston - 05, Shop - 12':'उदा. किंग्स्टन - 05, दुकान - 12',
    'Row Houses':'रो हाउस', '🏠 Transfer Ownership':'🏠 स्वामित्व हस्तांतरण', 'Transfer ownership':'स्वामित्व हस्तांतरण', '🏠 House Maintenance History':'🏠 मकान का मेंटेनेंस इतिहास', 'Current owner':'वर्तमान मालिक', 'Previous owner':'पूर्व मालिक', 'Inactive / Moved Out':'निष्क्रिय / स्थानांतरित',
    'All':'सभी', '2 BHK (Regal / Sunshine)':'2 बीएचके (रीगल / सनशाइन)',
    '2 BHK (Regal / A-1 / A-2)':'2 बीएचके (रीगल / A-1 / A-2)',
    '3 BHK (Kingston)':'3 बीएचके (किंग्स्टन)', 'Row House (Antillia / Empress / Royale)':'रो हाउस (एंटीलिया / एम्प्रेस / रोयाल)',
    'Member':'सदस्य', 'Unit Type':'इकाई प्रकार', 'House No.':'मकान नंबर', 'Contact':'संपर्क',
    '← Back to Members':'← सदस्यों पर वापस', 'Member Name':'सदस्य का नाम',
    '2 BHK':'2 बीएचके', '3 BHK':'3 बीएचके', 'Row House':'रो हाउस',
    'Add New Member':'नया सदस्य जोड़ें', 'Edit Member Profile':'सदस्य प्रोफ़ाइल संपादित करें',
    'Owner / Resident Photo':'मालिक / निवासी फ़ोटो', 'Click to upload photo':'फ़ोटो अपलोड करने के लिए क्लिक करें',
    'JPG, PNG, WEBP · Max 2 MB':'JPG, PNG, WEBP · अधिकतम 2 MB',
    'Full Name *':'पूरा नाम *', 'Unit Type *':'इकाई प्रकार *', 'Select type':'प्रकार चुनें',
    'Phone Number *':'फ़ोन नंबर *', 'Email':'ईमेल', 'Vehicle Registration No.':'वाहन पंजीकरण नंबर',
    'Emergency Contact Name':'आपातकालीन संपर्क नाम', 'Emergency Contact Phone':'आपातकालीन संपर्क फ़ोन',
    'Add Member':'सदस्य जोड़ें',
    'Please fill in all required fields (Name, House No., Unit Type, Phone).':'कृपया सभी आवश्यक फ़ील्ड भरें (नाम, मकान नंबर, इकाई प्रकार, फ़ोन)।',
    'Could not save member.':'सदस्य सहेजा नहीं जा सका।',
    'Permanently delete this member? This cannot be undone.':'इस सदस्य को स्थायी रूप से हटाएँ? इसे पूर्ववत नहीं किया जा सकता।',
    'Reactivate this member?':'इस सदस्य को पुनः सक्रिय करें?',
    '✅ Reactivate':'✅ पुनः सक्रिय करें', '🚫 Mark Inactive':'🚫 निष्क्रिय चिह्नित करें',
    'Active':'सक्रिय', 'Inactive':'निष्क्रिय',

    // ── Excel import (member) ──
    'Import Members from Excel':'एक्सेल से सदस्य इंपोर्ट करें',
    'Required columns:':'आवश्यक कॉलम:', 'Optional columns:':'वैकल्पिक कॉलम:',
    'Unit type auto-detection:':'इकाई प्रकार स्वतः पहचान:',
    'rows. Review below, then confirm to import into the live database.':'पंक्तियाँ। नीचे समीक्षा करें, फिर लाइव डेटाबेस में इंपोर्ट की पुष्टि करें।',
    'Found':'मिलीं', 'Sheet':'शीट', 'Type':'प्रकार', 'Phone':'फ़ोन',
    'Confirm & Import':'पुष्टि करें और इंपोर्ट करें',

    // ── Finances ──
    'Record society expenses and view financial summaries':'सोसायटी के खर्च दर्ज करें और वित्तीय सारांश देखें',
    '⬆ Import Expenses (XLSX)':'⬆ खर्च इंपोर्ट करें (XLSX)', 'Import Expenses from Excel':'एक्सेल से खर्च इंपोर्ट करें',
    'Click to select .xlsx file':'.xlsx फ़ाइल चुनने के लिए क्लिक करें',
    'Expenses will be previewed before import':'इंपोर्ट से पहले खर्चों का पूर्वावलोकन दिखेगा',
    'Total Maintenance Collected':'कुल मेंटेनेंस संग्रह', 'All maintenance payments received':'प्राप्त सभी मेंटेनेंस भुगतान',
    'Expenses — This Month':'खर्च — इस माह', 'Expenses — This Year':'खर्च — इस वर्ष',
    'Net Surplus / Deficit':'शुद्ध अधिशेष / घाटा', 'Maintenance + receivables − FY expenses':'संग्रह घटा वित्तीय वर्ष के खर्च',
    'Record New Expense':'नया खर्च दर्ज करें', 'Date *':'तारीख *', 'Category *':'श्रेणी *',
    'Select category':'श्रेणी चुनें', 'Enter new category name:':'नई श्रेणी का नाम दर्ज करें:',
    'Lift Maintenance':'लिफ्ट मेंटेनेंस', 'Labour Salary':'श्रमिक वेतन', 'Security Salary':'सुरक्षा वेतन',
    'Housekeeping Salary':'हाउसकीपिंग वेतन', 'Electricity':'बिजली', 'Water':'पानी',
    'Generator Maintenance':'जनरेटर मेंटेनेंस', 'Plumbing Repairs':'प्लंबिंग मरम्मत',
    'Electrical Repairs':'विद्युत मरम्मत', 'Pest Control':'कीट नियंत्रण', 'Garden / Landscaping':'उद्यान / भूदृश्य',
    'CCTV / Security System':'सीसीटीवी / सुरक्षा प्रणाली', 'Painting':'पेंटिंग', 'Stationery':'स्टेशनरी',
    'Bank Charges':'बैंक शुल्क', 'Audit Fees':'ऑडिट शुल्क', 'Festival / Events':'त्योहार / आयोजन', 'Others':'अन्य',
    'Houskeeping Item':'हाउसकीपिंग सामग्री', 'Housekeeping Item':'हाउसकीपिंग सामग्री',
    'Transport':'परिवहन', 'Waste Disposal':'कचरा निपटान', 'NOC charges':'एनओसी शुल्क',
    'Repairs':'मरम्मत', 'Miscellaneous':'विविध', 'Legal Fees':'कानूनी शुल्क', 'Insurance':'बीमा',
    'Description *':'विवरण *', 'Payee Name *':'भुगतानकर्ता नाम *', 'GST Number':'जीएसटी नंबर',
    '(optional)':'(वैकल्पिक)', 'Amount (₹) *':'राशि (₹) *', 'GST Amount (₹)':'जीएसटी राशि (₹)',
    'Bill Number':'बिल नंबर', 'Payment Method *':'भुगतान विधि *', 'Select method':'विधि चुनें',
    'Cash':'नकद', 'Cheque':'चेक', 'UPI':'यूपीआई', 'NEFT':'एनईएफटी', 'RTGS':'आरटीजीएस',
    'Demand Draft':'डिमांड ड्राफ्ट', 'Online':'ऑनलाइन',
    'Save Expense':'खर्च सहेजें', 'Edit Expense':'खर्च संपादित करें', 'Expense Records':'खर्च रिकॉर्ड',
    'Could not save expense. Please try again.':'खर्च सहेजा नहीं जा सका। कृपया पुनः प्रयास करें।',
    'Could not update expense.':'खर्च अपडेट नहीं हो सका।',
    'Could not find expense data. Please reload the page.':'खर्च डेटा नहीं मिला। कृपया पृष्ठ पुनः लोड करें।',
    'Loading expenses...':'खर्च लोड हो रहे हैं...',
    'All Months':'सभी माह', 'All Categories':'सभी श्रेणियाँ',
    'Voucher No.':'वाउचर नंबर', 'Date':'तारीख', 'Category':'श्रेणी', 'Description':'विवरण',
    'Payee':'भुगतानकर्ता', 'Amount':'राशि', 'Mode':'माध्यम', 'Payee / Vendor':'भुगतानकर्ता / विक्रेता',
    'Amount (₹)':'राशि (₹)',

    // ── Months ──
    'January':'जनवरी','February':'फ़रवरी','March':'मार्च','April':'अप्रैल','May':'मई','June':'जून',
    'July':'जुलाई','August':'अगस्त','September':'सितंबर','October':'अक्टूबर','November':'नवंबर','December':'दिसंबर',
    'Jan':'जन','Feb':'फ़र','Mar':'मार्च','Apr':'अप्रैल','Jun':'जून','Jul':'जुल','Aug':'अग','Sep':'सित','Oct':'अक्ट','Nov':'नव','Dec':'दिस',

    // ── Notices ──
    'Official announcements and circulars for residents':'निवासियों के लिए आधिकारिक घोषणाएँ और परिपत्र',
    '+ Post Notice':'+ सूचना पोस्ट करें', 'Post Notice':'सूचना पोस्ट करें', 'Post New Notice':'नई सूचना पोस्ट करें',
    'Posting…':'पोस्ट हो रहा है…', 'Edit Notice':'सूचना संपादित करें',
    'General':'सामान्य', 'Maintenance':'मेंटेनेंस', 'Events':'आयोजन', 'Emergency':'आपातकाल',
    'Finance':'वित्त', 'Legal':'कानूनी',
    'Notice Title *':'सूचना शीर्षक *', 'Notice Body *':'सूचना विवरण *', 'Visible To':'किसे दिखे',
    'All Residents':'सभी निवासी', 'Committee/Admin Only':'केवल समिति/एडमिन',
    'Valid Until (optional)':'इस तिथि तक मान्य (वैकल्पिक)', '📌 Pin this notice to the top':'📌 इस सूचना को ऊपर पिन करें',
    'Normal':'सामान्य', 'Urgent':'अत्यावश्यक',
    'Please fill in the required fields (Title, Body).':'कृपया आवश्यक फ़ील्ड भरें (शीर्षक, विवरण)।',
    'Please fill in the Title and Body.':'कृपया शीर्षक और विवरण भरें।',
    'Could not save notice.':'सूचना सहेजी नहीं जा सकी।', 'Delete this notice?':'इस सूचना को हटाएँ?',
    'Could not delete notice: ':'सूचना हटाई नहीं जा सकी: ',

    // ── Complaints ──
    'Log and track resident complaints to resolution':'निवासी शिकायतें दर्ज करें और समाधान तक ट्रैक करें',
    '+ Log Complaint':'+ शिकायत दर्ज करें', 'Log New Complaint':'नई शिकायत दर्ज करें',
    'Log Complaint':'शिकायत दर्ज करें',
    'Open':'खुली', 'Acknowledged':'स्वीकृत', 'In Progress':'प्रगति में', 'Resolved':'हल हो गई',
    'Closed':'बंद', 'Rejected':'अस्वीकृत',
    'Track and raise maintenance or service complaints':'मेंटेनेंस या सेवा शिकायतें दर्ज करें और ट्रैक करें',
    '+ Raise Complaint':'+ शिकायत दर्ज करें', 'New Complaint':'नई शिकायत',
    'Plumbing':'प्लंबिंग', 'Electrical':'विद्युत', 'Lift / Elevator':'लिफ्ट / एलिवेटर',
    'Common Area':'सामान्य क्षेत्र', 'Security':'सुरक्षा', 'Noise / Nuisance':'शोर / उपद्रव',
    'Noise / Disturbance':'शोर / बाधा', 'Parking':'पार्किंग',
    'Subject *':'विषय *', 'Subject / Title *':'विषय / शीर्षक *', 'Priority':'प्राथमिकता',
    'Low':'कम', 'Medium':'मध्यम', 'High':'उच्च', 'Critical':'अत्यधिक गंभीर',
    'Details *':'विवरण *', 'Detailed Description *':'विस्तृत विवरण *', 'Submit Complaint':'शिकायत सबमिट करें',
    "House Number (logging on a resident's behalf)":'मकान नंबर (निवासी की ओर से दर्ज)',
    'Please fill in category, subject, and details.':'कृपया श्रेणी, विषय और विवरण भरें।',
    'Could not submit complaint. Please try again.':'शिकायत सबमिट नहीं हो सकी। कृपया पुनः प्रयास करें।',
    'How was this resolved? (shown to the resident)':'यह कैसे हल हुआ? (निवासी को दिखाया जाएगा)',
    'Could not update status: ':'स्थिति अपडेट नहीं हो सकी: ',
    'Rate this repair from 1 (poor) to 5 (excellent):':'इस मरम्मत को 1 (खराब) से 5 (उत्कृष्ट) तक रेट करें:',
    'Please enter a number from 1 to 5.':'कृपया 1 से 5 तक कोई संख्या दर्ज करें।',
    'Any additional feedback? (optional — press OK to skip)':'कोई अतिरिक्त प्रतिक्रिया? (वैकल्पिक — छोड़ने के लिए OK दबाएँ)',
    'Any additional feedback? (optional)':'कोई अतिरिक्त प्रतिक्रिया? (वैकल्पिक)',
    'Could not submit rating: ':'रेटिंग सबमिट नहीं हो सकी: ',

    // ── Maintenance Payments (ledger) ──
    'Maintenance Payments':'मेंटेनेंस भुगतान',
    'Record of maintenance payments received, month by month':'माह-दर-माह प्राप्त मेंटेनेंस भुगतानों का रिकॉर्ड',
    '+ Add Payment':'+ भुगतान जोड़ें', 'Add Payment':'भुगतान जोड़ें', 'Edit Payment':'भुगतान संपादित करें',
    'Monthly Collections':'मासिक संग्रह',
    // ── Finance receivables, remark field, row popup, dashboard link ──
    'Dashboard':'डैशबोर्ड', 'Edit':'संपादित करें', 'Delete':'हटाएँ', 'Remark':'टिप्पणी', 'No remark':'कोई टिप्पणी नहीं',
    'Payment of previous dues':'पिछले बकाया का भुगतान', 'Other (type a remark)':'अन्य (टिप्पणी लिखें)', 'Type the remark':'टिप्पणी लिखें',
    '+ New Record':'+ नई प्रविष्टि', 'New Expense':'नया खर्च', 'New Receivable':'नई प्राप्ति',
    'Expenses':'खर्च', 'Receivables':'प्राप्तियाँ', 'Receivable Records':'प्राप्ति रिकॉर्ड',
    'Record New Receivable':'नई प्राप्ति दर्ज करें', 'Edit Receivable':'प्राप्ति संपादित करें', 'Save Receivable':'प्राप्ति सहेजें',
    'Receivables — This Year':'प्राप्तियाँ — इस वर्ष', 'Non-maintenance income':'मेंटेनेंस के अलावा आय',
    'Received From *':'किससे प्राप्त *', 'Received From':'किससे प्राप्त', 'House / Unit No.':'मकान / यूनिट नं.', 'Reference No.':'संदर्भ नं.',
    'Hall / Club House Booking':'हॉल / क्लब हाउस बुकिंग', 'Facility Charges':'सुविधा शुल्क',
    'Penalty / Late Fee':'जुर्माना / विलंब शुल्क', 'Donation':'दान', 'Bank Interest':'बैंक ब्याज', 'Advertisement / Hoarding':'विज्ञापन / होर्डिंग',
    'Rent / Lease':'किराया / पट्टा', 'Event / Sponsorship':'आयोजन / प्रायोजन', 'Refund Received':'वापसी प्राप्त',
    'All Categories':'सभी श्रेणियाँ', 'Maintenance + receivables − FY expenses':'मेंटेनेंस + प्राप्तियाँ − वित्त वर्ष का खर्च',
    'House No.':'मकान नं.', 'Bill book no.':'बिल बुक नं.', 'Receipt no.':'रसीद नं.', 'Voucher no.':'वाउचर नं.', 'Bill no.':'बिल नं.',
    'GST amount':'जीएसटी राशि', 'GST no.':'जीएसटी नं.', 'Reference no.':'संदर्भ नं.', 'House / Unit':'मकान / यूनिट',
    'Show password':'पासवर्ड दिखाएँ', 'Hide password':'पासवर्ड छिपाएँ',
    'Install app':'ऐप इंस्टॉल करें', 'Install the KGC app':'KGC ऐप इंस्टॉल करें', 'Install':'इंस्टॉल', 'Got it':'समझ गया',
    'Open it from your home screen, full screen.':'इसे होम स्क्रीन से पूरी स्क्रीन में खोलें।',
    'Duration':'अवधि', 'Starting quarter':'शुरुआती तिमाही', '1 quarter':'1 तिमाही', '2 quarters':'2 तिमाही', '3 quarters':'3 तिमाही',
    'Covers':'अवधि कवर', 'Full year':'पूरा वर्ष', 'Not recorded (older entry)':'दर्ज नहीं (पुरानी प्रविष्टि)',
    'Export CSV':'CSV निर्यात करें', 'What to export':'क्या निर्यात करना है', 'Payments received (list of entries)':'प्राप्त भुगतान (प्रविष्टियों की सूची)',
    'Paid & unpaid list (by house)':'भुगतान किए व बकाया की सूची (मकान अनुसार)', 'Period':'अवधि', 'Complete financial year':'पूरा वित्तीय वर्ष',
    'Quarter 1 (Apr–Jun)':'तिमाही 1 (अप्रैल–जून)', 'Quarter 2 (Jul–Sep)':'तिमाही 2 (जुलाई–सितंबर)', 'Quarter 3 (Oct–Dec)':'तिमाही 3 (अक्टूबर–दिसंबर)',
    'Quarter 4 (Jan–Mar)':'तिमाही 4 (जनवरी–मार्च)', 'Custom date range':'अपनी तिथि सीमा', 'Financial year':'वित्तीय वर्ष', 'From':'से', 'To':'तक', 'Show':'दिखाएँ',
    'All houses (with Paid / Unpaid status)':'सभी मकान (भुगतान / बकाया स्थिति सहित)', 'Paid only':'केवल भुगतान किए', 'Unpaid & part-paid only':'केवल बकाया व आंशिक भुगतान',
    'Download CSV':'CSV डाउनलोड करें', 'Import Payments from Excel':'एक्सेल से भुगतान इंपोर्ट करें',
    'Dues & Penalties':'बकाया और जुर्माना', 'Quarterly Dues':'त्रैमासिक बकाया',
    'Residents With Dues':'बकाया वाले निवासी', 'Total Maintenance Due':'कुल मेंटेनेंस बकाया',
    'Total Penalties':'कुल जुर्माना', 'Total Outstanding':'कुल बकाया',
    'House':'मकान', 'Resident':'निवासी', 'Total Owed':'कुल देय',
    'Paid':'भुगतान किया', 'Unpaid':'अवैतनिक', 'No residents found.':'कोई निवासी नहीं मिला।',
    'No dues on record.':'कोई बकाया दर्ज नहीं।', 'Could not update: ':'अपडेट नहीं हो सका: ',
    'Search name / house…':'नाम / मकान खोजें…',
    'Expected columns:':'अपेक्षित कॉलम:', 'Name':'नाम', 'House No':'मकान नंबर',
    'Bill Book':'बिल बुक', 'Receipt No.':'रसीद नंबर', 'Bill Book No.':'बिल बुक नंबर',
    'Payer Name *':'भुगतानकर्ता नाम *', 'Amount (Rs.) *':'राशि (₹) *', 'Reading file...':'फ़ाइल पढ़ी जा रही है...',
    'Date, payer name, and amount are required.':'तारीख, भुगतानकर्ता नाम और राशि आवश्यक हैं।',
    'Could not save.':'सहेजा नहीं जा सका।', 'Could not find entry. Reload the page.':'प्रविष्टि नहीं मिली। पृष्ठ पुनः लोड करें।',
    'Delete this payment entry? This cannot be undone.':'इस भुगतान प्रविष्टि को हटाएँ? इसे पूर्ववत नहीं किया जा सकता।',
    'Could not delete: ':'हटाया नहीं जा सका: ', 'No data rows found across all sheets.':'किसी भी शीट में डेटा पंक्तियाँ नहीं मिलीं।',

    // ── Pay UPI ──
    'Pay Maintenance':'मेंटेनेंस भुगतान करें', '← Back':'← वापस', 'UPI ID':'यूपीआई आईडी',
    'Payment History':'भुगतान इतिहास', 'Maintenance Payment Receipt':'मेंटेनेंस भुगतान रसीद',
    'Payment not found.':'भुगतान नहीं मिला।',

    // ── Generate bills / record payment modals ──
    'Generate Monthly Bills':'मासिक बिल जनरेट करें', 'Billing Month *':'बिलिंग माह *',
    'Financial Year *':'वित्तीय वर्ष *', 'Generate':'जनरेट करें', 'Generating…':'जनरेट हो रहा है…',
    'Please choose a billing month.':'कृपया बिलिंग माह चुनें।', 'Could not generate bills.':'बिल जनरेट नहीं हो सके।',
    'Record Payment':'भुगतान दर्ज करें', 'Amount Paid (₹) *':'भुगतान की गई राशि (₹) *',
    'Payment Date':'भुगतान तारीख', 'Payment Mode *':'भुगतान माध्यम *',
    'Bank Transfer — NEFT':'बैंक ट्रांसफर — एनईएफटी', 'Bank Transfer — IMPS':'बैंक ट्रांसफर — आईएमपीएस',
    'Bank Transfer — RTGS':'बैंक ट्रांसफर — आरटीजीएस', 'Bill No. (manual receipt number)':'बिल नंबर (मैनुअल रसीद नंबर)',
    'Cheque Number *':'चेक नंबर *', 'Cheque Date *':'चेक तारीख *', 'Bank Name':'बैंक का नाम',
    'UPI Transaction ID *':'यूपीआई ट्रांज़ैक्शन आईडी *', 'UPI App':'यूपीआई ऐप',
    'Bank Reference No. (UTR) *':'बैंक संदर्भ नंबर (UTR) *', 'DD Number':'डीडी नंबर', 'DD Date':'डीडी तारीख',
    'DD Bank Name':'डीडी बैंक का नाम', 'Remarks':'टिप्पणी',
    'Please enter a valid amount.':'कृपया एक मान्य राशि दर्ज करें।', 'Could not record payment.':'भुगतान दर्ज नहीं हो सका।',
    'Edit Bill':'बिल संपादित करें', 'Total Amount Due (₹) *':'कुल देय राशि (₹) *',
    'Amount Paid (₹)':'भुगतान की गई राशि (₹)', 'Waived':'माफ़', 'Total Due (₹)':'कुल देय (₹)',
    'Balance Due (₹)':'शेष देय (₹)', 'Please enter a valid amount due.':'कृपया एक मान्य देय राशि दर्ज करें।',
    'Could not update bill.':'बिल अपडेट नहीं हो सका।',

    // ── Facilities ──
    'Manage common areas, memberships and bookings':'सामान्य क्षेत्र, सदस्यता और बुकिंग प्रबंधित करें',
    'Overview':'अवलोकन', 'Gym & Membership':'जिम और सदस्यता', 'The Gymnasium':'जिम',
    'Located on the ground floor of the Club House':'क्लब हाउस के भूतल पर स्थित',
    'Hours':'समय', 'Area':'क्षेत्रफल', 'Equipment':'उपकरण', 'Trainers':'प्रशिक्षक',
    'Classes':'कक्षाएँ', 'Max Capacity':'अधिकतम क्षमता', 'View Membership Plans →':'सदस्यता योजनाएँ देखें →',
    'The Community Hall':'सामुदायिक हॉल', 'Capacity':'क्षमता', 'Bookable Hours':'बुकिंग योग्य समय',
    'Notice Required':'आवश्यक सूचना', 'Catering Kitchen':'कैटरिंग रसोई', 'Included':'शामिल',
    'Book the Hall →':'हॉल बुक करें →', 'Other Club House Amenities':'अन्य क्लब हाउस सुविधाएँ',
    'Available to all members — no prior booking required':'सभी सदस्यों के लिए उपलब्ध — पूर्व बुकिंग आवश्यक नहीं',
    'Library & Reading Room':'पुस्तकालय और वाचनालय',
    'Quiet space with 500+ books, newspapers, and magazines. Open 8 AM – 9 PM daily.':'500+ पुस्तकों, समाचार पत्रों और पत्रिकाओं के साथ शांत स्थान। प्रतिदिन सुबह 8 – रात 9 बजे तक खुला।',
    'Indoor Games Room':'इनडोर गेम्स रूम',
    'Billiards, carrom, table tennis, and chess — for residents of all age groups.':'बिलियर्ड्स, कैरम, टेबल टेनिस और शतरंज — सभी आयु वर्ग के निवासियों के लिए।',
    'Kids Play Zone':'बच्चों का खेल क्षेत्र',
    'Supervised indoor play area for children under 12. Foam flooring and safety equipment.':'12 वर्ष से कम उम्र के बच्चों के लिए निगरानी वाला इनडोर खेल क्षेत्र। फोम फ़्लोरिंग और सुरक्षा उपकरण।',
    'Café Lounge':'कैफ़े लाउंज',
    'Gym Membership Plans':'जिम सदस्यता योजनाएँ', 'Monthly':'मासिक', 'per month':'प्रति माह',
    'Full gym access':'पूर्ण जिम पहुँच', '2 group classes/week':'2 समूह कक्षाएँ/सप्ताह', 'Locker usage':'लॉकर उपयोग',
    'Quarterly':'त्रैमासिक', 'per 3 months · save ₹300':'प्रति 3 माह · ₹300 बचाएँ',
    '4 group classes/week':'4 समूह कक्षाएँ/सप्ताह', '1 free PT session':'1 निःशुल्क पीटी सत्र',
    'Annual':'वार्षिक', 'per year · save ₹2,100':'प्रति वर्ष · ₹2,100 बचाएँ',
    'Unlimited group classes':'असीमित समूह कक्षाएँ', 'Dedicated locker':'समर्पित लॉकर',
    '4 free PT sessions':'4 निःशुल्क पीटी सत्र', 'Guest pass (2/month)':'अतिथि पास (2/माह)',
    'Family':'परिवार', 'per year · up to 4 members':'प्रति वर्ष · 4 सदस्यों तक',
    'Full gym access all':'सभी के लिए पूर्ण जिम पहुँच', '6 free PT sessions':'6 निःशुल्क पीटी सत्र',
    'Kids yoga included':'बच्चों का योग शामिल', 'Enrol for Membership':'सदस्यता के लिए नामांकन करें',
    'Resident Name *':'निवासी नाम *', 'House Number *':'मकान नंबर *', 'Phone *':'फ़ोन *', 'Plan *':'योजना *',
    'Select plan':'योजना चुनें', 'Monthly — ₹800':'मासिक — ₹800', 'Quarterly — ₹2,100':'त्रैमासिक — ₹2,100',
    'Annual — ₹7,500':'वार्षिक — ₹7,500', 'Family Annual — ₹12,000':'पारिवारिक वार्षिक — ₹12,000',
    'Start Date *':'आरंभ तिथि *', 'Preferred Slot':'पसंदीदा समय', 'Any time':'किसी भी समय',
    'Early Morning (5:30–8:00 AM)':'सुबह जल्दी (5:30–8:00 AM)', 'Morning (8:00–11:00 AM)':'सुबह (8:00–11:00 AM)',
    'Afternoon (12:00–4:00 PM)':'दोपहर (12:00–4:00 PM)', 'Evening (4:00–7:00 PM)':'शाम (4:00–7:00 PM)',
    'Night (7:00–10:00 PM)':'रात (7:00–10:00 PM)', 'Enrol Now':'अभी नामांकन करें',
    '✓ Membership enrolment recorded successfully!':'✓ सदस्यता नामांकन सफलतापूर्वक दर्ज हुआ!',
    'Active Memberships':'सक्रिय सदस्यताएँ', 'House':'मकान', 'Plan':'योजना', 'Start Date':'आरंभ तिथि',
    'Valid Till':'मान्य तिथि तक', 'No memberships enrolled yet.':'अभी तक कोई सदस्यता नामांकित नहीं।',
    'Community Hall Booking':'सामुदायिक हॉल बुकिंग', 'New Booking Request':'नई बुकिंग अनुरोध',
    'Event Type *':'आयोजन प्रकार *', 'Select event type':'आयोजन प्रकार चुनें', 'Birthday Party':'जन्मदिन पार्टी',
    'Anniversary':'वर्षगाँठ', 'Pooja / Religious Ceremony':'पूजा / धार्मिक समारोह',
    'Kitty Party / Get-together':'किटी पार्टी / मिलन समारोह', 'Society Meeting':'सोसायटी बैठक',
    'Cultural Program':'सांस्कृतिक कार्यक्रम', 'Workshop / Seminar':'कार्यशाला / सेमिनार', 'Other':'अन्य',
    'Event Date *':'आयोजन तिथि *', 'Expected Guests *':'अपेक्षित अतिथि *', 'Start Time *':'आरंभ समय *',
    'Select time':'समय चुनें', 'End Time *':'समाप्ति समय *',
    'Additional Requirements / Notes':'अतिरिक्त आवश्यकताएँ / टिप्पणियाँ',
    '✓ Hall booking request submitted! You will be notified upon approval.':'✓ हॉल बुकिंग अनुरोध सबमिट हुआ! स्वीकृति पर आपको सूचित किया जाएगा।',
    'Submit Booking Request':'बुकिंग अनुरोध सबमिट करें', 'Upcoming Bookings':'आगामी बुकिंग',
    'Resident':'निवासी', 'Event':'आयोजन', 'Time':'समय', 'Guests':'अतिथि', 'No bookings yet.':'अभी तक कोई बुकिंग नहीं।',
    'Approve':'स्वीकृत करें', 'Reject':'अस्वीकृत करें', 'Mark completed':'पूर्ण चिह्नित करें',
    'Pending Approval':'स्वीकृति लंबित', 'Confirmed':'पुष्ट', 'Completed':'पूर्ण', 'Cancelled':'रद्द',
    'Enter reason for rejection (required):':'अस्वीकृति का कारण दर्ज करें (आवश्यक):',
    'Are you sure you want to cancel this booking?':'क्या आप वाकई यह बुकिंग रद्द करना चाहते हैं?',
    'Could not load bookings.':'बुकिंग लोड नहीं हो सकीं।', 'Could not update booking status.':'बुकिंग स्थिति अपडेट नहीं हो सकी।',
    'Could not submit booking. Please try again.':'बुकिंग सबमिट नहीं हो सकी। कृपया पुनः प्रयास करें।',

    // ── Resident home ──
    'Welcome':'स्वागत है', 'Good morning':'सुप्रभात', 'Good afternoon':'नमस्कार', 'Good evening':'शुभ संध्या',
    'Outstanding Balance':'बकाया राशि', 'Click to view bills':'बिल देखने के लिए क्लिक करें',
    'Open Complaints':'खुली शिकायतें', 'Click to view':'देखने के लिए क्लिक करें',
    '✏️ Edit Details':'✏️ विवरण संपादित करें', 'Recent Bills':'हाल के बिल',
    'My Maintenance Bills':'मेरे मेंटेनेंस बिल', 'Your payment history and outstanding dues':'आपका भुगतान इतिहास और बकाया राशि',
    'Total Outstanding':'कुल बकाया', 'Unpaid':'अवैतनिक', 'Partial':'आंशिक', 'Paid':'भुगतान किया',
    'Edit My Details':'मेरे विवरण संपादित करें', 'House Number':'मकान नंबर',
    'House number and unit type can only be changed by the admin.':'मकान नंबर और इकाई प्रकार केवल एडमिन द्वारा बदले जा सकते हैं।',
    'Primary Phone *':'प्राथमिक फ़ोन *', 'Alternate Phone':'वैकल्पिक फ़ोन', 'Email Address':'ईमेल पता',
    'Vehicle 1 Registration':'वाहन 1 पंजीकरण', 'Vehicle 2 Registration':'वाहन 2 पंजीकरण',
    'Primary phone number is required.':'प्राथमिक फ़ोन नंबर आवश्यक है।',
    'Your details have been updated successfully.':'आपके विवरण सफलतापूर्वक अपडेट हो गए।',
    'Could not load your profile data. Please try again.':'आपका प्रोफ़ाइल डेटा लोड नहीं हो सका। कृपया पुनः प्रयास करें।',
    'Could not save changes. Please try again.':'बदलाव सहेजे नहीं जा सके। कृपया पुनः प्रयास करें।',

    // ── Roles / misc labels ──
    'Administrator':'प्रशासक', 'President':'अध्यक्ष', 'Secretary':'सचिव', 'Treasurer':'कोषाध्यक्ष',
    'Manager':'प्रबंधक', 'Committee Member':'समिति सदस्य', 'Committee':'समिति', 'Admin':'एडमिन',
    'Surplus':'अधिशेष', 'Deficit':'घाटा', 'Today':'आज',
    'Billing Month':'बिलिंग माह', 'Financial Year':'वित्तीय वर्ष', 'Payment Mode':'भुगतान माध्यम',

    // ── Quarters ──
    'Q1 (Apr–Jun)':'ति1 (अप्रैल–जून)', 'Q2 (Jul–Sep)':'ति2 (जुलाई–सित)',
    'Q3 (Oct–Dec)':'ति3 (अक्ट–दिस)', 'Q4 (Jan–Mar)':'ति4 (जन–मार्च)'
  };

  const HINDI_FONT = "'Noto Sans Devanagari','Nirmala UI','Mangal',sans-serif";
  const SKIP_TAGS = new Set(['SCRIPT','STYLE','CODE','PRE','TEXTAREA','INPUT','SELECT','OPTION']);
  // OPTION is skipped by walker but handled separately below (their text needs translating too)

  let lang = localStorage.getItem('kgc_lang') || 'en';

  function translateTextNodes(root) {
    if (lang !== 'hi') return;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        let p = node.parentNode;
        while (p && p !== root.parentNode) {
          if (p.nodeType === 1 && (SKIP_TAGS.has(p.tagName) || p.classList.contains('no-translate')))
            return NodeFilter.FILTER_REJECT;
          p = p.parentNode;
        }
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(n => {
      if (n.__en !== undefined) return; // already translated
      const key = n.nodeValue.trim();
      if (DICT[key]) {
        n.__en = n.nodeValue;
        n.nodeValue = n.nodeValue.replace(key, DICT[key]);
      }
    });
  }

  // <option> elements and placeholders handled explicitly
  function translateOption(o) {
    if (o.__en === undefined && DICT[o.textContent.trim()]) {
      o.__en = o.textContent;
      o.textContent = DICT[o.textContent.trim()];
    }
  }
  function translatePlaceholder(el) {
    const ph = el.getAttribute && el.getAttribute('placeholder');
    if (!ph) return;
    const key = ph.trim();
    if (el.__enPh === undefined && DICT[key]) {
      el.__enPh = ph;
      el.setAttribute('placeholder', DICT[key]);
    }
  }
  function translateAttributes(root) {
    if (lang !== 'hi' || !root || root.nodeType !== 1) return;
    // the node itself may be an <option> or carry a placeholder
    if (root.tagName === 'OPTION') translateOption(root);
    if (root.hasAttribute && root.hasAttribute('placeholder')) translatePlaceholder(root);
    // and any descendants
    if (root.querySelectorAll) {
      root.querySelectorAll('option').forEach(translateOption);
      root.querySelectorAll('[placeholder]').forEach(translatePlaceholder);
    }
  }

  function restoreAll(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(n => {
      if (n.__en !== undefined) { n.nodeValue = n.__en; delete n.__en; }
    });
    root.querySelectorAll('option').forEach(o => {
      if (o.__en !== undefined) { o.textContent = o.__en; delete o.__en; }
    });
    root.querySelectorAll('[placeholder]').forEach(el => {
      if (el.__enPh !== undefined) { el.setAttribute('placeholder', el.__enPh); delete el.__enPh; }
    });
  }

  function applyLanguage() {
    const btn = document.getElementById('lang-toggle');
    if (lang === 'hi') {
      document.body.classList.add('lang-hi');
      document.documentElement.lang = 'hi';
      translateTextNodes(document.body);
      translateAttributes(document.body);
      if (btn) btn.textContent = 'English';
    } else {
      document.body.classList.remove('lang-hi');
      document.documentElement.lang = 'en';
      restoreAll(document.body);
      if (btn) btn.textContent = 'हिंदी';
    }
  }

  window.toggleLanguage = function () {
    lang = (lang === 'hi') ? 'en' : 'hi';
    localStorage.setItem('kgc_lang', lang);
    applyLanguage();
  };

  // Observe dynamically added content and translate it when in Hindi
  const observer = new MutationObserver(muts => {
    if (lang !== 'hi') return;
    muts.forEach(m => {
      m.addedNodes.forEach(node => {
        if (node.nodeType === 1) {         // element
          translateTextNodes(node);
          translateAttributes(node);
        } else if (node.nodeType === 3) {  // text
          const key = node.nodeValue && node.nodeValue.trim();
          if (key && DICT[key] && node.__en === undefined) {
            node.__en = node.nodeValue;
            node.nodeValue = node.nodeValue.replace(key, DICT[key]);
          }
        }
      });
    });
  });

  // Translate exact-match alert/confirm messages too
  const _alert = window.alert.bind(window);
  window.alert = function (msg) {
    if (lang === 'hi' && DICT[msg]) return _alert(DICT[msg]);
    return _alert(msg);
  };
  // Let toasts translate the same way alerts do
  window.__kgcTranslate = function (msg) {
    return (lang === 'hi' && DICT[msg]) ? DICT[msg] : msg;
  };
  const _confirm = window.confirm.bind(window);
  window.confirm = function (msg) {
    if (lang === 'hi' && DICT[msg]) return _confirm(DICT[msg]);
    return _confirm(msg);
  };

  function start() {
    applyLanguage();
    observer.observe(document.body, { childList: true, subtree: true });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();

/* ───── part 5 ───── */
/* ══════════════════════════════════════════════════════════════
   INSTALLABLE APP (PWA)
   - registers the service worker (offline opening, faster loads)
   - Android / desktop Chrome & Edge: uses the browser's install prompt
   - iPhone / iPad Safari: shows the "Add to Home Screen" steps
   - tells the user when the phone goes offline / comes back online
   ══════════════════════════════════════════════════════════════ */
(function () {
  var deferredPrompt = null;
  var isStandalone = function () {
    return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
  };
  var isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) ||
              (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  var DISMISS_KEY = 'kgc_install_dismissed';

  function recentlyDismissed() {
    try { var t = parseInt(localStorage.getItem(DISMISS_KEY) || '0', 10); return t && (Date.now() - t) < 14 * 864e5; }
    catch (e) { return false; }
  }
  // Can we offer installation right now?
  window.kgcCanInstall = function () { return !isStandalone() && (!!deferredPrompt || isIOS); };

  function refreshInstallUi() {
    var can = window.kgcCanInstall();
    var btn = document.getElementById('nav-install-btn');
    if (btn) btn.classList.toggle('show', can);
    var banner = document.getElementById('install-banner');
    if (banner) banner.classList.toggle('show', can && !recentlyDismissed() && window.innerWidth <= 768);
  }

  window.kgcInstallApp = function () {
    if (deferredPrompt) {
      var dp = deferredPrompt;
      dp.prompt();
      dp.userChoice.then(function (c) {
        deferredPrompt = null;
        if (c && c.outcome === 'accepted' && window.showToast) showToast('Installing the KGC app…', 'ok');
        refreshInstallUi();
      });
    } else if (isIOS) {
      document.getElementById('install-help-modal').classList.remove('hidden');
    }
  };
  window.closeInstallHelp = function () { document.getElementById('install-help-modal').classList.add('hidden'); };
  window.kgcDismissInstall = function () {
    try { localStorage.setItem(DISMISS_KEY, String(Date.now())); } catch (e) {}
    var b = document.getElementById('install-banner'); if (b) b.classList.remove('show');
  };

  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); deferredPrompt = e; refreshInstallUi(); });
  window.addEventListener('appinstalled', function () {
    deferredPrompt = null; refreshInstallUi();
    if (window.showToast) showToast('KGC app installed.', 'ok');
  });
  window.addEventListener('resize', refreshInstallUi);
  window.addEventListener('DOMContentLoaded', refreshInstallUi);

  // Offline / online notices
  window.addEventListener('offline', function () { if (window.showToast) showToast('You are offline. Society data needs an internet connection.', 'err'); });
  window.addEventListener('online',  function () { if (window.showToast) showToast('Back online.', 'ok'); });

  // Service worker (only on http/https, not when the file is opened directly)
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('sw.js').catch(function () { /* the site still works without it */ });
    });
  }
})();
