document.addEventListener('DOMContentLoaded', () => {
    
    // Elements
    const form = document.getElementById('evaluation-form');
    const perfTotalEl = document.getElementById('perf-total');
    const needTotalEl = document.getElementById('need-total');
    const commTotalEl = document.getElementById('comm-total');
    const grandTotalEl = document.getElementById('grand-total-val');
    const savedCountEl = document.getElementById('saved-count');
    const tableBody = document.getElementById('table-body');
    const emptyState = document.getElementById('empty-state');
    const tableContainer = document.querySelector('.table-container');
    
    // Radios
    const perfRadios = document.querySelectorAll('input[name^="p"]');
    const needRadios = document.querySelectorAll('input[name^="n"]');
    const commRadios = document.querySelectorAll('input[name^="c"]');
    
    // State
    let evaluations = [];
    try {
        evaluations = JSON.parse(localStorage.getItem('evaluations') || '[]') || [];
        if (!Array.isArray(evaluations)) evaluations = [];
    } catch (e) {
        console.error('خطأ في قراءة التقييمات المحفوظة:', e);
        evaluations = [];
    }
    let employeeData = [];
    let currentFilteredData = [];
    let activeEmpFilters = [];
    let currentFilteredEmpData = [];
    let currentMissingEmps = [];

    // توحيد الرقم الوظيفي للمقارنة (أرقام عربية/فارسية، Excel، مسافات)
    function normalizeEmpId(id) {
        if (id === null || id === undefined || id === '') return '';
        let str = String(id).trim();
        str = str.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
        str = str.replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
        if (/^\d+\.0+$/.test(str)) str = str.replace(/\.0+$/, '');
        return str;
    }

    function findEvalByEmpId(empId) {
        const norm = normalizeEmpId(empId);
        if (!norm) return -1;
        return evaluations.findIndex(ev => normalizeEmpId(ev.id) === norm);
    }

    function findEmployeeByEmpId(empId) {
        const norm = normalizeEmpId(empId);
        if (!norm) return null;
        return employeeData.find(e => normalizeEmpId(e['الرقم الوظيفي']) === norm) || null;
    }

    // --- منطق ملاحظة 2 ---
    // في الخارج + (أكاديمي/حالة خاصة/مسجون/اقترب تقاعده) → يُصنَّف ضمن الحالة الخاصة
    // في الخارج فقط → يُصنَّف «عادي»، مع إبقاء إمكانية استخراج «في الخارج» في التقرير
    const NOTES2_SPECIAL = ['أكاديمي', 'حالة خاصة', 'مسجون/ فقيد', 'اقترب تقاعده'];
    const NOTES2_ABROAD = 'في الخارج';
    const NOTES2_NORMAL = 'عادي';
    const NOTES2_ALL_TAGS = [NOTES2_ABROAD, ...NOTES2_SPECIAL, NOTES2_NORMAL];

    function parseNotes2Tags(notes2) {
        const raw = String(notes2 || NOTES2_NORMAL);
        const tags = [];
        NOTES2_ALL_TAGS.forEach(cat => {
            if (raw.includes(cat)) tags.push(cat);
        });
        if (tags.length === 0) tags.push(NOTES2_NORMAL);
        return tags;
    }

    function isNotes2Abroad(notes2) {
        const raw = String(notes2 || '').toLowerCase();
        return raw.includes('في الخارج') || raw.includes('بالخارج');
    }

    function hasNotes2SpecialCategory(notes2) {
        const tags = parseNotes2Tags(notes2);
        return NOTES2_SPECIAL.some(cat => tags.includes(cat));
    }

    function getEffectiveNotes2Categories(notes2) {
        const tags = parseNotes2Tags(notes2);
        const special = tags.filter(t => NOTES2_SPECIAL.includes(t));
        if (special.length > 0) return special;
        if (isNotes2Abroad(notes2)) return [NOTES2_NORMAL];
        if (tags.includes(NOTES2_NORMAL)) return [NOTES2_NORMAL];
        return [NOTES2_NORMAL];
    }

    function isEffectivelyNormalNotes2(notes2) {
        const effective = getEffectiveNotes2Categories(notes2);
        return effective.length === 1 && effective[0] === NOTES2_NORMAL;
    }

    function matchesNotes2Category(notes2, category) {
        if (category === NOTES2_ABROAD) return isNotes2Abroad(notes2);
        return getEffectiveNotes2Categories(notes2).includes(category);
    }

    function countNotes2Distribution(evals) {
        const counts = {};
        [...NOTES2_SPECIAL, NOTES2_NORMAL].forEach(c => { counts[c] = 0; });
        evals.forEach(ev => {
            getEffectiveNotes2Categories(ev.notes2).forEach(cat => {
                counts[cat] = (counts[cat] || 0) + 1;
            });
        });
        return counts;
    }

    const EMP_STORAGE_KEY = 'employeeData';
    const EMP_SESSION_KEY = 'employeeData_session';

    function computeEmployeeAge(dob) {
        if (!dob && dob !== 0) return '';
        try {
            let date;
            const num = Number(dob);
            if (!isNaN(num) && num > 10000 && num < 90000) {
                date = new Date(Math.round((num - 25569) * 86400 * 1000));
            } else {
                date = new Date(dob);
            }
            if (isNaN(date.getTime())) return '';
            return Math.abs(new Date(Date.now() - date.getTime()).getUTCFullYear() - 1970) || '';
        } catch (e) {
            return '';
        }
    }

    const EMP_EXCEL_KEYS = [
        'الرقم الوظيفي', 'الاسم', 'المسمى الوظيفي', 'الدائرة', 'القسم',
        'الجنس', 'نوع الوظيفة', 'تاريخ الميلاد', 'العمر', 'رقم الجوال',
        'البريد الالكتروني', 'نوع العقد'
    ];

    const EMP_HEADER_ALIASES = {
        'الرقم الوظيفي': ['الرقم الوظيفي', 'رقم وظيفي', 'الرقم', 'رقم الموظف', 'emp id', 'employee id', 'id', 'no', 'num'],
        'الاسم': ['الاسم', 'اسم الموظف', 'name', 'employee name'],
        'المسمى الوظيفي': ['المسمى الوظيفي', 'المسمى', 'الوظيفة', 'job title', 'title', 'position'],
        'الدائرة': ['الدائرة', 'الادارة', 'الإدارة', 'department', 'dept'],
        'القسم': ['القسم', 'section'],
        'الجنس': ['الجنس', 'gender', 'sex'],
        'نوع الوظيفة': ['نوع الوظيفة', 'نوع الوظيفه', 'job type'],
        'تاريخ الميلاد': ['تاريخ الميلاد', 'تاريخ الولادة', 'dob', 'birth date', 'birthdate'],
        'العمر': ['العمر', 'age'],
        'رقم الجوال': ['رقم الجوال', 'الجوال', 'الهاتف', 'mobile', 'phone'],
        'البريد الالكتروني': ['البريد الالكتروني', 'البريد الإلكتروني', 'الايميل', 'email', 'e-mail'],
        'نوع العقد': ['نوع العقد', 'العقد', 'contract']
    };

    function normalizeHeaderKey(key) {
        return String(key || '')
            .replace(/^\uFEFF/, '')
            .trim()
            .toLowerCase()
            .replace(/[أإآ]/g, 'ا')
            .replace(/ة/g, 'ه')
            .replace(/ى/g, 'ي')
            .replace(/[\u064B-\u065F]/g, '')
            .replace(/\s+/g, ' ');
    }

    function mapHeaderToField(header) {
        const norm = normalizeHeaderKey(header);
        if (!norm) return null;
        for (const [field, aliases] of Object.entries(EMP_HEADER_ALIASES)) {
            for (const a of aliases) {
                const aliasNorm = normalizeHeaderKey(a);
                if (!aliasNorm) continue;
                if (norm === aliasNorm) return field;
                if (aliasNorm.length >= 4 && (norm.includes(aliasNorm) || aliasNorm.includes(norm))) return field;
            }
        }
        return null;
    }

    function rowToArray(row) {
        if (Array.isArray(row)) return row;
        if (!row || typeof row !== 'object') return [];
        const keys = Object.keys(row).map(Number).filter(k => !isNaN(k)).sort((a, b) => a - b);
        if (keys.length) return keys.map(k => row[k]);
        return Object.values(row);
    }

    function rowLooksLikeHeader(row) {
        const arr = rowToArray(row);
        let hits = 0;
        arr.forEach(cell => { if (mapHeaderToField(cell)) hits++; });
        return hits >= 2;
    }

    function isEmployeeRow(emp) {
        const id = String(emp['الرقم الوظيفي'] || '').trim();
        const name = String(emp['الاسم'] || '').trim();
        return (id.length > 0) || (name.length > 1);
    }

    function buildEmpFromRowArray(row) {
        const arr = rowToArray(row);
        if (arr.length < 2) return null;
        const emp = {};
        EMP_EXCEL_KEYS.forEach((key, idx) => {
            if (arr[idx] !== undefined && arr[idx] !== null && String(arr[idx]).trim() !== '') {
                emp[key] = arr[idx];
            }
        });
        const clean = sanitizeEmployeeRecord(emp);
        return isEmployeeRow(clean) ? clean : null;
    }

    function parseEmployeeExcelRows(worksheet) {
        if (!worksheet || !worksheet['!ref']) return [];

        const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '', blankrows: false });
        if (!aoa || aoa.length === 0) return [];

        let headerRowIndex = -1;
        let colMap = {};

        for (let r = 0; r < Math.min(aoa.length, 15); r++) {
            const row = rowToArray(aoa[r]);
            const mapping = {};
            let matched = 0;
            row.forEach((cell, colIdx) => {
                const field = mapHeaderToField(cell);
                if (field && mapping[field] === undefined) {
                    mapping[field] = colIdx;
                    matched++;
                }
            });
            if (matched >= 1 && (mapping['الرقم الوظيفي'] !== undefined || mapping['الاسم'] !== undefined)) {
                headerRowIndex = r;
                colMap = mapping;
                if (matched >= 2) break;
            }
        }

        const records = [];

        if (headerRowIndex >= 0) {
            for (let i = headerRowIndex + 1; i < aoa.length; i++) {
                const row = rowToArray(aoa[i]);
                if (!row.length || rowLooksLikeHeader(row)) continue;
                const emp = {};
                Object.entries(colMap).forEach(([field, colIdx]) => {
                    const val = row[colIdx];
                    if (val !== undefined && val !== null && String(val).trim() !== '') {
                        emp[field] = val;
                    }
                });
                const clean = sanitizeEmployeeRecord(emp);
                if (isEmployeeRow(clean)) records.push(clean);
            }
            if (records.length > 0) return records;
        }

        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: '', blankrows: false });
        if (jsonData && jsonData.length > 0) {
            jsonData.forEach(raw => {
                const emp = {};
                Object.keys(raw || {}).forEach(key => {
                    if (String(key).startsWith('__')) return;
                    const field = mapHeaderToField(key);
                    if (field) emp[field] = raw[key];
                });
                const clean = sanitizeEmployeeRecord(emp);
                if (isEmployeeRow(clean)) records.push(clean);
            });
            if (records.length > 0) return records;
        }

        for (let i = 0; i < aoa.length; i++) {
            const row = rowToArray(aoa[i]);
            if (!row.length || rowLooksLikeHeader(row)) continue;
            const clean = buildEmpFromRowArray(row);
            if (clean) records.push(clean);
        }
        return records;
    }

    function parseEmployeeExcelWorkbook(workbook) {
        if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return [];
        let best = [];
        workbook.SheetNames.forEach(sheetName => {
            const ws = workbook.Sheets[sheetName];
            const rows = parseEmployeeExcelRows(ws);
            if (rows.length > best.length) best = rows;
        });
        return best;
    }

    function parseEmployeeCsvText(text) {
        if (!text || typeof XLSX === 'undefined') return [];
        try {
            const workbook = XLSX.read(text, { type: 'string', raw: false });
            return parseEmployeeExcelWorkbook(workbook);
        } catch (e) {
            console.error('CSV parse error:', e);
            return [];
        }
    }

    function handleEmployeeExcelFile(file) {
        if (!file) return;
        if (typeof XLSX === 'undefined') {
            alert('مكتبة Excel غير محمّلة. تأكد من الاتصال بالإنترنت وأعد تحميل الصفحة.');
            return;
        }

        const isCsv = /\.csv$/i.test(file.name);
        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                let records = [];
                if (isCsv) {
                    records = parseEmployeeCsvText(event.target.result);
                } else {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: true, cellText: false });
                    records = parseEmployeeExcelWorkbook(workbook);
                }

                if (records.length > 0) {
                    employeeData = records;
                    const saved = saveEmployeeDataToStorage();
                    pushToGitHub(true);
                    if (typeof renderEmpTable === 'function') renderEmpTable();
                    const empTab = document.querySelector('.tab-btn[data-target="emp-data-view"]');
                    if (empTab) empTab.click();
                    if (saved) {
                        alert('تم استيراد ' + employeeData.length + ' سجل بنجاح وحفظها!');
                    } else {
                        alert('تم قراءة ' + employeeData.length + ' سجل لكن فشل الحفظ المحلي. جرّب تصدير نسخة احتياطية فوراً.');
                    }
                } else {
                    alert('لم يُعثر على بيانات في الملف.\nتأكد أن الأعمدة تبدأ بـ: الرقم الوظيفي، الاسم، المسمى الوظيفي...\nأو أن الصفوف تبدأ برقم الموظف ثم الاسم.');
                }
            } catch (err) {
                console.error('خطأ استيراد الموظفين:', err);
                alert('خطأ في قراءة الملف: ' + (err.message || err) + '\nتأكد أن الملف بصيغة .xlsx أو .xls أو .csv');
            }
        };
        reader.onerror = function() {
            alert('تعذر قراءة الملف من الجهاز.');
        };
        if (isCsv) reader.readAsText(file, 'UTF-8');
        else reader.readAsArrayBuffer(file);
    }

    function setupEmployeeImport() {
        const input = document.getElementById('emp-excel-upload');
        if (!input) return;

        if (input.dataset.bound !== '1') {
            input.dataset.bound = '1';
            input.addEventListener('change', function(e) {
                const file = e.target.files && e.target.files[0];
                handleEmployeeExcelFile(file);
                e.target.value = '';
            });
        }

        const btn = document.getElementById('emp-excel-import-btn');
        if (btn && btn.dataset.bound !== '1') {
            btn.dataset.bound = '1';
            btn.addEventListener('click', function() {
                input.click();
            });
        }
    }

    function sanitizeEmployeeRecord(emp) {
        const clean = {};
        Object.keys(emp || {}).forEach(key => {
            let val = emp[key];
            if (val === undefined || val === null) return;
            if (val instanceof Date) {
                val = val.toISOString().split('T')[0];
            } else if (typeof val === 'object') {
                val = String(val);
            }
            clean[key] = val;
        });
        if (clean['الرقم الوظيفي']) {
            clean['الرقم الوظيفي'] = normalizeEmpId(clean['الرقم الوظيفي']);
        }
        if (clean['تاريخ الميلاد'] && !clean['العمر']) {
            const age = computeEmployeeAge(clean['تاريخ الميلاد']);
            if (age !== '') clean['العمر'] = age;
        }
        return clean;
    }

    function loadEmployeeDataFromStorage() {
        const sources = [
            () => localStorage.getItem(EMP_STORAGE_KEY),
            () => sessionStorage.getItem(EMP_SESSION_KEY)
        ];
        for (const getRaw of sources) {
            try {
                const raw = getRaw();
                if (!raw) continue;
                const parsed = JSON.parse(raw);
                if (!Array.isArray(parsed)) continue;
                return parsed.map(sanitizeEmployeeRecord);
            } catch (e) {
                console.error('خطأ في تحميل بيانات الموظفين:', e);
            }
        }
        return [];
    }

    function saveEmployeeDataToStorage(silent = false) {
        try {
            const sanitized = employeeData.map(sanitizeEmployeeRecord);
            const json = JSON.stringify(sanitized);
            localStorage.setItem(EMP_STORAGE_KEY, json);
            try {
                sessionStorage.setItem(EMP_SESSION_KEY, json);
            } catch (sessErr) {
                console.warn('تعذر الحفظ في sessionStorage:', sessErr);
            }
            employeeData = sanitized;
            touchDataTimestamp();
            pushToGitHub();
            return true;
        } catch (e) {
            console.error('خطأ في حفظ بيانات الموظفين:', e);
            if (!silent) {
                alert('تعذر حفظ بيانات الموظفين محلياً. قد يكون التخزين ممتلئاً — جرّب حذف بيانات قديمة أو متصفحاً آخر.');
            }
            return false;
        }
    }

    function reloadFromLocalStorage() {
        try {
            const rawEvals = localStorage.getItem('evaluations');
            if (rawEvals) {
                const parsed = JSON.parse(rawEvals);
                if (Array.isArray(parsed)) {
                    evaluations = migrateEvaluationList(parsed);
                }
            }
        } catch (e) {
            console.error('خطأ إعادة تحميل التقييمات:', e);
        }
        employeeData = loadEmployeeDataFromStorage();
    }

    function migrateEvaluationList(list) {
        return (list || []).map(ev => {
            if (ev.c1 === undefined && ev.n8 !== undefined) {
                ev.c1 = parseFloat(ev.n8) || 0;
            }
            ev.perfScore = (parseFloat(ev.p1)||0) + (parseFloat(ev.p2)||0) + (parseFloat(ev.p3)||0) + (parseFloat(ev.p4)||0) + (parseFloat(ev.p5)||0);
            ev.needScore = (parseFloat(ev.n1)||0) + (parseFloat(ev.n2)||0) + (parseFloat(ev.n3)||0) + (parseFloat(ev.n4)||0) + (parseFloat(ev.n5)||0) + (parseFloat(ev.n6)||0) + (parseFloat(ev.n7)||0);
            ev.commScore = parseFloat(ev.c1) || 0;
            ev.totalScore = ev.perfScore + ev.needScore;
            if (ev.id !== undefined && ev.id !== null && ev.id !== '') {
                ev.id = normalizeEmpId(ev.id);
            }
            return ev;
        });
    }

    function touchDataTimestamp() {
        localStorage.setItem('dataUpdatedAt', new Date().toISOString());
    }

    function pushToGitHub(immediate) {
        if (!window.GitHubSync) return;
        window.GitHubSync.scheduleSave({ evaluations, employeeData }, !!immediate);
    }

    function applyRemoteData(remote) {
        let changed = false;
        if (Array.isArray(remote.evaluations) && remote.evaluations.length > 0) {
            evaluations = migrateEvaluationList(remote.evaluations);
            changed = true;
        }
        if (Array.isArray(remote.employeeData) && remote.employeeData.length > 0) {
            employeeData = remote.employeeData.map(sanitizeEmployeeRecord);
            changed = true;
        }
        if (changed) {
            localStorage.setItem('evaluations', JSON.stringify(evaluations));
            localStorage.setItem(EMP_STORAGE_KEY, JSON.stringify(employeeData));
            sessionStorage.setItem(EMP_SESSION_KEY, JSON.stringify(employeeData));
            if (remote.updatedAt) localStorage.setItem('dataUpdatedAt', remote.updatedAt);
        }
        return changed;
    }

    function hasLocalData() {
        return evaluations.length > 0 || employeeData.length > 0;
    }

    function hasRemoteData(remote) {
        if (!remote) return false;
        return (remote.evaluations && remote.evaluations.length > 0) ||
            (remote.employeeData && remote.employeeData.length > 0);
    }

    async function syncFromGitHub(onlyIfLocalEmpty = false) {
        if (!window.GitHubSync || !window.GitHubSync.isConfigured()) {
            if (window.GitHubSync) window.GitHubSync.setStatus('التخزين المحلي', 'offline');
            return false;
        }

        reloadFromLocalStorage();
        const localHas = hasLocalData();

        if (onlyIfLocalEmpty && localHas) {
            window.GitHubSync.setStatus('محفوظ محلياً (' + evaluations.length + ' تقييم)', 'ok');
            return false;
        }

        const remote = await window.GitHubSync.loadAll();
        if (!hasRemoteData(remote)) {
            if (localHas) pushToGitHub(true);
            return false;
        }

        if (!localHas) {
            return applyRemoteData(remote);
        }

        const localTs = localStorage.getItem('dataUpdatedAt') || '';
        const remoteTs = remote.updatedAt || '';
        const remoteEvalCount = (remote.evaluations || []).length;
        const remoteEmpCount = (remote.employeeData || []).length;

        if (remoteTs && localTs && remoteTs > localTs) {
            if (remoteEvalCount >= evaluations.length) {
                return applyRemoteData(remote);
            }
            if (remoteEmpCount > employeeData.length) {
                employeeData = remote.employeeData.map(sanitizeEmployeeRecord);
                localStorage.setItem(EMP_STORAGE_KEY, JSON.stringify(employeeData));
                sessionStorage.setItem(EMP_SESSION_KEY, JSON.stringify(employeeData));
                return true;
            }
        }

        pushToGitHub(true);
        return false;
    }

    function refreshMissingEmpsView() {
        const compare = document.getElementById('compare-emp-btn');
        const addMissing = document.getElementById('add-missing-emp-btn');
        if (!compare || compare.style.display !== 'none') return;
        const evalIds = evaluations.map(e => normalizeEmpId(e.id)).filter(Boolean);
        currentMissingEmps = employeeData.filter(emp => {
            const empId = normalizeEmpId(emp['الرقم الوظيفي']);
            return empId && !evalIds.includes(empId);
        });
        renderEmpTable(currentMissingEmps);
        if (addMissing) {
            addMissing.style.display = currentMissingEmps.length > 0 ? 'inline-block' : 'none';
        }
    }

    evaluations = migrateEvaluationList(evaluations);
    localStorage.setItem('evaluations', JSON.stringify(evaluations));
    employeeData = loadEmployeeDataFromStorage();

    if (!localStorage.getItem('dataUpdatedAt') && hasLocalData()) {
        touchDataTimestamp();
    }

    function updateSavedCount() {
        if (savedCountEl) {
            savedCountEl.textContent = String(evaluations.length);
        }
    }

    updateSavedCount();
    setupEmployeeImport();

    window.addEventListener('beforeunload', () => {
        if (employeeData.length > 0) {
            try {
                const json = JSON.stringify(employeeData.map(sanitizeEmployeeRecord));
                localStorage.setItem(EMP_STORAGE_KEY, json);
                sessionStorage.setItem(EMP_SESSION_KEY, json);
            } catch (e) { /* ignore */ }
        }
        if (hasLocalData()) pushToGitHub(true);
    });

    // --- Tabs Logic ---
    const tabBtns = document.querySelectorAll('.tab-btn');
    const views = document.querySelectorAll('.view-section');

    function showStatsSubPanel(panelId) {
        document.querySelectorAll('.sub-tab-btn').forEach(b => {
            b.classList.toggle('active', b.getAttribute('data-subtarget') === panelId);
        });
        document.getElementById('stats-tables-panel').style.display = panelId === 'stats-tables-panel' ? 'block' : 'none';
        document.getElementById('stats-report-panel').style.display = panelId === 'stats-report-panel' ? 'block' : 'none';
        if (panelId === 'stats-tables-panel') {
            generateStatistics();
        } else if (panelId === 'stats-report-panel') {
            generateVisualReport();
        }
    }

    document.querySelectorAll('.sub-tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            showStatsSubPanel(btn.getAttribute('data-subtarget'));
        });
    });

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const externalUrl = btn.getAttribute('data-external');
            if (externalUrl) {
                window.location.href = externalUrl;
                return;
            }

            tabBtns.forEach(b => b.classList.remove('active'));
            views.forEach(v => v.style.display = 'none');
            
            btn.classList.add('active');
            const target = btn.getAttribute('data-target');
            document.getElementById(target).style.display = 'block';
            
            if (target === 'table-view') {
                renderTable();
            } else if (target === 'emp-data-view') {
                renderEmpTable();
            } else if (target === 'statistics-view') {
                showStatsSubPanel('stats-tables-panel');
            }
        });
    });

    // --- Statistics Generation Logic ---
    function generateStatistics() {
        const statsContainer = document.getElementById('stats-container');
        if (!statsContainer) return;

        statsContainer.innerHTML = ''; // Selectively generated cards below

        // Prepare combined data (evaluations + employee basic data where mapped)
        const activeEvalMap = {};
        evaluations.forEach(ev => activeEvalMap[ev.id] = ev);

        // Define a helper to render a table block
        const createStatsTable = (title, headers, dataRows) => {
            let html = `
                <div style="background: white; border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <h3 style="color: var(--primary-color); margin-top:0; font-size: 1.1rem; border-bottom: 2px solid var(--primary-color); padding-bottom: 8px;">${title}</h3>
                    <table style="width: 100%; margin-top: 10px; border-collapse: collapse;">
                        <thead>
                            <tr>${headers.map(h => `<th style="background:#f1f5f9; padding:8px; border: 1px solid var(--border-color); text-align: center; color: var(--text-color);">${h}</th>`).join('')}</tr>
                        </thead>
                        <tbody>
                            ${dataRows.map(row => `<tr>${row.map(cell => `<td style="padding:8px; border: 1px solid var(--border-color); text-align:center; color: var(--text-color);">${cell}</td>`).join('')}</tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            `;
            return html;
        };

        const calcPercent = (count, total) => {
            if (!total || total === 0) return '0%';
            return ((count / total) * 100).toFixed(1) + '%';
        };

        let htmlContent = '';

        // Added Overview section in Statistics (Numbers)
        const totalEmpsNum = employeeData.length;
        const totalEvalsNum = evaluations.length;
        const missingEmpsNum = Math.max(0, totalEmpsNum - totalEvalsNum);
        let validEvals = evaluations.filter(e => parseFloat(e.totalScore) > 0);
        let sumS = validEvals.reduce((sum, e) => sum + parseFloat(e.totalScore), 0);
        let avgS = validEvals.length > 0 ? (sumS / validEvals.length).toFixed(1) : 0;

        htmlContent += `
            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px;">
                <div style="background: white; padding: 15px; border-radius: 8px; border: 1px solid var(--border-color); text-align: center;">
                    <div style="color: #666; font-size: 0.9rem;">إجمالي الموظفين</div>
                    <div style="font-size: 1.8rem; font-weight: bold; color: var(--primary-color);">${totalEmpsNum || totalEvalsNum}</div>
                </div>
                <div style="background: white; padding: 15px; border-radius: 8px; border: 1px solid var(--border-color); text-align: center;">
                    <div style="color: #666; font-size: 0.9rem;">الموظفين المقيمين</div>
                    <div style="font-size: 1.8rem; font-weight: bold; color: #10b981;">${totalEvalsNum}</div>
                </div>
                <div style="background: white; padding: 15px; border-radius: 8px; border: 1px solid var(--border-color); text-align: center;">
                    <div style="color: #666; font-size: 0.9rem;">غير المقيمين</div>
                    <div style="font-size: 1.8rem; font-weight: bold; color: #ef4444;">${missingEmpsNum}</div>
                </div>
                <div style="background: white; padding: 15px; border-radius: 8px; border: 1px solid var(--border-color); text-align: center;">
                    <div style="color: #666; font-size: 0.9rem;">متوسط التقييم</div>
                    <div style="font-size: 1.8rem; font-weight: bold; color: #3b82f6;">${avgS} / 45</div>
                </div>
            </div>
        `;

        // 1. Stats by Job Type (نوع الوظيفة)
        let jobTypeCounts = {};
        employeeData.forEach(emp => {
            const jt = emp['نوع الوظيفة'] || 'غير محدد';
            jobTypeCounts[jt] = (jobTypeCounts[jt] || 0) + 1;
        });
        let jobTypeRows = Object.entries(jobTypeCounts).map(([type, count]) => [type, count, calcPercent(count, totalEmpsNum)]);
        jobTypeRows.push(["الإجمالي", totalEmpsNum, "100%"]);
        htmlContent += createStatsTable('إحصائية حسب نوع الوظيفة', ['نوع الوظيفة', 'العدد', 'النسبة'], jobTypeRows);

        // 2. Stats by Gender
        let genderCounts = {};
        employeeData.forEach(emp => {
            const g = emp['الجنس'] || 'ذكر'; // default assumption if blank
            genderCounts[g] = (genderCounts[g] || 0) + 1;
        });
        let genderRows = Object.entries(genderCounts).map(([g, c]) => [g, c, calcPercent(c, totalEmpsNum)]);
        genderRows.push(["الإجمالي", totalEmpsNum, "100%"]);
        htmlContent += createStatsTable('إحصائية حسب الجنس', ['الجنس', 'العدد', 'النسبة'], genderRows);

        // 3. Stats by Age Group (30-39, 40-49, 50-54, 55+) with Gender breakdown
        const calculateAgeLocal = (dobString) => {
            if (!dobString) return 0;
            let dob;
            if (!isNaN(Number(dobString)) && dobString > 10000 && dobString < 90000) {
                dob = new Date(Math.round((Number(dobString) - 25569) * 86400 * 1000));
            } else {
                dob = new Date(dobString);
            }
            if (isNaN(dob.getTime())) return 0;
            return Math.abs(new Date(Date.now() - dob.getTime()).getUTCFullYear() - 1970);
        };

        const ageGroups = {
            "من 30-39": { min: 30, max: 39, m: 0, f: 0 },
            "من 40-49": { min: 40, max: 49, m: 0, f: 0 },
            "من 50 إلى 54": { min: 50, max: 54, m: 0, f: 0 },
            "من 55 فأعلى": { min: 55, max: 200, m: 0, f: 0 },
            "غير متوفر/أخرى": { min: -1, max: 29, m: 0, f: 0 } // catch-all for missing or younger
        };

        employeeData.forEach(emp => {
            let age = parseInt(emp['العمر']) || calculateAgeLocal(emp['تاريخ الميلاد']);
            const gender = (emp['الجنس'] || 'ذكر') === 'ذكر' ? 'm' : 'f';
            
            for (const key in ageGroups) {
                if (age >= ageGroups[key].min && age <= ageGroups[key].max) {
                    ageGroups[key][gender]++;
                    break;
                }
            }
        });

        let ageRes = [];
        let tMale = 0, tFemale = 0, tAll = 0;
        for (const [groupName, data] of Object.entries(ageGroups)) {
            const sum = data.m + data.f;
            if(sum > 0 || groupName !== 'غير متوفر/أخرى') { // hide catch-all if 0
                ageRes.push([groupName, data.m, data.f, sum, calcPercent(sum, totalEmpsNum)]);
                tMale += data.m; tFemale += data.f; tAll += sum;
            }
        }
        ageRes.push(["الإجمالي", tMale, tFemale, tAll, "100%"]);
        htmlContent += createStatsTable('إحصائية حسب الأعمار والجنس', ['الفئة العمرية', 'ذكور', 'إناث', 'المجموع', 'النسبة الكلية'], ageRes);

        // إحصائية مكان التواجد
        const totalEmployees = employeeData.length;
        let outsideCountUI = 0;
        evaluations.forEach(ev => {
            if (isNotes2Abroad(ev.notes2)) outsideCountUI++;
        });
        const insideCountUI = totalEmployees - outsideCountUI;
        let locationRows = [
            ['في الخارج', outsideCountUI, calcPercent(outsideCountUI, totalEmployees)],
            ['في الداخل', insideCountUI, calcPercent(insideCountUI, totalEmployees)],
            ['الإجمالي', totalEmployees, "100%"]
        ];
        htmlContent += createStatsTable('إحصائية حسب مكان التواجد', ['مكان التواجد', 'العدد', 'النسبة'], locationRows);

        // 4. Stats by Note 2 Items
        const note2Counts = countNotes2Distribution(evaluations);
        const note2Categories = [...NOTES2_SPECIAL, NOTES2_NORMAL];
        
        let note2TotalUI = 0;
        let note2Rows = note2Categories.map(cat => {
            const count = note2Counts[cat] || 0;
            note2TotalUI += count;
            return [cat, count, calcPercent(count, evaluations.length)];
        });
        note2Rows.push(['الإجمالي', note2TotalUI, ""]);
        htmlContent += createStatsTable('إحصائية حسب ملاحظة 2 (من المقيمين)', ['البند', 'العدد', 'النسبة'], note2Rows);

        // إحصائية الاستمرارية (أ1)
        const p1Groups = {
            "الاستمرارية 0": 0,
            "الاستمرارية 1": 0,
            "الاستمرارية أكبر من 1": 0
        };
        evaluations.forEach(ev => {
            const p1 = parseInt(ev.p1) || 0;
            if (p1 === 0) p1Groups["الاستمرارية 0"]++;
            else if (p1 === 1) p1Groups["الاستمرارية 1"]++;
            else p1Groups["الاستمرارية أكبر من 1"]++;
        });
        
        let p1TotalUI = 0;
        let p1Rows = Object.entries(p1Groups).map(([g, count]) => {
            p1TotalUI += count;
            return [g, count, calcPercent(count, evaluations.length)];
        });
        p1Rows.push(['الإجمالي', p1TotalUI, "100%"]);
        htmlContent += createStatsTable('إحصائية الاستمرارية (أ1)', ['الفئة', 'العدد', 'النسبة'], p1Rows);

        // إحصائية معايير الحاجة بند 1 (n1)
        const n1Groups = {
            "الذين حصلوا على 0": 0,
            "الذين حصلوا على من 1 إلى 3": 0,
            "الذين حصلوا على من 4 فما فوق": 0
        };
        evaluations.forEach(ev => {
            const n1 = parseInt(ev.n1) || 0;
            if (n1 === 0) n1Groups["الذين حصلوا على 0"]++;
            else if (n1 <= 3) n1Groups["الذين حصلوا على من 1 إلى 3"]++;
            else n1Groups["الذين حصلوا على من 4 فما فوق"]++;
        });
        
        let n1TotalUI = 0;
        let n1Rows = Object.entries(n1Groups).map(([g, count]) => {
            n1TotalUI += count;
            return [g, count, calcPercent(count, evaluations.length)];
        });
        n1Rows.push(['الإجمالي', n1TotalUI, "100%"]);
        htmlContent += createStatsTable('حسب معايير الحاجة بند الحاجة 1', ['الفئة', 'العدد', 'النسبة'], n1Rows);

        // إحصائية عدد الأيام الوجاهي (ح2)
        const n2Groups = {
            "وجاهي 1 أو أقل": 0,
            "وجاهي من 2 إلى 4": 0,
            "وجاهي 5 أيام": 0
        };
        evaluations.forEach(ev => {
            const n2 = parseInt(ev.n2) || 0;
            if (n2 <= 1) n2Groups["وجاهي 1 أو أقل"]++;
            else if (n2 >= 2 && n2 <= 4) n2Groups["وجاهي من 2 إلى 4"]++;
            else n2Groups["وجاهي 5 أيام"]++;
        });

        let n2TotalUI = 0;
        let n2Rows = Object.entries(n2Groups).map(([g, count]) => {
            n2TotalUI += count;
            return [g, count, calcPercent(count, evaluations.length)];
        });
        n2Rows.push(['الإجمالي', n2TotalUI, "100%"]);
        htmlContent += createStatsTable('حسب عدد الأيام الوجاهي (ح2)', ['الفئة', 'العدد', 'النسبة'], n2Rows);

        // 8.5 Need 4 (N4) 
        const n4Groups = {
            "بدون 0": 0,
            "قليل 1": 0,
            "متوسط 2": 0,
            "كبير 3": 0
        };
        evaluations.forEach(ev => {
            const n4 = parseInt(ev.n4) || 0;
            if (n4 === 0) n4Groups["بدون 0"]++;
            else if (n4 === 1) n4Groups["قليل 1"]++;
            else if (n4 === 2) n4Groups["متوسط 2"]++;
            else n4Groups["كبير 3"]++;
        });

        let n4TotalUI = 0;
        let n4Rows = Object.entries(n4Groups).map(([g, count]) => {
            n4TotalUI += count;
            return [g, count, calcPercent(count, evaluations.length)];
        });
        n4Rows.push(['الإجمالي', n4TotalUI, "100%"]);
        htmlContent += createStatsTable('حسب حجم المعاملات (ح4)', ['الفئة', 'العدد', 'النسبة'], n4Rows);

        // --- New Feature: Need Score Groups (مجموع الحاجة) ---
        const needScoreGroups = {
            "أقل من 10": 0,
            "من 10-19": 0,
            "من 20-30": 0
        };
        evaluations.forEach(ev => {
            const ns = parseFloat(ev.needScore) || 0;
            if (ns < 10) needScoreGroups["أقل من 10"]++;
            else if (ns >= 10 && ns <= 19) needScoreGroups["من 10-19"]++;
            else needScoreGroups["من 20-30"]++;
        });
        let nsTotalUI = 0;
        let needScoreRows = Object.entries(needScoreGroups).map(([g, count]) => {
            nsTotalUI += count;
            return [g, count, calcPercent(count, evaluations.length)];
        });
        needScoreRows.push(['الإجمالي', nsTotalUI, "100%"]);
        htmlContent += createStatsTable('إحصائية مجموع معيار الحاجة', ['الفئة', 'العدد', 'النسبة'], needScoreRows);

        // 9. Performance Score with Note2 Normal
        const perfGroups = {
            "من 0-5": { min: 0, max: 5, count: 0 },
            "من 6-10": { min: 6, max: 10, count: 0 },
            "من 11-15": { min: 11, max: 15, count: 0 }
        };
        
        evaluations.forEach(ev => {
            if (isEffectivelyNormalNotes2(ev.notes2)) {
                const s = ev.perfScore || 0;
                for (const key in perfGroups) {
                    if (s >= perfGroups[key].min && s <= perfGroups[key].max) {
                        perfGroups[key].count++;
                        break;
                    }
                }
            }
        });
        
        let perfTotalUI = 0;
        let perfRows = Object.entries(perfGroups).map(([g, d]) => { perfTotalUI += d.count; return g; }); // temporary mapping
        perfRows = Object.entries(perfGroups).map(([g, d]) => [g, d.count, calcPercent(d.count, perfTotalUI)]);
        perfRows.push(['الإجمالي', perfTotalUI, "100%"]);
        htmlContent += createStatsTable('معيار الأداء (والملاحظة عادي)', ['الفئة', 'العدد', 'النسبة'], perfRows);

        // Inject computed HTML
        statsContainer.innerHTML = htmlContent;
    }

    // Function to generate and export stats
    const exportStats = (asZip = false) => {
        const calculateAgeLocalExport = (dobString) => {
            if (!dobString) return 0;
            let dob;
            if (!isNaN(Number(dobString)) && dobString > 10000 && dobString < 90000) {
                dob = new Date(Math.round((Number(dobString) - 25569) * 86400 * 1000));
            } else {
                dob = new Date(dobString);
            }
            if (isNaN(dob.getTime())) return 0;
            return Math.abs(new Date(Date.now() - dob.getTime()).getUTCFullYear() - 1970);
        };

        const wb = XLSX.utils.book_new();
        const zip = asZip ? new JSZip() : null;

        // Add sheet helper
        const addSheetToExport = (sheetName, ws) => {
            if (asZip) {
                const singleWb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(singleWb, ws, sheetName.substring(0, 31));
                const excelData = XLSX.write(singleWb, { bookType: 'xlsx', type: 'array' });
                zip.file(`${sheetName}.xlsx`, excelData);
            } else {
                XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
            }
        };

        // -------------------------------------------------------------
        // Sheet 0: ملخص الإحصائيات (Summary Statistics)
        // -------------------------------------------------------------
        let statsRows = [];
        const calcPercentExp = (count, total) => total ? ((count/total)*100).toFixed(1) + '%' : '0%';
        const totalEmps = employeeData.length;
        const totalEvals = evaluations.length;
        
        // 1. Stats by Job Type (نوع الوظيفة)
        let jobTypeCounts = {};
        employeeData.forEach(emp => {
            const jt = emp['نوع الوظيفة'] || 'غير محدد';
            jobTypeCounts[jt] = (jobTypeCounts[jt] || 0) + 1;
        });
        statsRows.push(["إحصائية حسب نوع الوظيفة"]);
        statsRows.push(["نوع الوظيفة", "العدد", "النسبة"]);
        Object.entries(jobTypeCounts).forEach(([type, count]) => statsRows.push([type, count, calcPercentExp(count, totalEmps)]));
        statsRows.push(["الإجمالي", employeeData.length, "100%"]);
        statsRows.push([]); // empty row spacer

        // 2. Stats by Gender
        let genderCounts = {};
        employeeData.forEach(emp => {
            const g = emp['الجنس'] || 'ذكر';
            genderCounts[g] = (genderCounts[g] || 0) + 1;
        });
        statsRows.push(["إحصائية حسب الجنس"]);
        statsRows.push(["الجنس", "العدد", "النسبة"]);
        Object.entries(genderCounts).forEach(([g, c]) => statsRows.push([g, c, calcPercentExp(c, totalEmps)]));
        statsRows.push(["الإجمالي", employeeData.length, "100%"]);
        statsRows.push([]);

        // 3. Stats by Age Group and Gender
        const ageGroupsExport = {
            "من 30-39": { min: 30, max: 39, m: 0, f: 0 },
            "من 40-49": { min: 40, max: 49, m: 0, f: 0 },
            "من 50 إلى 54": { min: 50, max: 54, m: 0, f: 0 },
            "من 55 فأعلى": { min: 55, max: 200, m: 0, f: 0 },
            "غير متوفر/أخرى": { min: -1, max: 29, m: 0, f: 0 }
        };

        employeeData.forEach(emp => {
            let age = parseInt(emp['العمر']) || calculateAgeLocalExport(emp['تاريخ الميلاد']);
            const gender = (emp['الجنس'] || 'ذكر') === 'ذكر' ? 'm' : 'f';
            
            for (const key in ageGroupsExport) {
                if (age >= ageGroupsExport[key].min && age <= ageGroupsExport[key].max) {
                    ageGroupsExport[key][gender]++;
                    break;
                }
            }
        });

        statsRows.push(["إحصائية حسب الأعمار والجنس"]);
        statsRows.push(["الفئة العمرية", "ذكور", "إناث", "المجموع", "النسبة للكل"]);
        let tMaleX = 0, tFemaleX = 0, tAllX = 0;
        for (const [groupName, data] of Object.entries(ageGroupsExport)) {
            const sum = data.m + data.f;
            if(sum > 0 || groupName !== 'غير متوفر/أخرى') {
                statsRows.push([groupName, data.m, data.f, sum, calcPercentExp(sum, totalEmps)]);
                tMaleX += data.m; tFemaleX += data.f; tAllX += sum;
            }
        }
        statsRows.push(["الإجمالي", tMaleX, tFemaleX, tAllX, "100%"]);
        statsRows.push([]);

        // إحصائية مكان التواجد
        const totalEmployeesExport = employeeData.length;
        let outsideCountExport = 0;
        evaluations.forEach(ev => {
            if (isNotes2Abroad(ev.notes2)) outsideCountExport++;
        });
        const insideCountExport = totalEmployeesExport - outsideCountExport;
        
        statsRows.push(["إحصائية حسب مكان التواجد"]);
        statsRows.push(["مكان التواجد", "العدد", "النسبة"]);
        statsRows.push(["في الخارج", outsideCountExport, calcPercentExp(outsideCountExport, totalEmployeesExport)]);
        statsRows.push(["في الداخل", insideCountExport, calcPercentExp(insideCountExport, totalEmployeesExport)]);
        statsRows.push(["الإجمالي", totalEmployeesExport, "100%"]);
        statsRows.push([]);

        // 4. Stats by Note 2 Items
        // 4. Stats by Note 2 Items (التصنيف الفعّال)
        const note2CountsExport = countNotes2Distribution(evaluations);
        const note2CategoriesExport = [...NOTES2_SPECIAL, NOTES2_NORMAL];
        
        let note2TotalExport = 0;
        statsRows.push(["إحصائية حسب ملاحظة 2"]);
        statsRows.push(["البند", "العدد", "النسبة"]);
        note2CategoriesExport.forEach(cat => {
            const count = note2CountsExport[cat] || 0;
            note2TotalExport += count;
            statsRows.push([cat, count, calcPercentExp(count, totalEvals)]);
        });
        statsRows.push(["الإجمالي", note2TotalExport, ""]);
        statsRows.push([]);

        // 9. Performance Score with Note2 Normal
        const perfGroupsExport = {
            "من 0-5": { min: 0, max: 5, count: 0 },
            "من 6-10": { min: 6, max: 10, count: 0 },
            "من 11-15": { min: 11, max: 15, count: 0 }
        };
        
        evaluations.forEach(ev => {
            if (isEffectivelyNormalNotes2(ev.notes2)) {
                const s = ev.perfScore || 0;
                for (const key in perfGroupsExport) {
                    if (s >= perfGroupsExport[key].min && s <= perfGroupsExport[key].max) {
                        perfGroupsExport[key].count++;
                        break;
                    }
                }
            }
        });
        
        let perfTotalExport = 0;
        Object.entries(perfGroupsExport).forEach(([g, d]) => { perfTotalExport += d.count; });
        statsRows.push(["معيار الأداء (والملاحظة عادي)"]);
        statsRows.push(["الفئة", "العدد", "النسبة"]);
        Object.entries(perfGroupsExport).forEach(([g, d]) => { statsRows.push([g, d.count, calcPercentExp(d.count, perfTotalExport)]); });
        statsRows.push(["الإجمالي", perfTotalExport, "100%"]);
        statsRows.push([]);

        addSheetToExport("ملخص الإحصائيات (1-4,9)", XLSX.utils.aoa_to_sheet(statsRows));

        // -------------------------------------------------------------
        // Detailed Sheets Helper
        // -------------------------------------------------------------
        const appendDetailedSheet = (sheetName, dataRows, extraHeaders = [], extraDataExtractors = [], addFullEval = false) => {
            const baseHeaders = ['الرقم الوظيفي', 'الاسم', 'الدائرة', 'القسم', 'تاريخ الميلاد', 'العمر', 'الجنس'];
            let evalHeaders = [];
            if (addFullEval) {
                evalHeaders = [
                    'أ1', 'أ2', 'أ3', 'أ4', 'أ5', 'مجموع الأداء',
                    'ح1', 'ح2', 'ح3', 'ح4', 'ح5', 'ح6', 'ح7', 'مجموع الحاجة',
                    'قرار اللجنة', 'المجموع الكلي', 'ملاحظات', 'ملاحظة2'
                ];
            }
            const allHeaders = [...baseHeaders, ...extraHeaders, ...evalHeaders];
            const rows = [allHeaders];

            dataRows.forEach(item => {
                let age = parseInt(item['العمر']) || calculateAgeLocalExport(item['تاريخ الميلاد'] || item.dob); // Handles both emp and eval
                let row = [
                    item['الرقم الوظيفي'] || item.id || '',
                    item['الاسم'] || item.name || '',
                    item['الدائرة'] || item.dept || '',
                    item['القسم'] || item.section || '',
                    item['تاريخ الميلاد'] || item.dob || '',
                    age,
                    item['الجنس'] || item.gender || ''
                ];
                extraDataExtractors.forEach(extractor => {
                    row.push(extractor(item));
                });

                if (addFullEval) {
                    const empId = normalizeEmpId(item['الرقم الوظيفي'] || item.id || '');
                    const ev = evaluations.find(x => normalizeEmpId(x.id) === empId);
                    if (ev) {
                        row.push(
                            ev.p1 ?? '', ev.p2 ?? '', ev.p3 ?? '', ev.p4 ?? '', ev.p5 ?? '', ev.perfScore ?? '',
                            ev.n1 ?? '', ev.n2 ?? '', ev.n3 ?? '', ev.n4 ?? '', ev.n5 ?? '', ev.n6 ?? '', ev.n7 ?? '', ev.needScore ?? '',
                            ev.c1 ?? '', ev.totalScore ?? '', ev.notes ?? '', ev.notes2 ?? ''
                        );
                    } else {
                        row.push('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
                    }
                }

                rows.push(row);
            });

            const ws = XLSX.utils.aoa_to_sheet(rows);
            
            // Apply Styles
            const borderStyle = {
                top: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } }
            };

            for (let i in ws) {
                if (i[0] === '!') continue;
                let cell = ws[i];
                
                cell.s = {
                    font: { name: "Simplified Arabic", sz: 12 },
                    border: borderStyle,
                    alignment: { horizontal: "center", vertical: "center", wrapText: true }
                };

                const rowNumStr = i.replace(/\D/g, '');
                if (rowNumStr === "1") {
                    cell.s.font.bold = true;
                    cell.s.fill = { fgColor: { rgb: "EAEAEA" } };
                }
            }

            if (!ws['!views']) ws['!views'] = [];
            ws['!views'].push({ rightToLeft: true });

            const wscols = [
                {wch: 15}, {wch: 35}, {wch: 20}, {wch: 20}, {wch: 15}, {wch: 10}, {wch: 10}
            ];
            extraHeaders.forEach(() => wscols.push({wch: 15}));
            if (addFullEval) {
                for (let i = 0; i < evalHeaders.length; i++) wscols.push({wch: 15});
            }
            ws['!cols'] = wscols;

            addSheetToExport(sheetName, ws);
        };

        // Sheet: جميع الأسماء
        appendDetailedSheet("جميع الموظفين (5)", employeeData, ['المسمى الوظيفي', 'نوع الوظيفة'], [e => e['المسمى الوظيفي'], e => e['نوع الوظيفة']], true);

        // Sheet: موظفي الداخل (الكل باستثناء في الخارج)
        appendDetailedSheet("موظفي الداخل", employeeData.filter(emp => {
            const empId = normalizeEmpId(emp['الرقم الوظيفي'] || emp.id || '');
            const ev = evaluations.find(x => normalizeEmpId(x.id) === empId);
            return !(ev && isNotes2Abroad(ev.notes2));
        }), ['المسمى الوظيفي', 'نوع الوظيفة'], [e => e['المسمى الوظيفي'], e => e['نوع الوظيفة']], true);

        // Sheet: موظفي في الخارج (استخراج منفصل)
        appendDetailedSheet("موظفي في الخارج", employeeData.filter(emp => {
            const empId = normalizeEmpId(emp['الرقم الوظيفي'] || emp.id || '');
            const ev = evaluations.find(x => normalizeEmpId(x.id) === empId);
            return ev && isNotes2Abroad(ev.notes2);
        }), ['المسمى الوظيفي', 'نوع الوظيفة'], [e => e['المسمى الوظيفي'], e => e['نوع الوظيفة']], true);

        // Sheet: خدمات
        appendDetailedSheet("خدمات (6)", employeeData.filter(e => String(e['نوع الوظيفة']).includes('خدمات')), ['نوع الوظيفة'], [e => e['نوع الوظيفة']], true);

        // Sheet: إداري
        appendDetailedSheet("إداري (7)", employeeData.filter(e => String(e['نوع الوظيفة']).includes('إداري') || String(e['نوع الوظيفة']).includes('اداري')), ['نوع الوظيفة'], [e => e['نوع الوظيفة']], true);

        // Sheet: العمر 58 فاكثر
        appendDetailedSheet("58 فأعلى (8)", employeeData.filter(e => {
            let age = parseInt(e['العمر']) || calculateAgeLocalExport(e['تاريخ الميلاد']);
            return age >= 58;
        }), [], [], true);

        // Sheet: معيار الحاجة بند 1 = 0 وملاحظة عادي (10)
        appendDetailedSheet("حاجة1صفر_عادي (10)", evaluations.filter(ev => parseInt(ev.n1) === 0 && isEffectivelyNormalNotes2(ev.notes2)), ['حاجة بند 1', 'ملاحظة 2'], [e => e.n1, e => e.notes2]);

        // Sheet: معيار الحاجة بند 1 > 3 (11)
        appendDetailedSheet("حاجة1 أكبرمن3 (11)", evaluations.filter(ev => parseInt(ev.n1) > 3), ['حاجة بند 1'], [e => e.n1]);

        // Sheets: التواجد الوجاهي
        appendDetailedSheet("وجاهي 1 أو أقل", evaluations.filter(ev => parseInt(ev.n2) <= 1), ['الوجاهي (ح2)'], [e => e.n2], true);
        appendDetailedSheet("وجاهي من 2 إلى 4", evaluations.filter(ev => parseInt(ev.n2) >= 2 && parseInt(ev.n2) <= 4), ['الوجاهي (ح2)'], [e => e.n2], true);
        appendDetailedSheet("وجاهي 5 أيام", evaluations.filter(ev => parseInt(ev.n2) >= 5), ['الوجاهي (ح2)'], [e => e.n2], true);

        // Sheets: حجم المعاملات ح4
        appendDetailedSheet("حجم المعاملات 0 بدون", evaluations.filter(ev => parseInt(ev.n4) === 0), ['حجم المعاملات (ح4)'], [e => e.n4], true);
        appendDetailedSheet("حجم المعاملات 1 قليل", evaluations.filter(ev => parseInt(ev.n4) === 1), ['حجم المعاملات (ح4)'], [e => e.n4], true);
        appendDetailedSheet("حجم المعاملات 2 متوسط", evaluations.filter(ev => parseInt(ev.n4) === 2), ['حجم المعاملات (ح4)'], [e => e.n4], true);
        appendDetailedSheet("حجم المعاملات 3 كبير", evaluations.filter(ev => parseInt(ev.n4) === 3), ['حجم المعاملات (ح4)'], [e => e.n4], true);

        // Sheets: مجموع الأداء
        appendDetailedSheet("الأداء 0 إلى 5", evaluations.filter(ev => ev.perfScore >= 0 && ev.perfScore <= 5), [], [], true);
        appendDetailedSheet("الأداء 6 إلى 10", evaluations.filter(ev => ev.perfScore >= 6 && ev.perfScore <= 10), [], [], true);
        appendDetailedSheet("الأداء 11 فما فوق", evaluations.filter(ev => ev.perfScore >= 11), [], [], true);

        // Sheets: مجموع الحاجة
        appendDetailedSheet("مجموع الحاجة أقل من 10", evaluations.filter(ev => ev.needScore < 10), [], [], true);
        appendDetailedSheet("مجموع الحاجة من 10-19", evaluations.filter(ev => ev.needScore >= 10 && ev.needScore <= 19), [], [], true);
        appendDetailedSheet("مجموع الحاجة من 20-30", evaluations.filter(ev => ev.needScore >= 20 && ev.needScore <= 30), [], [], true);

        // Sheet: جميع التقييمات كاملة للفصل
        appendDetailedSheet("جميع التقييمات", evaluations, [], [], true);

        // --- Custom Complex Sheets ---
        // 1. كشف الموظفين بدون عمل
        appendDetailedSheet("موظفين بدون عمل", evaluations.filter(ev => {
            return parseInt(ev.p1) === 0 && parseInt(ev.n1) === 0 && (ev.perfScore >= 0 && ev.perfScore <= 5);
        }), [], [], true);

        // 2. كشف حجم عمل قليل
        appendDetailedSheet("حجم عمل قليل", evaluations.filter(ev => {
            const n1 = parseInt(ev.n1) || 0;
            const n2 = parseInt(ev.n2) || 0;
            return (n1 >= 1 && n1 <= 3) && (n2 <= 1);
        }), [], [], true);

        // 3. كشف حجم عمل كبير
        appendDetailedSheet("حجم عمل كبير", evaluations.filter(ev => {
            const n1 = parseInt(ev.n1) || 0;
            const n2 = parseInt(ev.n2) || 0;
            return n1 > 3 && n2 > 3;
        }), [], [], true);

        // Sheets: معايير الأداء (الاستمرارية - أ1)
        appendDetailedSheet("استمرارية (أ1) صفر", evaluations.filter(ev => parseInt(ev.p1) === 0), [], [], true);
        appendDetailedSheet("استمرارية (أ1) واحد", evaluations.filter(ev => parseInt(ev.p1) === 1), [], [], true);
        appendDetailedSheet("استمرارية (أ1) أكبرمن1", evaluations.filter(ev => parseInt(ev.p1) > 1), [], [], true);

        // Sheets: معايير الحاجة بند 1 (n1) الجديدة 0 / 1-3 / 4+
        appendDetailedSheet("حاجة بند1 سجل 0", evaluations.filter(ev => parseInt(ev.n1) === 0), [], [], true);
        appendDetailedSheet("حاجة بند1 سجل 1الى3", evaluations.filter(ev => parseInt(ev.n1) >= 1 && parseInt(ev.n1) <= 3), [], [], true);
        appendDetailedSheet("حاجة بند1 سجل 4فمافوق", evaluations.filter(ev => parseInt(ev.n1) >= 4), [], [], true);

        // كشوفات تفصيلية بكامل المعلومات لكل فئة من فئات ملاحظة 2
        const note2CategoriesExportList = [NOTES2_ABROAD, ...NOTES2_SPECIAL, NOTES2_NORMAL];
        note2CategoriesExportList.forEach(cat => {
            const catFiltered = evaluations.filter(ev => matchesNotes2Category(ev.notes2, cat));
            if (catFiltered.length > 0) {
                const safeName = ("ملاحظة2_ " + cat).replace(/[\/\*\[\]\:\?]/g, '-').substring(0, 31);
                appendDetailedSheet(safeName, catFiltered, [], [], true);
            }
        });

        if (asZip) {
            zip.generateAsync({ type: "blob" }).then(function (content) {
                const url = window.URL.createObjectURL(content);
                const a = document.createElement('a');
                a.href = url;
                a.download = "إحصائيات_التقييم_كشوفات_منفصلة.zip";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            });
        } else {
            XLSX.writeFile(wb, "إحصائيات_التقييم_الشاملة.xlsx");
        }
    };

    // Export all stats to single Excel
    document.getElementById('export-stats-btn')?.addEventListener('click', () => exportStats(false));
    
    // Export all stats to ZIP (multiple Excels)
    document.getElementById('export-zip-btn')?.addEventListener('click', () => exportStats(true));

    // --- Score Calculation Logic ---
    function calculateScores() {
        let perfScore = 0;
        let needScore = 0;
        let commScore = 0;
        
        perfRadios.forEach(radio => {
            if (radio.checked) perfScore += parseFloat(radio.value);
        });
        
        needRadios.forEach(radio => {
            if (radio.checked) needScore += parseFloat(radio.value);
        });

        commRadios.forEach(radio => {
            if (radio.checked) commScore += parseFloat(radio.value);
        });
        
        perfTotalEl.textContent = perfScore;
        needTotalEl.textContent = needScore;
        commTotalEl.textContent = commScore;
        grandTotalEl.textContent = perfScore + needScore;
        
        return { perfScore, needScore, commScore, totalScore: perfScore + needScore };
    }

    perfRadios.forEach(radio => radio.addEventListener('change', calculateScores));
    needRadios.forEach(radio => radio.addEventListener('change', calculateScores));
    commRadios.forEach(radio => radio.addEventListener('change', calculateScores));

    window.resetScores = function() {
        localStorage.removeItem('unsavedFormState');
        setTimeout(calculateScores, 10);
    };

    // --- Save Evaluation ---
    
    // Auto-save form progress to prevent losing data if user leaves the page before submitting
    form.addEventListener('input', () => {
        const formData = new FormData(form);
        const dataObj = {};
        formData.forEach((val, key) => {
            if (dataObj[key] !== undefined) {
                if (!Array.isArray(dataObj[key])) {
                    dataObj[key] = [dataObj[key]];
                }
                dataObj[key].push(val);
            } else {
                dataObj[key] = val;
            }
        });
        localStorage.setItem('unsavedFormState', JSON.stringify(dataObj));
    });

    // Restore form progress on page load
    const restoreFormState = () => {
        const storedState = localStorage.getItem('unsavedFormState');
        if (storedState) {
            const dataObj = JSON.parse(storedState);
            Object.keys(dataObj).forEach(key => {
                const vals = Array.isArray(dataObj[key]) ? dataObj[key] : [dataObj[key]];
                const inputs = form.querySelectorAll(`[name="${key}"]`);
                if (inputs.length === 0) return;
                
                inputs.forEach(input => {
                    if (input.type === 'radio' || input.type === 'checkbox') {
                        input.checked = vals.includes(input.value);
                    } else if (input.multiple) {
                        Array.from(input.options).forEach(opt => {
                            opt.selected = vals.includes(opt.value);
                        });
                    } else {
                        input.value = vals[0] || '';
                    }
                });
            });
            setTimeout(calculateScores, 50); // Recalculate sums
        }
    };
    
    // Restore right after DOM load logic
    restoreFormState();

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        // Get Basic Info
        const empIdRaw = document.getElementById('emp-id').value;
        const empId = normalizeEmpId(empIdRaw) || empIdRaw.trim();
        const empName = document.getElementById('emp-name').value;
        const empTitle = document.getElementById('emp-title').value;
        const empDept = document.getElementById('emp-dept').value;
        const empSection = document.getElementById('emp-section').value;
        const empGender = document.getElementById('emp-gender') ? document.getElementById('emp-gender').value : 'ذكر';
        const empDob = document.getElementById('emp-dob').value;
        
        // Get Scores
        const scores = calculateScores();
        
        // Get individual answers for the table
        const getVal = (name) => {
            const el = document.querySelector(`input[name="${name}"]:checked`);
            return el ? el.value : '-';
        };

        const newEval = {
            id: empId,
            name: empName,
            title: empTitle,
            dept: empDept,
            section: empSection,
            gender: empGender || 'ذكر',
            dob: empDob,
            p1: getVal('p1'), p2: getVal('p2'), p3: getVal('p3'), p4: getVal('p4'), p5: getVal('p5'),
            perfScore: scores.perfScore,
            n1: getVal('n1'), n2: getVal('n2'), n3: getVal('n3'), n4: getVal('n4'), n5: getVal('n5'), n6: getVal('n6'), n7: getVal('n7'),
            needScore: scores.needScore,
            c1: getVal('c1'),
            commScore: scores.commScore,
            totalScore: scores.totalScore,
            notes: document.getElementById('emp-notes').value,
            notes2: Array.from(document.getElementById('emp-notes2').selectedOptions).map(o => o.value).join('، ') || 'عادي',
            timestamp: new Date().toISOString()
        };

        // Check if exists, replace or add
        const existingIndex = findEvalByEmpId(empId);
        if(existingIndex >= 0) {
            evaluations[existingIndex] = newEval;
            alert('تم تحديث بيانات الموظف بنجاح!');
        } else {
            evaluations.push(newEval);
            alert('تمت إضافة التقييم بنجاح!');
        }

        saveToLocal();
        updateSavedCount();
        refreshMissingEmpsView();
        form.reset();
        localStorage.removeItem('unsavedFormState');
        calculateScores();
    });

    // --- Search & Column Management State ---
    const defaultCols = [
        { key: "id", label: "الرقم الوظيفي", visible: true },
        { key: "name", label: "الاسم", visible: true, locked: true },
        { key: "p1", label: "أ1", visible: true, group: "perf" },
        { key: "p2", label: "أ2", visible: true, group: "perf" },
        { key: "p3", label: "أ3", visible: true, group: "perf" },
        { key: "p4", label: "أ4", visible: true, group: "perf" },
        { key: "p5", label: "أ5", visible: true, group: "perf" },
        { key: "perfScore", label: "مجموع الأداء", visible: true },
        { key: "n1", label: "ح1", visible: true, group: "need" },
        { key: "n2", label: "ح2", visible: true, group: "need" },
        { key: "n3", label: "ح3", visible: true, group: "need" },
        { key: "n4", label: "ح4", visible: true, group: "need" },
        { key: "n5", label: "ح5", visible: true, group: "need" },
        { key: "n6", label: "ح6", visible: true, group: "need" },
        { key: "n7", label: "ح7", visible: true, group: "need" },
        { key: "needScore", label: "مجموع الحاجة", visible: true },
        { key: "commScore", label: "قرار اللجنة", visible: true },
        { key: "totalScore", label: "الكلي (45)", visible: true },
        { key: "notes", label: "ملاحظات", visible: true },
        { key: "notes2", label: "ملاحظة2", visible: true },
        { key: "actions", label: "إجراءات", visible: true, locked: true }
    ];
    let colsState = JSON.parse(localStorage.getItem('colsState')) || defaultCols;
    colsState = colsState.filter(c => !['dept', 'section', 'gender', 'dob'].includes(c.key));
    if (colsState.length !== defaultCols.length) {
        colsState = defaultCols; // Reset if schema changes
        localStorage.setItem('colsState', JSON.stringify(colsState));
    }

    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', () => renderTable());
    }

    const colMgrBtn = document.getElementById('col-mgr-btn');
    const colDropdown = document.getElementById('col-dropdown');
    const colDropdownList = document.getElementById('col-dropdown-list');

    if (colMgrBtn && colDropdown) {
        colMgrBtn.addEventListener('click', () => {
            colDropdown.style.display = colDropdown.style.display === 'none' ? 'flex' : 'none';
        });
    }

    document.addEventListener('click', (e) => {
        if (colMgrBtn && colDropdown && !colMgrBtn.contains(e.target) && !colDropdown.contains(e.target)) {
            colDropdown.style.display = 'none';
        }
        
        // Also close emp column manager
        const empColMgrBtn = document.getElementById('emp-col-mgr-btn');
        const empColDropdown = document.getElementById('emp-col-dropdown');
        if (empColMgrBtn && empColDropdown && !empColMgrBtn.contains(e.target) && !empColDropdown.contains(e.target)) {
            empColDropdown.style.display = 'none';
        }
    });

    function renderColumnManager() {
        if (!colDropdownList) return;
        colDropdownList.innerHTML = '';
        colsState.forEach((col, index) => {
            if (col.locked) return; // Don't show locked columns in manager
            
            const div = document.createElement('div');
            div.className = 'col-item';
            div.innerHTML = `
                <input type="checkbox" id="col-chk-${index}" ${col.visible ? 'checked' : ''}>
                <label for="col-chk-${index}">${col.label}</label>
            `;
            const chk = div.querySelector('input');
            chk.addEventListener('change', () => {
                colsState[index].visible = chk.checked;
                localStorage.setItem('colsState', JSON.stringify(colsState));
                renderTable();
            });
            colDropdownList.appendChild(div);
        });
    }
    renderColumnManager();

    // --- Advanced Filter Logic ---
    const advFilterBtn = document.getElementById('adv-filter-btn');
    const advFilterModal = document.getElementById('adv-filter-modal');
    const closeFilterBtn = document.getElementById('close-filter-btn');
    const addFilterRuleBtn = document.getElementById('add-filter-rule-btn');
    const filterRulesContainer = document.getElementById('filter-rules-container');
    const applyFiltersBtn = document.getElementById('apply-filters-btn');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    const unevaluatedFilterChk = document.getElementById('unevaluated-filter-chk');
    
    let activeFilters = []; // Array of { key, min, max, type, value }
    let showUnevaluatedOnly = false;

    unevaluatedFilterChk.addEventListener('change', (e) => {
        showUnevaluatedOnly = e.target.checked;
    });

    const filterableFields = [
        { key: "name", label: "اسم الموظف", type: 'text' },
        { key: "notes", label: "الملاحظات", type: 'text' },
        { key: "notes2", label: "ملاحظة2", type: 'multiselect', options: ['في الخارج', 'أكاديمي', 'حالة خاصة', 'مسجون/ فقيد', 'اقترب تقاعده', 'عادي'] },
        { key: "dept", label: "الدائرة", type: 'text' },
        { key: "p1", label: "أ1 - الاستمرارية", type: 'number' },
        { key: "p2", label: "أ2 - الاستجابة", type: 'number' },
        { key: "p3", label: "أ3 - شكاوى", type: 'number' },
        { key: "p4", label: "أ4 - خدمات أخرى", type: 'number' },
        { key: "p5", label: "أ5 - تغيير المكان", type: 'number' },
        { key: "perfScore", label: "مجموع الأداء (15)", type: 'number' },
        { key: "n1", label: "ح1 - درجة الحاجة", type: 'number' },
        { key: "n2", label: "ح2 - الوجاهي", type: 'number' },
        { key: "n3", label: "ح3 - إلكترونيا", type: 'number' },
        { key: "n4", label: "ح4 - حجم المعاملات", type: 'number' },
        { key: "n5", label: "ح5 - مهمة فريدة", type: 'number' },
        { key: "n6", label: "ح6 - المؤهلات", type: 'number' },
        { key: "n7", label: "ح7 - موسمي", type: 'number' },
        { key: "needScore", label: "مجموع الحاجة (30)", type: 'number' },
        { key: "commScore", label: "قرار اللجنة (4)", type: 'number' },
        { key: "totalScore", label: "المجموع الكلي (45)", type: 'number' }
    ];

    function openFilterModal() {
        advFilterModal.style.display = 'flex';
        if (filterRulesContainer.children.length === 0 && activeFilters.length === 0) {
            addFilterRule();
        }
    }

    function closeFilterModal() {
        advFilterModal.style.display = 'none';
    }

    advFilterBtn.addEventListener('click', openFilterModal);
    closeFilterBtn.addEventListener('click', closeFilterModal);
    
    function addFilterRule() {
        const row = document.createElement('div');
        row.className = 'filter-rule-row';
        
        let optionsHtml = filterableFields.map(f => `<option value="${f.key}" data-type="${f.type}">${f.label}</option>`).join('');
        
        row.innerHTML = `
            <select class="filter-key">${optionsHtml}</select>
            <div class="filter-inputs" style="display: flex; gap: 0.5rem; flex: 1;">
                <input type="number" class="filter-min" placeholder="من (Min)" step="0.5">
                <input type="number" class="filter-max" placeholder="إلى (Max)" step="0.5">
            </div>
            <button type="button" class="remove-rule-btn" title="حذف الشرط">&times;</button>
        `;
        
        const keySelect = row.querySelector('.filter-key');
        const inputsContainer = row.querySelector('.filter-inputs');
        
        keySelect.addEventListener('change', () => {
            const selectedOpt = keySelect.options[keySelect.selectedIndex];
            if (selectedOpt.dataset.type === 'text') {
                inputsContainer.innerHTML = `<input type="text" class="filter-text" placeholder="يحتوي على النص..." style="flex: 1; width: 100%;">`;
            } else if (selectedOpt.dataset.type === 'multiselect') {
                inputsContainer.innerHTML = `
                    <select class="filter-multiselect" multiple style="flex: 1; width: 100%; height: 90px; padding: 0.5rem; font-family: inherit;">
                        <option value="عادي">عادي</option>
                        <option value="في الخارج">في الخارج</option>
                        <option value="أكاديمي">أكاديمي</option>
                        <option value="حالة خاصة">حالة خاصة</option>
                        <option value="مسجون/ فقيد">مسجون/ فقيد</option>
                        <option value="اقترب تقاعده">اقترب تقاعده</option>
                    </select>
                `;
            } else {
                inputsContainer.innerHTML = `
                    <input type="number" class="filter-min" placeholder="من (Min)" step="0.5" style="flex: 1; width: 100%;">
                    <input type="number" class="filter-max" placeholder="إلى (Max)" step="0.5" style="flex: 1; width: 100%;">
                `;
            }
        });
        
        // Trigger change to set correct initial inputs
        keySelect.dispatchEvent(new Event('change'));

        row.querySelector('.remove-rule-btn').addEventListener('click', () => {
            row.remove();
        });
        
        filterRulesContainer.appendChild(row);
    }
    
    addFilterRuleBtn.addEventListener('click', addFilterRule);

    applyFiltersBtn.addEventListener('click', () => {
        activeFilters = [];
        const rows = filterRulesContainer.querySelectorAll('.filter-rule-row');
        rows.forEach(row => {
            const keySelect = row.querySelector('.filter-key');
            const key = keySelect.value;
            const type = keySelect.options[keySelect.selectedIndex].dataset.type;
            
            if (type === 'text') {
                const textInput = row.querySelector('.filter-text');
                const textVal = textInput ? textInput.value.trim() : '';
                if (textVal !== '') {
                    activeFilters.push({
                        key: key,
                        type: 'text',
                        value: textVal
                    });
                }
            } else if (type === 'multiselect') {
                const selectElement = row.querySelector('.filter-multiselect');
                if (selectElement) {
                    const selectedVals = Array.from(selectElement.selectedOptions).map(o => o.value);
                    if (selectedVals.length > 0) {
                        activeFilters.push({
                            key: key,
                            type: 'multiselect',
                            values: selectedVals
                        });
                    }
                }
            } else {
                const minInput = row.querySelector('.filter-min');
                const maxInput = row.querySelector('.filter-max');
                const minStr = minInput ? minInput.value : '';
                const maxStr = maxInput ? maxInput.value : '';
                
                if (minStr !== '' || maxStr !== '') {
                    activeFilters.push({
                        key: key,
                        type: 'number',
                        min: minStr !== '' ? parseFloat(minStr) : -Infinity,
                        max: maxStr !== '' ? parseFloat(maxStr) : Infinity
                    });
                }
            }
        });
        
        // Highlight button if filters active
        let rulesCount = activeFilters.length + (showUnevaluatedOnly ? 1 : 0);
        if (rulesCount > 0) {
            advFilterBtn.classList.replace('btn-primary', 'btn-success');
            advFilterBtn.textContent = `تصفية متقدمة (${rulesCount})`;
        } else {
            advFilterBtn.classList.replace('btn-success', 'btn-primary');
            advFilterBtn.textContent = `تصفية متقدمة`;
        }
        
        closeFilterModal();
        renderTable();
    });

    clearFiltersBtn.addEventListener('click', () => {
        filterRulesContainer.innerHTML = '';
        unevaluatedFilterChk.checked = false;
        showUnevaluatedOnly = false;
        activeFilters = [];
        advFilterBtn.classList.replace('btn-success', 'btn-primary');
        advFilterBtn.textContent = `تصفية متقدمة`;
        closeFilterModal();
        renderTable();
    });

    // --- Synchronized Scrolling ---
    const topScrollContainer = document.getElementById('top-scroll-container');
    const topScrollContent = document.getElementById('top-scroll-content');
    
    tableContainer.addEventListener('scroll', () => {
        topScrollContainer.scrollLeft = tableContainer.scrollLeft;
    });
    topScrollContainer.addEventListener('scroll', () => {
        tableContainer.scrollLeft = topScrollContainer.scrollLeft;
    });

    // --- Table Rendering ---
    let currentSort = { key: 'timestamp', asc: false };

    // Helper to determine row color based on notes
    function getNoteColorInfo(notes, notes2) {
        const tags = parseNotes2Tags(notes2);
        const hasAcademic = tags.includes('أكاديمي');
        const hasAbroad = isNotes2Abroad(notes2);
        const hasSpecial = hasNotes2SpecialCategory(notes2);

        if (hasAcademic) {
            return { rank: 0, className: 'note-bg-green' };
        }
        if (tags.includes('حالة خاصة')) {
            return { rank: 1, className: 'note-bg-yellow' };
        }
        if (tags.includes('مسجون/ فقيد')) {
            return { rank: 2, className: 'note-bg-red' };
        }
        if (tags.includes(NOTES2_SPECIAL[3])) {
            return { rank: 3, className: 'note-bg-blue' };
        }
        // في الخارج فقط (بدون حالة خاصة) → يُعامل كـ «عادي» بدون تمييز لوني
        if (hasAbroad && !hasSpecial) {
            return { rank: 6, className: '' };
        }

        if (!notes || notes.trim() === '') return { rank: 6, className: '' };
        const lower = notes.toLowerCase();
        if (lower.includes('ليس لديه') || lower.includes('ليس لديها') || lower.includes('ليس عندها') || lower.includes('لا يوجد') || lower.includes('يمكن الاستغناء عنه')) {
            return { rank: 1, className: 'note-bg-red' }; // Red for no tasks / can be dispensed
        }
        if (lower.includes('تقاعد')) {
            return { rank: 2, className: 'note-bg-blue' }; // Blue for retirement
        }
        if (lower.includes('منتدب') || lower.includes('متطوع')) {
            return { rank: 3, className: 'note-bg-green' }; // Green for montadab / volunteer
        }
        if (lower.includes('إجازة') || lower.includes('مرضية') || lower.includes('اجازه')) {
            return { rank: 4, className: 'note-bg-yellow' }; // Yellow for leave
        }
        // General notes
        return { rank: 5, className: 'note-bg-gray' };
    }

    // Helper for Excel dates
    function formatDob(val) {
        if (!val) return '';
        const num = Number(val);
        // If it's a number resembling an Excel serial date (roughly 1927 to 2146)
        if (!isNaN(num) && num > 10000 && num < 90000 && !val.toString().includes('-') && !val.toString().includes('/')) {
            const date = new Date(Math.round((num - 25569) * 86400 * 1000));
            return date.toISOString().split('T')[0];
        }
        return val;
    }

    function renderTable() {
        // Build Thead dynamically based on visible columns
        const thead = document.querySelector('#evaluations-table thead');
        thead.innerHTML = '';
        tableBody.innerHTML = '';

        const perfCols = colsState.filter(c => c.group === 'perf' && c.visible);
        const needCols = colsState.filter(c => c.group === 'need' && c.visible);
        const hasGroups = perfCols.length > 0 || needCols.length > 0;

        let tr1 = document.createElement('tr');
        let tr2 = document.createElement('tr');

        colsState.forEach(col => {
            if (!col.visible) return;

            const setupSortableTh = (th, key) => {
                th.classList.add('sortable-th');
                th.setAttribute('data-sort', key);
                th.innerHTML = `${col.label} <span class="sort-icon"></span>`;
                if (currentSort.key === key) {
                    th.classList.add(currentSort.asc ? 'sort-asc' : 'sort-desc');
                }
                th.addEventListener('click', () => {
                    if (currentSort.key === key) {
                        currentSort.asc = !currentSort.asc;
                    } else {
                        currentSort.key = key;
                        currentSort.asc = false; // default descent for new sort
                    }
                    renderTable();
                });
            };

            if (col.group === 'perf') {
                if (perfCols[0].key === col.key) {
                    const th = document.createElement('th');
                    th.colSpan = perfCols.length;
                    th.textContent = "معايير الأداء (15)";
                    tr1.appendChild(th);
                }
                const th2 = document.createElement('th');
                setupSortableTh(th2, col.key);
                tr2.appendChild(th2);
            } else if (col.group === 'need') {
                if (needCols[0].key === col.key) {
                    const th = document.createElement('th');
                    th.colSpan = needCols.length;
                    th.textContent = "معايير الحاجة (30)";
                    tr1.appendChild(th);
                }
                const th2 = document.createElement('th');
                setupSortableTh(th2, col.key);
                tr2.appendChild(th2);
            } else {
                const th = document.createElement('th');
                th.rowSpan = hasGroups ? 2 : 1;
                
                if (col.key !== 'actions') {
                    setupSortableTh(th, col.key);
                } else {
                    th.textContent = col.label;
                }
                
                if (col.key === 'totalScore') th.classList.add('highlight-th');
                tr1.appendChild(th);
            }
        });

        thead.appendChild(tr1);
        if (hasGroups) thead.appendChild(tr2);

        // Helper for Arabic search normalization
        function normalizeText(text) {
            if (!text) return '';
            return text.toLowerCase()
                .replace(/[أإآ]/g, 'ا')
                .replace(/ة/g, 'ه')
                .replace(/ي/g, 'ى')
                .replace(/[\u064B-\u065F]/g, ''); // Remove diacritics
        }

        const query = normalizeText((searchInput && searchInput.value) ? searchInput.value.trim() : '');
        const searchTerms = query.split(/\s+/).filter(t => t.length > 0);
        
        // Sync data from employeeData
        if (employeeData && employeeData.length > 0) {
            evaluations.forEach(ev => {
                const emp = findEmployeeByEmpId(ev.id);
                if (emp) {
                    if (!ev.dob && emp['تاريخ الميلاد']) ev.dob = emp['تاريخ الميلاد'];
                    if (!ev.gender && emp['الجنس']) ev.gender = emp['الجنس'];
                    // Update dept and section to match employee data
                    if (emp['الدائرة']) ev.dept = emp['الدائرة'];
                    if (emp['القسم']) ev.section = String(emp['القسم']).replace(/\d+/g, '').trim();
                }
                // Strip numbers from any existing sections
                if (ev.section) {
                    ev.section = String(ev.section).replace(/\d+/g, '').trim();
                }
            });
            saveToLocal(); // persist synced data
        } else {
            // Strip numbers even if employeeData doesn't exist
            let changed = false;
            evaluations.forEach(ev => {
                if (ev.section && /\d/.test(ev.section)) {
                    ev.section = String(ev.section).replace(/\d+/g, '').trim();
                    changed = true;
                }
            });
            if (changed) saveToLocal();
        }

        const filteredEvals = evaluations.filter(ev => {
            // 1. Search Logic
            if (query.length > 0) {
                const name = normalizeText(ev.name || '');
                const notes = normalizeText(ev.notes || '');
                const notes2 = normalizeText(ev.notes2 || '');
                const dept = normalizeText(ev.dept || '');
                const combinedText = `${name} | ${notes} | ${notes2} | ${dept}`;
                
                // Allow matching the exact phrase or matching all individual words
                if (!combinedText.includes(query) && !searchTerms.every(term => combinedText.includes(term))) {
                    return false;
                }
            }
            
            // 1.5 Unevaluated Logic
            if (showUnevaluatedOnly) {
                const ts = parseFloat(ev.totalScore) || 0;
                const notes = (ev.notes || '').trim();
                // Not evaluated if score is 0 and notes are empty
                if (ts > 0 || notes !== '') {
                    return false;
                }
            }

            // 2. Advanced Filters Logic
            for (let filter of activeFilters) {
                if (filter.type === 'text') {
                    const textVal = normalizeText(ev[filter.key] || '');
                    const searchTerm = normalizeText(filter.value);
                    if (!textVal.includes(searchTerm)) return false;
                } else if (filter.type === 'multiselect') {
                    const hasMatch = filter.values.some(v => matchesNotes2Category(ev[filter.key], v));
                    if (!hasMatch) return false;
                } else {
                    let val = parseFloat(ev[filter.key]) || 0;
                    if (val < filter.min || val > filter.max) return false;
                }
            }
            
            return true;
        });
        
        // Save to state for exporting
        currentFilteredData = filteredEvals;
        
        // 3. Sorting Logic
        filteredEvals.sort((a, b) => {
            // First sort by note color if sorting by notes column
            if (currentSort.key === 'notes' || currentSort.key === 'notes2') {
                const rankA = getNoteColorInfo(a.notes, a.notes2).rank;
                const rankB = getNoteColorInfo(b.notes, b.notes2).rank;
                if (rankA !== rankB) {
                    return currentSort.asc ? rankA - rankB : rankB - rankA;
                }
            }

            let valA = a[currentSort.key];
            let valB = b[currentSort.key];
            
            // Handle undefined
            if (valA === undefined) valA = '';
            if (valB === undefined) valB = '';
            
            // Try numeric sort first if applicable
            let numA = parseFloat(valA);
            let numB = parseFloat(valB);
            
            if (!isNaN(numA) && !isNaN(numB)) {
                return currentSort.asc ? numA - numB : numB - numA;
            }
            
            // String sort
            valA = String(valA).toLowerCase();
            valB = String(valB).toLowerCase();
            if (valA < valB) return currentSort.asc ? -1 : 1;
            if (valA > valB) return currentSort.asc ? 1 : -1;
            return 0;
        });
        
        if (filteredEvals.length === 0) {
            emptyState.style.display = 'block';
            emptyState.innerHTML = searchTerms.length > 0 ? '<p>لا توجد نتائج مطابقة للبحث.</p>' : '<p>لا توجد بيانات محفوظة بعد. قم بإضافة التقييمات من "عرض النموذج" أو استيراد ملف إكسل.</p>';
            tableContainer.style.display = 'none';
            topScrollContainer.style.display = 'none';
            return;
        }
        
        emptyState.style.display = 'none';
        tableContainer.style.display = 'block';
        topScrollContainer.style.display = 'block';

        filteredEvals.forEach((ev) => {
            // Find original index for edit/delete functions
            const originalIndex = findEvalByEmpId(ev.id);
            
            const tr = document.createElement('tr');
            
            // Highlight row based on notes
            const colorInfo = getNoteColorInfo(ev.notes, ev.notes2);
            if (colorInfo.className) {
                tr.classList.add(colorInfo.className);
            }
            
            colsState.forEach(col => {
                if (!col.visible) return;
                const td = document.createElement('td');
                
                if (col.key === 'actions') {
                    td.innerHTML = `
                        <button class="action-btn edit-row-btn" onclick="editRow(${originalIndex})">تعديل</button>
                        <button class="action-btn del-row-btn" onclick="deleteRow(${originalIndex})">حذف</button>
                    `;
                } else if (col.key === 'dob') {
                    td.innerHTML = formatDob(ev[col.key] || '');
                } else if (col.key === 'id') {
                    td.innerHTML = `<strong>${ev[col.key] !== undefined ? ev[col.key] : '-'}</strong>`;
                } else if (col.key === 'name') {
                    const empName = ev.name || '-';
                    const dept = ev.dept || '-';
                    const section = ev.section || '-';
                    const gender = ev.gender || '-';
                    const dob = formatDob(ev.dob || '-');
                    td.innerHTML = `<div style="font-weight: bold; margin-bottom: 3px;">${empName}</div>
                                    <div style="font-size: 0.85rem; color: #666;">${dept} - ${section} - ${gender} - ${dob}</div>`;
                } else if (col.key === 'perfScore' || col.key === 'needScore' || col.key === 'commScore') {
                    td.innerHTML = `<strong>${ev[col.key] !== undefined ? ev[col.key] : '-'}</strong>`;
                    td.innerHTML = `<span style="font-weight:bold; color:var(--primary-color);">${ev[col.key]}</span>`;
                } else {
                    td.textContent = ev[col.key] !== undefined ? ev[col.key] : '-';
                    if (col.key === 'dob') td.dir = 'ltr';
                }
                tr.appendChild(td);
            });
            
            tableBody.appendChild(tr);
        });

        // Update filtered count
        const countContainer = document.getElementById('record-count-container');
        if (countContainer) {
            const countSpan = document.getElementById('filtered-count');
            if (filteredEvals.length > 0 || evaluations.length > 0) {
                countContainer.style.display = 'block';
                countSpan.textContent = filteredEvals.length;
            } else {
                countContainer.style.display = 'none';
            }
        }

        updateSavedCount();

        if (topScrollContent) {
            setTimeout(() => {
                const table = document.querySelector('.data-table');
                if (table) topScrollContent.style.width = table.offsetWidth + 'px';
            }, 50);
        }
    }

    window.editRow = function(index) {
        const ev = evaluations[index];
        
        // Switch to form tab
        document.querySelector('.tab-btn[data-target="form-view"]').click();
        
        // Fill basic info
        document.getElementById('emp-id').value = ev.id || '';
        document.getElementById('emp-name').value = ev.name || '';
        document.getElementById('emp-title').value = ev.title || '';
        document.getElementById('emp-dept').value = ev.dept || '';
        document.getElementById('emp-section').value = ev.section || '';
        if (document.getElementById('emp-gender')) document.getElementById('emp-gender').value = ev.gender || 'ذكر';
        document.getElementById('emp-dob').value = ev.dob || '';
        document.getElementById('emp-notes').value = ev.notes || '';
        
        // Multi-select for notes2
        const notes2Select = document.getElementById('emp-notes2');
        const vals = (ev.notes2 || 'عادي').split('، ').map(s => s.trim());
        
        vals.forEach(val => {
            if (val && val !== 'عادي' && !Array.from(notes2Select.options).find(o => o.value === val)) {
                const newOpt = document.createElement('option');
                newOpt.value = val;
                newOpt.textContent = val;
                notes2Select.appendChild(newOpt);
            }
        });

        Array.from(notes2Select.options).forEach(opt => {
            opt.selected = vals.includes(opt.value);
        });
        
        // Fill radios
        const setRadio = (name, val) => {
            const r = document.querySelector(`input[name="${name}"][value="${val}"]`);
            if (r) r.checked = true;
        };
        
        setRadio('p1', ev.p1); setRadio('p2', ev.p2); setRadio('p3', ev.p3); setRadio('p4', ev.p4); setRadio('p5', ev.p5);
        setRadio('n1', ev.n1); setRadio('n2', ev.n2); setRadio('n3', ev.n3); setRadio('n4', ev.n4); setRadio('n5', ev.n5); setRadio('n6', ev.n6); setRadio('n7', ev.n7);
        setRadio('c1', ev.c1 !== undefined ? ev.c1 : ev.n8);
        
        calculateScores();
    };

    window.deleteRow = function(index) {
        if(confirm('هل أنت متأكد من حذف هذا التقييم؟')) {
            evaluations.splice(index, 1);
            saveToLocal();
            renderTable();
            updateSavedCount();
        }
    };

    function saveToLocal() {
        localStorage.setItem('evaluations', JSON.stringify(evaluations));
        touchDataTimestamp();
        pushToGitHub();
    }

    // --- Delete All ---
    const deleteAllBtn = document.getElementById('delete-all-btn');
    if (deleteAllBtn) deleteAllBtn.addEventListener('click', () => {
        if(evaluations.length === 0) {
            alert('لا توجد تقييمات لحذفها.');
            return;
        }
        if(confirm('هل أنت متأكد تماماً من حذف جميع التقييمات المحفوظة؟ (هذا الإجراء لا يمكن التراجع عنه)')) {
            evaluations = [];
            saveToLocal();
            renderTable();
            updateSavedCount();
        }
    });

    // --- Export Template (Excel) ---
    const exportTemplateBtn = document.getElementById('export-template-btn');
    if (exportTemplateBtn) exportTemplateBtn.addEventListener('click', () => {
        // Create an empty worksheet with headers
        const headers = [
            "الرقم الوظيفي", "اسم الموظف", "المسمى الوظيفي", "الدائرة", "القسم", "تاريخ الميلاد",
            "الاستمرارية (0-3)", "الاستجابة للتعليمات (0-3)", "شكاوى (0-3)", "خدمات إضافية (0-3)", "تغيير مكانه (0-3)",
            "درجة الحاجة (0-6)", "أيام الوجاهي (0-6)", "انجاز إلكترونيا (0-3)", "حجم المعاملات (0-3)", "مهمة فريدة (0-3)", "مؤهلاته للمهام (0-6)", "عمل موسمي (0-3)", "قرار اللجنة (0-4)", "الملاحظات", "ملاحظة2"
        ];
        
        const ws = XLSX.utils.aoa_to_sheet([headers]);
        
        // Apply Styles
        const borderStyle = {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        };

        for (let i in ws) {
            if (i[0] === '!') continue;
            let cell = ws[i];
            cell.s = {
                font: { name: "Simplified Arabic", sz: 12, bold: true },
                border: borderStyle,
                fill: { fgColor: { rgb: "EAEAEA" } },
                alignment: { horizontal: "center", vertical: "center", wrapText: true }
            };
        }

        // Set Right-to-Left direction for the sheet
        if (!ws['!views']) ws['!views'] = [];
        ws['!views'].push({ rightToLeft: true });
        
        // Add some style/width to columns
        const wscols = [
            {wch: 15}, {wch: 25}, {wch: 20}, {wch: 20}, {wch: 20}, {wch: 15},
            {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15},
            {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 30}, {wch: 30}
        ];
        ws['!cols'] = wscols;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Template");
        
        // Generate file and trigger download
        XLSX.writeFile(wb, "evaluation_template_empty.xlsx");
    });

    // --- Export Data (Excel) ---
    const exportDataBtn = document.getElementById('export-data-btn');
    if (exportDataBtn) exportDataBtn.addEventListener('click', () => {
        if (currentFilteredData.length === 0) {
            alert('لا توجد بيانات لتصديرها (الجدول فارغ).');
            return;
        }

        const headers = [
            "الرقم الوظيفي", "اسم الموظف", "المسمى الوظيفي", "الدائرة", "القسم", "الجنس", "تاريخ الميلاد", "العمر",
            "أ1: الاستمرارية (0-3)", "أ2: الاستجابة للتعليمات (0-3)", "أ3: شكاوى (0-3)", "أ4: خدمات إضافية (0-3)", "أ5: تغيير مكانه (0-3)",
            "مجموع الأداء",
            "ح1: درجة الحاجة (0-6)", "ح2: أيام الوجاهي (0-6)", "ح3: انجاز إلكترونيا (0-3)", "ح4: حجم المعاملات (0-3)", "ح5: مهمة فريدة (0-3)", "ح6: مؤهلاته للمهام (0-6)", "ح7: عمل موسمي (0-3)", 
            "مجموع الحاجة",
            "قرار اللجنة (0-4)",
            "المجموع الكلي", 
            "الملاحظات",
            "ملاحظة2"
        ];

        const rows = [headers];
        
        // Helper to calculate age from DOB string YYYY-MM-DD
        const calculateAge = (dobString) => {
            if (!dobString) return '';
            const dob = new Date(dobString);
            if (isNaN(dob.getTime())) return '';
            const diff_ms = Date.now() - dob.getTime();
            const age_dt = new Date(diff_ms); 
            return Math.abs(age_dt.getUTCFullYear() - 1970);
        };

        currentFilteredData.forEach(ev => {
            const age = calculateAge(formatDob(ev.dob));
            const row = [
                ev.id || '', ev.name || '', ev.title || '', ev.dept || '', ev.section || '', ev.gender || 'ذكر', formatDob(ev.dob), age,
                ev.p1 ?? '', ev.p2 ?? '', ev.p3 ?? '', ev.p4 ?? '', ev.p5 ?? '',
                ev.perfScore ?? '',
                ev.n1 ?? '', ev.n2 ?? '', ev.n3 ?? '', ev.n4 ?? '', ev.n5 ?? '', ev.n6 ?? '', ev.n7 ?? '',
                ev.needScore ?? '',
                ev.c1 ?? '',
                ev.totalScore ?? '',
                ev.notes ?? '',
                ev.notes2 ?? ''
            ];
            rows.push(row);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        
        // Apply Styles
        const borderStyle = {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        };

        for (let i in ws) {
            if (i[0] === '!') continue;
            let cell = ws[i];
            
            cell.s = {
                font: { name: "Simplified Arabic", sz: 12 },
                border: borderStyle,
                alignment: { horizontal: "center", vertical: "center", wrapText: true }
            };
            
            // Make headers bold with a light background
            const rowNumStr = i.replace(/\D/g, '');
            const colLetterStr = i.replace(/\d/g, ''); // Extract column letter (e.g., L, T, U)

            if (rowNumStr === "1") {
                cell.s.font.bold = true;
                // Header colors for specific columns, else default light gray
                if (colLetterStr === 'N') {
                    cell.s.fill = { fgColor: { rgb: "B6D7A8" } }; // Darker green for header
                } else if (colLetterStr === 'V') {
                    cell.s.fill = { fgColor: { rgb: "FFE599" } }; // Darker yellow for header
                } else if (colLetterStr === 'W') {
                    cell.s.fill = { fgColor: { rgb: "D5A6BD" } }; // Darker purple for header
                } else {
                    cell.s.fill = { fgColor: { rgb: "EAEAEA" } };
                }
            } else {
                // Determine base background color for the row
                let rowBgColor = null;
                const dataIndex = parseInt(rowNumStr) - 2;
                if (dataIndex >= 0 && dataIndex < currentFilteredData.length) {
                    const ev = currentFilteredData[dataIndex];
                    const colorInfo = getNoteColorInfo(ev.notes, ev.notes2);
                    if (colorInfo.className === 'note-bg-acad-abroad') {
                        rowBgColor = "D8BFD8"; // Distinct Light Purple
                    } else if (colorInfo.className === 'note-bg-abroad') {
                        rowBgColor = "FFD6A5"; // Distinct Light Orange
                    } else if (colorInfo.className === 'note-bg-red') {
                        rowBgColor = "FFB3B3"; // Distinct Light Red
                    } else if (colorInfo.className === 'note-bg-blue') {
                        rowBgColor = "A3D1FF"; // Distinct Light Blue
                    } else if (colorInfo.className === 'note-bg-green') {
                        rowBgColor = "A3FFB8"; // Distinct Light Green
                    } else if (colorInfo.className === 'note-bg-yellow') {
                        rowBgColor = "FFF2A3"; // Distinct Light Yellow
                    } else if (colorInfo.className === 'note-bg-gray') {
                        rowBgColor = "E2E8F0"; // Distinct Light Gray
                    }
                }

                // Check specific target columns first for their particular color overriding the row color
                if (colLetterStr === 'N') {
                    cell.s.fill = { fgColor: { rgb: "D9EAD3" } }; // Light Green for perfScore
                } else if (colLetterStr === 'V') {
                    cell.s.fill = { fgColor: { rgb: "FFF2CC" } }; // Light Yellow for needScore
                } else if (colLetterStr === 'W') {
                    cell.s.fill = { fgColor: { rgb: "EAD1DC" } }; // Light Purple for commScore
                } else if (rowBgColor) {
                    cell.s.fill = { fgColor: { rgb: rowBgColor } };
                }
            }
        }

        // Set Right-to-Left direction for the sheet
        if (!ws['!views']) ws['!views'] = [];
        ws['!views'].push({ rightToLeft: true });

        const wscols = [
            {wch: 15}, {wch: 25}, {wch: 20}, {wch: 20}, {wch: 20}, {wch: 10}, {wch: 15}, {wch: 10},
            {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12},
            {wch: 15},
            {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12},
            {wch: 15},
            {wch: 15},
            {wch: 15},
            {wch: 40},
            {wch: 40}
        ];
        ws['!cols'] = wscols;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Evaluations");
        
        const dateStr = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `Evaluations_Export_${dateStr}.xlsx`);
    });

    // --- Import Excel Data ---
    const excelUpload = document.getElementById('excel-upload');
    if (excelUpload) excelUpload.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Read as array of arrays
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});
                
                if (jsonData.length > 1) {
                    const header = jsonData[0] || [];
                    let importCount = 0;
                    
                    const notesIndex = header.findIndex(h => typeof h === 'string' && (h.includes('ملاحظات') || h.includes('Notes')));
                    const actualNotesIndex = notesIndex >= 0 ? notesIndex : 19;
                    const notes2Index = header.findIndex(h => typeof h === 'string' && (h.includes('ملاحظة2') || h.includes('Notes2')));
                    const actualNotes2Index = notes2Index >= 0 ? notes2Index : 20;
                    
                    // Loop over rows starting from 1 (skip header)
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0 || !row[0]) continue; // Skip empty
                        
                        // Parse values based on template order:
                        // 0:ID, 1:Name, 2:Title, 3:Dept, 4:Section, 5:DOB, 6-10:Perf, 11-18:Need
                        const empId = normalizeEmpId(String(row[0] || '')) || String(row[0] || `EMP-${Date.now()}-${i}`).trim();
                        
                        let n1 = String(row[actualNotesIndex] || '');
                        let n2 = String(row[actualNotes2Index] || '');
                        let combinedLower = (n1 + ' ' + n2).toLowerCase();

                        let n2Options = [];
                        if (combinedLower.includes('خارج')) n2Options.push('في الخارج');
                        if (combinedLower.includes('أكاديمي') || combinedLower.includes('اكاديمي')) n2Options.push('أكاديمي');
                        if (combinedLower.includes('خاصة') || combinedLower.includes('خاصه')) n2Options.push('حالة خاصة');
                        if (combinedLower.includes('سجن') || combinedLower.includes('مسجون') || combinedLower.includes('فقيد') || combinedLower.includes('احتلال')) n2Options.push('مسجون/ فقيد');
                        if (combinedLower.includes('تقاعد')) n2Options.push('اقترب تقاعده');
                        
                        let defaultNotes2;
                        if (n2Options.length > 0) {
                            defaultNotes2 = [...new Set(n2Options)].join('، ');
                        } else {
                            defaultNotes2 = n2 || 'عادي';
                        }

                        const newEval = {
                            id: empId,
                            name: String(row[1] || 'غير محدد'),
                            title: String(row[2] || ''),
                            dept: String(row[3] || ''),
                            section: String(row[4] || ''),
                            dob: String(row[5] || ''),
                            p1: parseFloat(row[6])||0, p2: parseFloat(row[7])||0, p3: parseFloat(row[8])||0, p4: parseFloat(row[9])||0, p5: parseFloat(row[10])||0,
                            n1: parseFloat(row[11])||0, n2: parseFloat(row[12])||0, n3: parseFloat(row[13])||0, n4: parseFloat(row[14])||0, n5: parseFloat(row[15])||0, n6: parseFloat(row[16])||0, n7: parseFloat(row[17])||0,
                            c1: parseFloat(row[18])||0,
                            notes: String(row[actualNotesIndex] || ''),
                            notes2: defaultNotes2,
                            timestamp: new Date().toISOString()
                        };
                        
                        newEval.perfScore = newEval.p1 + newEval.p2 + newEval.p3 + newEval.p4 + newEval.p5;
                        newEval.needScore = newEval.n1 + newEval.n2 + newEval.n3 + newEval.n4 + newEval.n5 + newEval.n6 + newEval.n7;
                        newEval.commScore = newEval.c1;
                        newEval.totalScore = newEval.perfScore + newEval.needScore;
                        
                        const existingIndex = findEvalByEmpId(empId);
                        if(existingIndex >= 0) {
                            evaluations[existingIndex] = newEval;
                        } else {
                            evaluations.push(newEval);
                        }
                        importCount++;
                    }
                    
                    saveToLocal();
                    pushToGitHub(true);
                    updateSavedCount();
                    
                    document.querySelector('.tab-btn[data-target="table-view"]').click();
                    
                    alert(`تم استيراد ${importCount} تقييم بنجاح!`);
                } else {
                    alert('الملف فارغ أو لا يحتوي على بيانات.');
                }
            } catch (err) {
                console.error(err);
                alert('حدث خطأ أثناء قراءة الملف. تأكد من أنه بصيغة قالب التقييم الصحيحة.');
            }
            e.target.value = '';
        };
        
        reader.readAsArrayBuffer(file);
    });

    // --- Employee Data Logic --- //
    // Default columns for employee data
    const defaultEmpCols = [
        "الرقم الوظيفي", "الاسم", "المسمى الوظيفي", "الدائرة", "القسم", 
        "الجنس", "نوع الوظيفة", "تاريخ الميلاد", "العمر", "رقم الجوال", "البريد الالكتروني", "نوع العقد"
    ].map(k => ({ key: k, visible: true }));

    let empColsState = JSON.parse(localStorage.getItem('empColsState')) || defaultEmpCols;

    const empColMgrBtn = document.getElementById('emp-col-mgr-btn');
    const empColDropdown = document.getElementById('emp-col-dropdown');
    const empColDropdownList = document.getElementById('emp-col-dropdown-list');

    if(empColMgrBtn) {
        empColMgrBtn.addEventListener('click', () => {
            empColDropdown.style.display = empColDropdown.style.display === 'none' ? 'flex' : 'none';
        });
    }

    function renderEmpColumnManager() {
        if(!empColDropdownList) return;
        empColDropdownList.innerHTML = '';
        empColsState.forEach((col, index) => {
            if (col.key === 'الرقم الوظيفي' || col.key === 'الاسم') return; // locked
            const div = document.createElement('div');
            div.className = 'col-item';
            div.innerHTML = `
                <input type="checkbox" id="emp-col-chk-${index}" ${col.visible ? 'checked' : ''}>
                <label for="emp-col-chk-${index}">${col.key}</label>
            `;
            const chk = div.querySelector('input');
            chk.addEventListener('change', () => {
                empColsState[index].visible = chk.checked;
                localStorage.setItem('empColsState', JSON.stringify(empColsState));
                renderEmpTable();
            });
            empColDropdownList.appendChild(div);
        });
    }
    renderEmpColumnManager();

    if (employeeData.length > 0) {
        renderEmpTable();
    }

    // --- Advanced Filter for Employee Data ---
    const empAdvFilterBtn = document.getElementById('emp-adv-filter-btn');
    const empAdvFilterModal = document.getElementById('emp-adv-filter-modal');
    const closeEmpFilterBtn = document.getElementById('close-emp-filter-btn');
    const addEmpFilterRuleBtn = document.getElementById('add-emp-filter-rule-btn');
    const empFilterRulesContainer = document.getElementById('emp-filter-rules-container');
    const applyEmpFiltersBtn = document.getElementById('apply-emp-filters-btn');
    const clearEmpFiltersBtn = document.getElementById('clear-emp-filters-btn');

    const empFilterableFields = [
        { key: "الاسم", label: "الاسم", type: 'text' },
        { key: "الرقم الوظيفي", label: "الرقم الوظيفي", type: 'text' },
        { key: "المسمى الوظيفي", label: "المسمى الوظيفي", type: 'text' },
        { key: "الدائرة", label: "الدائرة", type: 'text' },
        { key: "القسم", label: "القسم", type: 'text' },
        { key: "الجنس", label: "الجنس", type: 'text' },
        { key: "نوع الوظيفة", label: "نوع الوظيفة", type: 'text' },
        { key: "تاريخ الميلاد", label: "تاريخ الميلاد", type: 'text' },
        { key: "العمر", label: "العمر", type: 'number' },
        { key: "رقم الجوال", label: "رقم الجوال", type: 'text' },
        { key: "البريد الالكتروني", label: "البريد الالكتروني", type: 'text' },
        { key: "نوع العقد", label: "نوع العقد", type: 'text' }
    ];

    function openEmpFilterModal() {
        if(empAdvFilterModal) empAdvFilterModal.style.display = 'flex';
        if (empFilterRulesContainer && empFilterRulesContainer.children.length === 0 && activeEmpFilters.length === 0) {
            addEmpFilterRule();
        }
    }

    function closeEmpFilterModal() {
        if(empAdvFilterModal) empAdvFilterModal.style.display = 'none';
    }

    if(empAdvFilterBtn) empAdvFilterBtn.addEventListener('click', openEmpFilterModal);
    if(closeEmpFilterBtn) closeEmpFilterBtn.addEventListener('click', closeEmpFilterModal);
    
    function addEmpFilterRule() {
        if(!empFilterRulesContainer) return;
        const row = document.createElement('div');
        row.className = 'filter-rule-row';
        
        let optionsHtml = empFilterableFields.map(f => `<option value="${f.key}" data-type="${f.type}">${f.label}</option>`).join('');
        
        row.innerHTML = `
            <select class="filter-key">${optionsHtml}</select>
            <div class="filter-inputs" style="display: flex; gap: 0.5rem; flex: 1;">
            </div>
            <button type="button" class="remove-rule-btn" title="حذف الشرط">&times;</button>
        `;
        
        const keySelect = row.querySelector('.filter-key');
        const inputsContainer = row.querySelector('.filter-inputs');
        
        keySelect.addEventListener('change', () => {
            const selectedOpt = keySelect.options[keySelect.selectedIndex];
            if (selectedOpt.dataset.type === 'text') {
                inputsContainer.innerHTML = `<input type="text" class="filter-text" placeholder="يحتوي على النص..." style="flex: 1; width: 100%;">`;
            } else {
                inputsContainer.innerHTML = `
                    <input type="number" class="filter-min" placeholder="من (Min)" step="1" style="flex: 1; width: 100%;">
                    <input type="number" class="filter-max" placeholder="إلى (Max)" step="1" style="flex: 1; width: 100%;">
                `;
            }
        });
        
        keySelect.dispatchEvent(new Event('change'));

        row.querySelector('.remove-rule-btn').addEventListener('click', () => {
            row.remove();
        });
        
        empFilterRulesContainer.appendChild(row);
    }
    
    if(addEmpFilterRuleBtn) addEmpFilterRuleBtn.addEventListener('click', addEmpFilterRule);

    if(applyEmpFiltersBtn) {
        applyEmpFiltersBtn.addEventListener('click', () => {
            activeEmpFilters = [];
            const rows = empFilterRulesContainer.querySelectorAll('.filter-rule-row');
            rows.forEach(row => {
                const keySelect = row.querySelector('.filter-key');
                const key = keySelect.value;
                const type = keySelect.options[keySelect.selectedIndex].dataset.type;
                
                if (type === 'text') {
                    const textInput = row.querySelector('.filter-text');
                    const textVal = textInput ? textInput.value.trim() : '';
                    if (textVal !== '') {
                        activeEmpFilters.push({ key: key, type: 'text', value: textVal });
                    }
                } else {
                    const minInput = row.querySelector('.filter-min');
                    const maxInput = row.querySelector('.filter-max');
                    const minStr = minInput ? minInput.value : '';
                    const maxStr = maxInput ? maxInput.value : '';
                    
                    if (minStr !== '' || maxStr !== '') {
                        activeEmpFilters.push({
                            key: key, type: 'number',
                            min: minStr !== '' ? parseFloat(minStr) : -Infinity,
                            max: maxStr !== '' ? parseFloat(maxStr) : Infinity
                        });
                    }
                }
            });
            
            let rulesCount = activeEmpFilters.length;
            if (rulesCount > 0) {
                empAdvFilterBtn.classList.replace('btn-primary', 'btn-success');
                empAdvFilterBtn.textContent = `تصفية متقدمة (${rulesCount})`;
            } else {
                empAdvFilterBtn.classList.replace('btn-success', 'btn-primary');
                empAdvFilterBtn.textContent = `تصفية متقدمة`;
            }
            
            closeEmpFilterModal();
            renderEmpTable(); // Without arguments, it will use employeeData but filtered by activeEmpFilters inside the function
        });
    }

    if(clearEmpFiltersBtn) {
        clearEmpFiltersBtn.addEventListener('click', () => {
            if(empFilterRulesContainer) empFilterRulesContainer.innerHTML = '';
            activeEmpFilters = [];
            empAdvFilterBtn.classList.replace('btn-success', 'btn-primary');
            empAdvFilterBtn.textContent = `تصفية متقدمة`;
            closeEmpFilterModal();
            renderEmpTable();
        });
    }

    // Normalizing text helper
    function normalizeText(text) {
        if (!text && text !== 0) return '';
        return String(text).toLowerCase()
            .replace(/[أإآ]/g, 'ا')
            .replace(/ة/g, 'ه')
            .replace(/ي/g, 'ى')
            .replace(/[\u064B-\u065F]/g, '');
    }

    function renderEmpTable(dataToRender) {
        let baseData = dataToRender || employeeData;
        
        // Apply Advanced Filters to Employee Data
        let filteredData = baseData;
        if (activeEmpFilters.length > 0) {
            filteredData = baseData.filter(emp => {
                for (let filter of activeEmpFilters) {
                    if (filter.type === 'text') {
                        const textVal = normalizeText(emp[filter.key] || '');
                        const searchTerm = normalizeText(filter.value);
                        if (!textVal.includes(searchTerm)) return false;
                    } else {
                        let val = parseFloat(emp[filter.key]) || 0;
                        if (val < filter.min || val > filter.max) return false;
                    }
                }
                return true;
            });
        }
        
        currentFilteredEmpData = filteredData;
        const tbody = document.getElementById('emp-data-body');
        if(!tbody) return;
        tbody.innerHTML = '';
        
        const visibleKeys = empColsState.filter(c => c.visible).map(c => c.key);
        
        if(filteredData.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${visibleKeys.length + 1}" style="text-align: center; padding: 2rem;">لا توجد بيانات مطابقة</td></tr>`;
            return;
        }

        // Rebuild Thead based on visible columns
        const thead = document.querySelector('#emp-data-table thead tr');
        if (thead) {
            thead.innerHTML = '';
            visibleKeys.forEach(key => {
                const th = document.createElement('th');
                th.textContent = key;
                thead.appendChild(th);
            });
            const actionTh = document.createElement('th');
            actionTh.textContent = 'إجراءات';
            thead.appendChild(actionTh);
        }
        
        filteredData.forEach((emp) => {
            const ind = employeeData.indexOf(emp); // Real index for editing
            const tr = document.createElement('tr');
            visibleKeys.forEach(key => {
                const td = document.createElement('td');
                if (key === 'الجنس' && !emp[key]) {
                    emp[key] = 'ذكر';
                }
                td.textContent = emp[key] || '';
                tr.appendChild(td);
            });
            const actionTd = document.createElement('td');
            actionTd.innerHTML = `
                <button class="action-btn rate-row-btn" onclick="rateEmpFromData(${ind})">تقييم</button>
                <button class="action-btn edit-row-btn" onclick="editEmpRow(${ind})">تعديل</button>
            `;
            tr.appendChild(actionTd);
            tbody.appendChild(tr);
        });

        // Update record count
        const countContainer = document.getElementById('emp-filtered-count');
        if(countContainer) countContainer.textContent = filteredData.length;
    }

    // Modal editing for Employee Data
    const editEmpModal = document.getElementById('edit-emp-modal');
    const closeEditEmpBtn = document.getElementById('close-edit-emp-btn');
    const cancelEmpEditBtn = document.getElementById('cancel-emp-edit-btn');
    const saveEmpDataBtn = document.getElementById('save-emp-data-btn');
    const editEmpFields = document.getElementById('edit-emp-fields');

    const closeEmpModal = () => editEmpModal.style.display = 'none';

    if(closeEditEmpBtn) closeEditEmpBtn.addEventListener('click', closeEmpModal);
    if(cancelEmpEditBtn) cancelEmpEditBtn.addEventListener('click', closeEmpModal);

    window.rateEmpFromData = function(index) {
        if (!employeeData[index]) return;
        const emp = employeeData[index];
        const evalIdx = findEvalByEmpId(emp['الرقم الوظيفي']);
        if (evalIdx >= 0) {
            window.editRow(evalIdx);
            return;
        }
        document.querySelector('.tab-btn[data-target="form-view"]').click();
        document.getElementById('emp-id').value = normalizeEmpId(emp['الرقم الوظيفي']) || emp['الرقم الوظيفي'] || '';
        document.getElementById('emp-name').value = emp['الاسم'] || '';
        document.getElementById('emp-title').value = emp['المسمى الوظيفي'] || '';
        document.getElementById('emp-dept').value = emp['الدائرة'] || '';
        document.getElementById('emp-section').value = emp['القسم'] || '';
        if (document.getElementById('emp-gender')) {
            document.getElementById('emp-gender').value = emp['الجنس'] || 'ذكر';
        }
        document.getElementById('emp-dob').value = emp['تاريخ الميلاد'] || '';
        document.getElementById('emp-notes').value = '';
        const notes2Select = document.getElementById('emp-notes2');
        Array.from(notes2Select.options).forEach(opt => { opt.selected = opt.value === 'عادي'; });
        form.querySelectorAll('input[type="radio"]').forEach(r => { r.checked = false; });
        calculateScores();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };

    window.editEmpRow = function(index) {
        if (!employeeData[index]) return;
        const emp = employeeData[index];
        document.getElementById('edit-emp-index').value = index;
        
        editEmpFields.innerHTML = '';
        
        // Define which keys we want to allow editing easily
        const keysToEdit = [
            { key: 'الاسم', type: 'text' },
            { key: 'المسمى الوظيفي', type: 'text' },
            { key: 'الدائرة', type: 'text' },
            { key: 'القسم', type: 'text' },
            { key: 'الجنس', type: 'select', opts: ['ذكر', 'أنثى'] },
            { key: 'تاريخ الميلاد', type: 'date' },
            { key: 'العمر', type: 'number' },
            { key: 'رقم الجوال', type: 'text' },
            { key: 'البريد الالكتروني', type: 'email' },
            { key: 'نوع العقد', type: 'text' },
            { key: 'نوع الوظيفة', type: 'text' }
        ];

        keysToEdit.forEach(field => {
            const val = emp[field.key] || (field.key === 'الجنس' ? 'ذكر' : '');
            let inputHtml = '';
            
            if (field.type === 'select') {
                const options = field.opts.map(opt => `<option value="${opt}" ${val === opt ? 'selected' : ''}>${opt}</option>`).join('');
                inputHtml = `<select id="emp_edit_${field.key}" name="${field.key}" style="width: 100%; padding: 0.8rem; border: 1px solid var(--border-color); border-radius: var(--radius-md);">${options}</select>`;
            } else {
                inputHtml = `<input type="${field.type}" id="emp_edit_${field.key}" name="${field.key}" value="${val}" style="width: 100%; padding: 0.8rem; border: 1px solid var(--border-color); border-radius: var(--radius-md);">`;
            }

            const div = document.createElement('div');
            div.className = 'input-group';
            div.innerHTML = `
                <label for="emp_edit_${field.key}">${field.key}</label>
                ${inputHtml}
            `;
            editEmpFields.appendChild(div);
        });

        editEmpModal.style.display = 'flex';
    };

    if(saveEmpDataBtn) {
        saveEmpDataBtn.addEventListener('click', () => {
            const index = document.getElementById('edit-emp-index').value;
            if (!employeeData[index]) return;
            const emp = employeeData[index];
            
            const inputs = editEmpFields.querySelectorAll('input, select');
            inputs.forEach(input => {
                emp[input.name] = input.value;
            });
            
            saveEmployeeDataToStorage();
            renderEmpTable();
            closeEmpModal();
        });
    }

    const compareBtn = document.getElementById('compare-emp-btn');
    const showAllBtn = document.getElementById('show-all-emp-btn');
    const addMissingBtn = document.getElementById('add-missing-emp-btn');

    if(compareBtn) {
        compareBtn.addEventListener('click', () => {
            const evalIds = evaluations.map(e => normalizeEmpId(e.id)).filter(Boolean);
            currentMissingEmps = employeeData.filter(emp => {
                const empId = normalizeEmpId(emp['الرقم الوظيفي']);
                return empId && !evalIds.includes(empId);
            });
            
            renderEmpTable(currentMissingEmps);
            compareBtn.style.display = 'none';
            showAllBtn.style.display = 'inline-block';
            
            if (currentMissingEmps.length > 0 && addMissingBtn) {
                addMissingBtn.style.display = 'inline-block';
            }
            
            alert(`يوجد ${currentMissingEmps.length} موظفين غير مقيمين.`);
        });
    }

    if(addMissingBtn) {
        addMissingBtn.addEventListener('click', () => {
            if (currentMissingEmps.length === 0) return;
            
            let addedCount = 0;
            currentMissingEmps.forEach(emp => {
                const newEval = {
                    id: normalizeEmpId(emp['الرقم الوظيفي']) || String(emp['الرقم الوظيفي']).trim(),
                    name: String(emp['الاسم'] || 'غير محدد'),
                    title: String(emp['المسمى الوظيفي'] || ''),
                    dept: String(emp['الدائرة'] || ''),
                    section: String(emp['القسم'] || ''),
                    dob: String(emp['تاريخ الميلاد'] || ''),
                    p1: 0, p2: 0, p3: 0, p4: 0, p5: 0,
                    n1: 0, n2: 0, n3: 0, n4: 0, n5: 0, n6: 0, n7: 0, c1: 0,
                    notes: '', notes2: '',
                    perfScore: 0, needScore: 0, commScore: 0, totalScore: 0,
                    timestamp: new Date().toISOString()
                };
                evaluations.push(newEval);
                addedCount++;
            });
            
            if (addedCount > 0) {
                saveToLocal();
                updateSavedCount();
                renderTable(); // Update evaluations table view
                
                alert(`تم إدراج ${addedCount} موظف بنجاح في جدول التقييمات!`);
            }
            
            // Reset view
            currentMissingEmps = [];
            addMissingBtn.style.display = 'none';
            showAllBtn.style.display = 'none';
            compareBtn.style.display = 'inline-block';
            renderEmpTable(employeeData);
        });
    }

    if(showAllBtn) {
        showAllBtn.addEventListener('click', () => {
            renderEmpTable(employeeData);
            compareBtn.style.display = 'inline-block';
            showAllBtn.style.display = 'none';
            if(addMissingBtn) addMissingBtn.style.display = 'none';
        });
    }

    setupEmployeeImport();

    const deleteEmpDataBtn = document.getElementById('delete-emp-data-btn');
    if (deleteEmpDataBtn) deleteEmpDataBtn.addEventListener('click', () => {
        if(confirm('هل أنت متأكد من حذف كافة بيانات الموظفين؟')) {
            employeeData = [];
            localStorage.removeItem(EMP_STORAGE_KEY);
            sessionStorage.removeItem(EMP_SESSION_KEY);
            touchDataTimestamp();
            pushToGitHub(true);
            renderEmpTable();
            alert('تم حذف بيانات الموظفين.');
        }
    });

    const exportEmpDataBtn = document.getElementById('export-emp-data-btn');
    if (exportEmpDataBtn) exportEmpDataBtn.addEventListener('click', () => {
        const dataToExport = currentFilteredEmpData && currentFilteredEmpData.length > 0 ? currentFilteredEmpData : employeeData;
        
        if (dataToExport.length === 0) {
            alert('لا توجد بيانات لتصديرها.');
            return;
        }

        const keys = ['الرقم الوظيفي', 'الاسم', 'المسمى الوظيفي', 'الدائرة', 'القسم', 'الجنس', 'نوع الوظيفة', 'تاريخ الميلاد', 'العمر', 'رقم الجوال', 'البريد الالكتروني', 'نوع العقد'];
        const rows = [keys];
        
        const calculateAge = (dobString) => {
            if (!dobString) return '';
            let dob;
            // Handle expected formats like Excel dates or strings
            if (!isNaN(Number(dobString)) && dobString > 10000 && dobString < 90000) {
                dob = new Date(Math.round((Number(dobString) - 25569) * 86400 * 1000));
            } else {
                dob = new Date(dobString);
            }
            if (isNaN(dob.getTime())) return '';
            const diff_ms = Date.now() - dob.getTime();
            const age_dt = new Date(diff_ms); 
            return Math.abs(age_dt.getUTCFullYear() - 1970);
        };

        dataToExport.forEach(emp => {
            if (!emp['العمر']) {
                emp['العمر'] = calculateAge(emp['تاريخ الميلاد']);
            }
            const row = keys.map(k => emp[k] || '');
            rows.push(row);
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, "بيانات الموظفين");
        XLSX.writeFile(wb, "بيانات_الموظفين_المصدرة.xlsx");
    });

    // --- Generate Visual Report (Chart.js) ---
    let reportCharts = [];
    function generateVisualReport() {
        // Destroy existing charts to prevent duplication
        reportCharts.forEach(c => c.destroy());
        reportCharts = [];

        document.getElementById('report-date').textContent = "تاريخ التقرير: " + new Date().toLocaleDateString('ar-EG');

        // Basic Counts
        const totalEmps = employeeData.length;
        const totalEvals = evaluations.length;
        const missingEmps = currentMissingEmps && currentMissingEmps.length > 0 ? currentMissingEmps.length : Math.max(0, totalEmps - totalEvals);
        
        let validEvalsWithScore = evaluations.filter(e => parseFloat(e.totalScore) > 0);
        let sumScores = validEvalsWithScore.reduce((sum, e) => sum + parseFloat(e.totalScore), 0);
        let avgScore = validEvalsWithScore.length > 0 ? (sumScores / validEvalsWithScore.length).toFixed(1) : 0;

        document.getElementById('rep-total-emps').textContent = totalEmps || totalEvals;
        document.getElementById('rep-eval-emps').textContent = totalEvals;
        document.getElementById('rep-missing-emps').textContent = missingEmps;
        document.getElementById('rep-avg-score').textContent = avgScore + " / 45";

        if(window.Chart) {
            // Register DataLabels Plugin globally if available
            if (window.ChartDataLabels) {
                Chart.register(window.ChartDataLabels);
            }
            
            Chart.defaults.font.family = "'Cairo', sans-serif";
            Chart.defaults.font.size = 20;
            Chart.defaults.font.weight = 'bold';
            if(!Chart.defaults.datasets) Chart.defaults.datasets = {};
            if(!Chart.defaults.datasets.bar) Chart.defaults.datasets.bar = {};
            Chart.defaults.datasets.bar.maxBarThickness = 70;
            
            // Common Datalabels config to show numbers clearly
            const commonDataLabels = {
                color: '#fff',
                font: { weight: 'bold', size: 24 },
                formatter: (value) => { return value > 0 ? value : ''; }
            };
            const barDataLabels = {
                color: '#000',
                anchor: 'end',
                align: 'end',
                offset: 4,
                textAlign: 'center',
                font: { weight: 'bold', size: 26 },
                formatter: (value) => { return value > 0 ? value : ''; }
            };

            // Chart 1: Total Scores Distribution
            const scoreGroups = { "ضعيف (0-15)": 0, "متوسط (16-25)": 0, "جيد (26-35)": 0, "ممتاز (36-45)": 0 };
            evaluations.forEach(ev => {
                let ts = parseFloat(ev.totalScore) || 0;
                if (ts <= 15) scoreGroups["ضعيف (0-15)"]++;
                else if (ts <= 25) scoreGroups["متوسط (16-25)"]++;
                else if (ts <= 35) scoreGroups["جيد (26-35)"]++;
                else scoreGroups["ممتاز (36-45)"]++;
            });

            const ctx1 = document.getElementById('chartGauges1').getContext('2d');
            reportCharts.push(new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: Object.keys(scoreGroups),
                    datasets: [{
                        data: Object.values(scoreGroups),
                        backgroundColor: ['#ef4444', '#f59e0b', '#3b82f6', '#10b981']
                    }]
                },
                options: { 
                    responsive: true, 
                    plugins: { 
                        legend: { display: false },
                        datalabels: barDataLabels
                    } 
                }
            }));

            // Chart 2: Job Types
            const jobCounts = { "إداري": 0, "خدمات": 0, "أخرى": 0 };
            (employeeData.length > 0 ? employeeData : evaluations).forEach(e => {
                let jobT = String(e['نوع الوظيفة'] || e.title || '').toLowerCase();
                if (jobT.includes('اداري') || jobT.includes('إداري')) jobCounts["إداري"]++;
                else if (jobT.includes('خدمات')) jobCounts["خدمات"]++;
                else jobCounts["أخرى"]++;
            });

            const ctx2 = document.getElementById('chartJobs').getContext('2d');
            reportCharts.push(new Chart(ctx2, {
                type: 'bar',
                data: {
                    labels: Object.keys(jobCounts),
                    datasets: [{
                        data: Object.values(jobCounts),
                        backgroundColor: ['#8b5cf6', '#6366f1', '#cbd5e1']
                    }]
                },
                options: { 
                    responsive: true, 
                    plugins: { 
                        legend: { display: false },
                        datalabels: barDataLabels
                    } 
                }
            }));

            // --- New Feature: Chart Gender ---
            const genderCountsRep = { "ذكر": 0, "أنثى": 0 };
            employeeData.forEach(e => {
                let g = e['الجنس'] || 'ذكر';
                if (g === 'أنثى') genderCountsRep["أنثى"]++;
                else genderCountsRep["ذكر"]++;
            });
            const ctxGender = document.getElementById('chartGender').getContext('2d');
            reportCharts.push(new Chart(ctxGender, {
                type: 'bar',
                data: {
                    labels: Object.keys(genderCountsRep),
                    datasets: [{
                        data: Object.values(genderCountsRep),
                        backgroundColor: ['#3b82f6', '#ec4899']
                    }]
                },
                options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
            }));

            // --- New Feature: Chart Location ---
            let outCount = 0;
            let inCount = 0;
            evaluations.forEach(ev => {
                if (isNotes2Abroad(ev.notes2)) outCount++;
            });
            inCount = totalEmps > 0 ? (totalEmps - outCount) : (totalEvals - outCount);
            if (inCount < 0) inCount = 0;
            const ctxLocation = document.getElementById('chartLocation').getContext('2d');
            reportCharts.push(new Chart(ctxLocation, {
                type: 'bar',
                data: {
                    labels: ['في الخارج', 'في الداخل'],
                    datasets: [{
                        data: [outCount, inCount],
                        backgroundColor: ['#f43f5e', '#14b8a6']
                    }]
                },
                options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
            }));

            // --- New Feature: Chart Perf Groups (0-5, 6-10, 11+) ---
            const cPerfGroups = { "0 إلى 5": 0, "6 إلى 10": 0, "11 فما فوق": 0 };
            evaluations.forEach(ev => {
                if (isEffectivelyNormalNotes2(ev.notes2)) {
                    const s = ev.perfScore || 0;
                    if (s <= 5) cPerfGroups["0 إلى 5"]++;
                    else if (s <= 10) cPerfGroups["6 إلى 10"]++;
                    else cPerfGroups["11 فما فوق"]++;
                }
            });
            const ctxPerf = document.getElementById('chartPerfGroups').getContext('2d');
            reportCharts.push(new Chart(ctxPerf, {
                type: 'bar',
                data: {
                    labels: Object.keys(cPerfGroups),
                    datasets: [{
                        data: Object.values(cPerfGroups),
                        backgroundColor: ['#ef4444', '#f59e0b', '#10b981']
                    }]
                },
                options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
            }));

            // --- New Feature: Chart P1 (الاستمرارية) ---
            const cP1Groups = { "استمرارية 0": 0, "استمرارية 1": 0, "أكبر من 1": 0 };
            evaluations.forEach(ev => {
                const p1 = parseInt(ev.p1) || 0;
                if (p1 === 0) cP1Groups["استمرارية 0"]++;
                else if (p1 === 1) cP1Groups["استمرارية 1"]++;
                else cP1Groups["أكبر من 1"]++;
            });
            const ctxP1 = document.getElementById('chartP1Groups').getContext('2d');
            reportCharts.push(new Chart(ctxP1, {
                type: 'bar',
                data: {
                    labels: Object.keys(cP1Groups),
                    datasets: [{
                        data: Object.values(cP1Groups),
                        backgroundColor: ['#f87171', '#fbbf24', '#34d399']
                    }]
                },
                options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
            }));

            // --- New Feature: Chart N1 (درجة الحاجة) ---
            const cN1Groups = { "سجل 0": 0, "سجل 1 إلى 3": 0, "سجل 4 فما فوق": 0 };
            evaluations.forEach(ev => {
                const n1 = parseInt(ev.n1) || 0;
                if (n1 === 0) cN1Groups["سجل 0"]++;
                else if (n1 <= 3) cN1Groups["سجل 1 إلى 3"]++;
                else cN1Groups["سجل 4 فما فوق"]++;
            });
            window.cP1Groups = cP1Groups; // Just for potential debugging outside
            
            // To render it we need a canvas inside the index.html. Since we don't know if we can edit the layout immediately without index.html context, I will add it dynamically if missing or just add a general one.
            let n1Canvas = document.getElementById('chartN1Groups');
            if (!n1Canvas) {
                const p1Container = document.getElementById('chartP1Groups').parentNode;
                const newContainer = document.createElement('div');
                newContainer.className = 'chart-card';
                newContainer.innerHTML = '<h3>حسب درجة الحاجة (ح1)</h3><div class="canvas-wrapper"><canvas id="chartN1Groups"></canvas></div>';
                p1Container.parentNode.insertBefore(newContainer, p1Container.nextSibling);
                n1Canvas = document.getElementById('chartN1Groups');
            }
            if (n1Canvas) {
                const ctxN1 = n1Canvas.getContext('2d');
                reportCharts.push(new Chart(ctxN1, {
                    type: 'bar',
                    data: {
                        labels: Object.keys(cN1Groups),
                        datasets: [{
                            data: Object.values(cN1Groups),
                            backgroundColor: ['#e11d48', '#f97316', '#059669']
                        }]
                    },
                    options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
                }));
            }

            // --- New Feature: Chart N2 and N4 ---
            const cN2Groups = { "وجاهي 1 أو أقل": 0, "وجاهي من 2 إلى 4": 0, "وجاهي 5 أيام": 0 };
            const cN4Groups = { "بدون 0": 0, "قليل 1": 0, "متوسط 2": 0, "كبير 3": 0 };
            
            evaluations.forEach(ev => {
                const n2 = parseInt(ev.n2) || 0;
                if (n2 <= 1) cN2Groups["وجاهي 1 أو أقل"]++;
                else if (n2 >= 2 && n2 <= 4) cN2Groups["وجاهي من 2 إلى 4"]++;
                else cN2Groups["وجاهي 5 أيام"]++;

                const n4 = parseInt(ev.n4) || 0;
                if (n4 === 0) cN4Groups["بدون 0"]++;
                else if (n4 === 1) cN4Groups["قليل 1"]++;
                else if (n4 === 2) cN4Groups["متوسط 2"]++;
                else cN4Groups["كبير 3"]++;
            });

            let n2Canvas = document.getElementById('chartN2Groups');
            if (!n2Canvas) {
                const n1Container = document.getElementById('chartN1Groups').parentNode.parentNode;
                const newContainer2 = document.createElement('div');
                newContainer2.className = 'chart-card';
                newContainer2.innerHTML = '<h3>التواجد الوجاهي (ح2)</h3><div class="canvas-wrapper"><canvas id="chartN2Groups"></canvas></div>';
                n1Container.parentNode.insertBefore(newContainer2, n1Container.nextSibling);
                n2Canvas = document.getElementById('chartN2Groups');
            }
            if (n2Canvas) {
                const ctxN2 = n2Canvas.getContext('2d');
                reportCharts.push(new Chart(ctxN2, {
                    type: 'bar',
                    data: {
                        labels: Object.keys(cN2Groups),
                        datasets: [{
                            data: Object.values(cN2Groups),
                            backgroundColor: ['#3b82f6', '#10b981', '#f59e0b']
                        }]
                    },
                    options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
                }));
            }

            let n4Canvas = document.getElementById('chartN4Groups');
            if (!n4Canvas) {
                const n2Container = document.getElementById('chartN2Groups').parentNode.parentNode;
                const newContainer4 = document.createElement('div');
                newContainer4.className = 'chart-card';
                newContainer4.innerHTML = '<h3>حجم المعاملات (ح4)</h3><div class="canvas-wrapper"><canvas id="chartN4Groups"></canvas></div>';
                n2Container.parentNode.insertBefore(newContainer4, n2Container.nextSibling);
                n4Canvas = document.getElementById('chartN4Groups');
            }
            if (n4Canvas) {
                const ctxN4 = n4Canvas.getContext('2d');
                reportCharts.push(new Chart(ctxN4, {
                    type: 'bar',
                    data: {
                        labels: Object.keys(cN4Groups),
                        datasets: [{
                            data: Object.values(cN4Groups),
                            backgroundColor: ['#64748b', '#8b5cf6', '#a855f7', '#d946ef']
                        }]
                    },
                    options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
                }));
            }

            // --- New Feature: Chart Need Score Groups ---
            const cNeedScoreGroups = { "أقل من 10": 0, "من 10-19": 0, "من 20-30": 0 };
            evaluations.forEach(ev => {
                const ns = parseFloat(ev.needScore) || 0;
                if (ns < 10) cNeedScoreGroups["أقل من 10"]++;
                else if (ns >= 10 && ns <= 19) cNeedScoreGroups["من 10-19"]++;
                else cNeedScoreGroups["من 20-30"]++;
            });
            let needScoreCanvas = document.getElementById('chartNeedScoreGroups');
            if (!needScoreCanvas) {
                const n4Container = document.getElementById('chartN4Groups').parentNode.parentNode;
                const newContainer5 = document.createElement('div');
                newContainer5.className = 'chart-card';
                newContainer5.innerHTML = '<h3>مجموع معيار الحاجة</h3><div class="canvas-wrapper"><canvas id="chartNeedScoreGroups"></canvas></div>';
                n4Container.parentNode.insertBefore(newContainer5, n4Container.nextSibling);
                needScoreCanvas = document.getElementById('chartNeedScoreGroups');
            }
            if (needScoreCanvas) {
                const ctxNS = needScoreCanvas.getContext('2d');
                reportCharts.push(new Chart(ctxNS, {
                    type: 'bar',
                    data: {
                        labels: Object.keys(cNeedScoreGroups),
                        datasets: [{
                            data: Object.values(cNeedScoreGroups),
                            backgroundColor: ['#ef4444', '#f59e0b', '#10b981']
                        }]
                    },
                    options: { responsive: true, plugins: { legend: { display: false }, datalabels: barDataLabels } }
                }));
            }

            // Chart 3: Notes2 Distribution (effective + abroad extraction)
            const note2Counts = countNotes2Distribution(evaluations);
            note2Counts[NOTES2_ABROAD] = evaluations.filter(ev => isNotes2Abroad(ev.notes2)).length;
            const note2ChartLabels = [...NOTES2_SPECIAL, NOTES2_NORMAL, NOTES2_ABROAD];
            const note2ChartData = note2ChartLabels.map(k => note2Counts[k] || 0);

                        const ctx3 = document.getElementById('chartNotes').getContext('2d');
            reportCharts.push(new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: note2ChartLabels,
                    datasets: [{
                        label: 'عدد الموظفين',
                        data: note2ChartData,
                        backgroundColor: '#6366f1'
                    }]
                },
                options: { 
                    responsive: true, 
                    scales: { y: { beginAtZero: true, ticks: { precision: 0 } } },
                    plugins: { legend: { display: false }, datalabels: barDataLabels }
                }
            }));
            
            // --- New Feature: Age Groups ---
            const calculateAgeLocalD = (dobString) => {
                if (!dobString) return 0;
                let dob;
                if (!isNaN(Number(dobString)) && dobString > 10000 && dobString < 90000) {
                    dob = new Date(Math.round((Number(dobString) - 25569) * 86400 * 1000));
                } else {
                    dob = new Date(dobString);
                }
                if (isNaN(dob.getTime())) return 0;
                return Math.abs(new Date(Date.now() - dob.getTime()).getUTCFullYear() - 1970);
            };

            const rAgeGroups = {
                "من 30-39": { m: 0, f: 0 },
                "من 40-49": { m: 0, f: 0 },
                "من 50 إلى 54": { m: 0, f: 0 },
                "من 55 فأعلى": { m: 0, f: 0 },
                "أخرى/غير متوفر": { m: 0, f: 0 }
            };

            employeeData.forEach(emp => {
                let age = parseInt(emp['العمر']) || calculateAgeLocalD(emp['تاريخ الميلاد']);
                const gender = (emp['الجنس'] || 'ذكر') === 'ذكر' ? 'm' : 'f';
                if (age >= 30 && age <= 39) rAgeGroups["من 30-39"][gender]++;
                else if (age >= 40 && age <= 49) rAgeGroups["من 40-49"][gender]++;
                else if (age >= 50 && age <= 54) rAgeGroups["من 50 إلى 54"][gender]++;
                else if (age >= 55) rAgeGroups["من 55 فأعلى"][gender]++;
                else rAgeGroups["أخرى/غير متوفر"][gender]++;
            });

            const ctxAge = document.getElementById('chartAgeGroups').getContext('2d');
            reportCharts.push(new Chart(ctxAge, {
                type: 'bar',
                data: {
                    labels: Object.keys(rAgeGroups),
                    datasets: [
                        { label: 'ذكور', data: Object.values(rAgeGroups).map(d => d.m), backgroundColor: '#3b82f6' },
                        { label: 'إناث', data: Object.values(rAgeGroups).map(d => d.f), backgroundColor: '#ec4899' }
                    ]
                },
                options: { 
                    responsive: true, 
                    scales: { y: { beginAtZero: true, ticks: { precision: 0 } } },
                    plugins: { legend: { position: 'bottom' }, datalabels: barDataLabels }
                }
            }));

            // Chart 4: Average Score per Department
            const deptScores = {};
            evaluations.forEach(ev => {
                let d = (ev.dept || 'غير محدد').trim();
                let s = parseFloat(ev.totalScore) || 0;
                if(!deptScores[d]) deptScores[d] = { sum: 0, count: 0 };
                deptScores[d].sum += s;
                deptScores[d].count++;
            });

            const deptLabels = [];
            const avgData = [];
            Object.entries(deptScores).forEach(([dept, data]) => {
                deptLabels.push(dept);
                avgData.push((data.sum / data.count).toFixed(1));
            });

            const ctx4 = document.getElementById('chartDepts').getContext('2d');
            reportCharts.push(new Chart(ctx4, {
                type: 'bar',
                data: {
                    labels: deptLabels,
                    datasets: [{
                        label: 'متوسط التقييم (من 45)',
                        data: avgData,
                        backgroundColor: '#10b981'
                    }]
                },
                options: { 
                    responsive: true, 
                    scales: { y: { beginAtZero: true, max: 45 } },
                    plugins: { legend: { display: false }, datalabels: barDataLabels }
                }
            }));
        }
    }

    document.getElementById('generate-report-btn')?.addEventListener('click', generateVisualReport);
    document.getElementById('print-report-btn')?.addEventListener('click', () => {
        document.body.classList.add('printing-report');
        window.print();
        document.body.classList.remove('printing-report');
    });

    // --- Settings & Backup Logic ---
    const settingsBtn = document.getElementById('settings-btn');
    const settingsModal = document.getElementById('settings-modal');
    const closeSettingsBtn = document.getElementById('close-settings');
    const exportBackupBtn = document.getElementById('export-backup-btn');
    const importBackupUpload = document.getElementById('import-backup-upload');

    function loadGitHubSettingsForm() {
        if (!window.GitHubSync) return;
        const s = window.GitHubSync.getSettingsForForm();
        const ownerEl = document.getElementById('gh-owner');
        const repoEl = document.getElementById('gh-repo');
        const branchEl = document.getElementById('gh-branch');
        const pathEl = document.getElementById('gh-filepath');
        const tokenEl = document.getElementById('gh-token');
        const tokenHint = document.getElementById('gh-token-hint');
        if (ownerEl) ownerEl.value = s.owner || '';
        if (repoEl) repoEl.value = s.repo || '';
        if (branchEl) branchEl.value = s.branch || 'main';
        if (pathEl) pathEl.value = s.filePath || 'data/survey-data.json';
        if (tokenEl) tokenEl.value = '';
        if (tokenHint) {
            tokenHint.textContent = s.hasToken
                ? '✓ يوجد رمز محفوظ في هذه الجلسة'
                : 'أدخل رمز GitHub لرفع البيانات على المستودع';
        }
    }

    const saveGhSettingsBtn = document.getElementById('save-gh-settings');
    if (saveGhSettingsBtn) {
        saveGhSettingsBtn.addEventListener('click', () => {
            if (!window.GitHubSync) return;
            window.GitHubSync.saveSettings({
                owner: document.getElementById('gh-owner')?.value,
                repo: document.getElementById('gh-repo')?.value,
                branch: document.getElementById('gh-branch')?.value,
                filePath: document.getElementById('gh-filepath')?.value,
                token: document.getElementById('gh-token')?.value
            });
            loadGitHubSettingsForm();
            alert('تم حفظ إعدادات GitHub.');
        });
    }

    const syncGhNowBtn = document.getElementById('sync-gh-now');
    if (syncGhNowBtn) {
        syncGhNowBtn.addEventListener('click', async () => {
            const changed = await syncFromGitHub();
            updateSavedCount();
            renderTable();
            if (employeeData.length > 0) renderEmpTable();
            if (changed) alert('تم تحديث البيانات من GitHub.');
            else if (window.GitHubSync?.hasToken()) alert('تمت المزامنة — البيانات المحلية هي الأحدث.');
            else alert('تم التحقق. أضف رمز GitHub لرفع البيانات.');
        });
    }

    if (settingsBtn && settingsModal) {
        settingsBtn.addEventListener('click', () => {
            loadGitHubSettingsForm();
            settingsModal.classList.remove('hidden');
            settingsModal.style.display = 'flex';
        });

        closeSettingsBtn.addEventListener('click', () => {
            settingsModal.classList.add('hidden');
            settingsModal.style.display = 'none';
        });

        // Close on outside click
        settingsModal.addEventListener('click', (e) => {
            if (e.target === settingsModal) {
                settingsModal.classList.add('hidden');
                settingsModal.style.display = 'none';
            }
        });
    }

    if (exportBackupBtn) {
        exportBackupBtn.addEventListener('click', () => {
            const backupData = {
                evaluations: evaluations || [],
                employeeData: employeeData || [],
                colsState: colsState || [],
                empColsState: empColsState || [],
                timestamp: new Date().toISOString()
            };

            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupData, null, 2));
            const downloadAnchorNode = document.createElement('a');
            downloadAnchorNode.setAttribute("href",     dataStr);
            downloadAnchorNode.setAttribute("download", `backup_evaluations_${new Date().toISOString().split('T')[0]}.json`);
            document.body.appendChild(downloadAnchorNode);
            downloadAnchorNode.click();
            downloadAnchorNode.remove();
        });
    }

    if (importBackupUpload) {
        importBackupUpload.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            if (confirm('استيراد النسخة الاحتياطية سيؤدي إلى استبدال كافة البيانات الحالية. هل أنت متأكد من الاستمرار؟')) {
                const reader = new FileReader();
                reader.onload = function(event) {
                    try {
                        const importedData = JSON.parse(event.target.result);
                        
                        if (importedData.evaluations) {
                            localStorage.setItem('evaluations', JSON.stringify(importedData.evaluations));
                        }
                        if (importedData.employeeData) {
                            const empJson = JSON.stringify(importedData.employeeData);
                            localStorage.setItem(EMP_STORAGE_KEY, empJson);
                            sessionStorage.setItem(EMP_SESSION_KEY, empJson);
                        }
                        if (importedData.colsState) {
                            localStorage.setItem('colsState', JSON.stringify(importedData.colsState));
                        }
                        if (importedData.empColsState) {
                            localStorage.setItem('empColsState', JSON.stringify(importedData.empColsState));
                        }
                        touchDataTimestamp();
                        
                        alert('تم استيراد النسخة الاحتياطية بنجاح! سيتم إعادة تحميل الصفحة لتطبيق التغييرات.');
                        window.location.reload();
                    } catch (err) {
                        console.error(err);
                        alert('حدث خطأ أثناء قراءة الملف. تأكد من أنه ملف نسخة احتياطية صالح (JSON).');
                    }
                    importBackupUpload.value = ''; // Reset input
                };
                reader.readAsText(file);
            } else {
                importBackupUpload.value = ''; // Reset input
            }
        });
    }

    function finishInit() {
        reloadFromLocalStorage();
        updateSavedCount();
        try {
            renderTable();
        } catch (renderErr) {
            console.error('خطأ عرض الجدول:', renderErr);
        }
        if (employeeData.length > 0) renderEmpTable();
        const tableTab = document.querySelector('.tab-btn[data-target="table-view"]');
        if (tableTab) tableTab.click();
    }

    finishInit();

    // مزامنة GitHub فقط إذا المحلي فارغ — لا تُمس البيانات المحلية تلقائياً
    syncFromGitHub(true).then((changed) => {
        if (changed) {
            updateSavedCount();
            renderTable();
            if (employeeData.length > 0) renderEmpTable();
        }
    }).catch((err) => console.error('مزامنة GitHub:', err));
});
