(function () {
    const SETTINGS_KEY = 'githubStorageSettings';
    const TOKEN_KEY = 'githubToken';
    let saveTimer = null;
    let statusEl = null;
    let fileSha = null;

    function getStoredSettings() {
        try {
            const raw = localStorage.getItem(SETTINGS_KEY);
            return raw ? JSON.parse(raw) : {};
        } catch (e) {
            return {};
        }
    }

    function getConfig() {
        const base = window.GITHUB_STORAGE || {};
        const stored = getStoredSettings();
        return {
            owner: (stored.owner || base.owner || '').trim(),
            repo: (stored.repo || base.repo || '').trim(),
            branch: (stored.branch || base.branch || 'main').trim() || 'main',
            filePath: (stored.filePath || base.filePath || 'data/survey-data.json').trim()
        };
    }

    function isConfigured() {
        const c = getConfig();
        return !!(c.owner && c.repo && c.filePath);
    }

    function hasToken() {
        return !!(sessionStorage.getItem(TOKEN_KEY) || '').trim();
    }

    function setStatus(text, state) {
        if (!statusEl) statusEl = document.getElementById('sync-status');
        if (!statusEl) return;
        statusEl.textContent = text;
        statusEl.className = 'sync-status sync-' + (state || 'idle');
        statusEl.style.display = 'inline-flex';
    }

    function rawUrl(cfg) {
        return 'https://raw.githubusercontent.com/' +
            encodeURIComponent(cfg.owner) + '/' +
            encodeURIComponent(cfg.repo) + '/' +
            encodeURIComponent(cfg.branch) + '/' +
            cfg.filePath.split('/').map(encodeURIComponent).join('/');
    }

    function apiContentsUrl(cfg) {
        return 'https://api.github.com/repos/' +
            encodeURIComponent(cfg.owner) + '/' +
            encodeURIComponent(cfg.repo) + '/contents/' +
            cfg.filePath.split('/').map(encodeURIComponent).join('/');
    }

    function saveSettings(settings) {
        const current = getStoredSettings();
        const next = {
            owner: (settings.owner || current.owner || '').trim(),
            repo: (settings.repo || current.repo || '').trim(),
            branch: (settings.branch || current.branch || 'main').trim() || 'main',
            filePath: (settings.filePath || current.filePath || 'data/survey-data.json').trim()
        };
        localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
        if (settings.token !== undefined) {
            const token = String(settings.token || '').trim();
            if (token) sessionStorage.setItem(TOKEN_KEY, token);
            else sessionStorage.removeItem(TOKEN_KEY);
        }
        return next;
    }

    function getSettingsForForm() {
        const cfg = getConfig();
        return {
            owner: cfg.owner,
            repo: cfg.repo,
            branch: cfg.branch,
            filePath: cfg.filePath,
            hasToken: hasToken()
        };
    }

    async function loadAll() {
        if (!isConfigured()) {
            setStatus('التخزين المحلي فقط', 'offline');
            return null;
        }
        try {
            setStatus('جاري التحميل من GitHub...', 'syncing');
            const cfg = getConfig();
            const res = await fetch(rawUrl(cfg) + '?t=' + Date.now());
            if (res.status === 404) {
                setStatus('متصل — لا يوجد ملف بيانات بعد', 'ok');
                fileSha = null;
                return null;
            }
            if (!res.ok) throw new Error('HTTP ' + res.status);
            const data = await res.json();
            setStatus('تم التحميل من GitHub', 'ok');
            return {
                evaluations: data.evaluations || [],
                employeeData: data.employeeData || [],
                updatedAt: data.updatedAt || ''
            };
        } catch (e) {
            console.error('GitHub load error:', e);
            setStatus('فشل التحميل من GitHub', 'error');
            return null;
        }
    }

    async function pushSave(payload) {
        if (!isConfigured()) {
            setStatus('التخزين المحلي فقط', 'offline');
            return false;
        }
        const token = sessionStorage.getItem(TOKEN_KEY);
        if (!token) {
            setStatus('محفوظ محلياً — أضف رمز GitHub للرفع', 'offline');
            return false;
        }
        try {
            setStatus('جاري الحفظ على GitHub...', 'syncing');
            const cfg = getConfig();
            const updatedAt = new Date().toISOString();
            const body = {
                evaluations: payload.evaluations || [],
                employeeData: payload.employeeData || [],
                updatedAt: updatedAt
            };
            const content = btoa(unescape(encodeURIComponent(JSON.stringify(body, null, 2))));

            if (!fileSha) {
                try {
                    const metaRes = await fetch(apiContentsUrl(cfg), {
                        headers: {
                            Authorization: 'Bearer ' + token,
                            Accept: 'application/vnd.github+json'
                        }
                    });
                    if (metaRes.ok) {
                        const meta = await metaRes.json();
                        fileSha = meta.sha;
                    } else if (metaRes.status !== 404) {
                        throw new Error('meta HTTP ' + metaRes.status);
                    }
                } catch (metaErr) {
                    console.warn('GitHub meta fetch:', metaErr);
                }
            }

            const putBody = {
                message: 'تحديث بيانات الاستبانة ' + updatedAt,
                content: content,
                branch: cfg.branch
            };
            if (fileSha) putBody.sha = fileSha;

            const putRes = await fetch(apiContentsUrl(cfg), {
                method: 'PUT',
                headers: {
                    Authorization: 'Bearer ' + token,
                    Accept: 'application/vnd.github+json',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(putBody)
            });

            if (!putRes.ok) {
                const errJson = await putRes.json().catch(() => ({}));
                throw new Error(errJson.message || ('HTTP ' + putRes.status));
            }
            const result = await putRes.json();
            if (result.content && result.content.sha) fileSha = result.content.sha;

            setStatus('تم الحفظ على GitHub', 'ok');
            return updatedAt;
        } catch (e) {
            console.error('GitHub save error:', e);
            setStatus('فشل الحفظ على GitHub', 'error');
            return false;
        }
    }

    function scheduleSave(payload, immediate) {
        if (!isConfigured()) return;
        if (saveTimer) clearTimeout(saveTimer);
        if (immediate) {
            pushSave(payload);
            return;
        }
        saveTimer = setTimeout(() => pushSave(payload), 1200);
    }

    window.GitHubSync = {
        isConfigured,
        hasToken,
        loadAll,
        pushSave,
        scheduleSave,
        saveSettings,
        getSettingsForForm,
        setStatus
    };
})();
