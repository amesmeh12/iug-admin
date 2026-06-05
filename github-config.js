/**
 * إعدادات حفظ البيانات على GitHub
 * يُكتشف المستودع تلقائياً من رابط GitHub Pages إن أمكن.
 *
 * لكتابة البيانات على GitHub تحتاج رمز وصول (PAT) بصلاحية repo
 * من: GitHub → Settings → Developer settings → Personal access tokens
 */
(function () {
    function detectFromPagesUrl() {
        const host = window.location.hostname || '';
        if (!host.endsWith('.github.io')) return {};
        const owner = host.replace('.github.io', '');
        const parts = window.location.pathname.split('/').filter(Boolean);
        const htmlIdx = parts.findIndex(p => p.endsWith('.html'));
        const repoParts = htmlIdx >= 0 ? parts.slice(0, htmlIdx) : parts;
        const repo = repoParts[0] || '';
        return { owner, repo };
    }

    const detected = detectFromPagesUrl();

    window.GITHUB_STORAGE = {
        owner: detected.owner || '',
        repo: detected.repo || '',
        branch: 'main',
        filePath: 'data/survey-data.json'
    };
})();
