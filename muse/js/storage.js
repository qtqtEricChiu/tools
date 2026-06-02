/**
 * MBolka Player - Storage & Theme
 * IndexedDB, directory handles, play stats, color extraction, theme logic
 */

// === IndexedDB 初始化 ===
async function initIDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_NAME, IDB_VERSION);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains('metadata')) {
                db.createObjectStore('metadata', { keyPath: 'key' });
            }
            if (!db.objectStoreNames.contains('errors')) {
                db.createObjectStore('errors', { keyPath: 'id', autoIncrement: true });
            }
            if (!db.objectStoreNames.contains('dirHandle')) {
                db.createObjectStore('dirHandle', { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains('stats')) {
                db.createObjectStore('stats', { keyPath: 'key' });
            }
        };
        req.onsuccess = (e) => { idb = e.target.result; resolve(); };
        req.onerror = () => { idb = null; resolve(); };
    });
}

async function cacheMetadata(key, data) {
    if (!idb) return;
    try {
        const tx = idb.transaction('metadata', 'readwrite');
        tx.objectStore('metadata').put({ key, data, timestamp: Date.now() });
    } catch(e) {}
}

async function getCachedMetadata(key) {
    if (!idb) return null;
    try {
        const tx = idb.transaction('metadata', 'readonly');
        const req = tx.objectStore('metadata').get(key);
        return new Promise(resolve => {
            req.onsuccess = () => {
                if (req.result && req.result.data) {
                    const cached = req.result.data;
                    // 核心修复：刷新后旧 Blob 失效，必须为缓存数据重新生成新的 Blob URL
                    if (cached.file) {
                        try { URL.revokeObjectURL(cached.url); } catch(e){}
                        cached.url = URL.createObjectURL(cached.file);
                    }
                    resolve(cached);
                } else {
                    resolve(null);
                }
            };
            req.onerror = () => resolve(null);
        });
    } catch(e) { return null; }
}

// 内存错误日志缓存（用于导出，避免重复解析 localStorage）
const _errorLogsCache = [];

async function logError(type, message, file) {
    try {
        // 写入 IndexedDB
        if (idb) {
            const tx = idb.transaction('errors', 'readwrite');
            tx.objectStore('errors').put({ type, message, file: file ? file.name : '', time: Date.now() });
        }
        // 写入内存缓存
        const entry = { type, message, file: file ? file.name : '', time: new Date().toISOString() };
        _errorLogsCache.push(entry);
        if (_errorLogsCache.length > 500) _errorLogsCache.shift();
        // 写入 localStorage
        try {
            localStorage.setItem('MBolka_ErrorLogs', JSON.stringify(_errorLogsCache));
        } catch(e) {}
    } catch(e) {}
}

// === 目录句柄持久化 (File System Access API) ===
async function saveDirectoryHandle(handle) {
    directoryHandle = handle;
    if (idb) {
        try {
            const tx = idb.transaction('dirHandle', 'readwrite');
            tx.objectStore('dirHandle').put({ id: 'main', handle: handle });
        } catch(e) {}
    }
}

async function loadDirectoryHandle() {
    if (!idb) return null;
    try {
        const tx = idb.transaction('dirHandle', 'readonly');
        const result = await new Promise(resolve => {
            const req = tx.objectStore('dirHandle').get('main');
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => resolve(null);
        });
        if (result && result.handle) {
            // 验证权限
            const opts = { mode: 'read' };
            if (await result.handle.queryPermission(opts) === 'granted' ||
                await result.handle.requestPermission(opts) === 'granted') {
                directoryHandle = result.handle;
                return result.handle;
            }
        }
    } catch(e) {}
    return null;
}

async function loadFromStoredDirectory() {
    const handle = await loadDirectoryHandle();
    if (!handle) return false;
    try {
        const files = [];
        await readDirHandleEntries(handle, files);
        if (files.length) {
            await processFiles(files);
            return true;
        }
    } catch(e) {
        logError('DIR_LOAD', e.message);
    }
    return false;
}

async function readDirHandleEntries(dirHandle, files) {
    for await (const entry of dirHandle.values()) {
        if (entry.kind === 'file') {
            files.push(await entry.getFile());
        } else if (entry.kind === 'directory') {
            await readDirHandleEntries(entry, files);
        }
    }
}

async function pickAndLoadFolder() {
    try {
        const handle = await window.showDirectoryPicker();
        await saveDirectoryHandle(handle);
        const files = [];
        await readDirHandleEntries(handle, files);
        if (files.length) {
            await processFiles(files);
        }
    } catch(e) {
        if (e.name !== 'AbortError') {
            showToast("❌ 无法访问文件夹，请重试");
        }
    }
}

// === 播放统计 ===
function recordPlay(song) {
    if (!song || !song.file) return;
    const key = song.file.name;
    if (!playStats[key]) {
        playStats[key] = { title: song.title, artist: song.artist, count: 0, lastPlay: 0, totalTime: 0 };
    }
    playStats[key].count++;
    playStats[key].lastPlay = Date.now();
    saveSettings();
}

function getTopSongs(limit = 10) {
    return Object.entries(playStats)
        .map(([key, val]) => ({ key, ...val }))
        .sort((a, b) => b.count - a.count)
        .slice(0, limit);
}

function getTotalListenTime() {
    return Object.values(playStats).reduce((sum, s) => sum + (s.totalTime || 0), 0);
}

// === 播放统计追踪 ===
let playStartTime = 0;
audio.addEventListener('play', () => { playStartTime = Date.now(); });
audio.addEventListener('pause', () => {
    if (playStartTime && currentIndex >= 0 && playlist[currentIndex]) {
        const elapsed = (Date.now() - playStartTime) / 1000;
        if (elapsed > 1) {
            const key = playlist[currentIndex].file.name;
            if (!playStats[key]) playStats[key] = { title: playlist[currentIndex].title, artist: playlist[currentIndex].artist, count: 0, lastPlay: 0, totalTime: 0 };
            playStats[key].totalTime += elapsed;
        }
    }
    playStartTime = 0;
});

const extractColor = (imgSrc) => {
    return new Promise((resolve) => {
        if (!imgSrc || imgSrc.startsWith('data:image/svg')) return resolve(null);
        const img = new Image();
        img.onload = () => {
            const cvs = document.createElement('canvas'); cvs.width = 32; cvs.height = 32;
            const ctx = cvs.getContext('2d', { willReadFrequently: true }); ctx.drawImage(img, 0, 0, 32, 32);
            try {
                const data = ctx.getImageData(0, 0, 32, 32).data; let r=0, g=0, b=0, count=0;
                for(let i=0; i<data.length; i+=16) { if(data[i]>20 && data[i]<235) { r+=data[i]; g+=data[i+1]; b+=data[i+2]; count++; } }
                resolve(count > 0 ? `rgb(${~~(r/count)},${~~(g/count)},${~~(b/count)})` : null);
            } catch(e) { resolve(null); }
        }; img.onerror = () => resolve(null); img.src = imgSrc;
    });
};
const getHueFromRgb = (rgbStr) => {
    if (!rgbStr) return 210; const match = rgbStr.match(/\d+/g); if (!match) return 210;
    let r = parseInt(match[0])/255, g = parseInt(match[1])/255, b = parseInt(match[2])/255;
    const max = Math.max(r,g,b), min = Math.min(r,g,b); let h = 0;
    if (max !== min) { const d = max - min; switch (max) { case r: h = (g-b)/d + (g<b?6:0); break; case g: h = (b-r)/d + 2; break; case b: h = (r-g)/d + 4; break; } h /= 6; } return h * 360;
};

// === 歌词设置应用 ===
function applyLrcSettings() {
    document.documentElement.style.setProperty('--lrc-font-size', `${cfg.lrcFontSize}px`);
    document.documentElement.style.setProperty('--lrc-line-height', cfg.lrcLineHeight);
    document.documentElement.style.setProperty('--lrc-align', cfg.lrcAlign);
    document.getElementById('lrcFontSizeVal').textContent = `${cfg.lrcFontSize}px`;
    document.getElementById('lrcLineHeightVal').textContent = cfg.lrcLineHeight;
    document.querySelectorAll('.lrc-align-btn').forEach(b => b.classList.toggle('active', b.dataset.align === cfg.lrcAlign));
}

// === 核心视觉与主题逻辑 ===
const applyThemeLogic = () => {
    let targetColor = cfg.defaultColor; let showImg = false, showColor = false, bgUrl = '';

    if (cfg.colorMode) targetColor = cfg.customBgImg ? (cfg.customBgColor || targetColor) : (currentAlbumColor || targetColor);
    document.documentElement.style.setProperty('--primary', targetColor);

    // v2.5-p2: WCAG 实时计算亮度，动态决定按钮前景色是反白还是用暗色
    const luminance = getLuminance(targetColor);
    const textOnPrimary = luminance < 140 ? '#ffffff' : '#0a0a1a';
    document.documentElement.style.setProperty('--text-on-primary', textOnPrimary);

    document.documentElement.style.setProperty('--bg-blur', `${cfg.blurAmt}px`);
    targetHue = getHueFromRgb(targetColor);

    if (cfg.customBgImg) { showImg = true; bgUrl = cfg.customBgImg; }
    else if (hasCurrentAlbumArt && !cfg.colorMode) { showImg = true; bgUrl = el.mainArt.src; }
    else showColor = true;

    if (showImg) {
        el.bgImg.style.backgroundImage = `url(${bgUrl})`;
        el.bgImg.classList.add('active');
        el.bgColor.classList.remove('active');
    } else if (showColor) {
        // v2.5: Canvas 流沙背景只需激活，颜色由 drawFlowingSand 实时渲染
        el.bgColor.classList.add('active');
        el.bgImg.classList.remove('active');
    }
};

const toggleDarkMode = () => {
    cfg.darkMode = !cfg.darkMode;
    document.body.classList.toggle('dark-mode', cfg.darkMode);
    updateDarkModeUI();
    saveSettings();
    showToast(cfg.darkMode ? "🌙 已开启深色/护眼模式" : "☀️ 已恢复标准模式");
};
function updateDarkModeUI() {
    const btn = document.getElementById('btnToggleDarkMode');
    if (btn) btn.textContent = cfg.darkMode ? '☀️ 标准模式' : '🌙 深色模式';
    document.body.classList.toggle('dark-mode', cfg.darkMode);
}

const toggleImmersiveMode = () => {
    isImmersiveMode = !isImmersiveMode;
    if (isImmersiveMode) {
        el.viewMain.classList.add('hidden'); el.viewImm.classList.remove('hidden');
        document.body.style.background = 'var(--bg-darker)';
        immCanvasCleared = false;
        showToast("🚀 已进入沉浸式音乐舱");
    } else {
        el.viewImm.classList.add('hidden'); el.viewMain.classList.remove('hidden');
        document.body.style.background = 'var(--bg-dark)';
        // 清理沉浸canvas - 修复canvas残留
        immCanvasCleared = true;
        const ctx = el.canvasImm.getContext('2d');
        ctx.clearRect(0, 0, el.canvasImm.width, el.canvasImm.height);
        particles = [];
        ripples = [];
    }
    updateFocusContext();
};

const toggleFullscreen = () => {
    if (!document.fullscreenElement) { document.documentElement.requestFullscreen().catch(e=>{}); showToast("⛶ 进入全屏"); }
    else { if (document.exitFullscreen) document.exitFullscreen(); showToast("⛶ 退出全屏"); }
};
