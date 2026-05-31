/**
 * MBolka Player - Ultimate Nexus v2.4.0
 * Main Application Logic
 */

// === Worker 内联定义 (Web Worker 解析元数据) ===
const workerCode = `
    self.onmessage = async function(e) {
        const { files, index } = e.data;
        const results = [];
        const batch = files.slice(index, index + 10);
        for (const file of batch) {
            const meta = { title: file.name.replace(/\\.[^/.]+$/, ""), artist: "未知" };
            results.push(meta);
        }
        self.postMessage({ results, batchSize: batch.length });
    };
`;
const workerBlob = new Blob([workerCode], { type: 'application/javascript' });
const workerUrl = URL.createObjectURL(workerBlob);
let metaWorker = null;

function initWorker() {
    if (metaWorker) return;
    try { metaWorker = new Worker(workerUrl); } catch(e) { metaWorker = null; }
}

// === 核心状态与全局变量 ===
const audio = new Audio(); audio.crossOrigin = "anonymous";
let audioCtx, analyser, source, dataArray;
let spectrumCtxMain;

let playlist = [];     // 当前活跃播放队列
let musicLibrary = []; // 🚀 新增：导入的完整本地音乐库（不因播放模式而缩水）
let lrcMap = new Map(), playHistory = [];
let currentIndex = -1, isPlaying = false, isShuffle = false, isRepeatOne = false, isImmersiveMode = false;
let parsedLyrics = [], isUserScrollingLyrics = false, lyricsScrollTimeout = null;
let gamepadConnected = false, prevPadBtns = [];
let lyricsOffset = 0; // 歌词时间偏移(秒)

// A-B 重复
let abMode = false, abPointA = null, abPointB = null;

// 视觉特效参数
let particles = [], ripples = [];
const MAX_PARTICLES = 120, MAX_RIPPLES = 12;
let emitterX = window.innerWidth / 2, emitterY = window.innerHeight / 2;
let currentHue = 210, targetHue = 210, visTime = 0;
let mouseX = window.innerWidth / 2, mouseY = window.innerHeight / 2;
let flowField = [];
let bassHistory = [];
let immCanvasCleared = false; // 防止沉浸canvas残留

// 播放引擎增强
let playbackRate = 1.0;
let preservesPitch = true;
let crossfadeEnabled = false;
let crossfadeDuration = 3; // 秒
let isFading = false; // 🚀 全局淡入淡出锁，防止 timeupdate 触发多重定时器崩溃

// 均衡器
let eqFilters = [];
let eqBands = [32, 64, 125, 250, 500, 1000, 2000, 4000, 8000, 16000];
let eqGains = new Array(10).fill(0); // dB

// 睡眠定时器
let sleepTimer = null;
let sleepEndTime = null;

// 性能模式
let performanceMode = false;
let targetFPS = 60;
let currentFPS = 60;
let fpsFrames = 0;
let fpsLastTime = performance.now();
let particleCount = MAX_PARTICLES; // 动态粒子数量

// 偏好配置
let cfg = {
    colorMode: false, customBgImg: null, customBgColor: null, blurAmt: 40,
    defaultColor: '#9ac8e2', darkMode: false, lrcFontSize: 18, lrcLineHeight: 2.2,
    lrcAlign: 'center', themePreset: null
};
let currentAlbumColor = null, hasCurrentAlbumArt = false;
let favorites = new Set();
let currentViewMode = 'list';
let ctxMenuTarget = -1;

// IndexedDB
let idb = null;
const IDB_NAME = 'MBolkaPlayerDB', IDB_VERSION = 2;

// 目录句柄持久化
let directoryHandle = null;

// 播放统计
let playStats = {};

// === 预设主题色 ===
const themePresets = [
    { name: '默认蓝', color: '#9ac8e2' },
    { name: '赛博朋克', color: '#ff00ff' },
    { name: '暖阳', color: '#ff8c42' },
    { name: '极光', color: '#00e5a0' },
    { name: '星夜', color: '#7c5cfc' },
    { name: '樱花', color: '#ff6b9d' },
    { name: '深海', color: '#00b4d8' },
    { name: '日落', color: '#ff6b35' },
    { name: '薄荷', color: '#48cae4' },
    { name: '玫瑰金', color: '#e8b4b8' },
];

// === DOM 引用映射 ===
const el = {
    viewMain: document.getElementById('view-main'), viewImm: document.getElementById('view-immersive'),
    bgImg: document.getElementById('bg-layer-img'), bgColor: document.getElementById('bg-layer-color'),
    canvasImm: document.getElementById('immersive-bg-canvas'), canvasMain: document.getElementById('spectrumCanvasMain'),
    folderIn: document.getElementById('folderInput'), btnLoad: document.getElementById('btnLoadFolder'),
    btnPlay: document.getElementById('btnPlay'), btnPrev: document.getElementById('btnPrev'), btnNext: document.getElementById('btnNext'),
    progAreaMain: document.getElementById('main-progArea'), progFillMain: document.getElementById('main-progFill'),
    progTipMain: document.getElementById('main-progTip'),
    timeCur: document.getElementById('timeCur'), timeTot: document.getElementById('timeTot'), volSlider: document.getElementById('volSlider'),
    mainColAlbum: document.getElementById('main-col-album'), mainTitle: document.getElementById('main-songTitle'), mainArtist: document.getElementById('main-songArtist'), mainArt: document.getElementById('main-artImg'),
    artBox: document.getElementById('artBox'),
    lrcView: document.getElementById('lrcViewport'), lrcPanel: document.getElementById('lrcPanel'), btnToggleLrc: document.getElementById('btnToggleLrc'),
    toast: document.getElementById('toastMsg'), loadBar: document.getElementById('loadBar'), loadWrap: document.getElementById('loadWrap'),
    btnMode: document.getElementById('btnModeToggle'),
    padStatus: document.getElementById('gamepadStatus'), fileInfo: document.getElementById('fileInfo'),
    settingsModal: document.getElementById('settingsModal'), playlistModal: document.getElementById('playlistModal'), plContainer: document.getElementById('playlistContainer'),
    contextMenu: document.getElementById('contextMenu'),
    emptyState: document.getElementById('emptyState'),
    coverWallContainer: document.getElementById('coverWallContainer'),
    fileInfoModal: document.getElementById('fileInfoModal'), fileInfoContent: document.getElementById('fileInfoContent'),
    helpModal: document.getElementById('helpModal'),
    btnFavQuick: document.getElementById('btnFavQuick'),
    btnPipQuick: document.getElementById('btnPipQuick'),

    // 沉浸模式
    immTrackCard: document.getElementById('imm-trackCard'), immArt: document.getElementById('imm-artImg'), immTitle: document.getElementById('imm-songTitle'), immArtist: document.getElementById('imm-songArtist'),
    immLrcCenter: document.getElementById('imm-lyricsCenter'), immCurrLine: document.getElementById('imm-currLine'), immNextLine: document.getElementById('imm-nextLine'),
    immProgArea: document.getElementById('imm-progArea'), immProgFill: document.getElementById('imm-progFill'), immProgTip: document.getElementById('imm-progTip'),
    immBtnPlay: document.getElementById('imm-btnPlay'), immBtnPrev: document.getElementById('imm-btnPrev'), immBtnNext: document.getElementById('imm-btnNext'), immBtnMode: document.getElementById('imm-btnMode'),
    immExitHint: document.getElementById('immExitHint'),
};

// === 工具函数 ===
const showToast = (msg, icon='') => { el.toast.innerHTML = `${icon} ${msg}`; el.toast.classList.add('show'); setTimeout(() => el.toast.classList.remove('show'), 2500); };
const formatTime = (sec) => { if (!sec || isNaN(sec)) return '0:00'; const m = Math.floor(sec / 60), s = Math.floor(sec % 60); return `${m}:${s.toString().padStart(2, '0')}`; };
const decodeText = (str) => { if (!str) return ''; let s = str.replace(/\\u([0-9a-fA-F]{4})/g, (m, g) => String.fromCharCode(parseInt(g, 16))); const txt = document.createElement("textarea"); txt.innerHTML = s; return txt.value; };

const saveSettings = () => {
    try {
        localStorage.setItem('MBolka_Cfg_v3', JSON.stringify({
            colorMode: cfg.colorMode, blurAmt: cfg.blurAmt, vol: audio.volume,
            isShuffle: isShuffle, isRepeatOne: isRepeatOne,
            customBgImg: cfg.customBgImg, customBgColor: cfg.customBgColor,
            darkMode: cfg.darkMode, lrcFontSize: cfg.lrcFontSize,
            lrcLineHeight: cfg.lrcLineHeight, lrcAlign: cfg.lrcAlign,
            themePreset: cfg.themePreset, playbackRate: playbackRate,
            preservesPitch: preservesPitch, crossfadeEnabled: crossfadeEnabled,
            crossfadeDuration: crossfadeDuration, performanceMode: performanceMode,
            eqGains: eqGains, lyricsOffset: lyricsOffset
        }));
        localStorage.setItem('MBolka_Favorites_v3', JSON.stringify([...favorites]));
        // Save play stats
        if (Object.keys(playStats).length) {
            localStorage.setItem('MBolka_Stats', JSON.stringify(playStats));
        }
    } catch(e){}
};
const loadSettings = () => {
    try {
        const stored = JSON.parse(localStorage.getItem('MBolka_Cfg_v3') || localStorage.getItem('MBolka_Cfg_v2'));
        if (stored) {
            cfg.colorMode = stored.colorMode ?? false;
            cfg.blurAmt = stored.blurAmt ?? 40;
            audio.volume = stored.vol ?? 0.7;
            isShuffle = stored.isShuffle ?? false;
            isRepeatOne = stored.isRepeatOne ?? false;
            cfg.customBgImg = stored.customBgImg ?? null;
            cfg.customBgColor = stored.customBgColor ?? null;
            cfg.darkMode = stored.darkMode ?? false;
            cfg.lrcFontSize = stored.lrcFontSize ?? 18;
            cfg.lrcLineHeight = stored.lrcLineHeight ?? 2.2;
            cfg.lrcAlign = stored.lrcAlign ?? 'center';
            cfg.themePreset = stored.themePreset ?? null;
            playbackRate = stored.playbackRate ?? 1.0;
            preservesPitch = stored.preservesPitch ?? true;
            crossfadeEnabled = stored.crossfadeEnabled ?? false;
            crossfadeDuration = stored.crossfadeDuration ?? 3;
            performanceMode = stored.performanceMode ?? false;
            eqGains = stored.eqGains ?? new Array(10).fill(0);
            lyricsOffset = stored.lyricsOffset ?? 0;
            el.volSlider.value = audio.volume;
            document.getElementById('blurSlider').value = cfg.blurAmt;
            document.getElementById('blurVal').textContent = `${cfg.blurAmt}px`;
            document.getElementById('lrcFontSizeSlider').value = cfg.lrcFontSize;
            document.getElementById('lrcLineHeightSlider').value = cfg.lrcLineHeight;
            document.getElementById('lrcFontSizeVal').textContent = `${cfg.lrcFontSize}px`;
            document.getElementById('lrcLineHeightVal').textContent = cfg.lrcLineHeight;
            applyLrcSettings();
            updateModeUI();
            updateSettingsUI();
            updateDarkModeUI();
            applyThemeLogic();
            if (cfg.themePreset) {
                document.documentElement.style.setProperty('--primary', cfg.themePreset);
            }
        }
        const favs = JSON.parse(localStorage.getItem('MBolka_Favorites_v3') || localStorage.getItem('MBolka_Favorites_v2'));
        if (favs) favorites = new Set(favs);
        const stats = JSON.parse(localStorage.getItem('MBolka_Stats') || '{}');
        if (stats) playStats = stats;
    } catch(e) { audio.volume = 0.7; }
};

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
                    // 🚀 核心修复：刷新后旧 Blob 失效，必须为缓存数据重新生成新的 Blob URL
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
    document.documentElement.style.setProperty('--bg-blur', `${cfg.blurAmt}px`);
    targetHue = getHueFromRgb(targetColor);

    if (cfg.customBgImg) { showImg = true; bgUrl = cfg.customBgImg; }
    else if (hasCurrentAlbumArt && !cfg.colorMode) { showImg = true; bgUrl = el.mainArt.src; }
    else showColor = true;

    if (showImg) { el.bgImg.style.backgroundImage = `url(${bgUrl})`; el.bgImg.classList.add('active'); el.bgColor.classList.remove('active'); }
    else if (showColor) { el.bgColor.style.background = `linear-gradient(135deg, ${targetColor} 0%, var(--bg-dark) 100%)`; el.bgColor.classList.add('active'); el.bgImg.classList.remove('active'); }
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

// === 文件解析与播放列表 ===
const parseMetadata = async (file) => {
    const key = `${file.name}_${file.size}_${file.lastModified}`;
    const cached = await getCachedMetadata(key);
    if (cached) return cached;

    return new Promise(resolve => {
        const url = URL.createObjectURL(file);
        const meta = { title: file.name.replace(/\.[^/.]+$/, ""), artist: "未知", album: "", url, file, error: false, lrcText: null };
        if (window.jsmediatags) {
            jsmediatags.read(file, {
                onSuccess: tag => {
                    if(tag.tags.title) meta.title = decodeText(tag.tags.title);
                    if(tag.tags.artist) meta.artist = decodeText(tag.tags.artist);
                    if(tag.tags.album) meta.album = decodeText(tag.tags.album);
                    // 内嵌歌词 (USLT/SYLT)
                    if(tag.tags.lyrics) {
                        meta.lrcText = decodeText(tag.tags.lyrics.lyrics || tag.tags.lyrics);
                    }
                    if(tag.tags.picture) {
                        let b64 = '';
                        const d = tag.tags.picture.data;
                        for(let i=0; i<d.length; i++) b64 += String.fromCharCode(d[i]);
                        meta.art = `data:${tag.tags.picture.format};base64,${window.btoa(b64)}`;
                    }
                    cacheMetadata(key, meta);
                    resolve(meta);
                },
                onError: () => {
                    cacheMetadata(key, meta);
                    resolve(meta);
                }
            });
        } else {
            cacheMetadata(key, meta);
            resolve(meta);
        }
    });
};

const renderPlaylist = () => {
    if (currentViewMode === 'coverwall') {
        renderCoverWall();
        return;
    }
    document.getElementById('playlistModalTitle').textContent = '播放列表';
    el.coverWallContainer.style.display = 'none';
    el.plContainer.style.display = 'flex';
    el.plContainer.innerHTML = '';
    playlist.forEach((s, i) => {
        const div = document.createElement('div');
        const isFav = favorites.has(s.file.name);
        let classes = 'pl-item focusable';
        if (i === currentIndex) classes += ' active';
        if (s.error) classes += ' error';
        div.className = classes;
        div.draggable = true;
        div.dataset.index = i;
        div.innerHTML = `<span class="pl-title">${s.title}</span><span style="font-size:12px;opacity:0.6;">${s.artist}</span><span class="favorite-btn ${isFav ? 'faved' : ''}" data-idx="${i}" title="收藏">${isFav ? '❤️' : '🤍'}</span>`;
        div.onclick = (e) => {
            if (e.target.classList.contains('favorite-btn')) {
                e.stopPropagation();
                toggleFavorite(i);
                return;
            }
            playAudio(i); closeAllModals();
        };
        div.oncontextmenu = (e) => { e.preventDefault(); ctxMenuTarget = i; showContextMenu(e.clientX, e.clientY); };

        // 拖拽排序 - 带插入线视觉反馈
        div.ondragstart = (e) => {
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', i.toString());
            div.classList.add('dragging');
            // 移除所有已有的插入线
            document.querySelectorAll('.drag-insert-line').forEach(l => l.remove());
        };
        div.ondragend = () => {
            div.classList.remove('dragging');
            document.querySelectorAll('.pl-item').forEach(d => d.classList.remove('drag-over'));
            document.querySelectorAll('.drag-insert-line').forEach(l => l.remove());
        };
        div.ondragover = (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            // 移除旧插入线
            document.querySelectorAll('.drag-insert-line').forEach(l => l.remove());
            // 判断插入位置（上半部=前，下半部=后）
            const rect = div.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            const insertLine = document.createElement('div');
            insertLine.className = 'drag-insert-line show';
            if (e.clientY < midY) {
                // 插入到当前项之前
                insertLine.style.top = (rect.top - el.plContainer.getBoundingClientRect().top + el.plContainer.scrollTop - 2) + 'px';
            } else {
                // 插入到当前项之后
                insertLine.style.top = (rect.bottom - el.plContainer.getBoundingClientRect().top + el.plContainer.scrollTop - 1) + 'px';
            }
            insertLine.style.left = '10px';
            insertLine.style.right = '10px';
            el.plContainer.style.position = 'relative';
            el.plContainer.appendChild(insertLine);
        };
        div.ondragleave = (e) => {
            // 只在真正离开时移除
            if (!div.contains(e.relatedTarget)) {
                document.querySelectorAll('.drag-insert-line').forEach(l => l.remove());
            }
        };
        div.ondrop = (e) => {
            e.preventDefault();
            document.querySelectorAll('.drag-insert-line').forEach(l => l.remove());
            const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
            const rect = div.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            let toIdx = i;
            if (e.clientY >= midY && fromIdx < i) toIdx++;
            if (e.clientY < midY && fromIdx > i) toIdx--;

            if (fromIdx !== toIdx && toIdx >= 0 && toIdx <= playlist.length) {
                const [moved] = playlist.splice(fromIdx, 1);
                playlist.splice(toIdx, 0, moved);
                if (currentIndex === fromIdx) currentIndex = toIdx;
                else if (currentIndex > fromIdx && currentIndex <= toIdx) currentIndex--;
                else if (currentIndex < fromIdx && currentIndex >= toIdx) currentIndex++;
                renderPlaylist();
                showToast("📋 播放列表已重排");
            }
            div.classList.remove('drag-over');
        };

        el.plContainer.appendChild(div);
    });
    updateFocusContext();
    const activeItem = el.plContainer.querySelector('.active');
    if(activeItem && el.playlistModal.classList.contains('open')) activeItem.scrollIntoView({ block: 'center', behavior: 'smooth' });
};

// 曲库渲染 - 按专辑/封面聚合
function renderCoverWall() {
    el.plContainer.style.display = 'none';
    el.coverWallContainer.style.display = 'flex';
    el.coverWallContainer.innerHTML = '';
    document.getElementById('playlistModalTitle').textContent = '曲库';

    const groups = new Map();
    playlist.forEach((s, i) => {
        const key = s.art || '__noart__';
        if (!groups.has(key)) {
            groups.set(key, { art: s.art || null, songs: [], firstIdx: i });
        }
        groups.get(key).songs.push(i);
    });

    const sorted = [...groups.entries()].sort((a, b) => b[1].songs.length - a[1].songs.length);

    sorted.forEach(([key, group]) => {
        const container = document.createElement('div');
        const firstSong = playlist[group.firstIdx];
        const hasActive = group.songs.includes(currentIndex);
        container.className = `cover-album-group ${hasActive ? 'active' : ''}`;

        const artDiv = document.createElement('div');
        artDiv.className = 'cover-album-art';

        if (group.art) {
            const img = document.createElement('img');
            img.src = group.art;
            artDiv.appendChild(img);
        } else {
            const noArt = document.createElement('div');
            noArt.className = 'cw-no-art';
            noArt.textContent = '🎵';
            artDiv.appendChild(noArt);
        }

        if (group.songs.length > 1) {
            const count = document.createElement('div');
            count.className = 'cover-album-count';
            count.textContent = group.songs.length;
            artDiv.appendChild(count);
        }

        const info = document.createElement('div');
        info.className = 'cover-album-info';
        const albumName = group.art ? (firstSong.album || firstSong.title) : '无封面';
        const nameEl = document.createElement('div');
        nameEl.className = 'cover-album-name';
        nameEl.textContent = albumName.length > 15 ? albumName.slice(0,14)+'…' : albumName;
        const artistEl = document.createElement('div');
        artistEl.className = 'cover-album-artist';
        artistEl.textContent = `${group.songs.length} 首 · ${firstSong.artist}`;

        info.appendChild(nameEl);
        info.appendChild(artistEl);

        container.appendChild(artDiv);
        container.appendChild(info);

        container.onclick = () => {
            playAudio(group.firstIdx);
            closeAllModals();
        };
        container.oncontextmenu = (e) => {
            e.preventDefault();
            if (group.songs.length === 1) {
                ctxMenuTarget = group.firstIdx;
                showContextMenu(e.clientX, e.clientY);
            }
        };

        el.coverWallContainer.appendChild(container);
    });

    updateFocusContext();
}

// === 右键菜单 ===
function showContextMenu(x, y) {
    el.contextMenu.style.left = `${x}px`;
    el.contextMenu.style.top = `${y}px`;
    el.contextMenu.classList.add('show');
}
function hideContextMenu() {
    el.contextMenu.classList.remove('show');
}

el.contextMenu.addEventListener('click', (e) => {
    const action = (e.target && e.target.closest ? e.target.closest('.ctx-item')?.dataset.action : null);
    if (!action) return;
    hideContextMenu();
    switch(action) {
        case 'play':
            if (ctxMenuTarget >= 0) playAudio(ctxMenuTarget);
            break;
        case 'info':
            if (ctxMenuTarget >= 0) showFileInfo(ctxMenuTarget);
            break;
        case 'favorite':
            if (ctxMenuTarget >= 0) toggleFavorite(ctxMenuTarget);
            break;
        case 'remove':
            if (ctxMenuTarget >= 0) removeFromPlaylist(ctxMenuTarget);
            break;
        case 'clear':
            clearPlaylist();
            break;
    }
});

document.addEventListener('click', (e) => {
    if (!el.contextMenu.contains(e.target)) hideContextMenu();
});

function showFileInfo(idx) {
    const song = playlist[idx];
    if (!song) return;
    const file = song.file;
    const info = `
        <div><b>标题:</b> ${song.title}</div>
        <div><b>艺术家:</b> ${song.artist}</div>
        ${song.album ? `<div><b>专辑:</b> ${song.album}</div>` : ''}
        <div><b>文件名:</b> ${file.name}</div>
        <div><b>文件大小:</b> ${(file.size / 1048576).toFixed(2)} MB</div>
        <div><b>格式:</b> ${file.name.split('.').pop().toUpperCase()}</div>
        <div><b>时长:</b> ${audio.duration ? formatTime(audio.duration) : '未知'}</div>
        <div><b>有封面:</b> ${song.art ? '是' : '否'}</div>
        <div><b>收藏:</b> ${favorites.has(file.name) ? '❤️ 已收藏' : '否'}</div>
    `;
    el.fileInfoContent.innerHTML = info;
    el.fileInfoModal.classList.add('open');
    updateFocusContext();
}

function removeFromPlaylist(idx) {
    if (idx < 0 || idx >= playlist.length) return;
    const song = playlist[idx];
    playlist.splice(idx, 1);
    if (idx === currentIndex) {
        currentIndex = -1;
        if (playlist.length > 0) playAudio(0);
    } else if (idx < currentIndex) {
        currentIndex--;
    }
    renderPlaylist();
    showToast(`🗑️ 已移除: ${song.title}`);
}

function clearPlaylist() {
    playlist = [];
    currentIndex = -1;
    audio.pause();
    setPlayState(false);
    audio.src = '';
    el.mainTitle.textContent = 'MBolka Player Ultimate';
    el.mainArtist.textContent = '等待载入音乐...';
    document.title = 'MBolka Player';
    renderPlaylist();
    updateEmptyState();
    showToast("🚫 播放列表已清空");
}

function toggleFavorite(idx) {
    const song = playlist[idx];
    if (!song) return;
    const key = song.file.name;
    if (favorites.has(key)) {
        favorites.delete(key);
    } else {
        favorites.add(key);
    }
    saveSettings();
    renderPlaylist();
    updateFavQuickBtn();
}

// 更新首页收藏快捷按钮
function updateFavQuickBtn() {
    if (!el.btnFavQuick) return;
    const song = playlist[currentIndex];
    if (song && favorites.has(song.file.name)) {
        el.btnFavQuick.classList.add('faved');
    } else {
        el.btnFavQuick.classList.remove('faved');
    }
}

// 更新首页画中画快捷按钮
function updatePipQuickBtn() {
    if (!el.btnPipQuick) return;
    if (pipWindow && !pipWindow.closed) {
        el.btnPipQuick.classList.add('pip-active');
    } else {
        el.btnPipQuick.classList.remove('pip-active');
    }
}

function updateEmptyState() {
    el.emptyState.style.display = playlist.length === 0 ? 'flex' : 'none';
}

// === 拖拽文件夹支持 (仅在以下场景生效: 主界面空白区、空状态区域、播放列表区域) ===
let dragCounter = 0;
document.addEventListener('dragenter', (e) => {
    e.preventDefault(); e.stopPropagation();
    dragCounter++;
    // 显示拖拽视觉反馈
    if (dragCounter === 1 && playlist.length === 0) {
        el.emptyState.style.opacity = '0.5';
    }
});

document.addEventListener('dragleave', (e) => {
    e.preventDefault(); e.stopPropagation();
    dragCounter--;
    if (dragCounter === 0) {
        el.emptyState.style.opacity = '';
    }
});

document.addEventListener('dragover', (e) => {
    e.preventDefault(); e.stopPropagation();
    // 安全检测：target 可能不是 Element（如拖到文档边缘或文本节点）
    const target = e.target && e.target.closest ? e.target : null;
    // 防止在模态框、按钮、专辑封面上拖放
    if (target && (target.closest('.modal-overlay') || target.closest('.art-box') ||
        target.closest('.btn-group-main') || target.closest('.vis-canvas-container'))) {
        e.dataTransfer.dropEffect = 'none';
        return;
    }
    e.dataTransfer.dropEffect = 'copy';
});

document.addEventListener('drop', async (e) => {
    e.preventDefault(); e.stopPropagation();
    dragCounter = 0;
    el.emptyState.style.opacity = '';

    // 禁止在模态框、按钮区、专辑封面上拖放（安全检测 target 类型）
    const dropTarget = e.target && e.target.closest ? e.target : null;
    if (dropTarget && (dropTarget.closest('.modal-overlay') || dropTarget.closest('.btn-group-main') ||
        dropTarget.closest('.vis-canvas-container') || dropTarget.closest('.art-box'))) {
        return;
    }

    // 如果正在加载中，拒绝新的拖入
    if (isLoadingFiles) {
        showToast("⚠️ 正在加载中，请稍候再拖入");
        return;
    }

    const items = e.dataTransfer.items;
    if (!items) return;
    const files = [];
    for (let item of items) {
        if (item.kind === 'file') {
            const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
            if (entry && entry.isDirectory) {
                await readDirEntries(entry, files);
            } else {
                files.push(item.getAsFile());
            }
        }
    }
    if (files.length) processFiles(files);
});

async function readDirEntries(dirEntry, files) {
    const entries = await new Promise(resolve => dirEntry.createReader().readEntries(resolve));
    for (let entry of entries) {
        if (entry.isFile) {
            files.push(await new Promise(resolve => entry.file(resolve)));
        } else if (entry.isDirectory) {
            await readDirEntries(entry, files);
        }
    }
}

// 全局拖拽锁，防止重复加载
let isLoadingFiles = false;

async function processFiles(files) {
    if (isLoadingFiles) {
        showToast("⚠️ 正在处理中，请稍候...");
        return;
    }
    isLoadingFiles = true;

    // 释放旧的Blob URL防止内存泄漏
    releaseAllBlobUrls();

    playlist = []; lrcMap.clear();
    const audios = [];
    files.forEach(f => {
        // 排除系统隐藏文件 (Mac ._xxx, Win Thumbs.db 等)
        if (f.name.startsWith('.') || f.name.startsWith('._')) return;

        const ext = f.name.slice(f.name.lastIndexOf('.')).toLowerCase();
        if (['.mp3','.flac','.wav','.m4a','.ogg','.aac','.wma','.opus'].includes(ext)) audios.push(f);
        else if (ext === '.lrc') lrcMap.set(f.name.replace('.lrc','').toLowerCase(), f);
        else if (ext === '.cue') parseCueFile(f);
    });
    if (!audios.length) {
        isLoadingFiles = false;
        return showToast("⚠️ 未发现音频");
    }

    el.loadWrap.classList.add('show');
    el.loadBar.style.width = '0%';
    const totalCount = audios.length;

    // 使用并发批处理加速首批加载 (同时解析最多6首)
    const initLen = Math.min(20, totalCount);
    const initBatchSize = 6;
    for (let batchStart = 0; batchStart < initLen; batchStart += initBatchSize) {
        const batchEnd = Math.min(batchStart + initBatchSize, initLen);
        const batchPromises = [];
        for (let i = batchStart; i < batchEnd; i++) {
            batchPromises.push(parseMetadataWrapped(audios[i]));
        }
        try {
            const results = await Promise.all(batchPromises);
            playlist.push(...results);
        } catch(batchErr) {
            // 单个批次失败不应阻塞整体加载
            logError('BATCH_INIT', `首批批次解析失败: ${batchErr.message}`, null);
        }
        el.loadBar.style.width = `${((batchEnd) / totalCount) * 100}%`;
    }

    updateEmptyState();
    showToast(`🚀 首批 ${initLen} 首就绪，随机开播`);
    isShuffle = true; updateModeUI(); renderPlaylist(); await playAudio(Math.floor(Math.random() * playlist.length));

    if (totalCount > initLen) {
        // 后续批次用更小的并发避免卡顿
        let curr = initLen;
        
        // 🚀 曲库增量刷新防抖：每解析完一批后，最多每秒刷新一次曲库面板
        let coverLibRefreshTimer = null;
        const debouncedCoverLibRefresh = () => {
            if (coverLibRefreshTimer) clearTimeout(coverLibRefreshTimer);
            coverLibRefreshTimer = setTimeout(() => {
                // 检查用户是否打开了曲库面板
                const coverLibPanel = document.querySelector('.cover-library-panel');
                if (coverLibPanel) {
                    const grid = document.getElementById('coverLibGrid');
                    const searchEl = document.getElementById('coverLibSearch');
                    if (grid) {
                        const filter = searchEl ? searchEl.value : '';
                        if (coverLibSortMode === 'artist') {
                            renderArtistGrid(grid, filter);
                        } else if (coverLibSortMode === 'recent') {
                            renderRecentGrid(grid, filter);
                        } else {
                            renderAlbumGrid(grid, filter);
                        }
                    }
                }
                // 同时刷新封面墙（播放列表内的封面视图）
                if (currentViewMode === 'coverwall') {
                    renderCoverWall();
                }
            }, 1000);
        };
        
        const parseRemaining = async () => {
            if (curr >= totalCount) {
                el.loadBar.style.width = '100%';
                setTimeout(() => {
                    el.loadWrap.classList.remove('show');
                    
                    // 🚀 核心改动：全库与队列双向初始化
                    musicLibrary = [...playlist]; 
                    
                    renderPlaylist();
                    debouncedCoverLibRefresh(); // 最后一次完整刷新
                    showToast(`✅ 全库 ${totalCount} 首加载完毕`);
                    isLoadingFiles = false;
                }, 800);
                return;
            }
            const batchEnd = Math.min(curr + 5, totalCount);
            const batchPromises = [];
            for (let i = curr; i < batchEnd; i++) {
                batchPromises.push(parseMetadataWrapped(audios[i]));
            }
            try {
                const results = await Promise.all(batchPromises);
                playlist.push(...results);
                // 🚀 增量同步到 musicLibrary，让曲库在加载过程中就能显示
                musicLibrary = [...playlist];
            } catch(batchErr) {
                // 单批次失败不阻塞后续加载
                logError('BATCH_REMAINING', `剩余批次解析失败: ${batchErr.message}`, null);
            }
            curr = batchEnd;
            el.loadBar.style.width = `${(curr / totalCount) * 100}%`;

            if (curr % 50 === 0) renderPlaylist();
            
            // 🚀 防抖刷新曲库面板（如果用户正开着看）
            debouncedCoverLibRefresh();

            // 使用setTimeout让出主线程，避免卡顿
            setTimeout(() => parseRemaining(), 50);
        };
        setTimeout(() => parseRemaining(), 100);
    } else {
        setTimeout(() => el.loadWrap.classList.remove('show'), 500);
        musicLibrary = [...playlist];
        isLoadingFiles = false;
    }
}

// 释放所有Blob URL，防止内存泄漏
let loadedUrls = [];
function releaseAllBlobUrls() {
    loadedUrls.forEach(url => {
        try { URL.revokeObjectURL(url); } catch(e) {}
    });
    loadedUrls = [];
}

// 包装parseMetadata以追踪URL + 超时熔断
const _originalParseMetadata = parseMetadata;
const parseMetadataWrapped = async function(file) {
    try {
        // 1.5秒超时熔断：防止损坏文件导致jsmediatags永久pending
        const result = await Promise.race([
            _originalParseMetadata(file),
            new Promise((resolve) => setTimeout(() => {
                console.warn(`文件解析超时，降级处理: ${file.name}`);
                const fallbackMeta = {
                    title: file.name.replace(/\.[^/.]+$/, ""),
                    artist: "未知",
                    album: "",
                    url: URL.createObjectURL(file),
                    file: file,
                    error: false,
                    lrcText: null
                };
                resolve(fallbackMeta);
            }, 1500))
        ]);
        if (result && result.url && !loadedUrls.includes(result.url)) {
            loadedUrls.push(result.url);
        }
        return result;
    } catch(e) {
        logError('PARSE_META', `解析失败: ${e.message}`, file);
        return {
            title: file.name.replace(/\.[^/.]+$/, ""),
            artist: "未知",
            album: "",
            url: URL.createObjectURL(file),
            file: file,
            error: true,
            lrcText: null
        };
    }
};
// 替换全局引用 - processFiles中直接使用parseMetadataWrapped
// 因为parseMetadata是const无法重新赋值，我们在processFiles调用处改为parseMetadataWrapped

el.folderIn.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    await processFiles(files);
});

// === CUE 分轨支持 ===
async function parseCueFile(cueFile) {
    try {
        const text = await new Promise(resolve => {
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.readAsText(cueFile, 'utf-8');
        });
        const lines = text.split(/\r?\n/);
        let currentFile = null;
        let currentTitle = null;
        let currentArtist = null;
        const tracks = [];

        for (let line of lines) {
            line = line.trim();
            if (line.startsWith('FILE ')) {
                const match = line.match(/FILE\s+"(.+)"\s+(\w+)/i);
                if (match) currentFile = match[1];
            }
            if (line.startsWith('TITLE ') && currentFile) {
                const match = line.match(/TITLE\s+"(.+)"/i);
                if (match) currentTitle = match[1];
            }
            if (line.startsWith('PERFORMER ') && currentFile) {
                const match = line.match(/PERFORMER\s+"(.+)"/i);
                if (match) currentArtist = match[1];
            }
            if (line.startsWith('INDEX 01 ')) {
                if (currentFile && currentTitle) {
                    const timeMatch = line.match(/INDEX 01 (\d+):(\d+):(\d+)/);
                    if (timeMatch) {
                        tracks.push({
                            file: currentFile,
                            title: currentTitle,
                            artist: currentArtist || '未知',
                            startTime: parseInt(timeMatch[1])*60 + parseInt(timeMatch[2]) + parseInt(timeMatch[3])/75
                        });
                    }
                }
                currentTitle = null;
            }
        }
        if (tracks.length > 0) {
            if (!window._cueTracks) window._cueTracks = {};
            window._cueTracks[cueFile.name] = tracks;
            showToast(`📑 CUE 分轨: ${tracks.length} 个曲目已解析`, "🎯");
        }
    } catch(e) {
        logError('CUE_PARSE', e.message, cueFile);
    }
}

// === 歌词引擎 ===
function parseLyricText(text) {
    const result = [];
    text.split(/\r?\n/).forEach(line => {
        const times = line.match(/\[\d{2}:\d{2}(\.\d{2,3})?\]/g);
        if (times) {
            const txt = decodeText(line.replace(/\[.*?\]/g, '').trim());
            if (txt) {
                times.forEach(t => {
                    const match = t.match(/\[(\d{2}):(\d{2})(?:\.(\d{2,3}))?\]/);
                    if (match) {
                        const ms = match[3] ? parseInt(match[3].padEnd(3,'0')) : 0;
                        result.push({ time: parseInt(match[1])*60 + parseInt(match[2]) + ms/1000, text: txt });
                    }
                });
            }
        }
    });
    result.sort((a,b) => a.time - b.time);
    return result;
}

const loadLrc = async (song) => {
    parsedLyrics = []; el.lrcView.innerHTML = ''; el.immCurrLine.textContent = ''; el.immNextLine.textContent = '';
    let lrcText = null;

    // 优先内嵌歌词
    if (song.lrcText) {
        lrcText = song.lrcText;
    } else {
        // 其次同名LRC文件
        const base = song.file.name.replace(/\.[^/.]+$/, "").toLowerCase(); let file = lrcMap.get(base);
        if(!file) for (let [k, v] of lrcMap.entries()) if (k.includes(base) || base.includes(k)) { file = v; break; }
        if (file) {
            lrcText = await new Promise(resolve => {
                const r = new FileReader();
                r.onload = e => {
                    try { resolve(new TextDecoder('utf-8', {fatal: true}).decode(e.target.result)); }
                    catch {
                        const r2 = new FileReader();
                        r2.onload = ev => resolve(ev.target.result);
                        r2.readAsText(file, 'gbk');
                    }
                };
                r.readAsArrayBuffer(file);
            });
        }
    }

    if (!lrcText) {
        el.lrcPanel.style.display = 'none'; el.btnToggleLrc.classList.remove('active');
        el.immLrcCenter.classList.add('hidden');
        return;
    }

    parsedLyrics = parseLyricText(lrcText);

    if(parsedLyrics.length) {
        el.lrcPanel.style.display = 'flex'; el.btnToggleLrc.classList.add('active');
        el.immLrcCenter.classList.remove('hidden');
        parsedLyrics.forEach((l) => {
            const d = document.createElement('div'); d.className = 'lrc-line'; d.textContent = l.text;
            d.onclick = () => { audio.currentTime = l.time + lyricsOffset; syncLyrics(true); };
            el.lrcView.appendChild(d);
        });
    } else {
        el.lrcPanel.style.display = 'none'; el.btnToggleLrc.classList.remove('active');
        el.immLrcCenter.classList.add('hidden');
    }
};

// 歌词偏移调整
function adjustLyricsOffset(delta) {
    lyricsOffset += delta;
    lyricsOffset = Math.round(lyricsOffset * 100) / 100; // 保留2位小数
    saveSettings();
    showToast(`⏱ 歌词偏移: ${lyricsOffset > 0 ? '+' : ''}${lyricsOffset.toFixed(1)}秒`);
    syncLyrics(true);
}

const handleUserScroll = () => { isUserScrollingLyrics = true; clearTimeout(lyricsScrollTimeout); lyricsScrollTimeout = setTimeout(() => { isUserScrollingLyrics = false; syncLyrics(); }, 2000); };
el.lrcView.addEventListener('wheel', handleUserScroll, {passive: true}); el.lrcView.addEventListener('touchmove', handleUserScroll, {passive: true});

const syncLyrics = (force = false) => {
    if(!parsedLyrics.length) return;
    const cur = audio.currentTime - lyricsOffset;
    let activeIdx = -1;
    for (let i = 0; i < parsedLyrics.length; i++) { if (cur >= parsedLyrics[i].time - 0.2) activeIdx = i; else break; }

    if (el.lrcPanel.style.display !== 'none') {
        const lines = el.lrcView.querySelectorAll('.lrc-line');
        lines.forEach((line, i) => {
            if (i === activeIdx) {
                if (!line.classList.contains('active')) {
                    line.classList.add('active');
                    if (!isUserScrollingLyrics || force) {
                        const offset = line.offsetTop - el.lrcView.offsetTop - (el.lrcView.clientHeight / 2) + (line.clientHeight / 2);
                        el.lrcView.scrollTo({ top: offset, behavior: 'smooth' });
                    }
                }
            } else line.classList.remove('active');
        });
    }

    if (activeIdx !== -1 && !el.immLrcCenter.classList.contains('hidden')) {
        const curTxt = parsedLyrics[activeIdx].text;
        const nextTxt = activeIdx+1 < parsedLyrics.length ? parsedLyrics[activeIdx+1].text : '';
        if(el.immCurrLine.textContent !== curTxt) { el.immCurrLine.style.opacity=0; setTimeout(()=>{ el.immCurrLine.textContent=curTxt; el.immCurrLine.style.opacity=1; }, 200); }
        if(el.immNextLine.textContent !== nextTxt) { el.immNextLine.style.opacity=0; setTimeout(()=>{ el.immNextLine.textContent=nextTxt; el.immNextLine.style.opacity=1; }, 200); }
    }
};

// === 进度条悬停歌词预览 ===
function getLyricAtTime(time) {
    if (!parsedLyrics.length) return null;
    const adjustedTime = time - lyricsOffset;
    let best = null;
    for (let i = 0; i < parsedLyrics.length; i++) {
        if (parsedLyrics[i].time <= adjustedTime) best = parsedLyrics[i];
        else break;
    }
    return best;
}

function setupProgressHover(progArea, progTip) {
    progArea.addEventListener('mousemove', (e) => {
        const rect = progArea.getBoundingClientRect();
        const pct = (e.clientX - rect.left) / rect.width;
        const time = pct * (audio.duration || 0);
        const lyric = getLyricAtTime(time);
        if (lyric) {
            progTip.textContent = `🎤 ${lyric.text} [${formatTime(time)}]`;
            progTip.classList.add('show');
        } else if (audio.duration) {
            progTip.textContent = formatTime(time);
            progTip.classList.add('show');
        }
    });
    progArea.addEventListener('mouseleave', () => {
        progTip.classList.remove('show');
    });
}
setupProgressHover(el.progAreaMain, el.progTipMain);
setupProgressHover(el.immProgArea, el.immProgTip);

// === 专辑封面滑动切歌 ===
let swipeStartX = 0, swipeStartY = 0;
el.artBox.addEventListener('touchstart', (e) => {
    swipeStartX = e.touches[0].clientX;
    swipeStartY = e.touches[0].clientY;
}, { passive: true });

el.artBox.addEventListener('touchend', (e) => {
    const dx = e.changedTouches[0].clientX - swipeStartX;
    const dy = e.changedTouches[0].clientY - swipeStartY;
    if (Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 50) {
        el.artBox.classList.add(dx > 0 ? 'swipe-right' : 'swipe-left');
        setTimeout(() => el.artBox.classList.remove('swipe-right', 'swipe-left'), 400);
        if (dx > 0) goPrev();
        else goNext();
    }
});

// === 均衡器引擎 ===
function initEQ() {
    if (!audioCtx) initVis();
    if (eqFilters.length > 0) return; // 已初始化

    // 断开原有连接
    try { source.disconnect(); } catch(e) {}

    eqFilters = [];
    let prevNode = source;

    for (let i = 0; i < 10; i++) {
        const filter = audioCtx.createBiquadFilter();
        filter.type = 'peaking';
        filter.frequency.value = eqBands[i];
        filter.Q.value = 1.0;
        filter.gain.value = eqGains[i];
        prevNode.connect(filter);
        prevNode = filter;
        eqFilters.push(filter);
    }
    prevNode.connect(analyser);
    analyser.connect(audioCtx.destination);
}

function setEQBand(bandIdx, gainDb) {
    eqGains[bandIdx] = gainDb;
    if (eqFilters[bandIdx]) {
        eqFilters[bandIdx].gain.value = gainDb;
    }
    saveSettings();
}

function setEQPreset(preset) {
    const presets = {
        'flat': [0,0,0,0,0,0,0,0,0,0],
        'pop': [3,2,1,0,-1,-1,0,1,2,3],
        'rock': [4,3,2,0,-2,-1,1,2,3,4],
        'classical': [4,3,1,0,-1,-2,0,1,2,3],
        'vocal': [-2,-1,0,2,3,2,1,0,-1,-2],
        'bass': [6,5,3,1,0,-1,-2,-1,0,1],
        'electronic': [5,3,0,-2,-3,-1,2,4,5,4],
        'jazz': [3,2,1,0,1,1,0,-1,-1,0]
    };
    const gains = presets[preset] || presets['flat'];
    eqGains = gains;
    eqFilters.forEach((f, i) => {
        if (f) f.gain.value = gains[i];
    });
    // 更新UI
    for (let i = 0; i < 10; i++) {
        const slider = document.getElementById(`eq-band-${i}`);
        const val = document.getElementById(`eq-val-${i}`);
        if (slider) slider.value = gains[i];
        if (val) val.textContent = `${gains[i] > 0 ? '+' : ''}${gains[i]}dB`;
    }
    saveSettings();
    showToast(`🎛 均衡器: ${preset}`);
}

// === 播放速度/升降调控制 ===
function setPlaybackRate(rate) {
    playbackRate = rate;
    audio.playbackRate = rate;
    audio.preservesPitch = preservesPitch;
    const el = document.getElementById('speedVal');
    if (el) el.textContent = `${rate.toFixed(2)}x`;
    saveSettings();
}

function togglePitchPreserve() {
    preservesPitch = !preservesPitch;
    audio.preservesPitch = preservesPitch;
    const btn = document.getElementById('btnTogglePitch');
    if (btn) btn.textContent = preservesPitch ? '🔒 保持音调' : '🎵 允许变调';
    saveSettings();
    showToast(preservesPitch ? '已锁定音调' : '已允许升降调');
}

// === 终极双向锁定淡入淡出引擎 ===
function setupCrossfade() {
    audio.addEventListener('timeupdate', () => {
        // 如果未开启、单曲循环或列表少于2首，不触发
        if (!crossfadeEnabled || isRepeatOne || playlist.length < 2) return;
        
        const remaining = audio.duration - audio.currentTime;
        
        // 只有当进入切歌临界区，且当前 [没有] 处于淡入淡出状态时，才触发一次
        if (remaining <= crossfadeDuration && !isFading && remaining > 0.5) {
            isFading = true; // 立刻上锁
            triggerFadeOut();
        }
    });
}

function triggerFadeOut() {
    const userVolume = parseFloat(el.volSlider.value); // 获取用户设定的音量
    const step = 0.05; // 每次音量递减的幅度
    
    // 动态计算定时器的时间间隔，确保在指定的 crossfadeDuration 内刚好淡出到 0
    const stepsCount = userVolume / step;
    const intervalTime = stepsCount > 0 ? (crossfadeDuration * 1000) / stepsCount : 100;
    
    const fadeOutInterval = setInterval(() => {
        if (audio.volume > step) {
            audio.volume = Math.max(0, audio.volume - step);
        } else {
            clearInterval(fadeOutInterval);
            audio.volume = 0;
            
            // 自动切歌
            goNext();
            
            // 触发新歌淡入
            triggerFadeIn(userVolume);
        }
    }, intervalTime);
}

function triggerFadeIn(targetVolume) {
    audio.volume = 0;
    const step = 0.05;
    const fadeInDuration = 1.5; // 淡入固定为 1.5 秒，听感最自然
    const stepsCount = targetVolume / step;
    const intervalTime = stepsCount > 0 ? (fadeInDuration * 1000) / stepsCount : 100;
    
    let fadeInInterval = null;
    
    // 🚀 核心保护：必须等音频真正开始播放（playing 事件）后才启动淡入定时器
    // 防止歌曲因网络/磁盘缓冲延迟导致提前淡入完成，随后爆音
    const onPlaying = () => {
        audio.removeEventListener('playing', onPlaying);
        
        fadeInInterval = setInterval(() => {
            if (audio.volume < targetVolume - step) {
                audio.volume = Math.min(targetVolume, audio.volume + step);
            } else {
                clearInterval(fadeInInterval);
                audio.volume = targetVolume; // 确保音量完全恢复
                isFading = false; // 彻底解开状态锁，迎接下一首
            }
        }, intervalTime);
    };
    
    // 如果已经处于 playing 状态（如手动切歌后被重置），直接启动
    if (!audio.paused && audio.currentTime > 0 && audio.readyState >= 2) {
        onPlaying();
    } else {
        audio.addEventListener('playing', onPlaying, { once: false });
        // 🚀 兜底保护：如果 5 秒内还没 playing，强制启动淡入防止永久静音
        setTimeout(() => {
            audio.removeEventListener('playing', onPlaying);
            if (isFading && audio.volume === 0) {
                audio.volume = targetVolume;
                isFading = false;
            }
        }, 5000);
    }
}

// === 播放控制 ===
const playAudio = async (idx) => {
    if (!playlist[idx]) return;
    
    // 🚀 核心修复：手动切歌时，必须强制打断并释放所有正在运行的淡入淡出状态
    isFading = false;
    audio.volume = parseFloat(el.volSlider.value); // 立即恢复为用户设定的标准音量
    
    if (currentIndex !== idx) { playHistory.push(idx); currentIndex = idx; }
    const song = playlist[idx];
    audio.src = song.url;

    // 同步信息到双界面
    el.mainTitle.textContent = el.immTitle.textContent = song.title;
    el.mainArtist.textContent = el.immArtist.textContent = song.artist;
    document.title = `${song.title} - ${song.artist}`;
    el.fileInfo.innerHTML = `📄 ${song.file.name} <span style="opacity:0.5; margin-left:10px;">(${(song.file.size/1048576).toFixed(2)} MB)</span>`;

    hasCurrentAlbumArt = !!song.art;
    if (hasCurrentAlbumArt) {
        el.mainArt.src = el.immArt.src = song.art;
        currentAlbumColor = await extractColor(song.art);
        // 设置专辑环境光阴影CSS变量
        if (currentAlbumColor) {
            document.documentElement.style.setProperty('--album-color', currentAlbumColor + '80');
        }
        el.mainColAlbum.classList.remove('no-art'); el.immTrackCard.classList.remove('no-art');
    } else {
        el.mainArt.src = el.immArt.src = "";
        currentAlbumColor = null;
        document.documentElement.style.setProperty('--album-color', 'rgba(0,0,0,0.5)');
        el.mainColAlbum.classList.add('no-art'); el.immTrackCard.classList.add('no-art');
    }

    applyThemeLogic(); await loadLrc(song); renderPlaylist();
    recordPlay(song);
    // 更新首页收藏按钮状态
    updateFavQuickBtn();
    // 更新画中画按钮状态
    updatePipQuickBtn();

    // 应用播放速度
    audio.playbackRate = playbackRate;
    audio.preservesPitch = preservesPitch;

    if ('mediaSession' in navigator) {
        navigator.mediaSession.metadata = new MediaMetadata({
            title: song.title, artist: song.artist,
            artwork: song.art ? [{ src: song.art, sizes: '512x512', type: 'image/jpeg' }] : []
        });
        navigator.mediaSession.setActionHandler('play', togglePlay);
        navigator.mediaSession.setActionHandler('pause', togglePlay);
        navigator.mediaSession.setActionHandler('previoustrack', goPrev);
        navigator.mediaSession.setActionHandler('nexttrack', goNext);
        navigator.mediaSession.setActionHandler('seekto', (d) => {
            if (d.seekTime) audio.currentTime = d.seekTime;
        });
        navigator.mediaSession.setActionHandler('seekbackward', (d) => {
            audio.currentTime = Math.max(0, audio.currentTime - (d.seekOffset || 10));
        });
        navigator.mediaSession.setActionHandler('seekforward', (d) => {
            audio.currentTime = Math.min(audio.duration, audio.currentTime + (d.seekOffset || 10));
        });
        // PositionState 实时同步
        if ('setPositionState' in navigator.mediaSession) {
            navigator.mediaSession.setPositionState({
                duration: audio.duration || 0,
                playbackRate: playbackRate,
                position: audio.currentTime || 0
            });
        }
    }

    try {
        await audio.play();
        setPlayState(true);
        if(!audioCtx) { initVis(); initEQ(); }
    } catch(e) {
        setPlayState(false);
        showToast("❌ 播放受阻");
    }
};

const setPlayState = (playing) => {
    isPlaying = playing;
    el.btnPlay.textContent = el.immBtnPlay.textContent = playing ? '⏸' : '▶';
    if(playing && !audioCtx) { initVis(); initEQ(); }
};

const togglePlay = () => {
    if (!playlist.length) return el.btnLoad.click();
    if (isPlaying) audio.pause();
    else audio.play();
    setPlayState(!isPlaying);
    createRipple(window.innerWidth/2, window.innerHeight/2);
};

const goNext = () => {
    if(!playlist.length) return;
    if (isRepeatOne) {
        audio.currentTime = 0;
        audio.play();
        setPlayState(true);
        return;
    }
    playAudio(isShuffle ? Math.floor(Math.random()*playlist.length) : (currentIndex + 1) % playlist.length);
    createExplosion(window.innerWidth*0.8, window.innerHeight/2, 2);
};

const goPrev = () => {
    if(!playlist.length) return;
    if (isRepeatOne) {
        audio.currentTime = 0;
        audio.play();
        setPlayState(true);
        return;
    }
    if(isShuffle) {
        playHistory.pop();
        playAudio(playHistory.length ? playHistory.pop() : Math.floor(Math.random()*playlist.length));
    } else playAudio((currentIndex - 1 + playlist.length) % playlist.length);
    createExplosion(window.innerWidth*0.2, window.innerHeight/2, 2);
};

audio.addEventListener('error', async (e) => {
    const song = playlist[currentIndex];
    if (song) {
        song.error = true;
        await logError('PLAY_ERROR', `解码失败: ${audio.error ? audio.error.code : 'unknown'}`, song.file);
        renderPlaylist();
        showToast(`❌ 解码失败: ${song.title}，自动跳过`, "⚠️");
    }
    setTimeout(() => goNext(), 500);
});

// A-B 重复模式
let abLongPressTimer = null;
function startABMode() {
    abMode = true; abPointA = null; abPointB = null;
    el.btnPlay.classList.add('ab-active');
    el.immBtnPlay.classList.add('ab-active');
    hideABMarkers();
    showToast("🔁 A-B重复模式: 请先设置A点 (点击进度条)", "🎯");
}
function cancelABMode() {
    abMode = false; abPointA = null; abPointB = null;
    el.btnPlay.classList.remove('ab-active');
    el.immBtnPlay.classList.remove('ab-active');
    hideABMarkers();
    showToast("A-B重复模式已取消");
}

// 更新AB标记点位置
function updateABMarkers() {
    const duration = audio.duration;
    if (!duration) return;

    const markersA = [document.getElementById('main-abMarkerA'), document.getElementById('imm-abMarkerA')];
    const markersB = [document.getElementById('main-abMarkerB'), document.getElementById('imm-abMarkerB')];
    const ranges = [document.getElementById('main-abRange'), document.getElementById('imm-abRange')];

    if (abPointA !== null) {
        const aPct = (abPointA / duration) * 100;
        markersA.forEach(m => {
            if (m) { m.style.display = 'block'; m.style.left = aPct + '%'; }
        });
    }
    if (abPointB !== null) {
        const bPct = (abPointB / duration) * 100;
        markersB.forEach(m => {
            if (m) { m.style.display = 'block'; m.style.left = bPct + '%'; }
        });
        // 显示AB范围
        ranges.forEach(r => {
            if (r) {
                r.style.display = 'block';
                const aPct = (abPointA / duration) * 100;
                r.style.left = aPct + '%';
                r.style.width = (bPct - aPct) + '%';
            }
        });
    }
}

function hideABMarkers() {
    document.querySelectorAll('.ab-marker, .ab-range').forEach(el => el.style.display = 'none');
}

function setupABLongPress(btn) {
    btn.addEventListener('mousedown', () => {
        abLongPressTimer = setTimeout(() => {
            if (!abMode) startABMode();
            else cancelABMode();
        }, 800);
    });
    btn.addEventListener('mouseup', () => { clearTimeout(abLongPressTimer); });
    btn.addEventListener('mouseleave', () => { clearTimeout(abLongPressTimer); });
    btn.addEventListener('touchstart', (e) => {
        abLongPressTimer = setTimeout(() => {
            if (!abMode) startABMode();
            else cancelABMode();
        }, 800);
    });
    btn.addEventListener('touchend', () => { clearTimeout(abLongPressTimer); });
    btn.addEventListener('touchmove', () => { clearTimeout(abLongPressTimer); });
}
setupABLongPress(el.btnPlay);
setupABLongPress(el.immBtnPlay);

function handleABSeek(e, container) {
    if (!audio.duration) return;
    const clickTime = ((e.clientX - container.getBoundingClientRect().left) / container.offsetWidth) * audio.duration;
    if (abMode) {
        if (abPointA === null) {
            abPointA = clickTime;
            updateABMarkers();
            showToast(`A点已设置: ${formatTime(abPointA)}`);
        } else if (abPointB === null) {
            abPointB = clickTime;
            if (abPointB < abPointA) [abPointA, abPointB] = [abPointB, abPointA];
            updateABMarkers();
            showToast(`B点已设置: ${formatTime(abPointB)} - A-B重复开始`);
            audio.currentTime = abPointA;
            audio.play();
        } else {
            abPointA = clickTime; abPointB = null;
            hideABMarkers();
            updateABMarkers();
            showToast(`A点重新设置: ${formatTime(abPointA)}`);
        }
        return;
    }
    audio.currentTime = clickTime;
    createExplosion(e.clientX, e.clientY, 1.5);
}

el.btnPlay.onclick = el.immBtnPlay.onclick = (e) => {
    if (abMode) return;
    togglePlay();
};
el.btnNext.onclick = el.immBtnNext.onclick = goNext;
el.btnPrev.onclick = el.immBtnPrev.onclick = goPrev;

// === 终极丝滑进度条点击与拖拽引擎 ===
let isProgressDragging = false;

function bindProgressBar(progArea, progFill, timeDisplayEl, isMain) {
    if (!progArea || !progFill) return;

    const updateVisuals = (clientX) => {
        if (!audio.duration) return 0;
        const rect = progArea.getBoundingClientRect();
        
        // 计算点击位置占进度条的比例
        let pct = (clientX - rect.left) / rect.width;
        pct = Math.max(0, Math.min(1, pct)); // 严格限制在 0 ~ 1 之间
        
        // 实时平滑更新 UI 进度条
        progFill.style.width = (pct * 100) + '%';
        
        // 拖拽时实时更新当前数字时间，体验更跟手
        if (timeDisplayEl) {
            timeDisplayEl.textContent = formatTime(pct * audio.duration);
        }
        return pct;
    };

    const handleStart = (e) => {
        // 如果开启了 A-B 循环，则拦截拖拽，走 A-B 选点逻辑
        if (abMode) { handleABSeek(e, progArea); return; } 
        
        isProgressDragging = true;
        const clientX = e.type.includes('touch') ? e.touches[0].clientX : e.clientX;
        updateVisuals(clientX);
    };

    const handleMove = (e) => {
        if (!isProgressDragging || abMode) return;
        
        // 阻止默认行为（比如防止拖拽时意外选中文字）
        if (e.cancelable) e.preventDefault(); 
        const clientX = e.type.includes('touch') ? e.touches[0].clientX : e.clientX;
        updateVisuals(clientX);
    };

    const handleEnd = (e) => {
        if (!isProgressDragging || abMode) return;
        isProgressDragging = false;
        
        const clientX = e.type.includes('touch') ? e.changedTouches[0].clientX : e.clientX;
        const pct = updateVisuals(clientX);
        
        // 松手瞬间，真正修改音频的播放进度
        audio.currentTime = pct * audio.duration;
        
        // 如果是沉浸模式，在点击位置生成特效
        if (!isMain && typeof createExplosion === 'function') {
            createExplosion(clientX, progArea.getBoundingClientRect().top + 10, 1.5);
        }
    };

    // 绑定鼠标事件 (PC 端)
    progArea.addEventListener('mousedown', handleStart);
    document.addEventListener('mousemove', handleMove);
    document.addEventListener('mouseup', handleEnd);

    // 绑定触摸事件 (移动端/平板)
    progArea.addEventListener('touchstart', handleStart, { passive: false });
    document.addEventListener('touchmove', handleMove, { passive: false });
    document.addEventListener('touchend', handleEnd);
}

// 🚀 分别为主界面和沉浸界面的进度条挂载引擎
bindProgressBar(el.progAreaMain, el.progFillMain, el.timeCur, true);
// 注意：如果沉浸模式的时间文字ID叫 immTimeCur，需要确保它存在
bindProgressBar(el.immProgArea, el.immProgFill, document.getElementById('immTimeCur'), false);


// === 必须修改 audio.ontimeupdate 以防止系统时间覆盖拖拽进度 ===
audio.ontimeupdate = () => {
    if (audio.duration) {
        // 🚀 只有当用户 [没有在拖拽] 时，才让系统自增更新进度条
        if (!isProgressDragging) {
            const pct = `${(audio.currentTime / audio.duration) * 100}%`;
            if (el.progFillMain) el.progFillMain.style.width = pct;
            if (el.immProgFill) el.immProgFill.style.width = pct;
            
            if (el.timeCur) el.timeCur.textContent = formatTime(audio.currentTime);
            const immTimeCur = document.getElementById('immTimeCur');
            if (immTimeCur) immTimeCur.textContent = formatTime(audio.currentTime);
        }
        
        syncLyrics();
        
        // A-B 重复判定
        if (abMode && abPointA !== null && abPointB !== null) {
            if (audio.currentTime >= abPointB) audio.currentTime = abPointA;
        }
        
        // 同步系统媒体中心 (Media Session)
        if ('mediaSession' in navigator && 'setPositionState' in navigator.mediaSession) {
            navigator.mediaSession.setPositionState({
                duration: audio.duration, 
                playbackRate: playbackRate, 
                position: audio.currentTime
            });
        }
    }
};

audio.onloadedmetadata = () => {
    el.timeTot.textContent = formatTime(audio.duration);
    const immTimeTot = document.getElementById('immTimeTot');
    if (immTimeTot) immTimeTot.textContent = formatTime(audio.duration);
    if (abMode) updateABMarkers();
};

audio.onended = () => {
    if (isRepeatOne) {
        audio.currentTime = 0;
        audio.play();
        return;
    }
    goNext();
};

el.volSlider.oninput = (e) => { audio.volume = e.target.value; saveSettings(); };
const adjustVolume = (delta) => { audio.volume = Math.max(0, Math.min(1, audio.volume + delta)); el.volSlider.value = audio.volume; saveSettings(); };

// === 睡眠定时器 ===
function setSleepTimer(minutes) {
    if (sleepTimer) clearTimeout(sleepTimer);
    if (minutes === 0) {
        sleepTimer = null; sleepEndTime = null;
        updateSleepTimerUI();
        showToast("🌙 睡眠定时已取消");
        return;
    }
    const ms = minutes * 60 * 1000;
    sleepEndTime = Date.now() + ms;
    sleepTimer = setTimeout(() => {
        audio.pause();
        setPlayState(false);
        sleepTimer = null;
        sleepEndTime = null;
        updateSleepTimerUI();
        showToast("🌙 睡眠定时结束，已停止播放");
    }, ms);
    updateSleepTimerUI();
    showToast(`🌙 睡眠定时: ${minutes} 分钟后停止`);
}

function updateSleepTimerUI() {
    const display = document.getElementById('sleepTimerDisplay');
    if (!display) return;
    if (sleepTimer && sleepEndTime) {
        const remaining = Math.max(0, sleepEndTime - Date.now());
        const mins = Math.floor(remaining / 60000);
        const secs = Math.floor((remaining % 60000) / 1000);
        const timeStr = mins > 0 ? `${mins}:${secs.toString().padStart(2,'0')}` : `${secs}s`;
        display.textContent = `🌙 ${timeStr}`;
        display.className = 'sleep-timer-display active';
        // 最后1分钟红色闪烁
        if (remaining <= 60000 && remaining > 0) {
            display.style.color = (Math.floor(Date.now() / 500) % 2 === 0) ? '#ff6b6b' : 'var(--primary)';
        }
    } else {
        display.textContent = '🌙 定时';
        display.className = 'sleep-timer-display';
        display.style.color = '';
    }
}
// 每秒更新
setInterval(updateSleepTimerUI, 1000);

// === 导出/导入 ===
function exportPlaylist(format = 'json') {
    if (!playlist.length) return showToast("⚠️ 播放列表为空");
    let content, filename, mime;

    if (format === 'm3u') {
        content = '#EXTM3U\n';
        playlist.forEach((s, i) => {
            content += `#EXTINF:${Math.floor(audio.duration || 0)},${s.artist} - ${s.title}\n`;
            content += `${s.file.name}\n`;
        });
        filename = 'MBolka_Playlist.m3u';
        mime = 'audio/x-mpegurl';
    } else {
        const data = playlist.map(s => ({
            title: s.title,
            artist: s.artist,
            album: s.album,
            fileName: s.file.name,
            fileSize: s.file.size,
            favorite: favorites.has(s.file.name)
        }));
        content = JSON.stringify(data, null, 2);
        filename = 'MBolka_Playlist.json';
        mime = 'application/json';
    }

    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
    showToast(`📥 已导出: ${filename}`);
}

function importPlaylist(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = JSON.parse(e.target.result);
            if (Array.isArray(data)) {
                // JSON导入 - 只能恢复元数据信息
                showToast(`📥 已导入 ${data.length} 条记录（需要重新加载音频文件）`);
            }
        } catch {
            showToast("⚠️ 导入失败：格式不正确");
        }
    };
    reader.readAsText(file);
}

// === 全文搜索 ===
function searchPlaylist(query) {
    if (!query.trim()) {
        renderPlaylist();
        return;
    }
    const q = query.toLowerCase().trim();
    document.getElementById('playlistModalTitle').textContent = `🔍 搜索: "${q}"`;
    el.coverWallContainer.style.display = 'none';
    el.plContainer.style.display = 'flex';
    el.plContainer.innerHTML = '';

    const results = playlist.filter((s, i) => {
        return s.title.toLowerCase().includes(q) ||
               s.artist.toLowerCase().includes(q) ||
               (s.album && s.album.toLowerCase().includes(q)) ||
               s.file.name.toLowerCase().includes(q);
    });

    if (!results.length) {
        el.plContainer.innerHTML = '<div style="color:var(--text-sub); text-align:center; padding:20px;">未找到匹配的歌曲</div>';
        return;
    }

    results.forEach((s) => {
        const i = playlist.indexOf(s);
        const div = document.createElement('div');
        const isFav = favorites.has(s.file.name);
        let classes = 'pl-item focusable';
        if (i === currentIndex) classes += ' active';
        div.className = classes;
        div.dataset.index = i;
        div.innerHTML = `<span class="pl-title">${s.title}</span><span style="font-size:12px;opacity:0.6;">${s.artist}</span><span class="favorite-btn ${isFav ? 'faved' : ''}" data-idx="${i}">${isFav ? '❤️' : '🤍'}</span>`;
        div.onclick = (e) => {
            if (e.target.classList.contains('favorite-btn')) { e.stopPropagation(); toggleFavorite(i); return; }
            playAudio(i); closeAllModals();
        };
        el.plContainer.appendChild(div);
    });
    updateFocusContext();
}

// === 画中画 (Document Picture-in-Picture) v2.2 重构 ===
let pipWindow = null;

async function togglePip() {
    if (pipWindow) {
        pipWindow.close();
        pipWindow = null;
        updatePipQuickBtn();
        return;
    }

    if (!('documentPictureInPicture' in window)) {
        showToast("⚠️ 浏览器不支持画中画功能");
        return;
    }

    try {
        pipWindow = await window.documentPictureInPicture.requestWindow({
            width: 400, height: 280
        });

        const song = playlist[currentIndex];
        const title = song ? song.title : 'MBolka Player';
        const artist = song ? song.artist : '';
        const artSrc = (song && song.art) ? song.art : '';
        const hasLrc = parsedLyrics.length > 0;
        const isFaved = song ? favorites.has(song.file.name) : false;

        // 1. 复制主窗口所有样式表到 PiP 窗口
        const pipHead = pipWindow.document.head;
        const pipBody = pipWindow.document.body;
        pipBody.innerHTML = ''; // 清空

        [...document.styleSheets].forEach(sheet => {
            try {
                const newStyle = pipWindow.document.createElement('style');
                const rules = [...sheet.cssRules].map(r => r.cssText).join('\n');
                newStyle.textContent = rules;
                pipHead.appendChild(newStyle);
            } catch(e) {
                // 跨域样式表忽略
                try {
                    if (sheet.href) {
                        const link = pipWindow.document.createElement('link');
                        link.rel = 'stylesheet';
                        link.href = sheet.href;
                        pipHead.appendChild(link);
                    }
                } catch(e2) {}
            }
        });

        // 额外注入PiP专用样式
        const pipExtraStyle = pipWindow.document.createElement('style');
        pipExtraStyle.textContent = `
            :root { --primary: ${cfg.defaultColor || '#9ac8e2'}; }
            body { margin: 0; padding: 0; overflow: hidden; background: #0a0a1a; }
            @media (min-aspect-ratio: 2/1) {
                .pip-container { flex-direction: row !important; }
                .pip-lyrics-wrap { flex-direction: row !important; gap: 12px !important; text-align: left !important; padding: 16px 24px !important; }
                .pip-line-current { font-size: clamp(15px, 2.5vw, 22px) !important; }
                .pip-line-next { font-size: clamp(11px, 1.5vw, 15px) !important; }
                .pip-fallback { flex-direction: row !important; }
            }
        `;
        pipHead.appendChild(pipExtraStyle);

        // 2. 构建PiP HTML结构（将歌词容器和降级容器都写死在DOM里）
        pipBody.innerHTML = `
            <div class="pip-container">
                <div class="pip-bg" id="pipBg"></div>

                <!-- 歌词容器 (默认隐藏) -->
                <div class="pip-lyrics-wrap" id="pipLyricsWrap" style="display: none;">
                    <div class="pip-line-current" id="pipCurrLine"></div>
                    <div class="pip-line-next" id="pipNextLine"></div>
                </div>
                
                <!-- 无歌词降级容器 (默认隐藏) -->
                <div class="pip-fallback" id="pipFallback" style="display: none;">
                    <div class="pip-vinyl" id="pipVinylWrap">
                        <!-- 封面图片将在JS中动态插入/替换 -->
                    </div>
                    <div class="pip-fallback-info">
                        <div class="pip-fallback-title" id="pipFallbackTitle"></div>
                        <div class="pip-fallback-artist" id="pipFallbackArtist"></div>
                    </div>
                </div>

                <div class="pip-controls-overlay">
                    <div class="pip-track-info" id="pipTrackInfo"></div>
                    <div class="pip-btn-group">
                        <button class="pip-btn" id="pipPrev" title="上一首">⏮</button>
                        <button class="pip-btn pip-play-btn" id="pipPlay" title="播放/暂停">▶</button>
                        <button class="pip-btn" id="pipNext" title="下一首">⏭</button>
                        <button class="pip-btn pip-fav-btn" id="pipFav" title="收藏">❤️</button>
                    </div>
                </div>

                <div class="pip-progress-bar">
                    <div class="pip-progress-fill" id="pipProgFill" style="width:0%"></div>
                </div>
            </div>
        `;

        // 3. 直接绑定PiP内部按钮事件到主窗口函数
        pipWindow.document.getElementById('pipPrev').onclick = () => goPrev();
        pipWindow.document.getElementById('pipPlay').onclick = () => {
            togglePlay();
            // 同步更新按钮文字
            const btn = pipWindow.document.getElementById('pipPlay');
            if (btn && !pipWindow.closed) {
                btn.textContent = isPlaying ? '⏸' : '▶';
            }
        };
        pipWindow.document.getElementById('pipNext').onclick = () => goNext();
        pipWindow.document.getElementById('pipFav').onclick = () => {
            if (currentIndex >= 0) {
                toggleFavorite(currentIndex);
                const s = playlist[currentIndex];
                const favBtn = pipWindow.document.getElementById('pipFav');
                if (favBtn && s && !pipWindow.closed) {
                    if (favorites.has(s.file.name)) {
                        favBtn.classList.add('faved');
                    } else {
                        favBtn.classList.remove('faved');
                    }
                }
            }
        };

        // 4. 更新PiP界面的核心函数
        let pipLastCurr = '', pipLastNext = '';
        const updatePipUI = () => {
            if (!pipWindow || pipWindow.closed) {
                pipWindow = null;
                updatePipQuickBtn();
                return;
            }
            try {
                const s = playlist[currentIndex];
                if (!s) return;

                const progress = audio.duration ? (audio.currentTime / audio.duration) * 100 : 0;
                
                // 1. 进度条与播放按钮
                const progFill = pipWindow.document.getElementById('pipProgFill');
                if (progFill) progFill.style.width = progress + '%';
                const playBtn = pipWindow.document.getElementById('pipPlay');
                if (playBtn) playBtn.textContent = isPlaying ? '⏸' : '▶';

                // 2. 动态更新背景和封面 (解决封面不刷新的问题)
                const bg = pipWindow.document.getElementById('pipBg');
                const vinylWrap = pipWindow.document.getElementById('pipVinylWrap');
                if (s.art) {
                    if (bg) bg.style.backgroundImage = `url('${s.art}')`;
                    if (vinylWrap && vinylWrap.innerHTML.indexOf(s.art) === -1) {
                        vinylWrap.innerHTML = `<img src="${escapeHTML(s.art)}">`;
                    }
                } else {
                    if (bg) bg.style.backgroundImage = 'none';
                    if (vinylWrap && vinylWrap.innerHTML.indexOf('🎵') === -1) {
                        vinylWrap.innerHTML = `<div style="width:100%;height:100%;background:linear-gradient(135deg,#1a1a1a,#333);display:flex;align-items:center;justify-content:center;font-size:24px;">🎵</div>`;
                    }
                }

                // 3. 判断当前是否有歌词，动态切换两套UI的显示状态 (解决有无歌词切换失效的问题)
                const hasLrc = parsedLyrics.length > 0;
                const lyricsWrap = pipWindow.document.getElementById('pipLyricsWrap');
                const fallbackWrap = pipWindow.document.getElementById('pipFallback');
                
                if (lyricsWrap) lyricsWrap.style.display = hasLrc ? 'flex' : 'none';
                if (fallbackWrap) fallbackWrap.style.display = hasLrc ? 'none' : 'flex';

                // 4. 更新文本信息
                if (hasLrc) {
                    const currEl = pipWindow.document.getElementById('pipCurrLine');
                    const nextEl = pipWindow.document.getElementById('pipNextLine');
                    const curLrcIdx = parsedLyrics.findIndex(l => l.time > audio.currentTime - lyricsOffset);
                    const currLrc = curLrcIdx > 0 ? parsedLyrics[curLrcIdx - 1].text : (parsedLyrics.length ? parsedLyrics[parsedLyrics.length-1]?.text : '');
                    const nextLrc = curLrcIdx > 0 && curLrcIdx < parsedLyrics.length ? parsedLyrics[curLrcIdx].text : '';

                    // 添加防闪烁机制
                    if (currEl && currLrc !== pipLastCurr) {
                        currEl.classList.add('fade-out');
                        setTimeout(() => {
                            if (!pipWindow || pipWindow.closed) return;
                            currEl.textContent = currLrc || '';
                            currEl.classList.remove('fade-out');
                        }, 200);
                        pipLastCurr = currLrc;
                    }
                    if (nextEl && nextLrc !== pipLastNext) {
                        nextEl.classList.add('fade-out');
                        setTimeout(() => {
                            if (!pipWindow || pipWindow.closed) return;
                            nextEl.textContent = nextLrc || '';
                            nextEl.classList.remove('fade-out');
                        }, 200);
                        pipLastNext = nextLrc;
                    }
                } else {
                    const ftEl = pipWindow.document.getElementById('pipFallbackTitle');
                    if (ftEl && ftEl.textContent !== s.title) ftEl.textContent = s.title;
                    const faEl = pipWindow.document.getElementById('pipFallbackArtist');
                    if (faEl && faEl.textContent !== s.artist) faEl.textContent = s.artist;
                }

                // 控制栏信息
                const trackInfo = pipWindow.document.getElementById('pipTrackInfo');
                if (trackInfo) trackInfo.textContent = s.title + ' - ' + s.artist;
                
            } catch(e) {
                // 如果出现 DOM 异常，清理引用
                pipWindow = null;
                updatePipQuickBtn();
            }
        };

        // 5. 启动定时同步 (每500ms)
        let pipSyncInterval = setInterval(() => {
            if (!pipWindow || pipWindow.closed) {
                clearInterval(pipSyncInterval);
                pipWindow = null;
                updatePipQuickBtn();
                return;
            }
            updatePipUI();
        }, 500);

        // 初始化一次
        updatePipUI();

        pipWindow.addEventListener('pagehide', () => {
            clearInterval(pipSyncInterval);
            pipWindow = null;
            updatePipQuickBtn();
        });

        updatePipQuickBtn();
        showToast("📺 画中画已开启");

    } catch(e) {
        showToast("❌ 画中画启动失败");
        pipWindow = null;
    }
}

// HTML转义辅助函数
function escapeHTML(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// === 可视化引擎 (增强版，灰阶频谱) ===
class Particle {
    constructor(x, y, vx, vy, color, size) { this.x = x; this.y = y; this.vx = vx; this.vy = vy; this.color = color; this.size = size; this.life = 1; }
    update() { this.x += this.vx; this.y += this.vy; this.life -= 0.02; this.size *= 0.95; this.vx *= 0.95; this.vy *= 0.95; return this.life > 0; }
    draw(ctx) { ctx.save(); ctx.globalAlpha = this.life; ctx.fillStyle = this.color; ctx.beginPath(); ctx.arc(this.x, this.y, this.size, 0, Math.PI*2); ctx.fill(); ctx.restore(); }
}
class Ripple {
    constructor(x, y, color) { this.x = x; this.y = y; this.radius = 5; this.color = color; this.life = 1; }
    update() { this.radius += 10; this.life -= 0.04; return this.life > 0; }
    draw(ctx) { ctx.save(); ctx.globalAlpha = this.life; ctx.strokeStyle = this.color; ctx.lineWidth = 2; ctx.beginPath(); ctx.arc(this.x, this.y, this.radius, 0, Math.PI*2); ctx.stroke(); ctx.restore(); }
}

const createExplosion = (x, y, intensity) => {
    if(!isImmersiveMode) return;
    const count = Math.floor(10 * intensity);
    for (let i=0; i<count; i++) {
        const ang = Math.random() * Math.PI*2, spd = 1 + Math.random()*4*intensity;
        const size = 1.5 + Math.random()*4;
        const gray = 150 + Math.floor(Math.random() * 105);
        particles.push(new Particle(x, y, Math.cos(ang)*spd, Math.sin(ang)*spd, `rgb(${gray},${gray},${gray})`, size));
    }
};
const createRipple = (x, y) => {
    if(isImmersiveMode) {
        ripples.push(new Ripple(x, y, 'rgba(200,200,200,0.6)'));
        for(let i=0;i<3;i++) {
            const gray = 180 + Math.floor(Math.random() * 75);
            particles.push(new Particle(x, y, (Math.random()-0.5)*1.5, (Math.random()-0.5)*1.5, `rgb(${gray},${gray},${gray})`, 1+Math.random()*2));
        }
    }
};

document.addEventListener('mousemove', (e) => {
    mouseX = e.clientX; mouseY = e.clientY;
    if (isImmersiveMode && isPlaying && Math.random() < 0.25) {
        const gray = 160 + Math.floor(Math.random() * 95);
        particles.push(new Particle(mouseX, mouseY, (Math.random()-0.5)*1.5, (Math.random()-0.5)*1.5, `rgb(${gray},${gray},${gray})`, 1+Math.random()*3.5));
    }
});

document.addEventListener('touchmove', (e) => {
    if (isImmersiveMode && isPlaying) {
        const touch = e.touches[0];
        mouseX = touch.clientX; mouseY = touch.clientY;
        if (Math.random() < 0.4) {
            const gray = 150 + Math.floor(Math.random() * 105);
            particles.push(new Particle(mouseX, mouseY, (Math.random()-0.5)*2, (Math.random()-0.5)*2, `rgb(${gray},${gray},${gray})`, 1.5+Math.random()*4));
        }
    }
}, { passive: true });

document.addEventListener('click', (e) => {
    if (isImmersiveMode) {
        const ct = e.target && e.target.closest ? e.target : null;
        if (!ct || (!ct.closest('button') && !ct.closest('.progress-area'))) {
            createRipple(e.clientX, e.clientY);
        }
    }
});

const initVis = () => {
    try {
        audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        analyser = audioCtx.createAnalyser(); analyser.fftSize = 256;
        source = audioCtx.createMediaElementSource(audio);
        source.connect(analyser); analyser.connect(audioCtx.destination);
        dataArray = new Uint8Array(analyser.frequencyBinCount);
        spectrumCtxMain = el.canvasMain.getContext('2d');
        renderVisLoop();
    } catch(e) {}
};

let lastFrameTime = 0;
const renderVisLoop = (timestamp) => {
    requestAnimationFrame(renderVisLoop);

    // FPS 监测
    fpsFrames++;
    if (timestamp - fpsLastTime >= 1000) {
        currentFPS = Math.round(fpsFrames / ((timestamp - fpsLastTime) / 1000));
        fpsFrames = 0;
        fpsLastTime = timestamp;

        // 自适应粒子密度：FPS低于30时减少粒子
        if (currentFPS < 30 && particleCount > 30) {
            particleCount = Math.max(30, particleCount - 10);
        } else if (currentFPS > 55 && particleCount < MAX_PARTICLES) {
            particleCount = Math.min(MAX_PARTICLES, particleCount + 5);
        }
    }

    // 性能模式帧率控制
    const frameInterval = performanceMode ? 1000 / 30 : 1000 / targetFPS;
    if (timestamp - lastFrameTime < frameInterval) return;
    lastFrameTime = timestamp;

    if (!analyser) return;
    analyser.getByteFrequencyData(dataArray);

    if (isImmersiveMode && !immCanvasCleared) {
        const cvs = el.canvasImm, ctx = cvs.getContext('2d');
        if (cvs.width !== window.innerWidth || cvs.height !== window.innerHeight) {
            cvs.width = window.innerWidth; cvs.height = window.innerHeight;
            const cols = Math.ceil(cvs.width / 20);
            const rows = Math.ceil(cvs.height / 20);
            flowField = new Array(cols * rows).fill(0);
        }
        const W = cvs.width, H = cvs.height;
        ctx.clearRect(0, 0, W, H);
        visTime += 0.008;

        if (cfg.colorMode) {
            let diff = targetHue - currentHue;
            if (diff > 180) diff -= 360;
            else if (diff < -180) diff += 360;
            currentHue += diff * 0.04;
        } else {
            currentHue = (currentHue + 0.15) % 360;
        }
        if (currentHue < 0) currentHue += 360;

        if (isPlaying) {
            const bassAvg = (dataArray[0] + dataArray[1] + dataArray[2] + dataArray[3]) / 4;
            const midAvg = (dataArray.slice(4, 20).reduce((a,b)=>a+b,0)) / 16;
            const highAvg = (dataArray.slice(20, 64).reduce((a,b)=>a+b,0)) / 44;
            bassHistory.push(bassAvg);
            if (bassHistory.length > 30) bassHistory.shift();
            const smoothBass = bassHistory.reduce((a,b)=>a+b,0) / bassHistory.length;

            // 灰阶背景
            const bgGrad = ctx.createRadialGradient(W*0.3, H*0.3, 0, W*0.5, H*0.5, Math.max(W,H)*0.7);
            bgGrad.addColorStop(0, `rgba(30,30,40,${0.2 + smoothBass/255*0.2})`);
            bgGrad.addColorStop(0.4, `rgba(20,20,30,${0.1 + midAvg/255*0.12})`);
            bgGrad.addColorStop(0.7, 'rgba(10,10,20,0.06)');
            bgGrad.addColorStop(1, 'rgba(3,3,10,1)');
            ctx.fillStyle = bgGrad;
            ctx.fillRect(0, 0, W, H);

            // 中心发光核心
            const coreX = W * 0.5 + Math.sin(visTime * 0.3) * W * 0.05;
            const coreY = H * 0.5 + Math.cos(visTime * 0.4) * H * 0.05;
            const coreGrad = ctx.createRadialGradient(coreX, coreY, 0, coreX, coreY, 350 + smoothBass*1.5);
            coreGrad.addColorStop(0, `rgba(180,190,210,${0.4 + smoothBass/255*0.3})`);
            coreGrad.addColorStop(0.3, 'rgba(140,150,180,0.2)');
            coreGrad.addColorStop(0.7, 'rgba(80,90,110,0.05)');
            coreGrad.addColorStop(1, 'transparent');
            ctx.fillStyle = coreGrad;
            ctx.fillRect(0, 0, W, H);

            // 底部频谱弧线
            ctx.save();
            const arcY = H * 0.82;
            const arcRadius = W * 0.45;
            ctx.beginPath();
            const totalPoints = 48;
            const ampScale = 0.5 + smoothBass/255 * 1.2;
            for (let i = 0; i <= totalPoints; i++) {
                const angle = Math.PI + (i / totalPoints) * Math.PI;
                const idx = Math.floor(i / totalPoints * 32);
                const val = dataArray[Math.min(idx, 63)] / 255;
                const r = arcRadius + val * 120 * ampScale;
                const x = W/2 + Math.cos(angle) * r;
                const y = arcY + Math.sin(angle) * r * 0.5;
                if (i === 0) ctx.moveTo(x, y);
                else ctx.lineTo(x, y);
            }
            ctx.strokeStyle = 'rgba(200,200,210,0.7)';
            ctx.lineWidth = 3;
            ctx.shadowColor = 'rgba(200,200,210,0.6)';
            ctx.shadowBlur = 20;
            ctx.stroke();
            ctx.strokeStyle = 'rgba(220,220,230,0.4)';
            ctx.lineWidth = 1;
            ctx.shadowBlur = 8;
            ctx.stroke();
            ctx.restore();

            // 两侧频谱柱 (灰阶)
            const barCount = 32;
            const barWidth = W * 0.012;
            const barGap = W * 0.004;
            const barMaxH = H * 0.25;
            const barBaseY = H * 0.88;

            for (let side = 0; side < 2; side++) {
                const startX = side === 0 ? W * 0.08 : W * 0.92 - barWidth * barCount;
                for (let i = 0; i < barCount; i++) {
                    const val = dataArray[i * 2] / 255;
                    const h = val * barMaxH * (0.6 + smoothBass/255*0.8);
                    const x = startX + i * (barWidth + barGap);
                    const y = barBaseY - h;
                    const grad = ctx.createLinearGradient(x, y, x, barBaseY);
                    const lightness = 60 + i * 1.2 + val * 25;
                    grad.addColorStop(0, `rgba(${lightness+40},${lightness+40},${lightness+50},${0.6+val*0.4})`);
                    grad.addColorStop(0.5, `rgba(${lightness},${lightness},${lightness+10},0.5)`);
                    grad.addColorStop(1, `rgba(${lightness-20},${lightness-20},${lightness-10},0.2)`);
                    ctx.fillStyle = grad;
                    ctx.beginPath();
                    ctx.roundRect(x, y, barWidth, h, [barWidth/2, barWidth/2, 0, 0]);
                    ctx.fill();
                    if (val > 0.6) {
                        ctx.fillStyle = `rgba(255,255,255,${val})`;
                        ctx.beginPath();
                        ctx.arc(x + barWidth/2, y, 3, 0, Math.PI*2);
                        ctx.fill();
                    }
                }
            }

            // 顶部细线频谱
            ctx.save();
            ctx.beginPath();
            const topY = H * 0.06;
            const topW = W * 0.7;
            const topX = W * 0.15;
            for (let i = 0; i <= 64; i++) {
                const val = dataArray[i] / 255;
                const x = topX + (i / 64) * topW;
                const y = topY - val * 40;
                if (i === 0) ctx.moveTo(x, y);
                else ctx.lineTo(x, y);
            }
            ctx.strokeStyle = 'rgba(180,180,200,0.35)';
            ctx.lineWidth = 1.5;
            ctx.shadowColor = 'rgba(180,180,200,0.2)';
            ctx.shadowBlur = 8;
            ctx.stroke();
            ctx.restore();

            // 散布光点
            for (let i = 0; i < 18; i++) {
                const dotX = W * 0.15 + i * W * 0.04 + Math.sin(visTime*1.5 + i)*10;
                const dotY = H * 0.15 + Math.cos(visTime*1.2 + i*0.7)*H*0.12;
                const dotIdx = Math.floor(i / 18 * 63);
                const dotVal = dataArray[dotIdx] / 255;
                const alpha = 0.15 + dotVal * 0.5;
                const size = 1.5 + dotVal * 4;
                const gl = 180 + i * 3;
                ctx.fillStyle = `rgba(${gl},${gl},${gl+10},${alpha})`;
                ctx.beginPath();
                ctx.arc(dotX, dotY, size, 0, Math.PI*2);
                ctx.fill();
            }

            if (smoothBass > 200 && Math.random() < 0.25) {
                const ex = W*0.2 + Math.random()*W*0.6;
                const ey = H*0.3 + Math.random()*H*0.4;
                createExplosion(ex, ey, 1 + smoothBass/255*2);
            }
        } else {
            const stillGrad = ctx.createRadialGradient(W*0.5, H*0.4, 0, W*0.5, H*0.5, Math.max(W,H)*0.6);
            stillGrad.addColorStop(0, 'rgba(40,40,50,0.12)');
            stillGrad.addColorStop(1, 'rgba(3,3,10,1)');
            ctx.fillStyle = stillGrad;
            ctx.fillRect(0, 0, W, H);
        }

        if (particles.length > particleCount) particles.splice(0, particles.length - particleCount);
        if (ripples.length > MAX_RIPPLES) ripples.shift();
        particles = particles.filter(p => { if(p.update()){ p.draw(ctx); return true; } return false; });
        ripples = ripples.filter(r => { if(r.update()){ r.draw(ctx); return true; } return false; });
    } else if (!isImmersiveMode) {
        // 主界面频谱 - 灰阶
        const cvs = el.canvasMain;
        cvs.width = cvs.offsetWidth; cvs.height = cvs.offsetHeight;
        spectrumCtxMain.clearRect(0,0, cvs.width, cvs.height);
        const w = (cvs.width / (analyser.frequencyBinCount/2)); let x = 0;
        for(let i=0; i<analyser.frequencyBinCount/2; i++) {
            const h = (dataArray[i]/255) * cvs.height;
            const lightness = 140 + (i * 1.5);
            spectrumCtxMain.fillStyle = `rgba(${lightness},${lightness},${lightness+15},0.6)`;
            spectrumCtxMain.fillRect(x, cvs.height - h, w-1.5, h); x += w;
        }
    }
};

// === 模态与 UI 控制 ===
// 剥离纯同步关闭逻辑，防止视图过渡嵌套崩溃
const _closeModalsSync = () => {
    el.settingsModal.classList.remove('open');
    el.playlistModal.classList.remove('open');
    el.fileInfoModal.classList.remove('open');
    el.helpModal.classList.remove('open');
    
    // 🚀 新增：移除静态曲库的 open 类（触发淡出/缩小动画）
    const coverLibModal = document.getElementById('coverLibraryModal');
    if (coverLibModal) coverLibModal.classList.remove('open');
    
    // 清理其他真正动态的详情面板
    const statsPanel = document.querySelector('.stats-grid');
    if (statsPanel) statsPanel.closest('.modal-overlay').remove();
    const detailPanel = document.querySelector('.album-detail-panel');
    if (detailPanel) detailPanel.closest('.modal-overlay').remove();
    updateFocusContext();
};

const closeAllModals = () => {
    _closeModalsSync(); 
};

// 安全封装视图过渡
const safeTransition = (fn) => {
    if (document.startViewTransition) {
        document.startViewTransition(() => { _closeModalsSync(); fn(); });
    } else {
        _closeModalsSync(); fn();
    }
};

// === 顶级强力绑定顶部金刚键 ===
const bindBtn = (id, fn) => {
    const el = document.getElementById(id);
    if (el) {
        // 先移除旧绑定防重复，再用监听器强绑
        el.onclick = null; 
        el.addEventListener('click', (e) => {
            e.stopPropagation(); // 阻止任何可能的冒泡拦截
            fn();
        });
    }
};

bindBtn('btnLoadFolder', () => {
    if (window.showDirectoryPicker) pickAndLoadFolder();
    else el.folderIn.click();
});

bindBtn('btnToggleLrc', () => {
    const isH = el.lrcPanel.style.display === 'none';
    el.lrcPanel.style.display = isH ? 'flex' : 'none';
    el.btnToggleLrc.classList.toggle('active', isH);
    if(isH) syncLyrics(true);
});

bindBtn('btnToggleList', () => {
    safeTransition(() => {
        el.playlistModal.classList.add('open');
        currentViewMode = 'list';
        renderPlaylist();
        const searchInput = document.getElementById('searchInput');
        if (searchInput) searchInput.value = '';
    });
});

bindBtn('btnSettings', () => {
    safeTransition(() => {
        el.settingsModal.classList.add('open');
        renderThemePresets();
        renderEQPanel();
        updateFocusContext();
    });
});

bindBtn('btnCoverLibrary', () => safeTransition(showCoverLibrary));
bindBtn('btnShowStats', () => safeTransition(showStatsPanel));
bindBtn('btnFavQuick', () => { if (currentIndex >= 0) toggleFavorite(currentIndex); });
bindBtn('btnPipQuick', togglePip);

// 绑定关闭按钮
document.getElementById('btnCloseList').onclick = closeAllModals;
document.getElementById('btnCloseFileInfo').onclick = closeAllModals;
document.getElementById('btnCloseHelp').onclick = closeAllModals;
document.getElementById('btnCloseSettings').onclick = closeAllModals;
document.querySelectorAll('.modal-overlay').forEach(m => m.onclick = (e) => { if(e.target === m) closeAllModals(); });

document.getElementById('btnShowAll').onclick = () => {
    currentViewMode = 'list';
    document.getElementById('playlistModalTitle').textContent = '播放列表';
    
    // 🚀 核心改动：如果当前队列少于媒体库（如处于专辑播放中），点击全部一键恢复全库播放
    if (playlist.length < musicLibrary.length) {
        // 1. 记住当前正在播放的歌曲的唯一标识 (用文件名)
        const currentPlayingSong = playlist[currentIndex];
        const currentFileName = currentPlayingSong ? currentPlayingSong.file.name : null;
        
        // 2. 恢复全库
        playlist = [...musicLibrary];
        
        // 3. 🚀 核心修复：在新列表里重新寻找这首歌的 Index
        if (currentFileName) {
            const newIndex = playlist.findIndex(s => s.file.name === currentFileName);
            if (newIndex !== -1) {
                currentIndex = newIndex; // 纠正 Index，状态完美对齐！
            }
        }
        
        renderPlaylist();
        showToast("📋 已恢复播放全部歌曲", "🎶");
    } else {
        renderPlaylist();
    }
};

document.getElementById('btnShowFavorites').onclick = () => {
    currentViewMode = 'list';
    document.getElementById('playlistModalTitle').textContent = '❤️ 收藏';
    el.plContainer.innerHTML = '';
    el.coverWallContainer.style.display = 'none';
    el.plContainer.style.display = 'flex';
    playlist.forEach((s, i) => {
        if (!favorites.has(s.file.name)) return;
        const div = document.createElement('div');
        div.className = `pl-item focusable ${i === currentIndex ? 'active' : ''}`;
        div.dataset.index = i;
        div.innerHTML = `<span class="pl-title">${s.title}</span><span style="font-size:12px;opacity:0.6;">${s.artist}</span><span class="favorite-btn faved" data-idx="${i}">❤️</span>`;
        div.onclick = (e) => {
            if (e.target.classList.contains('favorite-btn')) { e.stopPropagation(); toggleFavorite(i); return; }
            playAudio(i); closeAllModals();
        };
        el.plContainer.appendChild(div);
    });
    if (!el.plContainer.children.length) el.plContainer.innerHTML = '<div style="color:var(--text-sub); text-align:center; padding:20px;">还没有收藏任何歌曲</div>';
};

document.getElementById('btnShowCoverWall').onclick = () => {
    currentViewMode = 'coverwall';
    renderCoverWall();
};

document.getElementById('btnEnterImmersive').onclick = toggleImmersiveMode;
document.getElementById('btnExitImmersive').onclick = toggleImmersiveMode;

el.viewImm.addEventListener('dblclick', (e) => {
    if (e.target === el.viewImm || e.target.closest('.imm-wrapper') && !e.target.closest('.imm-bottom') && !e.target.closest('.imm-topbar') && !e.target.closest('.imm-icon-btn')) {
        toggleImmersiveMode();
    }
});

let immSwipeY = 0;
el.viewImm.addEventListener('touchstart', (e) => {
    immSwipeY = e.touches[0].clientY;
}, { passive: true });
el.viewImm.addEventListener('touchend', (e) => {
    const dy = e.changedTouches[0].clientY - immSwipeY;
    if (dy > 100 && Math.abs(dy) > Math.abs(e.changedTouches[0].clientX - immSwipeY)) {
        toggleImmersiveMode();
    }
});
el.immExitHint.onclick = toggleImmersiveMode;

function updateModeUI() {
    let icon, label;
    if (isRepeatOne) {
        icon = '🔂'; label = '单曲循环';
    } else if (isShuffle) {
        icon = '🔀'; label = '随机';
    } else {
        icon = '⇄'; label = '顺序';
    }
    el.btnMode.innerHTML = `${icon} ${label}`;
    el.immBtnMode.innerHTML = `${icon} ${label}`;
    el.btnMode.classList.toggle('active', isShuffle || isRepeatOne);
    el.immBtnMode.classList.toggle('active', isShuffle || isRepeatOne);
}
function updateSettingsUI() {
    const btn = document.getElementById('btnToggleColorMode');
    btn.textContent = cfg.colorMode ? '关闭取色模式 (Y/C)' : '开启取色模式 (Y/C)';
    btn.style.color = cfg.colorMode ? 'var(--primary)' : '';
    btn.style.borderColor = cfg.colorMode ? 'var(--primary)' : '';
}

const cyclePlayMode = () => {
    if (!isShuffle && !isRepeatOne) {
        isShuffle = true; isRepeatOne = false;
    } else if (isShuffle && !isRepeatOne) {
        isShuffle = false; isRepeatOne = true;
    } else {
        isShuffle = false; isRepeatOne = false;
    }
    updateModeUI();
    saveSettings();
    const label = isRepeatOne ? '单曲循环' : (isShuffle ? '随机播放' : '顺序播放');
    showToast(`🔄 ${label}`, isRepeatOne ? '🔂' : (isShuffle ? '🔀' : '⇄'));
};
el.btnMode.onclick = el.immBtnMode.onclick = cyclePlayMode;

const toggleColorMode = () => {
    cfg.colorMode = !cfg.colorMode;
    updateSettingsUI();
    applyThemeLogic();
    saveSettings();
    showToast(cfg.colorMode ? "已开启取色跟随" : "已关闭取色跟随", "🎨");
};
document.getElementById('btnToggleColorMode').onclick = toggleColorMode;
document.getElementById('btnToggleDarkMode').onclick = toggleDarkMode;
document.getElementById('blurSlider').oninput = function() {
    cfg.blurAmt = this.value;
    document.getElementById('blurVal').textContent = `${this.value}px`;
    applyThemeLogic(); saveSettings();
};

document.getElementById('lrcFontSizeSlider').oninput = function() {
    cfg.lrcFontSize = parseInt(this.value);
    applyLrcSettings(); saveSettings();
};
document.getElementById('lrcLineHeightSlider').oninput = function() {
    cfg.lrcLineHeight = parseFloat(this.value);
    applyLrcSettings(); saveSettings();
};
document.querySelectorAll('.lrc-align-btn').forEach(btn => {
    btn.onclick = () => {
        cfg.lrcAlign = btn.dataset.align;
        applyLrcSettings(); saveSettings();
    };
});

// 主题预设
function renderThemePresets() {
    const grid = document.getElementById('themePresetGrid');
    grid.innerHTML = '';
    themePresets.forEach((t, i) => {
        const div = document.createElement('div');
        div.className = 'theme-preset';
        div.style.background = `linear-gradient(135deg, ${t.color}, rgba(0,0,0,0.3))`;
        div.title = t.name;
        div.innerHTML = '<span class="check">✓</span>';
        if (cfg.themePreset === t.color) div.classList.add('active');
        div.onclick = () => {
            cfg.themePreset = t.color;
            document.documentElement.style.setProperty('--primary', t.color);
            applyThemeLogic();
            saveSettings();
            renderThemePresets();
            showToast(`🎨 主题: ${t.name}`);
        };
        grid.appendChild(div);
        const label = document.createElement('div');
        label.className = 'theme-label';
        label.textContent = t.name;
        grid.appendChild(label);
    });
}

// 均衡器UI渲染
function renderEQPanel() {
    const eqContainer = document.getElementById('eqPanelContainer');
    if (!eqContainer) return;
    eqContainer.innerHTML = '';

    // 预设按钮
    const presetDiv = document.createElement('div');
    presetDiv.className = 'eq-preset-btns';
    const presets = ['flat','pop','rock','classical','vocal','bass','electronic','jazz'];
    presets.forEach(p => {
        const btn = document.createElement('button');
        btn.className = 'eq-preset-btn';
        btn.textContent = p.charAt(0).toUpperCase() + p.slice(1);
        btn.onclick = () => {
            setEQPreset(p);
            document.querySelectorAll('.eq-preset-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        };
        presetDiv.appendChild(btn);
    });
    eqContainer.appendChild(presetDiv);

    const panel = document.createElement('div');
    panel.className = 'eq-panel';

    const labels = ['32Hz','64Hz','125Hz','250Hz','500Hz','1kHz','2kHz','4kHz','8kHz','16kHz'];
    for (let i = 0; i < 10; i++) {
        const band = document.createElement('div');
        band.className = 'eq-band';
        band.innerHTML = `
            <label>${labels[i]}</label>
            <input type="range" id="eq-band-${i}" min="-12" max="12" step="0.5" value="${eqGains[i]}" class="focusable">
            <span class="eq-val" id="eq-val-${i}">${eqGains[i] > 0 ? '+' : ''}${eqGains[i]}dB</span>
        `;
        panel.appendChild(band);
    }
    eqContainer.appendChild(panel);

    // 绑定滑块事件
    for (let i = 0; i < 10; i++) {
        document.getElementById(`eq-band-${i}`).oninput = function() {
            const val = parseFloat(this.value);
            document.getElementById(`eq-val-${i}`).textContent = `${val > 0 ? '+' : ''}${val}dB`;
            setEQBand(i, val);
        };
    }

    // 播放速度/音调
    const spDiv = document.createElement('div');
    spDiv.className = 'drawer-box';
    spDiv.style.marginTop = '20px';
    spDiv.innerHTML = `
        <div class="drawer-title">⏩ 播放速度与音调</div>
        <div class="speed-pitch-row">
            <label>速度</label>
            <input type="range" id="speedSlider" min="0.5" max="2.0" step="0.05" value="${playbackRate}" class="focusable">
            <span class="val" id="speedVal">${playbackRate.toFixed(2)}x</span>
        </div>
        <button class="btn-glass focusable" id="btnTogglePitch" style="width:100%;justify-content:center;margin-top:8px;">${preservesPitch ? '🔒 保持音调' : '🎵 允许变调'}</button>
    `;
    eqContainer.appendChild(spDiv);

    document.getElementById('speedSlider').oninput = function() {
        setPlaybackRate(parseFloat(this.value));
    };
    document.getElementById('btnTogglePitch').onclick = togglePitchPreserve;

    // 淡入淡出
    const cfDiv = document.createElement('div');
    cfDiv.className = 'drawer-box';
    cfDiv.style.marginTop = '20px';
    cfDiv.innerHTML = `
        <div class="drawer-title">🔄 淡入淡出切歌</div>
        <div style="display:flex;align-items:center;gap:12px;">
            <button class="btn-glass focusable" id="btnToggleCrossfade" style="flex:1;justify-content:center;">${crossfadeEnabled ? '✅ 已开启' : '⏸ 关闭'}</button>
            <span style="font-size:13px;color:var(--text-sub);">时长:</span>
            <input type="range" id="crossfadeSlider" min="1" max="8" step="0.5" value="${crossfadeDuration}" style="width:100px;">
            <span id="crossfadeVal" style="font-size:13px;color:var(--primary);">${crossfadeDuration}s</span>
        </div>
    `;
    eqContainer.appendChild(cfDiv);

    document.getElementById('btnToggleCrossfade').onclick = function() {
        crossfadeEnabled = !crossfadeEnabled;
        this.textContent = crossfadeEnabled ? '✅ 已开启' : '⏸ 关闭';
        saveSettings();
        showToast(crossfadeEnabled ? '淡入淡出已开启' : '淡入淡出已关闭');
    };
    document.getElementById('crossfadeSlider').oninput = function() {
        crossfadeDuration = parseFloat(this.value);
        document.getElementById('crossfadeVal').textContent = `${crossfadeDuration}s`;
        saveSettings();
    };

    // 性能模式
    const perfDiv = document.createElement('div');
    perfDiv.className = 'drawer-box';
    perfDiv.style.marginTop = '20px';
    perfDiv.innerHTML = `
        <div class="drawer-title">⚡ 性能模式</div>
        <button class="btn-glass focusable" id="btnTogglePerf" style="width:100%;justify-content:center;">${performanceMode ? '⚡ 节能模式 (30fps)' : '🚀 全性能模式 (60fps)'}</button>
    `;
    eqContainer.appendChild(perfDiv);
    document.getElementById('btnTogglePerf').onclick = function() {
        performanceMode = !performanceMode;
        this.textContent = performanceMode ? '⚡ 节能模式 (30fps)' : '🚀 全性能模式 (60fps)';
        saveSettings();
        showToast(performanceMode ? '已切换节能模式 (30fps)' : '已切换全性能模式 (60fps)');
    };
}

document.getElementById('btnSetBg').onclick = () => document.getElementById('bgInput').click();
document.getElementById('bgInput').onchange = (e) => {
    const f = e.target.files[0];
    if(f) {
        const r = new FileReader();
        r.onload = async (ev) => {
            cfg.customBgImg = ev.target.result;
            cfg.customBgColor = await extractColor(cfg.customBgImg);
            applyThemeLogic(); saveSettings();
            showToast("🖼️ 自定义背景应用成功");
        };
        r.readAsDataURL(f);
    }
};
document.getElementById('btnClearBg').onclick = () => {
    cfg.customBgImg = null; cfg.customBgColor = null;
    applyThemeLogic(); saveSettings();
    showToast("🗑️ 已恢复默认");
};

// === 统计面板 ===
function showStatsPanel() {
    closeAllModals();
    const modal = document.createElement('div');
    modal.className = 'modal-overlay open';
    modal.innerHTML = `
        <div class="modal-content" style="width:550px;max-height:85vh;">
            <div class="modal-header">
                <h2 style="font-size:20px;">📊 音乐统计</h2>
                <button class="btn-glass" id="btnCloseStats" style="padding:6px 12px;">关闭</button>
            </div>
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-value">${playlist.length}</div>
                    <div class="stat-label">曲库总数</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${favorites.size}</div>
                    <div class="stat-label">收藏歌曲</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${formatTime(getTotalListenTime())}</div>
                    <div class="stat-label">总听歌时长</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${Object.keys(playStats).length}</div>
                    <div class="stat-label">播放过歌曲</div>
                </div>
            </div>
            <div class="drawer-title" style="margin-top:16px;">🏆 最爱Top10</div>
            <div class="stat-top-list" id="statTopList"></div>
        </div>
    `;
    document.body.appendChild(modal);

    const topList = getTopSongs(10);
    const listEl = modal.querySelector('#statTopList');
    topList.forEach((s, i) => {
        const item = document.createElement('div');
        item.className = 'stat-top-item';
        item.innerHTML = `
            <span class="rank">${i + 1}</span>
            <span class="song-name">${s.title} - ${s.artist}</span>
            <span class="play-count">${s.count}次</span>
        `;
        listEl.appendChild(item);
    });

    modal.querySelector('#btnCloseStats').onclick = () => modal.remove();
    modal.onclick = (e) => { if (e.target === modal) modal.remove(); };
}

// === 曲库独立面板 (v2.4.0 静态重构) ===
let coverLibSortMode = 'album'; // album / artist / recent

function showCoverLibrary() {
    closeAllModals();
    
    const modal = document.getElementById('coverLibraryModal');
    if (!modal) return;
    
    // 🚀 核心：像列表弹窗一样，通过添加 open 触发完美的 CSS 渐变与弹性放大动画
    modal.classList.add('open');

    // 初始化事件监听（只在第一次打开时绑定，防止重复绑定造成的内存泄露）
    if (!modal.dataset.init) {
        modal.dataset.init = "true";

        // 标签切换
        modal.querySelectorAll('.cover-lib-tab').forEach(tab => {
            tab.onclick = () => {
                modal.querySelectorAll('.cover-lib-tab').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                coverLibSortMode = tab.dataset.mode;
                renderCoverLibGrid(modal.querySelector('#coverLibSearch').value);
            };
        });

        // 搜索框输入
        modal.querySelector('#coverLibSearch').addEventListener('input', (e) => {
            renderCoverLibGrid(e.target.value);
        });

        // 关闭按钮
        modal.querySelector('#btnCloseCoverLib').onclick = closeAllModals;

        // 导入
        modal.querySelector('#btnImportPlaylist').onclick = () => {
            const input = document.createElement('input');
            input.type = 'file'; input.accept = '.json';
            input.onchange = (e) => {
                if (e.target.files[0]) importPlaylist(e.target.files[0]);
            };
            input.click();
        };
        
        // 允许点击遮罩层关闭
        modal.onclick = (e) => { if (e.target === modal) closeAllModals(); };
    }

    // 渲染网格
    renderCoverLibGrid();
    updateFocusContext();
}

function renderCoverLibGrid(filter = '') {
    const modal = document.getElementById('coverLibraryModal');
    if (!modal) return;
    const grid = modal.querySelector('#coverLibGrid');
    if (!grid) return;
    grid.innerHTML = '';

    if (coverLibSortMode === 'artist') {
        renderArtistGridStatic(grid, filter);
    } else if (coverLibSortMode === 'recent') {
        renderRecentGridStatic(grid, filter);
    } else {
        renderAlbumGridStatic(grid, filter, modal);
    }
}

// 高性能切片渲染引擎
function renderGridChunkedStatic(grid, entries, createCardFn) {
    let i = 0;
    const chunkSize = 12;
    function renderNextChunk() {
        const end = Math.min(i + chunkSize, entries.length);
        for (; i < end; i++) {
            const card = createCardFn(entries[i], i);
            grid.appendChild(card);
        }
        if (i < entries.length) {
            requestAnimationFrame(renderNextChunk);
        }
    }
    requestAnimationFrame(renderNextChunk);
}

function renderAlbumGridStatic(grid, filter, modal) {
    const groups = new Map();
    musicLibrary.forEach((s, i) => {
        const key = s.art || '__noart__';
        if (!groups.has(key)) groups.set(key, { art: s.art || null, album: s.album || '未知专辑', artist: s.artist, songs: [], firstIdx: i });
        groups.get(key).songs.push(i);
    });
    let entries = [...groups.entries()];
    if (filter) {
        const q = filter.toLowerCase();
        entries = entries.filter(([k, g]) => g.album.toLowerCase().includes(q) || g.artist.toLowerCase().includes(q));
    }
    entries.sort((a, b) => b[1].songs.length - a[1].songs.length);

    renderGridChunkedStatic(grid, entries, ([key, group], idx) => {
        const card = createCoverCard(group, idx, 'album');
        card.onclick = () => showAlbumDetail(group, modal);
        return card;
    });
}

function renderArtistGridStatic(grid, filter) {
    const groups = new Map();
    musicLibrary.forEach((s, i) => {
        const key = s.artist;
        if (!groups.has(key)) groups.set(key, { artist: s.artist, art: s.art || null, songs: [], firstIdx: i });
        groups.get(key).songs.push(i);
    });
    let entries = [...groups.entries()];
    if (filter) {
        const q = filter.toLowerCase();
        entries = entries.filter(([k, g]) => k.toLowerCase().includes(q));
    }
    entries.sort((a, b) => b[1].songs.length - a[1].songs.length);

    renderGridChunkedStatic(grid, entries, ([key, group], idx) => {
        const card = createCoverCard(group, idx, 'artist');
        card.onclick = () => { playAudio(group.firstIdx); closeAllModals(); };
        return card;
    });
}

function renderRecentGridStatic(grid, filter) {
    const entries = musicLibrary.map((s, i) => ({
        art: s.art, album: s.album || '未知', artist: s.artist,
        title: s.title, songs: [i], firstIdx: i, isSingle: true
    }));
    let filtered = entries;
    if (filter) {
        const q = filter.toLowerCase();
        filtered = entries.filter(e => e.title.toLowerCase().includes(q) || e.artist.toLowerCase().includes(q));
    }
    filtered = filtered.slice(-50).reverse();

    renderGridChunkedStatic(grid, filtered, (group, idx) => {
        const card = document.createElement('div');
        card.className = 'cover-lib-card focusable';
        card.tabIndex = 0;
        card.innerHTML = `
            <div class="art-wrap">
                ${group.art ? `<img src="${group.art}" loading="lazy">` : '<div style="width:100%;height:100%;background:rgba(255,255,255,0.05);display:flex;align-items:center;justify-content:center;font-size:32px;">🎵</div>'}
                <div class="play-all-btn">▶</div>
            </div>
            <div class="album-name">${(group.title || '').length > 12 ? (group.title||'').slice(0,11)+'…' : (group.title||'')}</div>
            <div class="album-meta">${group.artist}</div>
        `;
        card.onclick = () => { playAudio(group.firstIdx); closeAllModals(); };
        return card;
    });
}

function createCoverCard(group, idx, type) {
    const card = document.createElement('div');
    // 🚀 核心修改：加上 focusable 类名与 tabIndex，支持手柄/键盘焦点导航
    card.className = 'cover-lib-card focusable';
    card.tabIndex = 0;
    if (type === 'artist') card.classList.add('artist-card');

    const albumName = type === 'artist' ? group.artist : (group.album || '未知专辑');
    const metaText = type === 'artist' ? `${group.songs.length}首` : `${group.artist} · ${group.songs.length}首`;

    card.innerHTML = `
        <div class="art-wrap" style="--cover-transition-name: cover-${idx};">
            ${group.art ? `<img src="${group.art}" loading="lazy">` : '<div style="width:100%;height:100%;background:rgba(255,255,255,0.05);display:flex;align-items:center;justify-content:center;font-size:36px;">🎵</div>'}
            ${group.songs.length > 1 ? `<span class="song-count-badge">${group.songs.length}</span>` : ''}
            <div class="vinyl-slip"></div>
            <div class="play-all-btn">▶</div>
        </div>
        <div class="album-name" title="${escapeHTML(albumName)}">${albumName.length > 14 ? albumName.slice(0,13)+'…' : albumName}</div>
        <div class="album-meta">${metaText}</div>
    `;
    return card;
}

// 专辑详情面板
function showAlbumDetail(group, parentModal) {
    const detailModal = document.createElement('div');
    detailModal.className = 'modal-overlay open';
    detailModal.style.zIndex = '1001';

    detailModal.innerHTML = `
        <div class="album-detail-panel">
            <div class="album-detail-header">
                <div class="album-detail-cover">
                    ${group.art ? `<img src="${group.art}">` : '<div style="width:100%;height:100%;background:rgba(255,255,255,0.05);display:flex;align-items:center;justify-content:center;font-size:64px;">🎵</div>'}
                </div>
                <div class="album-detail-info">
                    <div class="album-detail-name">${escapeHTML(group.album || '未知专辑')}</div>
                    <div class="album-detail-artist">${escapeHTML(group.artist)}</div>
                    <div class="album-detail-meta">共 ${group.songs.length} 首歌曲</div>
                    <div class="album-detail-actions">
                        <button class="btn-glass focusable" id="btnPlayAlbum" style="background:var(--primary);color:#000;border:none;">▶ 播放整张专辑</button>
                        <button class="btn-glass focusable" id="btnCloseAlbumDetail">关闭</button>
                    </div>
                </div>
            </div>
            <div class="album-detail-tracks" id="albumDetailTracks"></div>
        </div>
    `;
    document.body.appendChild(detailModal);

    // 渲染曲目列表
    const tracksEl = detailModal.querySelector('#albumDetailTracks');
    group.songs.forEach((songIdx, trackNum) => {
        const s = musicLibrary[songIdx];
        const track = document.createElement('div');
        // 🚀 核心修改：加上 focusable 和 tabIndex
        track.className = `album-detail-track focusable${songIdx === currentIndex ? ' active' : ''}`;
        track.tabIndex = 0;
        track.innerHTML = `
            <span class="track-num">${trackNum + 1}</span>
            <span class="track-title">${escapeHTML(s.title)}</span>
            <span class="track-dur">${s.artist}</span>
        `;
        track.onclick = () => {
            playAudio(songIdx);
            detailModal.remove();
            if (parentModal && parentModal.classList) parentModal.classList.remove('open');
        };
        tracksEl.appendChild(track);
    });

    detailModal.querySelector('#btnPlayAlbum').onclick = () => {
        const albumQueue = group.songs.map(idx => musicLibrary[idx]);
        playlist = albumQueue;
        currentIndex = 0;
        isShuffle = false; 
        isRepeatOne = false;
        updateModeUI(); 
        saveSettings();
        playAudio(0);
        renderPlaylist();
        detailModal.remove();
        if (parentModal && parentModal.classList) parentModal.classList.remove('open');
        showToast(`🎵 正在播放专辑: ${group.album || '未知'}`);
    };

    detailModal.querySelector('#btnCloseAlbumDetail').onclick = () => detailModal.remove();
    detailModal.onclick = (e) => { if (e.target === detailModal) detailModal.remove(); };
    
    updateFocusContext();
}

// === 全域焦点控制系统 ===
let focusableElements = [];
let currentFocusIndex = -1;
let lastNavTime = 0;

const updateFocusContext = () => {
    focusableElements.forEach(el => el.classList.remove('gamepad-focus'));
    currentFocusIndex = -1;

    // 🚀 核心：检测动态生成的专辑详情面板与静态曲库弹窗
    const albumDetailPanel = document.querySelector('.album-detail-panel');
    const coverLibModal = document.getElementById('coverLibraryModal');

    if (albumDetailPanel) {
        // 如果专辑详情打开，焦点锁定在详情内的按钮和歌曲行上
        focusableElements = Array.from(albumDetailPanel.querySelectorAll('.focusable'));
    } else if (coverLibModal && coverLibModal.classList.contains('open')) {
        // 如果曲库打开，焦点锁定在 Tab 按钮和专辑卡片上
        focusableElements = Array.from(coverLibModal.querySelectorAll('.focusable'));
    } else if (el.helpModal.classList.contains('open')) {
        focusableElements = Array.from(el.helpModal.querySelectorAll('.focusable'));
    } else if (el.fileInfoModal.classList.contains('open')) {
        focusableElements = Array.from(el.fileInfoModal.querySelectorAll('.focusable'));
    } else if (el.settingsModal.classList.contains('open')) {
        focusableElements = Array.from(el.settingsModal.querySelectorAll('.focusable'));
    } else if (el.playlistModal.classList.contains('open')) {
        const container = currentViewMode === 'coverwall' ? el.coverWallContainer : el.plContainer;
        const modalFocus = Array.from(el.playlistModal.querySelectorAll('.focusable'));
        const listFocus = Array.from(container.querySelectorAll('.focusable'));
        focusableElements = [...modalFocus, ...listFocus];
    } else if (isImmersiveMode) {
        focusableElements = Array.from(el.viewImm.querySelectorAll('.focusable'));
    } else {
        focusableElements = Array.from(el.viewMain.querySelectorAll('.focusable'));
    }
};

// 🚀 升级为智能 2D 空间导航系统
const moveFocus2D = (dir) => {
    if (focusableElements.length === 0) return;
    
    // 如果当前没有任何焦点，默认激活第一个
    if (currentFocusIndex === -1) {
        setFocus(0);
        return;
    }

    const currentEl = focusableElements[currentFocusIndex];
    const currentRect = currentEl.getBoundingClientRect();
    const curX = currentRect.left + currentRect.width / 2;
    const curY = currentRect.top + currentRect.height / 2;

    let bestTargetIdx = -1;
    let minScore = Infinity;

    focusableElements.forEach((targetEl, idx) => {
        if (idx === currentFocusIndex) return;

        const targetRect = targetEl.getBoundingClientRect();
        const tarX = targetRect.left + targetRect.width / 2;
        const tarY = targetRect.top + targetRect.height / 2;

        const dx = tarX - curX;
        const dy = tarY - curY;

        // 过滤绝对不符合方向要求的元素（比如按"右"时，过滤掉所有在左边的元素）
        let isOpposite = false;
        switch(dir) {
            case 'left': if (dx >= 0) isOpposite = true; break;
            case 'right': if (dx <= 0) isOpposite = true; break;
            case 'up': if (dy >= 0) isOpposite = true; break;
            case 'down': if (dy <= 0) isOpposite = true; break;
        }
        if (isOpposite) return;

        // 二维欧式几何加权评分：主要方向距离 + 垂直偏离惩罚
        let score = 0;
        if (dir === 'left' || dir === 'right') {
            score = Math.abs(dx) + Math.abs(dy) * 2.5;
        } else {
            score = Math.abs(dy) + Math.abs(dx) * 2.5;
        }

        if (score < minScore) {
            minScore = score;
            bestTargetIdx = idx;
        }
    });

    if (bestTargetIdx !== -1) {
        setFocus(bestTargetIdx);
    }
};

// 辅助设置聚焦函数，带滚动条视口自动跟随
function setFocus(idx) {
    if (currentFocusIndex >= 0 && focusableElements[currentFocusIndex]) {
        focusableElements[currentFocusIndex].classList.remove('gamepad-focus');
    }
    currentFocusIndex = idx;
    const target = focusableElements[currentFocusIndex];
    target.classList.add('gamepad-focus');
    
    // 🚀 让滚轮自动平滑跟随焦点，防止焦点卡到屏幕外面
    if (target.scrollIntoViewIfNeeded) {
        target.scrollIntoViewIfNeeded(false);
    } else {
        target.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
}

// 保留旧 moveFocus 以兼容可能存在的调用，但内部委托给 2D 寻路
const moveFocus = (direction) => {
    if (direction === 1) moveFocus2D('down');
    else moveFocus2D('up');
};

const activateFocus = () => {
    if (currentFocusIndex >= 0 && focusableElements[currentFocusIndex]) {
        focusableElements[currentFocusIndex].click();
    } else {
        togglePlay();
    }
};

updateFocusContext();

// === 键盘映射 ===
window.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' && e.target.type !== 'range') return;
    switch(e.key.toLowerCase()) {
        case ' ': case 'enter': e.preventDefault(); activateFocus(); break;
        
        // 🚀 WASD与方向键全面升级为2D空间导航
        case 'w': case 'arrowup': e.preventDefault(); moveFocus2D('up'); break;
        case 's': case 'arrowdown': e.preventDefault(); moveFocus2D('down'); break;
        case 'a': case 'arrowleft': e.preventDefault(); moveFocus2D('left'); break;
        case 'd': case 'arrowright': e.preventDefault(); moveFocus2D('right'); break;

        case 'j': audio.currentTime -= 10; break;
        case 'k': audio.currentTime += 10; break;
        case 'i': toggleImmersiveMode(); break;
        case 'm': case 'r': cyclePlayMode(); break;
        case 'c': case 'y': toggleColorMode(); break;
        case 'd': toggleDarkMode(); break;
        case 'l': el.btnToggleLrc.click(); break;
        case 'p':
            closeAllModals();
            el.playlistModal.classList.add('open');
            currentViewMode = 'list';
            renderPlaylist();
            break;
        case 'f': toggleFullscreen(); break;
        case 't': showStatsPanel(); break;
        case 'g': showCoverLibrary(); break;
        case 'q': togglePip(); break;
        case '?':
            e.preventDefault();
            closeAllModals();
            el.helpModal.classList.add('open');
            updateFocusContext();
            break;
        case 'escape':
            if (el.helpModal.classList.contains('open')) { el.helpModal.classList.remove('open'); }
            else if (el.fileInfoModal.classList.contains('open')) { el.fileInfoModal.classList.remove('open'); }
            else if (el.settingsModal.classList.contains('open') || el.playlistModal.classList.contains('open')) closeAllModals();
            else if (isImmersiveMode) toggleImmersiveMode();
            break;
    }
});

// === 移动端手势 ===
let gestureStartX = 0, gestureStartY = 0, gestureStartTime = 0;

document.addEventListener('touchstart', (e) => {
    if (e.touches.length === 1) {
        gestureStartX = e.touches[0].clientX;
        gestureStartY = e.touches[0].clientY;
        gestureStartTime = Date.now();
    }
}, { passive: true });

document.addEventListener('touchend', (e) => {
    const ct = e.target && e.target.closest ? e.target : null;
    if (ct && (ct.closest('button') || ct.closest('input') || ct.closest('.progress-area') || ct.closest('.modal-overlay'))) return;
    const dx = e.changedTouches[0].clientX - gestureStartX;
    const dy = e.changedTouches[0].clientY - gestureStartY;
    const dt = Date.now() - gestureStartTime;

    // 双击检测 (300ms内两次点击)
    if (Math.abs(dx) < 20 && Math.abs(dy) < 20 && dt < 300) {
        const screenW = window.innerWidth;
        if (e.changedTouches[0].clientX < screenW * 0.4) {
            // 左侧双击 - 快退10秒
            audio.currentTime = Math.max(0, audio.currentTime - 10);
            showToast("⏪ 快退 10秒");
        } else {
            // 右侧双击 - 快进10秒
            audio.currentTime = Math.min(audio.duration, audio.currentTime + 10);
            showToast("⏩ 快进 10秒");
        }
    }

    // 上下滑动调节音量 (非沉浸模式，非按钮区域)
    if (!isImmersiveMode && Math.abs(dy) > 30 && Math.abs(dy) > Math.abs(dx) && Math.abs(dx) < 50) {
        if (dy > 0) adjustVolume(-0.05);
        else adjustVolume(0.05);
    }

    // 沉浸模式左右长滑切歌
    if (isImmersiveMode && Math.abs(dx) > 80 && Math.abs(dx) > Math.abs(dy)) {
        if (dx > 0) goPrev();
        else goNext();
    }
});

// === 手柄映射 ===
window.addEventListener("gamepadconnected", () => { gamepadConnected = true; el.padStatus.innerHTML = `🎮 手柄已连接 | 使用摇杆导航`; el.padStatus.style.color = 'var(--primary)'; });
window.addEventListener("gamepaddisconnected", () => { gamepadConnected = false; el.padStatus.innerHTML = `⌨️ 等待手柄接入...`; el.padStatus.style.color = 'var(--text-sub)'; });

const pollGamepad = () => {
    if (!gamepadConnected) { requestAnimationFrame(pollGamepad); return; }
    const pad = navigator.getGamepads()[0];
    if (pad) {
        const btns = pad.buttons.map(b => b.pressed);

        if (btns[0] && !prevPadBtns[0]) activateFocus();
        if (btns[1] && !prevPadBtns[1]) {
            if (el.helpModal.classList.contains('open')) el.helpModal.classList.remove('open');
            else if (el.fileInfoModal.classList.contains('open')) el.fileInfoModal.classList.remove('open');
            else if (el.settingsModal.classList.contains('open') || el.playlistModal.classList.contains('open')) closeAllModals();
            else if (isImmersiveMode) toggleImmersiveMode();
        }
        if (btns[2] && !prevPadBtns[2]) cyclePlayMode();
        if (btns[3] && !prevPadBtns[3]) toggleColorMode();
        if (btns[4] && !prevPadBtns[4]) goPrev();
        if (btns[5] && !prevPadBtns[5]) goNext();
        if (btns[8] && !prevPadBtns[8]) toggleFullscreen();
        if (btns[9] && !prevPadBtns[9]) { closeAllModals(); el.settingsModal.classList.add('open'); renderThemePresets(); renderEQPanel(); updateFocusContext(); }

        if (pad.buttons[6].pressed) adjustVolume(-0.02);
        if (pad.buttons[7].pressed) adjustVolume(0.02);

        if (btns[12] && !prevPadBtns[12]) { moveFocus2D('up'); }
        if (btns[13] && !prevPadBtns[13]) { moveFocus2D('down'); }
        if (btns[14] && !prevPadBtns[14]) { moveFocus2D('left'); }
        if (btns[15] && !prevPadBtns[15]) { moveFocus2D('right'); }

        let stickX = pad.axes[0], stickY = pad.axes[1];
        if (Date.now() - lastNavTime > 200) {
            if (stickY < -0.5) { moveFocus2D('up'); lastNavTime = Date.now(); }
            else if (stickY > 0.5) { moveFocus2D('down'); lastNavTime = Date.now(); }
            else if (stickX < -0.5) { moveFocus2D('left'); lastNavTime = Date.now(); }
            else if (stickX > 0.5) { moveFocus2D('right'); lastNavTime = Date.now(); }
        }

        prevPadBtns = btns;
    }
    requestAnimationFrame(pollGamepad);
};
requestAnimationFrame(pollGamepad);

// === 错误边界与日志导出 ===
// 注意：logError 已在文件顶部定义，这里不再二次赋值（否则会触发 TypeError 导致应用崩溃）

// 全局错误捕获
window.addEventListener('error', (e) => {
    logError('JS_ERROR', `${e.message} at ${e.filename}:${e.lineno}`);
});

// 导出错误日志
function exportErrorLogs() {
    const logs = JSON.parse(localStorage.getItem('MBolka_ErrorLogs') || '[]');
    if (!logs.length) return showToast("📋 暂无错误日志");
    const blob = new Blob([JSON.stringify(logs, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `MBolka_ErrorLogs_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("📋 错误日志已导出");
}

// === 命名空间封装 ===
window.MBolka = {
    togglePlay, goNext, goPrev, cyclePlayMode,
    toggleImmersiveMode, toggleFullscreen, toggleDarkMode,
    togglePip, showCoverLibrary, showStatsPanel,
    searchPlaylist, exportPlaylist, importPlaylist,
    adjustLyricsOffset, setSleepTimer, exportErrorLogs,
    toggleFavorite,
    get playlist() { return playlist; },
    get currentIndex() { return currentIndex; },
    get isPlaying() { return isPlaying; },
    get favorites() { return favorites; },
};

// === 初始化 ===
window.addEventListener('load', async () => {
    await initIDB();
    loadSettings();
    updateEmptyState();
    updateDarkModeUI();
    applyLrcSettings();
    setupCrossfade();
    updateFavQuickBtn();

    // 从localStorage恢复错误日志到内存缓存
    try {
        const stored = JSON.parse(localStorage.getItem('MBolka_ErrorLogs') || '[]');
        if (stored.length) _errorLogsCache.push(...stored);
    } catch(e) {}

    // 尝试从持久化目录加载
    const loaded = await loadFromStoredDirectory();
    if (!loaded) {
        // 没有持久化目录，显示空状态
        updateEmptyState();
    }

    // 如果设置了EQ增益，初始化EQ
    if (eqGains.some(g => g !== 0) && audioCtx) {
        initEQ();
    }

    // === 统一事件绑定 (从 index.html 内联 script 迁移至此，确保加载时序一致) ===

    // 搜索功能
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => searchPlaylist(e.target.value));
    }

    // 导出按钮
    const btnExportM3U = document.getElementById('btnExportM3U');
    if (btnExportM3U) btnExportM3U.onclick = () => exportPlaylist('m3u');
    const btnExportJSON = document.getElementById('btnExportJSON');
    if (btnExportJSON) btnExportJSON.onclick = () => exportPlaylist('json');

    // 画中画按钮 (设置面板内) - 注意: btnCoverLibrary/btnShowStats/btnFavQuick/btnPipQuick 已在模态与UI控制段绑定
    const btnTogglePip = document.getElementById('btnTogglePip');
    if (btnTogglePip) btnTogglePip.onclick = togglePip;

    // 歌词偏移按钮
    const btnLrcOffsetMinus = document.getElementById('btnLrcOffsetMinus');
    if (btnLrcOffsetMinus) btnLrcOffsetMinus.onclick = () => adjustLyricsOffset(-0.5);
    const btnLrcOffsetPlus = document.getElementById('btnLrcOffsetPlus');
    if (btnLrcOffsetPlus) btnLrcOffsetPlus.onclick = () => adjustLyricsOffset(0.5);
    const btnLrcOffsetMinusFine = document.getElementById('btnLrcOffsetMinusFine');
    if (btnLrcOffsetMinusFine) btnLrcOffsetMinusFine.onclick = () => adjustLyricsOffset(-0.1);
    const btnLrcOffsetPlusFine = document.getElementById('btnLrcOffsetPlusFine');
    if (btnLrcOffsetPlusFine) btnLrcOffsetPlusFine.onclick = () => adjustLyricsOffset(0.1);

    // 导出错误日志
    const btnExportLogs = document.getElementById('btnExportLogs');
    if (btnExportLogs) btnExportLogs.onclick = exportErrorLogs;

    // 睡眠定时器按钮
    const btnSleep15 = document.getElementById('btnSleep15');
    if (btnSleep15) btnSleep15.onclick = () => setSleepTimer(15);
    const btnSleep30 = document.getElementById('btnSleep30');
    if (btnSleep30) btnSleep30.onclick = () => setSleepTimer(30);
    const btnSleep60 = document.getElementById('btnSleep60');
    if (btnSleep60) btnSleep60.onclick = () => setSleepTimer(60);
    const btnSleepCancel = document.getElementById('btnSleepCancel');
    if (btnSleepCancel) btnSleepCancel.onclick = () => setSleepTimer(0);
});
