/**
 * MBolka Player - Ultimate Nexus v2.8.0
 * Worker, globals, config, presets, DOM refs
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
let musicLibrary = []; // 导入的完整本地音乐库（不因播放模式而缩水）
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

// v2.5: 流沙波动相位
let sandPhaseA = 0, sandPhaseB = 0, sandPhaseC = 0;

// v2.5-p2: 三角函数查表法 (LUT) — 用极小的内存换取零CPU浮点运算
// 128 点全圆查表，覆盖沉浸模式 48 点半弧、频谱 64 点环形等高频场景
const LUT_SIZE = 128;
const SIN_TABLE = [], COS_TABLE = [];
for (let i = 0; i < LUT_SIZE; i++) {
    const angle = (i / LUT_SIZE) * Math.PI * 2;
    SIN_TABLE.push(Math.sin(angle));
    COS_TABLE.push(Math.cos(angle));
}
// 便捷函数：将角度映射到查表索引
function lutSin(angle) {
    let idx = Math.round((angle % (Math.PI * 2) + Math.PI * 2) % (Math.PI * 2) / (Math.PI * 2) * LUT_SIZE) % LUT_SIZE;
    return SIN_TABLE[idx];
}
function lutCos(angle) {
    let idx = Math.round((angle % (Math.PI * 2) + Math.PI * 2) % (Math.PI * 2) / (Math.PI * 2) * LUT_SIZE) % LUT_SIZE;
    return COS_TABLE[idx];
}

// 播放引擎增强
let playbackRate = 1.0;
let preservesPitch = true;
let crossfadeEnabled = false;
let crossfadeDuration = 3; // 秒
let isFading = false; // 全局淡入淡出锁，防止 timeupdate 触发多重定时器崩溃

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
    lrcAlign: 'center', themePreset: null,
    energySavingEnabled: true // v2.7: 画中画启动时自动优化主界面性能
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
