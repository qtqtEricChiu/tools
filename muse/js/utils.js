/**
 * MBolka Player - Utilities
 * showToast, formatTime, decodeText, saveSettings, loadSettings
 */

const showToast = (msg, icon='') => { el.toast.innerHTML = `${icon} ${msg}`; el.toast.classList.add('show'); setTimeout(() => el.toast.classList.remove('show'), 2500); };
const formatTime = (sec) => { if (!sec || isNaN(sec)) return '0:00'; const m = Math.floor(sec / 60), s = Math.floor(sec % 60); return `${m}:${s.toString().padStart(2, '0')}`; };
const decodeText = (str) => { if (!str) return ''; let s = str.replace(/\\u([0-9a-fA-F]{4})/g, (m, g) => String.fromCharCode(parseInt(g, 16))); const txt = document.createElement("textarea"); txt.innerHTML = s; return txt.value; };

const saveSettings = () => {
    try {
        localStorage.setItem('MBolka_Cfg_v3', JSON.stringify({
            // 核心修复：只保存滑块的物理数值，防止保存淡入淡出时的临时"0"音量
            colorMode: cfg.colorMode, blurAmt: cfg.blurAmt, vol: parseFloat(el.volSlider.value),
            isShuffle: isShuffle, isRepeatOne: isRepeatOne,
            customBgImg: cfg.customBgImg, customBgColor: cfg.customBgColor,
            darkMode: cfg.darkMode, lrcFontSize: cfg.lrcFontSize,
            lrcLineHeight: cfg.lrcLineHeight, lrcAlign: cfg.lrcAlign,
            themePreset: cfg.themePreset, playbackRate: playbackRate,
            preservesPitch: preservesPitch, crossfadeEnabled: crossfadeEnabled,
            crossfadeDuration: crossfadeDuration, performanceMode: performanceMode,
            eqGains: eqGains, lyricsOffset: lyricsOffset,
            energySavingEnabled: cfg.energySavingEnabled
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
            cfg.energySavingEnabled = stored.energySavingEnabled ?? true;
            // v2.7: 恢复节能开关UI状态
            const esToggle = document.getElementById('energySavingToggle');
            if (esToggle) esToggle.checked = cfg.energySavingEnabled;
            // 同时初始化双端滑块
            el.volSlider.value = audio.volume;
            if (el.immVolSlider) el.immVolSlider.value = audio.volume;
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
