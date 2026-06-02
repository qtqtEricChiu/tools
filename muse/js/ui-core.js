/**
 * MBolka Player - UI Core
 * Modals, button bindings, theme presets, EQ panel, stats
 */

// === 模态与 UI 控制 ===
// 剥离纯同步关闭逻辑，防止视图过渡嵌套崩溃
// 核心重构：引入 isSwitching 锁，完美恢复回弹动画
const _closeModalsSync = (isSwitching = false) => {
    // 寻找当前处于打开状态的弹窗
    const openModals = document.querySelectorAll('.modal-overlay.open');
    
    openModals.forEach(m => {
        m.classList.remove('open');
        
        // 只有当"切换窗口"时，为了防止两个大弹窗重叠，才让旧窗口瞬间消失（禁用 transition）
        if (isSwitching) {
            m.style.transition = 'none';
            const content = m.querySelector('.modal-content');
            if (content) content.style.transition = 'none';
            
            // 100ms 后立刻恢复 transition，绝不干扰下一次打开
            setTimeout(() => {
                m.style.transition = '';
                if (content) content.style.transition = '';
            }, 100);
        }
    });

    // 清理其他动态生成的面板
    const statsPanel = document.querySelector('.stats-grid');
    if (statsPanel) statsPanel.closest('.modal-overlay').remove();
    const detailPanel = document.querySelector('.album-detail-panel');
    if (detailPanel) detailPanel.closest('.modal-overlay').remove();
    
    updateFocusContext();
};

// 正常关闭弹窗（点击Close、Esc、手柄B）：彻底恢复原本极具动感的 CSS 淡出和回弹缩小动画！
const closeAllModals = () => {
    _closeModalsSync(false); // 传参 false：保留完整的过渡动画
};

// === 统一弹窗栈关闭管理器 (完美支持 LIFO 后进先出) ===
function handleGlobalClose() {
    // 扫描页面上所有当前处于打开状态的弹窗（包括动态创建和静态隐藏的）
    const activeModals = Array.from(document.querySelectorAll('.modal-overlay')).filter(m => {
        return m.classList.contains('open') || (m.style.display !== 'none' && document.body.contains(m));
    });

    if (activeModals.length > 0) {
        // 永远只关闭位于最上层的那个弹窗（数组的最后一项）
        const topModal = activeModals[activeModals.length - 1];
        
        // 区分静态和动态弹窗进行关闭
        const staticIds = ['playlistModal', 'settingsModal', 'fileInfoModal', 'helpModal', 'coverLibraryModal'];
        if (staticIds.includes(topModal.id)) {
            topModal.classList.remove('open');
        } else {
            topModal.remove(); // 动态生成的模态框（如统计、专辑详情）直接从DOM移除
        }
        
        updateFocusContext(); // 刷新焦点
        return true; // 成功关闭了一个弹窗
    }
    return false; // 当前没有打开的弹窗
}

// 切换窗口：旧窗口瞬间消失，新窗口优雅回弹展开
const safeTransition = (fn) => {
    _closeModalsSync(true); // 传参 true：让旧窗口瞬间消失，防止两个大弹窗叠在一起
    
    // 给浏览器 50ms 缓冲，确保旧窗口的 transition 禁用完全生效
    setTimeout(() => {
        if (document.startViewTransition) {
            document.startViewTransition(() => fn());
        } else {
            fn();
        }
    }, 50);
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
    
    // 核心改动：如果当前队列少于媒体库（如处于专辑播放中），点击全部一键恢复全库播放
    if (playlist.length < musicLibrary.length) {
        // 1. 记住当前正在播放的歌曲的唯一标识 (用文件名)
        const currentPlayingSong = playlist[currentIndex];
        const currentFileName = currentPlayingSong ? currentPlayingSong.file.name : null;
        
        // 2. 恢复全库
        playlist = [...musicLibrary];
        
        // 3. 核心修复：在新列表里重新寻找这首歌的 Index
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

// === v2.5: 预设主题色统一渲染与多场景同步引擎 ===
function renderThemePresets() {
    const grid = document.getElementById('themePresetGrid');
    if (!grid) return;
    grid.innerHTML = '';

    themePresets.forEach((t) => {
        const card = document.createElement('div');
        // 将外层卡片声明为 focusable，实现 2D 空间完美寻路
        card.className = `theme-preset-card focusable${cfg.themePreset === t.color ? ' active' : ''}`;
        card.tabIndex = 0;
        card.innerHTML = `
            <div class="theme-color-circle" style="background: linear-gradient(135deg, ${t.color}, rgba(0,0,0,0.4))"></div>
            <div class="theme-preset-label">${t.name}</div>
        `;
        card.onclick = () => {
            applyThemeColorAction(t.color, t.name);
        };
        grid.appendChild(card);
    });
    updateFocusContext();
}

// 核心行动：多场景一键联动同步
function applyThemeColorAction(color, name) {
    cfg.themePreset = color;
    cfg.defaultColor = color;
    document.documentElement.style.setProperty('--primary', color);

    // 1. 计算并设置氛围发光色 (16进制转带透明度的RGBA)
    const rgb = hexToRgb(color);
    if (rgb) {
        const glowStr = `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.5)`;
        document.documentElement.style.setProperty('--primary-glow', glowStr);
        document.documentElement.style.setProperty('--album-color', `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.4)`);
    }

    applyThemeLogic();
    saveSettings();
    renderThemePresets();

    // 2. 多场景 A：画中画悬浮窗颜色实时无缝同步
    if (pipWindow && !pipWindow.closed) {
        try {
            pipWindow.document.documentElement.style.setProperty('--primary', color);
            const pipFill = pipWindow.document.getElementById('pipProgFill');
            if (pipFill) pipFill.style.boxShadow = `0 0 8px rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.4)`;
        } catch (e) { }
    }

    if (name) showToast(`🎨 主题已应用: ${name}`);
}

// 辅助函数：16进制转RGB
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}

// v2.5-p2: WCAG 相对亮度计算 (心理学灰度公式)
function getLuminance(colorStr) {
    if (!colorStr) return 255;
    let r = 255, g = 255, b = 255;

    if (colorStr.startsWith('#')) {
        const rgb = hexToRgb(colorStr);
        if (rgb) { r = rgb.r; g = rgb.g; b = rgb.b; }
    } else if (colorStr.startsWith('rgb')) {
        const match = colorStr.match(/\d+/g);
        if (match) {
            r = parseInt(match[0]);
            g = parseInt(match[1]);
            b = parseInt(match[2]);
        }
    }
    // 心理学灰度公式: L = 0.299R + 0.587G + 0.114B
    return (0.299 * r + 0.587 * g + 0.114 * b);
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
