/**
 * MBolka Player - Gamepad & Navigation
 * Focus control, 2D navigation, keyboard mapping, mobile gestures, gamepad
 */

// === 全域焦点控制系统 ===
let focusableElements = [];
let currentFocusIndex = -1;
let lastNavTime = 0;

const updateFocusContext = () => {
    focusableElements.forEach(el => el.classList.remove('gamepad-focus'));
    currentFocusIndex = -1;

    // 核心：检测动态生成的专辑详情面板与静态曲库弹窗
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

// 升级为智能 2D 空间导航系统
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
    
    // 让滚轮自动平滑跟随焦点，防止焦点卡到屏幕外面
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
        
        // WASD与方向键全面升级为2D空间导航
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
            if (document.fullscreenElement) {
                document.exitFullscreen();
            } else {
                // 优先尝试关闭最上层弹窗，如果没有弹窗打开，才执行退出沉浸模式
                const closed = handleGlobalClose();
                if (!closed && isImmersiveMode) {
                    toggleImmersiveMode();
                }
            }
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
        // 手柄 B 键（退回键）一键适配全域弹窗关闭
        if (btns[1] && !prevPadBtns[1]) {
            const closed = handleGlobalClose();
            if (!closed && isImmersiveMode) {
                toggleImmersiveMode();
            }
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
