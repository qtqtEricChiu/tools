/**
 * MBolka Player - Picture-in-Picture & Energy Saving
 * PiP window management, energy saving state machine
 */

// === 画中画 (Document Picture-in-Picture) v2.2 重构 ===
let pipWindow = null;
let isEnergySaving = false; // v2.7: 节能模式状态机

// v2.7: 进入节能模式 — 暂停渲染、粒子、歌词降频
function enterEnergySaving() {
    if (isEnergySaving) return;
    isEnergySaving = true;
    // 清空粒子池
    particles.length = 0;
    ripples.length = 0;
    // 清空流沙背景 Canvas
    const ctx = el.bgColor.getContext('2d');
    if (ctx) ctx.clearRect(0, 0, el.bgColor.width, el.bgColor.height);
    // 降低歌词同步频率
    if (lrcTimer) clearInterval(lrcTimer);
    lrcTimer = setInterval(() => syncLyrics(true), 500);
    // CSS 休眠主界面
    const wrapper = document.querySelector('.player-wrapper');
    if (wrapper) wrapper.classList.add('pip-standby');
}

// v2.7: 退出节能模式
function exitEnergySaving() {
    if (!isEnergySaving) return;
    isEnergySaving = false;
    // 恢复歌词高频同步
    if (lrcTimer) { clearInterval(lrcTimer); lrcTimer = null; }
    // 恢复 CSS
    const wrapper = document.querySelector('.player-wrapper');
    if (wrapper) wrapper.classList.remove('pip-standby');
}

let lrcTimer = null; // v2.7: 歌词降频定时器句柄

async function togglePip() {
    if (pipWindow) {
        pipWindow.close();
        pipWindow = null;
        exitEnergySaving();
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

        // v2.7: 仅在用户启用节能开关时，PiP 激活后主窗口进入节电状态
        if (cfg.energySavingEnabled) {
            enterEnergySaving();
            showToast("⚡ 主界面已进入节能模式", "📺");
        }

        // PiP 窗口关闭监听 — 用户点 × 关闭时也要恢复主窗口
        pipWindow.addEventListener('pagehide', () => {
            pipWindow = null;
            exitEnergySaving();
            updatePipQuickBtn();
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

                // 🚀 v2.8.7: PiP歌词 — 当前原文 + 下一句原文（与沉浸舱一致）
                if (hasLrc) {
                    const currEl = pipWindow.document.getElementById('pipCurrLine');
                    const nextEl = pipWindow.document.getElementById('pipNextLine');

                    // 计算当前句和下一句索引
                    const curLrcIdx = parsedLyrics.findIndex(l => l.time > audio.currentTime - lyricsOffset);
                    const currLrc = curLrcIdx > 0 ? parsedLyrics[curLrcIdx - 1] : (parsedLyrics.length ? parsedLyrics[parsedLyrics.length-1] : null);
                    const nextLrc = (curLrcIdx > 0 && curLrcIdx < parsedLyrics.length) ? parsedLyrics[curLrcIdx] : null;

                    // 当前句原文
                    const currTxt = currLrc ? (currLrc.original || currLrc.text) : '';
                    // 下一句原文（不是翻译！）
                    const nextTxt = nextLrc ? (nextLrc.original || nextLrc.text) : '';

                    // 第一行：当前句原文
                    if (currEl && currTxt !== pipLastCurr) {
                        currEl.classList.add('fade-out');
                        setTimeout(() => {
                            if (!pipWindow || pipWindow.closed) return;
                            currEl.textContent = currTxt || '';
                            currEl.classList.remove('fade-out');
                            currEl.style.opacity = '1';
                        }, 200);
                        pipLastCurr = currTxt;
                    }
                    // 第二行：下一句原文
                    if (nextEl && nextTxt !== pipLastNext) {
                        nextEl.classList.add('fade-out');
                        setTimeout(() => {
                            if (!pipWindow || pipWindow.closed) return;
                            nextEl.textContent = nextTxt || '';
                            nextEl.classList.remove('fade-out');
                            nextEl.style.opacity = nextTxt ? '0.6' : '0';
                        }, 200);
                        pipLastNext = nextTxt;
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
