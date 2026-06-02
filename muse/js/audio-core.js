/**
 * MBolka Player - Audio Core
 * Lyrics engine, EQ, speed, crossfade, play control, AB repeat, progress, volume, sleep, export, search
 */

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
    // v2.5-p2: 缓存 rect，只在 mouseenter 时刷新
    let hoverRect = null;
    progArea.addEventListener('mouseenter', () => { hoverRect = progArea.getBoundingClientRect(); });
    progArea.addEventListener('mousemove', (e) => {
        const rect = hoverRect || progArea.getBoundingClientRect();
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
        hoverRect = null;
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
    
    // 核心保护：必须等音频真正开始播放（playing 事件）后才启动淡入定时器
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
        // 兜底保护：如果 5 秒内还没 playing，强制启动淡入防止永久静音
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
    
    // 核心修复：手动切歌时，必须强制打断并释放所有正在运行的淡入淡出状态
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

    // v2.5-p2: 缓存 rect 避免 mousemove 中反复触发布局重排
    let cachedRect = null;

    const updateVisuals = (clientX) => {
        if (!audio.duration) return 0;
        const rect = cachedRect || progArea.getBoundingClientRect();
        
        // 计算点击位置占进度条的比例
        let pct = (clientX - rect.left) / rect.width;
        pct = Math.max(0, Math.min(1, pct)); // 严格限制在 0 ~ 1 之间
        
        // v2.7: 进度条改用 transform:scaleX 避免布局重排
        progFill.style.transform = `scaleX(${pct})`;
        
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
        cachedRect = progArea.getBoundingClientRect(); // 按下时缓存一次
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
            const rect = cachedRect || progArea.getBoundingClientRect();
            createExplosion(clientX, rect.top + 10, 1.5);
        }
        cachedRect = null; // 释放缓存
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

// 分别为主界面和沉浸界面的进度条挂载引擎
bindProgressBar(el.progAreaMain, el.progFillMain, el.timeCur, true);
// 注意：如果沉浸模式的时间文字ID叫 immTimeCur，需要确保它存在
bindProgressBar(el.immProgArea, el.immProgFill, document.getElementById('immTimeCur'), false);


// === 必须修改 audio.ontimeupdate 以防止系统时间覆盖拖拽进度 ===
audio.ontimeupdate = () => {
    if (audio.duration) {
        // v2.7: 进度条改用 transform:scaleX 避免布局重排
        if (!isProgressDragging) {
            const pct = audio.currentTime / audio.duration;
            if (el.progFillMain) el.progFillMain.style.transform = `scaleX(${pct})`;
            if (el.immProgFill) el.immProgFill.style.transform = `scaleX(${pct})`;
            
            if (el.timeCur) el.timeCur.textContent = formatTime(audio.currentTime);
            const immTimeCur = document.getElementById('immTimeCur');
            if (immTimeCur) immTimeCur.textContent = formatTime(audio.currentTime);
        }
        
        // v2.7: 节能模式下歌词走低频定时器，跳过此高频回调
        if (!isEnergySaving) syncLyrics();
        
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
const adjustVolume = (delta) => { 
    audio.volume = Math.max(0, Math.min(1, audio.volume + delta)); 
    el.volSlider.value = audio.volume; 
    // 新增：同步手柄调音到沉浸滑块
    if (el.immVolSlider) el.immVolSlider.value = audio.volume; 
    saveSettings(); 
};

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
