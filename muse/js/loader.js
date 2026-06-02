/**
 * MBolka Player - File Loader
 * Metadata parsing, playlist rendering, context menu, drag/drop, CUE
 */

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

// === 核心修改：用状态机完美驱动"空态"与"播放态"的物理隔离 ===
function updateEmptyState() {
    const grid = document.querySelector('.content-grid');
    if (!grid) return;
    
    const isEmpty = playlist.length === 0;
    
    if (isEmpty) {
        grid.classList.add('is-empty');
        
        // 体验金加项：点击空状态区域，直接等同于点击"载入音乐"按钮，极其符合直觉！
        el.emptyState.onclick = () => {
            if (el.btnLoad) el.btnLoad.click();
        };
    } else {
        grid.classList.remove('is-empty');
        el.emptyState.onclick = null;
    }
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
        
        // 曲库增量刷新防抖：每解析完一批后，最多每秒刷新一次曲库面板
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
                    
                    // 核心改动：全库与队列双向初始化
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
                // 增量同步到 musicLibrary，让曲库在加载过程中就能显示
                musicLibrary = [...playlist];
            } catch(batchErr) {
                // 单批次失败不阻塞后续加载
                logError('BATCH_REMAINING', `剩余批次解析失败: ${batchErr.message}`, null);
            }
            curr = batchEnd;
            el.loadBar.style.width = `${(curr / totalCount) * 100}%`;

            if (curr % 50 === 0) renderPlaylist();
            
            // 防抖刷新曲库面板（如果用户正开着看）
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
