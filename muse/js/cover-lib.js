/**
 * MBolka Player - Cover Library
 * Standalone cover library panel, album detail
 */

// === 曲库独立面板 (v2.4.0 静态重构) ===
let coverLibSortMode = 'album'; // album / artist / recent

function showCoverLibrary() {
    closeAllModals();
    
    const modal = document.getElementById('coverLibraryModal');
    if (!modal) return;
    
    // 核心：像列表弹窗一样，通过添加 open 触发完美的 CSS 渐变与弹性放大动画
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
        renderArtistGrid(grid, filter);
    } else if (coverLibSortMode === 'recent') {
        renderRecentGrid(grid, filter);
    } else {
        renderAlbumGrid(grid, filter);
    }
}

// 核心：增量切片渲染引擎（每次只渲染12个，保持曲库展开时60fps丝滑）
function renderGridChunked(grid, entries, createCardFn) {
    let i = 0;
    const chunkSize = 12;
    function renderNextChunk() {
        const end = Math.min(i + chunkSize, entries.length);
        for (; i < end; i++) {
            const card = createCardFn(entries[i], i);
            grid.appendChild(card);
        }
        if (i < entries.length) requestAnimationFrame(renderNextChunk);
    }
    requestAnimationFrame(renderNextChunk);
}

function renderAlbumGrid(grid, filter) {
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

    // 应用切片渲染
    renderGridChunked(grid, entries, ([key, group], idx) => {
        const card = createCoverCard(group, idx, 'album');
        card.onclick = () => showAlbumDetail(group, modal);
        return card;
    });
}

function renderArtistGrid(grid, filter) {
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

    // 应用切片渲染
    renderGridChunked(grid, entries, ([key, group], idx) => {
        const card = createCoverCard(group, idx, 'artist');
        card.onclick = () => { playAudio(group.firstIdx); modal.remove(); };
        return card;
    });
}

function renderRecentGrid(grid, filter) {
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

    // 应用切片渲染
    renderGridChunked(grid, filtered, (group, idx) => {
        const card = document.createElement('div');
        card.className = 'cover-lib-card';
        card.innerHTML = `
            <div class="art-wrap">
                ${group.art ? `<img src="${group.art}" loading="lazy">` : '<div style="width:100%;height:100%;background:rgba(255,255,255,0.05);display:flex;align-items:center;justify-content:center;font-size:32px;">🎵</div>'}
                <div class="play-all-btn">▶</div>
            </div>
            <div class="album-name">${(group.title || '').length > 12 ? (group.title||'').slice(0,11)+'…' : (group.title||'')}</div>
            <div class="album-meta">${group.artist}</div>
        `;
        card.onclick = () => { playAudio(group.firstIdx); modal.remove(); };
        return card;
    });
}

function createCoverCard(group, idx, type) {
    const card = document.createElement('div');
    // 核心修改：加上 focusable 类名与 tabIndex，支持手柄/键盘焦点导航
    card.className = 'cover-lib-card focusable';
    card.tabIndex = 0;
    if (type === 'artist') card.classList.add('artist-card');

    const albumName = type === 'artist' ? group.artist : (group.album || '未知专辑');
    const metaText = type === 'artist' ? `${group.songs.length}首` : `${group.artist} · ${group.songs.length}首`;

    card.innerHTML = `
        <div class="art-wrap">
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

// 专辑详情面板 (v2.3.2 完美手柄适配版)
function showAlbumDetail(group, parentModal) {
    const detailModal = document.createElement('div');
    detailModal.className = 'modal-overlay open';
    detailModal.style.zIndex = '1001';

    // 核心改动 1：为操作按钮加上 focusable 类名与 tabindex="0"
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
                        <button class="btn-glass focusable" id="btnPlayAlbum" style="background:var(--primary);color:var(--text-on-primary);border:none;" tabindex="0">▶ 播放整张专辑</button>
                        <button class="btn-glass focusable" id="btnCloseAlbumDetail" tabindex="0">关闭</button>
                    </div>
                </div>
            </div>
            <div class="album-detail-tracks" id="albumDetailTracks"></div>
        </div>
    `;
    document.body.appendChild(detailModal);

    // 核心改动 2：立即更新焦点上下文，让手柄焦点瞬间"吸附"进入详情弹窗中
    updateFocusContext();

    // 渲染曲目列表
    const tracksEl = detailModal.querySelector('#albumDetailTracks');
    group.songs.forEach((songIdx, trackNum) => {
        const s = musicLibrary[songIdx]; // 从全库取数
        const track = document.createElement('div');
        
        // 核心改动 3：为每一行歌曲注入 focusable 与 tabIndex
        track.className = `album-detail-track focusable${songIdx === currentIndex ? ' active' : ''}`;
        track.tabIndex = 0;
        
        track.innerHTML = `
            <span class="track-num">${trackNum + 1}</span>
            <span class="track-title">${escapeHTML(s.title)}</span>
            <span class="track-dur">${s.artist}</span>
        `;
        
        // 鼠标/手柄确认点击播放
        track.onclick = () => {
            playAudio(songIdx);
            detailModal.remove();
            parentModal.remove();
        };
        tracksEl.appendChild(track);
    });

    // 播放整张专辑
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
        parentModal.remove();
        showToast(`🎵 正在播放专辑: ${group.album || '未知'}`);
    };

    // 关闭详情（核心改动 4：关闭时，必须重新扫描，让焦点优雅退回到曲库面板中）
    detailModal.querySelector('#btnCloseAlbumDetail').onclick = () => {
        detailModal.remove();
        updateFocusContext(); 
    };
    
    detailModal.onclick = (e) => { 
        if (e.target === detailModal) {
            detailModal.remove();
            updateFocusContext(); 
        }
    };
}
