/**
 * MBolka Player - Visualizer
 * Particle system, object pooling, flowing sand, renderVisLoop
 */

// === 可视化引擎 (增强版，灰阶频谱) v2.5-p2: 对象池化零GC ===
// 粒子对象池
const particlePool = [];
const MAX_POOL = 150;
class Particle {
    constructor() { this.reset(); }
    reset() { this.x = 0; this.y = 0; this.vx = 0; this.vy = 0; this.color = ''; this.size = 0; this.life = 0; this.active = false; return this; }
    init(x, y, vx, vy, color, size) {
        this.x = x; this.y = y; this.vx = vx; this.vy = vy; this.color = color; this.size = size; this.life = 1; this.active = true; return this;
    }
    update() { this.x += this.vx; this.y += this.vy; this.life -= 0.02; this.size *= 0.95; this.vx *= 0.95; this.vy *= 0.95; return this.life > 0; }
    draw(ctx) { if(!this.active) return; const prevAlpha = ctx.globalAlpha; ctx.globalAlpha = this.life; ctx.fillStyle = this.color; ctx.beginPath(); ctx.arc(this.x, this.y, this.size, 0, Math.PI*2); ctx.fill(); ctx.globalAlpha = prevAlpha; }
    kill() { this.active = false; }
}
// 🚀 v2.8.2+: 根据硬件动态调整池大小
const HARDWARE_CONCURRENCY = navigator.hardwareConcurrency || 4;
const ADJUSTED_MAX_POOL = HARDWARE_CONCURRENCY <= 4 ? 80 : 150;
// 预分配池
for (let i = 0; i < ADJUSTED_MAX_POOL; i++) particlePool.push(new Particle());

function acquireParticle(x, y, vx, vy, color, size) {
    for (let p of particlePool) if (!p.active) return p.init(x, y, vx, vy, color, size);
    // 池耗尽兜底：扩展池
    const p = new Particle().init(x, y, vx, vy, color, size);
    particlePool.push(p);
    return p;
}

// Ripple 对象池
const ripplePool = [];
const MAX_RIPPLE_POOL = 20;
class Ripple {
    constructor() { this.reset(); }
    reset() { this.x = 0; this.y = 0; this.radius = 5; this.color = ''; this.life = 0; this.active = false; return this; }
    init(x, y, color) {
        this.x = x; this.y = y; this.radius = 5; this.color = color; this.life = 1; this.active = true; return this;
    }
    update() { this.radius += 10; this.life -= 0.04; return this.life > 0; }
    draw(ctx) { if(!this.active) return; const prevAlpha = ctx.globalAlpha; ctx.globalAlpha = this.life; ctx.strokeStyle = this.color; ctx.lineWidth = 2; ctx.beginPath(); ctx.arc(this.x, this.y, this.radius, 0, Math.PI*2); ctx.stroke(); ctx.globalAlpha = prevAlpha; }
    kill() { this.active = false; }
}
for (let i = 0; i < MAX_RIPPLE_POOL; i++) ripplePool.push(new Ripple());

function acquireRipple(x, y, color) {
    for (let r of ripplePool) if (!r.active) return r.init(x, y, color);
    const r = new Ripple().init(x, y, color);
    ripplePool.push(r);
    return r;
}

const createExplosion = (x, y, intensity) => {
    if(!isImmersiveMode) return;
    const count = Math.floor(10 * intensity);
    for (let i=0; i<count; i++) {
        const ang = Math.random() * Math.PI*2, spd = 1 + Math.random()*4*intensity;
        const size = 1.5 + Math.random()*4;
        const gray = 150 + Math.floor(Math.random() * 105);
        particles.push(acquireParticle(x, y, Math.cos(ang)*spd, Math.sin(ang)*spd, `rgb(${gray},${gray},${gray})`, size));
    }
};
const createRipple = (x, y) => {
    if(isImmersiveMode) {
        ripples.push(acquireRipple(x, y, 'rgba(200,200,200,0.6)'));
        for(let i=0;i<3;i++) {
            const gray = 180 + Math.floor(Math.random() * 75);
            particles.push(acquireParticle(x, y, (Math.random()-0.5)*1.5, (Math.random()-0.5)*1.5, `rgb(${gray},${gray},${gray})`, 1+Math.random()*2));
        }
    }
};

document.addEventListener('mousemove', (e) => {
    if (isEnergySaving) return; // v2.7: 节能模式不产生粒子
    mouseX = e.clientX; mouseY = e.clientY;
    if (isImmersiveMode && isPlaying && Math.random() < 0.25) {
        const gray = 160 + Math.floor(Math.random() * 95);
        particles.push(acquireParticle(mouseX, mouseY, (Math.random()-0.5)*1.5, (Math.random()-0.5)*1.5, `rgb(${gray},${gray},${gray})`, 1+Math.random()*3.5));
    }
});

document.addEventListener('touchmove', (e) => {
    if (isEnergySaving) return; // v2.7: 节能模式不产生粒子
    if (isImmersiveMode && isPlaying) {
        const touch = e.touches[0];
        mouseX = touch.clientX; mouseY = touch.clientY;
        if (Math.random() < 0.4) {
            const gray = 150 + Math.floor(Math.random() * 105);
            particles.push(acquireParticle(mouseX, mouseY, (Math.random()-0.5)*2, (Math.random()-0.5)*2, `rgb(${gray},${gray},${gray})`, 1.5+Math.random()*4));
        }
    }
}, { passive: true });

document.addEventListener('click', (e) => {
    if (isEnergySaving) return; // v2.7: 节能模式跳过涟漪
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

// v2.5: 流沙流动渲染 - 极低分辨率 Canvas + CSS 强力模糊实现流体质感
function drawFlowingSand() {
    const cvs = el.bgColor;
    if (!cvs || !cvs.classList.contains('active')) return;

    // 保持画布极低分辨率（64x64），实现极致性能
    if (cvs.width !== 64 || cvs.height !== 64) {
        cvs.width = 64;
        cvs.height = 64;
    }

    const ctx = cvs.getContext('2d');
    ctx.clearRect(0, 0, 64, 64);

    const hue = currentHue;

    // 结合音频节奏：如果 Bass 很强，流沙波动速度会加快
    const bass = dataArray ? (dataArray[0] + dataArray[1]) / 2 : 0;
    const speedFactor = 1.0 + (bass / 255) * 1.5;

    sandPhaseA += 0.005 * speedFactor;
    sandPhaseB += 0.008 * speedFactor;
    sandPhaseC += 0.003 * speedFactor;

    // 1. 绘制底层（暗调沙）
    ctx.fillStyle = `hsl(${hue}, 40%, 12%)`;
    ctx.fillRect(0, 0, 64, 64);

    // 2. 绘制中层沙浪 A
    ctx.fillStyle = `hsl(${(hue - 20 + 360) % 360}, 55%, 18%)`;
    ctx.beginPath();
    ctx.moveTo(0, 64);
    for (let x = 0; x <= 64; x += 4) {
        const y = 32 + Math.sin((x / 10) + sandPhaseA) * 12;
        ctx.lineTo(x, y);
    }
    ctx.lineTo(64, 64);
    ctx.fill();

    // 3. 绘制中层沙浪 B
    ctx.fillStyle = `hsl(${(hue + 20) % 360}, 50%, 15%)`;
    ctx.beginPath();
    ctx.moveTo(0, 64);
    for (let x = 0; x <= 64; x += 4) {
        const y = 40 + Math.cos((x / 12) - sandPhaseB) * 15;
        ctx.lineTo(x, y);
    }
    ctx.lineTo(64, 64);
    ctx.fill();

    // 4. 绘制顶层亮沙 C（随节奏膨胀更明显）
    const peakAmp = (bass / 255) * 8;
    ctx.fillStyle = `hsl(${hue}, 65%, 22%)`;
    ctx.beginPath();
    ctx.moveTo(0, 64);
    for (let x = 0; x <= 64; x += 4) {
        const y = 48 + Math.sin((x / 8) + sandPhaseC) * (6 + peakAmp);
        ctx.lineTo(x, y);
    }
    ctx.lineTo(64, 64);
    ctx.fill();
}

// === 核心重构：全域 60FPS 色音同步视觉主循环 ===
const renderVisLoop = (timestamp) => {
    requestAnimationFrame(renderVisLoop);

    // 1. FPS 监测与性能自适应
    fpsFrames++;
    if (timestamp - fpsLastTime >= 1000) {
        currentFPS = Math.round(fpsFrames / ((timestamp - fpsLastTime) / 1000));
        fpsFrames = 0;
        fpsLastTime = timestamp;

        if (currentFPS < 30 && particleCount > 30) {
            particleCount = Math.max(30, particleCount - 10);
        } else if (currentFPS > 55 && particleCount < MAX_PARTICLES) {
            particleCount = Math.min(MAX_PARTICLES, particleCount + 5);
        }
    }

    const frameInterval = performanceMode ? 1000 / 30 : 1000 / targetFPS;
    if (timestamp - lastFrameTime < frameInterval) return;
    lastFrameTime = timestamp;

    if (!analyser) return;
    analyser.getByteFrequencyData(dataArray);

    // v2.7: 节能模式 — isEnergySaving 激活且非沉浸时跳过全部绘制保持 rAF
    if (isEnergySaving && !isImmersiveMode) {
        return;
    }
    // 兼容 v2.6: PiP 激活但节能开关未开时仍保持渲染
    if (pipWindow && !pipWindow.closed && !isImmersiveMode) {
        return;
    }

    visTime += 0.008;

    // 2. 核心同步：不论在哪个界面，统一执行 60 帧无缝色相（Hue）过渡计算
    if (isPlaying) {
        if (cfg.colorMode) {
            let diff = targetHue - currentHue;
            if (diff > 180) diff -= 360;
            else if (diff < -180) diff += 360;
            currentHue += diff * 0.04; // 每帧平滑位移 4%
            if (currentHue < 0) currentHue += 360;
        } else {
            currentHue = (currentHue + 0.15) % 360; // 自动虹彩旋转
        }
    }

    // 3. 分视角渲染
    if (isImmersiveMode && !immCanvasCleared) {
        // === 沉浸模式渲染 ===
        const cvs = el.canvasImm, ctx = cvs.getContext('2d');
        if (cvs.width !== window.innerWidth || cvs.height !== window.innerHeight) {
            cvs.width = window.innerWidth; cvs.height = window.innerHeight;
            const cols = Math.ceil(cvs.width / 20);
            const rows = Math.ceil(cvs.height / 20);
            flowField = new Array(cols * rows).fill(0);
        }
        const W = cvs.width, H = cvs.height;
        ctx.clearRect(0, 0, W, H);

        if (isPlaying) {
            const bassAvg = (dataArray[0] + dataArray[1] + dataArray[2] + dataArray[3]) / 4;
            const midAvg = (dataArray.slice(4, 20).reduce((a,b)=>a+b,0)) / 16;
            const highAvg = (dataArray.slice(20, 64).reduce((a,b)=>a+b,0)) / 44;
            bassHistory.push(bassAvg);
            if (bassHistory.length > 30) bassHistory.shift();
            const smoothBass = bassHistory.reduce((a,b)=>a+b,0) / bassHistory.length;

            // 沉浸光晕背景
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
                const x = W/2 + lutCos(angle) * r;
                const y = arcY + lutSin(angle) * r * 0.5;
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

            // 两侧频谱柱
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
        // === 主界面渲染 ===
        const cvs = el.canvasMain;
        cvs.width = cvs.offsetWidth; cvs.height = cvs.offsetHeight;
        spectrumCtxMain.clearRect(0,0, cvs.width, cvs.height);
        const w = (cvs.width / (analyser.frequencyBinCount/2)); let x = 0;
        for(let i=0; i<analyser.frequencyBinCount/2; i++) {
            const h = (dataArray[i]/255) * cvs.height;
            // 彩色随主题色相渐变
            const amp = dataArray[i]/255;
            spectrumCtxMain.fillStyle = `hsla(${currentHue}, 75%, ${45 + amp * 20}%, ${0.5 + amp * 0.4})`;
            spectrumCtxMain.fillRect(x, cvs.height - h, w-1.5, h); x += w;
        }

        // 核心修改：在主界面也激活 60 帧取色流沙渲染，实现绝对一致的取色效率！
        if (cfg.colorMode) {
            drawFlowingSand();
        } else {
            // 如果关闭了取色模式，擦除主背景Canvas，让 CSS 的静态预设主题渐变显露出来
            const bgCtx = el.bgColor.getContext('2d');
            if (bgCtx) bgCtx.clearRect(0, 0, el.bgColor.width, el.bgColor.height);
        }
    }
};
