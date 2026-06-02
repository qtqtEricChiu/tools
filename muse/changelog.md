# MBolka Player 更新日志

---

## v2.8.9 (2026-06-03)

### 🔥 LRC 歌词解析引擎全面重写

**根因**：单语 LRC 被误识别为双语（isTranslationLine 长度规则假阳性、元数据行漏跳过），双语 LRC 存在漏识别（pendingOriginal 已翻译后后续行被静默丢弃）。

**修复方案**：
- **预扫描中文计数**：全文件 < 3 个中文字 → 纯单语守卫，跳过所有翻译检测
- **元数据行检测扩展**：新增英文制作信息（Lyrics by / Composed by / Arranged by / Produced by）、版权行（TME享有 / 著作权 / Copyright）、创作信息行（0~10s 内含 ` - ` 模式）
- **移除粗糙的长度规则**（Rule 4：txt.length < original.length * 0.5）— 主要假阳性来源
- **pendingOrig 生命周期修正**：原文被赋翻译后立即 `pendingOrig = null`，后续行作为新原文
- **双重保护**：翻译行检测时若 `pendingOrig.translation` 已存在 → 强制 `pendingOrig = null`

### 🔥 Crossfade 引擎 — Web Audio API 全面重写

**旧架构**：rAF + `audio.volume` 指数动画 → 帧精度 ±16ms，可能掉帧

**新架构**（Chrome 148）：
- **固定双槽位 A/B**（`audio` + `cfAudioB`）— 永不交换元素
- **AudioContext lazy-init**（首次需要时创建，符合自动播放策略）
- **GainNode 精确定时斜坡**：`gain.exponentialRampToValueAtTime()` — 亚毫秒精度，零掉帧
- **提前 50% 预加载**：`setTimeout` 在当前歌曲 50% 时预加载被动槽
- **Semaphore 防重叠**：`cfTransitionId` 事务版本号 + `cfAirLocked` 门禁锁
- **3 秒超时回退**：预加载超时自动降级为 `goNext()` 直接切歌

---

## v2.8.8 (2026-06-03)

### 🎯 歌词垂直居中修复

**根因分析**：4个独立根因导致歌词定位偏移——(1) CSS `padding:50%` 基于宽度而非高度；(2) `offsetTop`/`clientHeight` 在CSS过渡期间值不稳定；(3) `transform:scale(1.05)` 未纳入视觉高度计算；(4) 未强制重排就读取布局属性。

**修复方案**：
- 强制重排 (`void line.offsetHeight`) 确保读取最新布局
- 使用 `getBoundingClientRect()` 替代 `offsetTop`/`clientHeight`
- 纳入 `scale(1.05)` 因子修正视觉高度
- CSS `.lrc-viewport` padding 从 `50%` 改为 `20px`（固定值）
- 添加 `window resize` 监听器（150ms防抖），窗口变化时重新居中

### 🔗 双语歌词解析增强

**问题**：链式LRC格式中，翻译行与下一句原文同时间戳，原有分组算法将翻译和原文放入同一 group，导致配对错误。

**修复**：移除按时间戳分组逻辑，改为逐行处理。新增 `isTranslationLine()` 启发式检测函数（4条规则）：
1. 时间戳差 < 0.1s → 翻译
2. 本行含中文，上行不含 → 翻译
3. 纯中文 vs 纯ASCII 组合 → 强信号
4. 长度 < 上行50% → 翻译

### 🎨 下一句歌词中间模糊厚度

- `.lrc-line.active + .lrc-line` 模糊值：`blur(0px) !important` → `blur(1px)`
- 视觉层次：当前行(0px) > 下一句(1px) > 其他行(2px)

### 🎮 B键高优先级退出所有浮窗

- 新增 `closeSettings()` / `closePlaylist()` / `closeCoverLibrary()` / `closeHelp()` / `closeFileInfo()` 专用关闭函数
- `handleGlobalClose()` 按浮窗ID路由到对应关闭函数（带动画+焦点恢复+设置保存）
- `pollGamepad` 中 B键最高优先级拦截，先于所有其他逻辑

### 🕹️ 手柄适配逻辑全面重写

**焦点系统**：
- `updateFocusContext()` 按浮窗类型收集所有可交互元素（含 `input[range]`/`checkbox`/`select`）
- `moveFocus2D()` 增强：跳过不可见元素、视口外惩罚(+10000)、元素面积奖励

**按钮映射重写**：
| 按钮 | v2.8.7 | v2.8.8 |
|------|--------|--------|
| A | 确认/播放 | 元素感知确认（滑块→微调、复选框→切换、下拉→聚焦） |
| B | 退出 | 退出（含滑块微调退出） |
| X | 收藏 | 播放/暂停（全局） |
| Y | 循环模式 | 沉浸模式切换 |
| LB | 上一首 | 设置中切换选项卡 / 上一首 |
| RB | 下一首 | 设置中切换选项卡 / 下一首 |
| LT | 降低音量 | 快退 5秒 |
| RT | 增加音量 | 快进 5秒 |
| L3 | — | 焦点模式切换 |

- 新增滑块微调模式：焦点在 `range` 输入框上按A进微调，方向键调整值，B退出

### 🎨 沉浸模式修复

- `.imm-wrapper` 添加 `background:transparent` + `isolation:auto` 修复 backdrop-filter 模糊被阻断
- 非双语歌词沉浸模式显示：当前句原文 + 下一句原文

---

## v2.8.7 (2026-06-03)

### 🔧 空翻译检测 — 避免误用下句原文

**根因**：v2.8.6 链式解析中 `group.entries.filter()` 会过滤空字符串，如果同名时间戳首行为空（翻译空置），过滤后的 `texts[0]` 变成了本句原文而非上一句翻译，导致将"下一句原文"错误赋值给上句的 `translation`。

**修复**：在过滤前增加 `hasTranslation` 检查（`group.entries.length > 1 && group.entries[0] && group.entries[0].trim()`），只有翻译非空时才赋值。翻译空置时 `isBilingual` 保持 `false`，歌词以单行显示。

### 📝 翻译行字号缩小

- 激活翻译行：`font-size: 0.95em` → `0.75em`，视觉层次明显（原文/翻译分层清晰）

### 📺 PiP 歌词回退

- `pipCurrLine` = 当前句原文，`pipNextLine` = 下一句原文（不显示翻译）
- 与沉浸舱行为保持一致

### 🎵 Crossfade 引擎系统性重写

**问题诊断**：
1. `audio.src = nextAudio.src` 导致音频重新加载，丢失淡入效果
2. `isFading` 布尔锁在多个异步路径中容易遗漏重置
3. `setInterval` + 固定步长导致音质步进感
4. `playing` 事件触发时机不稳定

**v2.8.7 新架构**：

| 组件 | 变更 |
|------|------|
| **状态机** | `CrossfadeState { IDLE, PRELOADING, FADING, COMPLETED }` 替代 `isFading` 布尔锁 |
| **音频池** | `audioPool = [audio, nextAudio]`，Crossfade 完成后 swap 活跃元素 |
| **rAF 动画** | `performCrossfade()` 使用 `requestAnimationFrame` + `performance.now()` 驱动 |
| **指数曲线** | 淡入淡出使用 `Math.pow(progress, 2)` 替代线性，人耳感知更自然 |
| **async/await** | `triggerCrossfadeV2()` 显式 `await nextAudio.load()` 预加载 |
| **旧版删除** | 移除 `triggerFadeOutLegacy()` 和 `triggerFadeInLegacy()`，降级走 `goNext()` |

**关键修复点**：
- `finishCrossfadeV2()`: 直接 swap audio 元素而非复制 src
- `playAudio()`: 手动切歌时取消 rAF + 停止池中音频 + 重置 `crossfadeState`
- `goNext()`: 用 `crossfadeState !== IDLE` 替代 `isFading` 检查

### ™️ 其他

- **版本号更新**：`index.html`、`app.js`、版权区域统一为 v2.8.7

---

## v2.8.6 (2026-06-03)

### 🔧 双语LRC链式解析 — 专项修复

**根因**：v2.8.5 的 Map 分组算法将同时间戳的两行视为"同一句的原文+翻译"，但实际 LRC 格式是**链式结构**——同时间戳第一行 = **上一句原文的翻译**，第二行 = **本句原文**。导致"上一句翻译被合并到下一句原文"的配对错误。

**修复**：`parseLyricText()` 完全重写为三阶段链式解析：

1. **顺序分组**：按时间戳分组，保持同时间戳内行序
2. **元数据过滤**：`isMetaLine()` 识别制作信息行（曲、编曲、词、TME版权 等），重置翻译链
3. **链式赋值**：`pendingOriginal` 追踪上一句原文，本组第一行→赋值给上一句的 `translation`，本组最后一行→新的 `original`
4. **最终兜底**：链尾未配对的原文保持单语

**链式算法示意**：
```
[00:02.69] Cross my heart          ← 原文1, pendingOriginal="Cross my heart"
[00:03.84] 胸前画十字 郑重起誓     → 赋值给原文1.translation ✓
[00:03.84] we'll always            ← 原文2, pendingOriginal="we'll always"
[00:04.70] 我们会永远              → 赋值给原文2.translation ✓
```

### 🎬 沉浸模式 + PiP 双行回退

- **沉浸模式**：`immCurrLine`=原文，`immNextLine`=翻译（不再预告下一句），翻译行 `opacity` 从 0.7 提升到 **0.85**
- **PiP 画中画**：`pipCurrLine`=原文，`pipNextLine`=翻译（不再显示下一句歌词），翻译行 `opacity` 设为 **0.85**，无翻译时隐藏

### 📋 CSS 清理

- **`immersive.css`**：移除重复 `.imm-subtitle-line` 过渡规则，翻译行颜色从 `rgba(255,255,255,0.5)` 调整为 `0.6`

### ™️ 其他

- **版本号更新**：`index.html`、`app.js`、版权区域统一为 v2.8.6

---

## v2.8.5 (2026-06-03)

### 🐛 双语LRC解析重写

- **Map分组算法**：`parseLyricText()` 从"两遍扫描+容差合并"重构为 Map 精确分组（按 `time.toFixed(3)` 做键），同时间戳第一行=原文、第二行=翻译，彻底解决"上一句翻译被错误合并到下一句原文"的问题
- **`text` 字段兼容**：有翻译时 `text = original + \n + translation`，进度条预览自然显示双语

### 🎬 沉浸模式歌词动画回退 + 延迟修复

- **回退简单动画**：从 CSS `switching`/`active` 类+`transitionend` 回归简洁的 `el.style.opacity=0` + `setTimeout` 动画（旧版 final1.0 验证有效的方案）
- **修复更新延迟**：`syncLyrics` 的 `activeIdx` 计算从顺序扫描 `-0.2` 偏移改为**逆序精确匹配**（从后往前找到第一个 `cur >= time`），消除约 1 秒的更新延迟
- **不再显示下一句**：沉浸模式 `immNextLine` 改为显示翻译文本（有双语时）或无文本（单语时），不再提前展示下一句歌词
- **移除 `lrcTransitionLock`**：清理 v2.8.4 引入的过渡锁变量及相关兜底计时器

### 📐 新增歌词垂直对齐模式

- **两种模式可选**：垂直居中（默认）和偏上显示（QQ音乐风格，距顶部30%）
- **设置面板新增切换按钮**：`btnLrcAlignCenter` / `btnLrcAlignTop`，带视觉激活状态
- **`updateLrcAlignUI()` 函数**：同步按钮状态 + 切换 `.lyrics-align-top` CSS 类
- **持久化存储**：`lyricsAlignMode` 写入 `localStorage`（`saveSettings` / `loadSettings`）

### 🎵 Crossfade 跨曲修复

- **`getNextTrackIndex()` 统一选择**：提取独立函数，Shuffle 模式排除当前歌曲随机选（与 `goNext` 一致），顺序模式取下一首
- **`triggerCrossfade()` 改用统一逻辑**：从 `Math.floor(Math.random()*length)` 改为 `getNextTrackIndex()`，交叉淡入淡出的歌曲 = 下一首实际播放的歌曲
- **`goNext()` 防重入**：`isFading` 时直接 return，阻止 crossfade 末尾再次触发切歌导致混音
- **`finishCrossfade()` 状态管理**：切换前清除 `audio.onended = null`，播放后延迟 1 秒重新绑定
- **防御性检查**：`checkCrossfade` 新增 `isNaN(remaining)` / `audio.paused` / `audio.ended` 三项检查

### 📋 CSS 调整

- **`base-layout.css`**：双语翻译样式简化（`opacity:0.5→0.9` 过渡），新增 `.lyrics-align-top` 偏上模式样式
- **`immersive.css`**：移除 v2.8.4 的 `.switching`/`.active` CSS 过渡类，移除 `.imm-original`/`.imm-translation` 双 span 样式

### ™️ 其他

- **版本号更新**：`index.html` 标题、`app.js` 文件头、版权区域版本号统一为 v2.8.5

---

## v2.8.4 (2026-06-03)

### 🐛 双语LRC解析重构

- **两遍扫描算法**：`parseLyricText()` 从"向前看一行"改为"两遍扫描"——先提取所有时间戳条目，再按时间戳分组合并（容差0.02秒），彻底消除同时间戳双语歌词被拆分为独立条目导致的重复行问题
- **字段一致性**：合并后 `text` 保留原文（兼容单语场景），`original`/`translation` 分别存储原文和翻译
- **歌词面板**、**沉浸模式**、**进度条预览**三者数据完全对齐，不再出现显示错乱

### 🎬 沉浸模式歌词动画防频闪

- **CSS transition 替代硬编码 setTimeout**：`el.style.opacity = 0` + `setTimeout` 改为 CSS `switching`/`active` 类切换，利用 `transitionend` 事件精确完成动画
- **过渡锁 `lrcTransitionLock`**：防止快速切歌时多个 setTimeout 堆叠导致的透明度抖动
- **300ms 兜底保护**：若 transition 未触发（元素隐藏等极端情况），自动强制完成动画
- **新增CSS样式**：`immersive.css` 中 `.imm-subtitle-line.switching`（opacity:0 + translateY）、`.imm-subtitle-line.active`（opacity:1）过渡类

### ⚡ 节能模式状态机重构

- **位标志状态机**：从布尔状态 `isEnergySaving` + 多个标记变量重构为 `EnergyMode` 位标志（`NONE/ONE_CLICK/PIP_TEMP/FRAME_LIMIT/VISIBILITY`），支持多模式叠加共存
- **`enterEnergySaving(mode)` / `exitEnergySaving(mode)`**：按位 OR 进入、按位 AND NOT 退出，精准控制每种模式的启停
- **`applyEnergySaving(enable, triggerMode)`**：实际执行节能操作（清空粒子、释放流场、降频歌词），仅在首次进入/完全退出时触发
- **一键节能 + 画中画叠加**：开启一键节能后打开画中画，关闭画中画后一键节能保持；`visibilitychange` 只退出 `VISIBILITY` 模式不误退其他
- **向后兼容**：旧变量 `isEnergySaving`、`pipTempEnergySaving`、`oneClickEnergySaving` 保留并同步更新

### 🎵 曲库浮窗关闭逻辑修复

- **播放整张专辑**：从 `detailModal.remove(); parentModal.remove()` 改为 `safeTransition()` 统一关闭，消除 DOM 状态残留导致的浮窗无法二次打开问题
- **曲目点击播放**：从 `detailModal.remove(); parentModal.remove()` 改为 `closeAllModals()` 标准关闭流程
- **关闭详情动画**：点击遮罩关闭统一走 `closeDetail()` 函数，确保 `parentModal` 焦点恢复

### 🎮 手柄B键退出增强

- **handleGlobalClose z-index 排序**：从简单数组最后一项（DOM顺序）改为按 CSS `z-index` 排序，正确识别视觉最上层浮窗
- **动态弹窗动画移除**：动态创建的弹窗（统计、专辑详情）关闭时先移除 `.open` 类触发过渡动画，400ms 后再移除 DOM
- **焦点智能恢复**：关闭后延迟 50ms 调用 `updateFocusContext()`，确保焦点正确退回下层

### 📋 updateFocusContext 增强

- **z-index 优先级检测**：按 `z-index` 排序活跃浮窗，优先聚焦最上层浮窗内的 `.focusable` 元素
- **无适配元素兜底**：若浮窗内无 `.focusable`，自动聚焦 `.modal-content`

### ™️ 其他

- **版权署名**：设置页面底部 `co-created with` 添加 QClaw
- **版本号更新**：`index.html` 标题、`app.js` 文件头、版权区域版本号统一为 v2.8.4

---

## v2.8.3 (2026-06-02)

### 🌏 双语LRC歌词支持

- **智能双语检测**：`parseLyricText()` 自动识别同时间戳的双语歌词行（原文+翻译），合并为单条记录
- **主界面显示优化**：原文正常高亮，翻译行以较小字号、较低透明度显示；激活时翻译跟随提升
- **沉浸模式双语适配**：原文大字高亮，翻译行紧随其后（稍小但清晰），支持画中画模式响应式适配
- **进度条预览合并**：`getLyricAtTime()` 返回 `"原文 | 翻译"` 格式，进度条悬停时显示双语预览
- **新增 `escapeHtml()` 辅助函数**：防止歌词内容中的HTML标签被解析为DOM元素

### 🎵 交叉淡入淡出引擎重构

- **真正的双轨交叉播放**：使用 `requestAnimationFrame` 实现当前歌曲淡出、下一首淡入的平滑过渡
- **预加载机制**：提前创建第二个 `Audio` 元素加载下一首，确保无缝切换
- **高精度进度检测**：每100ms通过 `performance.now()` 检查播放进度，避免 `timeupdate` 精度不足问题
- **降级兼容方案**：预加载失败时自动回退到旧版 `triggerFadeOutLegacy()` 逻辑
- **手动切歌保护**：`playAudio()` 中强制取消正在进行的交叉淡入淡出，立即恢复标准音量
- **修复 `once` 参数**：`triggerFadeInLegacy()` 中 `{ once: false }` 修正为 `{ once: true }`
- **新增 `crossfadeRafId` 全局变量**：管理交叉淡入淡出的 `requestAnimationFrame` ID

### 📋 其他

- **版本号更新**：`index.html` 标题和 `app.js` 文件头更新至 v2.8.3
- **新增CSS样式**：
  - `base-layout.css`：`.lrc-line.bilingual`、`.lrc-original`、`.lrc-translation` 样式
  - `immersive.css`：`.imm-original`、`.imm-translation` 沉浸模式双语歌词样式
  - `components.css`：`.crossfade-indicator` 交叉淡入淡出状态指示器

---

## v2.8.2 (2026-06-02)

### ⚡ 节能板块整合重构

- **统一节能板块**：将散落的节能功能整合到"⚡ 节能模式"黄色高亮板块
- **🔋 一键节能**：去除所有可视化动效，保持正常亮度（不添加 `pip-standby` CSS 暗黑类）
- **🎬 画面节能**：仅将动画帧率降至 30fps，保留视觉效果（替代旧性能模式）
- **📺 临时节能**：启动画中画时自动优化主界面性能（画中画关闭后自动退出）
- **移除旧 UI**：EQ 面板下的性能模式按钮已整合，不再单独显示

### 🔧 画中画节能状态同步修复

- **`pipTempEnergySaving` 标记**：区分"画中画临时节能"和"手动节能"，关闭画中画时自动退出临时节能
- **`visibilitychange` 修复**：标签页返回时检查画中画状态，防止意外退出节能模式
- **`exitEnergySaving` 防御**：画中画运行时阻止退出临时节能模式

### 📝 设置页面命名规范统一

- **去除技术术语**：移除"核心视觉引擎"、"UI 视觉引擎"等描述
- **统一格式**：`emoji + 简要功能名 +（快捷键）`
- **12 处命名优化**："封面取色"、"主题色"、"背景图片"、"均衡器"、"快捷键指南"等
- **移除设置中画中画按钮**：仅保留主界面播放器上的画中画入口

### 🚀 性能优化深度改进

- **Page Visibility API**：页面不可见时暂停高频渲染循环，降至 500ms 间隔心跳
- **GPU 优化**：`Particle.draw()` / `Ripple.draw()` 用 `globalAlpha` 缓存替代 `save()/restore()`，减少 overdraw
- **内存优化**：粒子对象池根据 `navigator.hardwareConcurrency` 动态调整（低端设备 80 vs 高端 150）

### 📋 配置兼容性

- 旧版 `performanceMode` → 自动映射到 `cfg.frameEnergyEnabled`
- 旧版 `energySavingEnabled` → 自动映射到 `cfg.pipEnergyEnabled`
- 设置数据自动迁移，无需手动重新配置

---

## v2.8.1 (2026-06-02)

### 🔧 画中画节能模式加强

- **exitEnergySaving 防御性检查**：防止在画中画激活时意外退出节能模式
- **togglePip 状态管理优化**：关闭画中画时根据用户设置决定是否保持节能模式，画中画启动时强制进入节能模式
- **健康检查兜底同步**：定时器关闭时也遵循节能设置

### 🎬 专辑详情动画修复

- **双重 requestAnimationFrame**：确保模态框 `.open` 类分帧添加，触发 CSS 弹入动画
- **关闭动画优化**：先移除 `.open` 类触发退出动画，400ms 后再移除 DOM 节点

### 🎮 手柄适配完善

- **EQ 预设按钮**：添加 `.focusable` 类和 `tabIndex`，支持手柄/键盘导航
- **activateFocus 增强**：区分专辑曲目、EQ 预设按钮等元素类型做精准处理
- **按钮映射优化**：X 键→收藏，Y 键→循环播放模式（更符合直觉）
- **手柄提示徽章扩展**：新增画中画按钮和设置按钮的手柄提示

### 📋 设置浮窗菜单重构

- **新排序**：文件 → 显示 → 音频 → 音效 → 歌词 → 其它 → 系统
- **自定义背景整合**：从"其它"区域移至"显示"区域紧跟主题色
- **快捷键指南更新**：同步 X/Y 键映射变更

---

## v2.8.0 (2026-06-01)

### ⚡ 节能模式精确控制 + 沉浸舱完整停止

- **渲染守卫重构**：节能ON时全局跳过所有绘制（含沉浸舱），不再有沉浸模式渲染泄漏
- **沉浸舱强制退出**：`enterEnergySaving()` 自动退出沉浸模式，停止所有动画特效并释放 `flowField`、`particles`、`ripples` 内存
- **主频谱Canvas清空**：节能激活时同时清空主界面频谱显示
- **节能关 = 完整主界面**：关闭节能开关时，PiP激活也不影响主界面频谱、取色背景、流沙渲染

### ⌨️ 键盘快捷键全面升级

| 快捷键 | 功能 | 说明 |
|---|---|---|
| **Ctrl+O** | 打开文件夹载入音乐 | 最高效的载入方式 |
| **U / F** | 收藏/取消收藏当前歌曲 | 快速标记喜爱的歌 |
| **/** | 聚焦搜索播放列表 | 类似 YouTube 体验 |
| **Shift+F** | 全屏开关 | F键改用于收藏 |
| **Alt+T** | 睡眠定时器快速菜单 | 弹窗式秒设定时 |
| **Shift+Esc** | 一键关闭所有弹窗 | 强制回到主界面 |

### 🎮 手柄体验升级 — 按键指示器

- **手柄接入时自动注入 ⓐⓑⓧⓨ 徽章**：播放按钮标注 ⓐ，关闭按钮标注 ⓑ，模式切换标注 ⓧ 等
- **断开手柄自动清除**：所有徽章随手柄断开而移除
- **曲库-专辑详情 B键**：`handleGlobalClose` LIFO栈式关闭，B键先关详情再关曲库
- **手柄专属帮助面板**：帮助面板新增完整的游戏手柄操作指引表

### 🎨 设置页面 UI 重构

- **载入音乐移至设置页最顶端**：主页面更简洁，设置页首行醒目的 `📁 打开文件夹` 按钮
- **专辑封面取色模式提升为视觉核心**：独立取色引擎板块，状态标签实时显示激活状态，配色预览条
- **预设主题色归入 UI 视觉引擎**：取色模式开启时覆盖所有预设主题，关闭时预设主题复活
- **深色模式 + 背景模糊独立为视觉调节板块**

### 🩹 修复

- **专辑详情弹入动画**：`.album-detail-panel` 新增 `scale(0.9)→scale(1)` 弹簧动画，与所有弹窗一致
- **专辑详情手柄导航**：`updateFocusContext` 正确检测 `.album-detail-panel` 内 `focusable` 元素
- **空状态提示更新**：移除已不存在的"载入音乐"按钮引用，改为 Ctrl+O 提示

---

## v2.7.0-preview2 (2026-06-01)

### 🧹 内存优化 — 杜绝泄漏

- **P0: Blob URL 彻底回收**：`releaseAllBlobUrls()` 新增强制遍历 `playlist` 和 `musicLibrary`，释放所有残留的 `blob:` URL，每次载入新文件夹时旧音乐资源完全回收，内存不再随切换堆积
- **P1: 沉浸模式流场释放**：退出沉浸模式时将 `flowField` 置为空数组，大数组立即进入 GC 可回收状态
- **P1: PiP 定时器兜底清理**：`pipSyncInterval` 提升为模块级变量，关闭 PiP 时显式 `clearInterval`；新增 10 秒健康检查 `pipHealthCheck`，即使 PiP 窗口被操作系统外部关闭也能彻底清除定时器
- **P2: 播放历史限制**：`playHistory` 数组最大长度限制为 200，防止长时间播放后历史无限增长
- **P3: 统计面板闭包清理**：`showStatsPanel` 关闭时将 `modal` 引用置为 `null`，防止闭包循环引用

### 🩹 修复

- **曲库专辑详情页面**：修复 `renderAlbumGrid`/`renderArtistGrid`/`renderRecentGrid` 中 `modal` 变量作用域丢失问题，`modal` 现在作为参数显式传递，专辑详情面板恢复正常，点击专辑卡片可正确打开详情

---

## v2.6.0 (2026-06-01)

### 🎬 弹窗关闭动画回归 — 渐进式消失

- **`_closeModalsSync(isSwitching)` 双模式关闭**：新增 `isSwitching` 参数区分"切换弹窗"与"普通关闭"
  - `isSwitching=true`（切换弹窗）：禁用 `transition` + 强制重绘，0ms 瞬间消失
  - `isSwitching=false`（普通关闭）：保留 CSS transition，弹窗优雅淡出，恢复关闭动画体验
- **`closeAllModals()` 动画关闭**：传 `false` 让所有弹窗保留过渡动画
- **`safeTransition(fn)` 瞬切**：传 `true` 在弹窗切换时无动画快速过渡

### 👻 封面卡片消除残影

- **移除 `view-transition-name`**：从 `.cover-lib-card .art-wrap` 中移除 CSS `view-transition-name`，彻底消除曲库切换时封面图片的幽灵残影问题

### 🔈 音量系统修复

- **音量保存值修正**：`saveSettings` 中 `vol` 字段从 `audio.volume` 改为 `parseFloat(el.volSlider.value)`，避免保存淡入淡出过程中临时降低的音量值
- **沉浸模式音量滑块同步**：`loadSettings` 和 `adjustVolume` 中新增 `el.immVolSlider` 同步，确保沉浸舱音量滑块与主界面保持一致

### 🪶 画中画微待机模式 (Tiny Standby Mode)

- **主窗口低功耗休眠**：PiP 画中画激活时，主窗口自动添加 `.pip-standby` class
  - 整体透明度降至 0.35，禁止所有鼠标交互（`pointer-events: none`）
  - 专辑封面动画冻结（`animation: none`）
  - 背景层透明度锐减至 0.15
  - Canvas 完全隐藏（`opacity: 0`）
  - 仅 PiP 控制按钮保持可交互并附带呼吸发光脉冲动画
- **渲染层断电**：`renderVisLoop` 中 PiP 激活时，非沉浸模式直接跳过所有主窗口渲染（`return`），大幅节省 CPU/GPU 算力
- **pagehide 清理**：监听 `pagehide` 事件，PiP 窗口关闭时自动移除 `.pip-standby` 状态并更新按钮

---

## v2.5.0-release (2026-06-01)

### 🚀 全域弹窗栈控制器 — LIFO 后进先出完美退出

- **`handleGlobalClose()` 统一关闭管理器**：扫描页面上所有 `.modal-overlay.open` 弹窗（包括静态 HTML 和动态创建的），永远只关闭最上层（LIFO 栈顶）的那一个
- **键盘 Esc 全适配**：优先尝试关闭最上层弹窗，无弹窗打开时才退出沉浸模式；全屏状态优先退出全屏
- **手柄 B 键全适配**：与 Esc 共享同一逻辑，不再需要手动枚举每个弹窗类型
- **完美层级退回**：曲库 → 专辑详情 → 按 B/Esc 先关详情 → 再按 B/Esc 关曲库，层层递进
- **100% 向前兼容**：未来新增的任何弹窗自动获得手柄 B 键和键盘 Esc 退出支持

### 🌊 全域 60FPS 色调同步 — 主页面流沙背景实时取色

- **`renderVisLoop` 核心重构**：色相（Hue）过渡计算提升到函数顶部，不分支、不分界面，不论在主界面还是沉浸模式都统一以 60 帧平滑推进 `currentHue`
- **主页面流沙激活**：`cfg.colorMode` 开启时，主界面每帧调用 `drawFlowingSand()` 实时渲染低分辨率 Canvas 流沙背景，与沉浸舱毫秒级同步变色
- **关闭取色优雅降级**：关闭取色模式时自动清除背景 Canvas，让 CSS 静态预设主题渐变平滑显现
- **帧率控制优化**：`lastFrameTime` 独立作用域，性能模式和 30fps 节能模式正确生效

### 🔧 Bug 修复

- **`.btn-mode.active` 深色反白**：模式切换按钮激活态 `color` 从硬编码 `#000` 改为 `var(--text-on-primary)`，深色主题（深海/星夜）下自动反白
- **弹窗切换视觉残留**：`_closeModalsSync` 临时禁用 `transition` + 强制重绘 `offsetHeight`，实现弹窗 0ms 瞬间消失，消除"曲库→列表"约 1 秒重叠残留
- **专辑详情手柄全适配**：按钮注入 `tabindex="0"`，打开/关闭时主动 `updateFocusContext()` 拉焦点入/退弹窗
- **曲库增量加载恢复**：`renderGridChunked` 每次渲染 12 张卡片，`requestAnimationFrame` 分批递进，彻底解决曲库展开卡顿
- **函数名对齐**：`renderAlbumGridStatic` → `renderAlbumGrid` 等三函数与调用方统一命名

### 🎮 手柄体验增强

- **曲库 B 键退出**：手柄 B 键通过 `handleGlobalClose()` 统一处理，曲库打开时一键退回主页

---

## v2.5.0-preview2 (2026-05-31)

### ⚡ 性能优化四剑客 — CPU 算力节约 + 消除微卡顿

#### 1. 粒子对象池化 (Object Pooling) — 消除 GC 抖动
- **Particle / Ripple 类重构**：从每次 `new` 创建改为对象池模式，程序启动时预分配 150 个 Particle + 20 个 Ripple 实例
- **`acquireParticle()` / `acquireRipple()`**：取代 `new Particle()` 和 `new Ripple()`，从池中激活空闲实例，池耗尽时动态扩展
- **原地迭代替代 `filter`**：`particles.filter(...)` 和 `ripples.filter(...)` 改为 for 循环原地 compact，不再每帧创建新数组
- **`kill()` 方法**：粒子死亡时仅设置 `active = false` 归还池中，零对象创建、零 GC 垃圾

#### 2. 布局抖动消除 — 缓存 DOM 几何属性
- **`bindProgressBar` 重构**：`mousedown/touchstart` 时调用一次 `getBoundingClientRect()` 存入 `cachedRect`，拖拽过程中 `handleMove` 直接使用缓存值，`handleEnd` 时释放缓存
- **`setupProgressHover` 重构**：`mouseenter` 时缓存 rect，`mousemove` 期间使用缓存值，`mouseleave` 时释放

#### 3. 查表法 (LUT) 替代三角函数
- **128 点全圆查表**：`SIN_TABLE[]` / `COS_TABLE[]` 预计算 128 个等分角的三角函数值
- **`lutSin(angle)` / `lutCos(angle)`**：通过角度映射索引直接取值，替代 `Math.sin/cos`
- **沉浸模式频谱弧线**：49 次/帧的 `Math.cos/sin` 调用全部替换为查表

#### 4. GPU 图层升格 — `will-change` 减少重绘
- **`.view-container`**：添加 `will-change: transform, opacity`
- **`.modal-content`**：添加 `will-change: transform`
- **`.player-wrapper`**：添加 `will-change: transform`（独立合成层，避免 backdrop-filter 像素着色器重复计算）

### 🎨 WCAG 无障碍对比度 — `--text-on-primary` 动态反色

- **`getLuminance(colorStr)` 函数**：基于心理学相对亮度公式 `L = 0.299R + 0.587G + 0.114B`，支持 `#hex` 和 `rgb()` 两种格式
- **`applyThemeLogic()` 增强**：每次设置 `--primary` 后自动计算亮度，`luminance < 140` 时 `--text-on-primary` 设为 `#ffffff`（反白），否则为 `#0a0a1a`（深色）
- **CSS 统一替换**：`.btn-glass.active`、`.btn-play`、`.btn-play:hover`、`.cover-lib-tab.active` 的 `color` 从硬编码 `#000` 改为 `var(--text-on-primary)`
- **HTML/JS 内联样式替换**：`btnLoadFolder` 和 `btnPlayAlbum` 的 `color:#000` 改为 `color:var(--text-on-primary)`
- **`:root` 默认值**：`--text-on-primary: #0a0a1a`（匹配默认蓝色主题的浅色背景）

---

## v2.5.0-preview (2026-05-31)

### 🌊 彩色动态流沙背景 — 极低分辨率 Canvas + CSS 强力模糊
- **"高性能秘诀"实现**：`#bg-layer-color` 从 `<div>` 重构为 `<canvas>`，保持 64×64 物理分辨率，每帧使用三层正弦波叠加绘制同色系流体色块
- **CSS 流体质感**：`filter: blur(80px) contrast(1.2)` + `transform: scale(1.1)` 由浏览器 GPU 硬件加速完成放大和模糊，实现如丝顺滑的"彩色流沙/极光"质感
- **音频节奏联动**：Bass 频段（`dataArray[0] + dataArray[1]`）实时影响沙浪速度（最大 2.5 倍）和浪尖高度（`peakAmp`），鼓点强时沙浪翻涌
- **`drawFlowingSand()` 函数**：三层沙浪（暗调底层 + 正弦波中层 A + 余弦波中层 B + 顶层亮沙 C），颜色自动跟随 `currentHue`，配合 `sandPhaseA/B/C` 独立相位
- **性能零负担**：64×64 Canvas 每帧计算量几乎为 0，Windows 11 / Android 17 上接近 0% CPU 占用

### 🎨 预设主题色重构 — 聚焦卡片 + 多场景全域联动
- **卡片化重构**：旧 `.theme-preset` 色块 + `.theme-label` 分离结构 → 新 `.theme-preset-card` 一体化卡片，包含 `.theme-color-circle` 圆块 + `.theme-preset-label` 标签
- **2D 空间聚焦适配**：卡片声明为 `focusable` + `tabIndex=0`，手柄摇杆/十字键可精准导航，`.gamepad-focus` 时边框发光 + `scale(1.08)` + `--primary-glow` 阴影
- **`applyThemeColorAction()`**：统一主题应用入口，同步设置 `--primary` / `--primary-glow` / `--album-color`
- **画中画实时同步**：切换主题时 `pipWindow` 内的 `--primary` 和进度条发光色无缝实时变色
- **`hexToRgb()`**：16 进制颜色转 RGB 辅助函数，支撑 RGBA 发光计算
- **旧样式清理**：移除 `.theme-preset`、`.theme-preset .check`、`.theme-label` 等碎片化样式

### 🌈 主界面频谱彩色渐变
- **灰阶→极光**：主界面频谱柱从死板 `rgba(gray, gray, gray)` 改为 `hsla(currentHue, 75%, ...)` 动态色相渐变，幅值越高越亮越不透明
- 频谱颜色完全跟随当前专辑封面提取色或预设主题色，视觉一致性达到顶峰

### 🔧 `applyThemeLogic` 适配 Canvas
- Canvas 背景不再设置 `style.background`，只需 `classList.add('active')` 激活，颜色由 `drawFlowingSand` 实时渲染



### 🎮 2D 空间导航系统 — 手柄/键盘完美适配曲库网格
- **智能空间寻路算法 `moveFocus2D`**：取代旧版一维循环 `moveFocus`。通过 `getBoundingClientRect()` 计算每个可聚焦元素在屏幕上的物理坐标，按下方向键/摇杆时以加权欧式几何距离算法（主方向距离 + 垂直偏离惩罚系数 2.5）自动计算最近目标，实现网格化 2D 空间导航
- **焦点滚动跟随 `setFocus`**：焦点切换时自动调用 `scrollIntoView({ block: 'nearest', behavior: 'smooth' })`，确保焦点框永远可见
- **键盘 WASD + 方向键**：升级为独立方向映射——`W/↑` 向上、`S/↓` 向下、`A/←` 向左、`D/→` 向右
- **手柄摇杆 4 方向**：左摇杆 X/Y 轴独立判断方向（±0.5 死区），200ms 防抖，取代旧的 `moveFocus(-1/1)` 模糊映射
- **手柄十字键 (D-Pad)**：`btn[12]` 上、`btn[13]` 下、`btn[14]` 左、`btn[15]` 右，全部映射到 `moveFocus2D`

### 🏗️ 曲库弹窗静态化重构 — 动画 100% 统一
- **静态 HTML 化**：曲库弹窗从 `document.createElement` 动态生成改为写入 `index.html` 的 `#coverLibraryModal`，共享 `modal-overlay.open` 类名触发的 CSS 过渡动画
- **动画一致性**：打开时遮罩 `opacity: 0→1` + 面板 `scale(0.9)→scale(1.0)` 弹性弹簧动效，关闭时平滑缩小退场，与列表弹窗/设置弹窗像素级一致
- **焦点扫描升级**：`updateFocusContext` 新增 `album-detail-panel` 和 `#coverLibraryModal.open` 检测，打开曲库或专辑详情时自动切换焦点上下文
- **事件绑定防泄漏**：`modal.dataset.init` 标记确保事件只绑定一次
- **`_closeModalsSync`**：新增静态曲库的 `classList.remove('open')` 关闭逻辑

### 🎯 焦点元素全面标注
- **曲库 Tab 标签**：`cover-lib-tab` 添加 `focusable` + `tabindex="0"`
- **专辑卡片 `createCoverCard`**：`cover-lib-card` 添加 `focusable` + `tabIndex = 0`
- **专辑详情曲目行**：`album-detail-track` 添加 `focusable` + `tabIndex = 0`
- **专辑详情按钮**：`btnPlayAlbum` / `btnCloseAlbumDetail` 添加 `focusable`

### 🎤 歌词面板视觉梯队
- **下一行去模糊**：`.lrc-line.active + .lrc-line` 使用 CSS 相邻兄弟选择器，让紧跟在激活行后的那一行歌词 `filter: blur(0px) !important` + `opacity: 0.75` + `scale(0.98)`，形成清晰的视觉阶梯（当前行完全清晰 → 下一行清晰 → 其余行模糊）

### 🔧 淡入淡出引擎增强
- **`playing` 事件保护**：`triggerFadeIn` 不再盲目启动淡入定时器，改为监听 `audio.playing` 事件，确保音频真正开始播放后才逐步提升音量，防止歌曲因缓冲延迟导致爆音
- **5 秒兜底保护**：如果 `playing` 事件在 5 秒内仍未触发，强制恢复音量 + 解锁 `isFading`，防止永久静音

---

## v2.3.1 (2026-05-31)

### 🔤 全局字体优化
- **新增 OPPO Sans 4.0 优先级字体**：`body` 和 `.pip-container` 的 `font-family` 首位添加 `'OPPO Sans 4.0'`，系统已安装该字体时优先渲染，呈现更精致的文字质感

### 🏷️ 文案统一
- **「封面库」统一改名为「曲库」**：`index.html` 快捷键帮助面板、`app.js` 注释/弹窗标题/渲染函数注释等全部替换

### 🔧 细节修正
- **版权信息补充**：设置页面底部版权添加 CodeBuddy 协作署名
- **版本号同步**：`index.html` 标题、`app.js` 顶部版本号、版权区域版本号统一为 v2.3.1

---

## v2.3.0 (2026-05-31)

### 🚨 关键 Bug 修复（视图层级穿透 + 进度条冲突 + 数据丢失）

- **修复沉浸视图遮罩导致顶部按钮无法点击（终极破案）**：`#view-immersive.hidden` 仅设 `opacity: 0` + `pointer-events: none`，在部分 Chromium 核心浏览器中依然产生隐形事件拦截。现补充 `visibility: hidden`（彻底剔除渲染树）+ `z-index: -1`（强行沉底），同时 `#view-main` 和 `#view-immersive` 分别设置 `z-index: 10/20`，确保视图堆叠上下文绝对正确
- **修复进度条点击/拖拽冲突**：旧版 `setupDraggableProgress` 和 `onclick` 点击事件同时存在，鼠标松手瞬间触发 `mouseup` + `click` 两次修改进度，互相打架。现已统一为 `bindProgressBar` 引擎，整合 mousedown→mousemove→mouseup 流水线，点击即极短拖拽，完美兼容
- **修复播放整张专辑后全库数据丢失**：原 `playlist = albumQueue` 直接覆盖内存，其余歌曲永远消失，封面库随之崩溃。现引入 `musicLibrary`（全库只读容器）+ `playlist`（临时播放队列）双轨制，专辑播放仅修改 `playlist`，封面库始终从 `musicLibrary` 读取

### ✨ 新增功能

- **全库/队列双轨制**：`musicLibrary` 永久保存全部导入歌曲，封面库、搜索、统计永远访问全库；`playlist` 仅负责当前播放队列
- **一键恢复全库播放**：点击播放列表「📋 全部」时，若检测到队列被缩减（如处于专辑播放中），自动从 `musicLibrary` 恢复全部歌曲，提示「已恢复播放全部歌曲」
- **进度条沉浸模式时间同步**：`ontimeupdate` 新增 `immTimeCur`/`immTimeTot` 时间文本同步，沉浸模式进度条数字随拖拽实时更新
- **进度条拖拽防文字选中**：`handleMove` 中增加 `e.cancelable && e.preventDefault()`，防止拖拽时意外选中页面文字

### 🔧 架构优化

- `processFiles` 加载完毕时同步执行 `musicLibrary = [...playlist]`
- `showCoverLibrary` 中 `renderAlbumGrid`/`renderArtistGrid`/`renderRecentGrid` 遍历源从 `playlist` 改为 `musicLibrary`
- `showAlbumDetail` 播放专辑从 `musicLibrary[idx]` 取数据而非 `playlist[idx]`
- `audio.ontimeupdate` 增加安全检测（`if (el.progFillMain)`），防止 null 引用

---

## v2.2.3 (2026-05-31)

### 🚨 关键 Bug 修复（View Transitions 嵌套崩溃 + PiP 状态切换失效）

- **修复 View Transitions 嵌套崩溃（终极破案）**：当点击"列表"或"设置"时，`document.startViewTransition()` 回调内调用了 `closeAllModals()`，而后者再次调用 `startViewTransition()`。浏览器绝对不允许嵌套视图过渡，导致 `::view-transition` 全屏透明伪元素卡在屏幕最顶层永远不消失（"死玻璃"效应），阻挡所有鼠标点击。现已重构为 `_closeModalsSync` 纯同步关闭 + `safeTransition` 安全封装，每次最多只触发一次视图过渡
- **修复 PiP 状态切换失效**：原代码在打开画中画时用模板字符串 `${hasLrc ? ... : ...}` 写死 DOM 结构，导致切歌后状态改变（有歌词→无歌词）时找不到对应节点。现改为两套 UI 都写死在 DOM 里（`pipLyricsWrap` + `pipFallback`），`updatePipUI` 每 500ms 根据 `parsedLyrics.length` 动态切换 `display`，并强制刷新封面 `src`
- **修复 PiP 封面不刷新**：`pipBg` 和 `pipVinylWrap` 现在使用 `id` 选择器精准定位，每次定时器触发都会检查并更新 `backgroundImage` 和 `innerHTML`

### ✨ 新增功能

- **丝滑进度条拖拽**：主页和沉浸模式进度条支持鼠标拖拽和触摸滑动。拖拽时实时更新进度和时间数字，松手瞬间切入目标位置。`isProgressDragging` 防冲突标志位防止 `ontimeupdate` 和拖拽同时写入导致滑块抽搐
- **粒子爆炸反馈**：拖拽进度条松手时触发 `createExplosion`，提供视觉回馈

### 🔧 架构优化

- **CSS 层级提升**：`.header` z-index 从 100 提升至 9999，确保导航栏永远可点击
- **模态框关闭统一**：所有关闭按钮（`btnCloseFileInfo`/`btnCloseHelp` 等）统一使用 `closeAllModals`，同步清理动态生成的 cover-library/stats/detail 面板
- **事件绑定收口**：删除 load 初始化中与模态段重复的 `btnCoverLibrary`/`btnShowStats`/`btnFavQuick`/`btnPipQuick` 绑定

---

## v2.2.2 (2026-05-31)

### 🚨 关键 Bug 修复（应用假死崩溃）
- **修复 `logError` 二次赋值导致 JS 编译崩溃**：`logError` 在文件顶部已声明为 `async function`，末尾再次 `logError = ...` 触发 `TypeError: Assignment to constant variable`，导致整个 app.js 编译中断，页面完全假死。现已合并为单一函数，移除重复赋值
- **修复 `e.target.closest` TypeError**：拖拽文件到浏览器边缘或悬停文本节点时，`e.target` 不是 Element，调用 `.closest()` 抛出异常。所有 `closest` 调用增加可选链和安全类型检测
- **强化分批加载 try-catch 屏障**：首批/剩余批次的 `Promise.all` 增加 try-catch，单个批次解析失败不再阻塞后续加载，所有歌曲都能被加载

### 🔧 架构优化
- **统一事件绑定**：原 `index.html` 底部内联 `<script>` 中的 `typeof xxx === 'function'` 脆弱绑定全部迁移到 `app.js` 的 `load` 初始化中，确保加载时序一致，消除函数未定义的竞态风险
- `toggleFavorite` 暴露到 `window.MBolka` 命名空间

---

## v2.2.1 (2026-05-31)

### 🎛 UI/UX 改进
- **收藏❤️和画中画📺按钮移至中心控制区**：从顶部 header 移到 `btn-group-main`，分别放在上一曲左侧和下一曲右侧，按钮风格改为圆形控件（`.btn-ctrl`）
- **CSS 层级修复**：`.header` z-index 提升至 100，添加 `pointer-events: auto`；`.load-strip-container` 添加 `z-index: 5`
- **画中画重构**：PiP 内部按钮直接绑定主窗口函数引用（`goPrev()` / `togglePlay()` / `goNext()`），不再依赖 BroadcastChannel；样式表复制改用 `document.styleSheets` 逐一拷贝 CSS rules
- **文件加载增强**：排除系统隐藏文件（`.` 和 `._` 开头）；解析超时熔断 1.5 秒自动降级

### ✨ 动效进阶
- **歌词动态模糊**：非激活行 `filter: blur(2px)` + `scale(0.95)`，激活行完全清晰 + 放大，过渡曲线 `cubic-bezier(0.2, 0.8, 0.2, 1)`
- **专辑环境光阴影**：通过 `--album-color` CSS 变量动态设置封面阴影颜色，配合 `ambientBreathe` 呼吸动画
- **按钮微距回馈**：`.btn-glass` 和 `.btn-ctrl` 添加 `:active { transform: scale(0.92) }` 物理按压感
- **View Transitions API**：弹窗打开/关闭均用 `document.startViewTransition()` 包裹，获得原生 App 级展开/收起动画

---

## v2.2.0 (2026-05-31)

### 🔧 Bug 修复与体验优化
- **移除专辑封面滑动切歌提示**：`← 滑动切歌 →` 文字已移除，不再干扰封面观感
- **修复加载文件夹卡顿**：改为并发批处理（6首/批）+ `setTimeout` 让出主线程，避免UI冻结
- **拖拽场景限制**：明确仅在主界面空白区/空状态区允许拖入文件夹，专辑封面/按钮/模态框区域禁用
- **拖拽排序视觉反馈**：播放列表拖拽时显示蓝色插入线，精确定位插入位置
- **加载条优化**：渐变主色 + 圆角末端，从左到右更直观
- **睡眠定时器倒计时**：底部状态栏实时显示剩余分钟:秒数，最后1分钟红色闪烁
- **A-B 段落重复视觉标记**：进度条上显示红色 A/B 标记点和半透明区间范围
- **空状态引导增强**：脉冲呼吸动画引导用户，非Chrome浏览器提示使用按钮

### 🖼️ 封面库全面增强
- **多维度聚合切换**：支持按专辑 / 按艺术家 / 最近添加三种视图
- **专辑详情面板**：点击专辑卡片展开，显示大封面 + 完整曲目列表 + "播放整张专辑"按钮
- **黑胶唱片动效**：Hover时从封面侧边滑出旋转的黑胶唱片
- **艺术家视图**：圆形头像展示，点击直接播放

### 📺 画中画全面重构
- **动态模糊背景**：专辑封面提取 + `blur(50px)` + 呼吸动画，色调随歌曲变化
- **两行歌词排版**：当前行大字高亮带文字发光，下一行小字半透明，切换时 `translateY` 淡入淡出过渡
- **悬停控制栏**：默认纯净歌词，鼠标悬停浮现控制按钮（上一首/播放暂停/下一首/快速收藏）
- **极简进度条**：底部 3px 主题色进度条，随播放实时推进
- **无歌词降级UI**：纯音乐/无歌词时显示旋转黑胶唱片 + 歌名/艺术家
- **响应式形态适配**：宽扁形态切换为单行横排布局，竖排形态恢复居中两行
- **CSS样式同步**：自动克隆主界面 `<style>` 和 `<link>` 到 PiP 窗口

### ⚡ 性能与健壮性
- **粒子性能自适应**：FPS 实时监测，低于 30 时自动减少粒子数
- **内存泄漏修复**：加载新文件夹前释放旧 Blob URL，防止内存爆炸
- **命名空间封装**：`window.MBolka` 暴露核心 API
- **错误日志持久化**：全局 `window.onerror` 捕获 + localStorage 持久化 + 设置中一键导出
- **并发加载优化**：`Promise.all` 批量解析元数据，加载速度提升约 3 倍

### 🎛 细节打磨
- **歌词偏移精细调**：新增 ±0.1s 按钮，实现精确校准
- **首页金刚键**：收藏 ❤️ 和画中画 📺 快捷按钮添加在导航栏两侧
- **收藏状态联动**：首页收藏按钮与播放列表收藏状态实时同步

---

## v2.1.0 (2026-05-31)

### 🔧 Bug 修复
- **修复沉浸Canvas残留**：退出沉浸模式后，Canvas粒子频谱不再卡在背景中
- **频谱改为灰阶**：彩色频谱改为现代简约的灰阶设计，更符合设计调性
- **修复右上角按钮遮挡**：调整z-index层级，确保四个操作按钮始终可点击
- **空态页拖拽修复**：空状态不再干扰拖拽文件夹功能
- **滑动切歌提示优化**：减少对专辑封面观感的影响

### 📁 项目结构优化
- **CSS/JS拆分**：将原单文件拆分为 `index.html` + `css/style.css` + `js/app.js`
- 代码组织更清晰，便于维护和扩展

### 🔒 本地库强化 (Library & Storage)
- **目录句柄持久化**：使用 File System Access API 的 `showDirectoryPicker()` + IndexedDB 持久化目录句柄，下次打开网页自动恢复音乐库
- **全文搜索**：播放列表顶部增加搜索框，支持对标题/艺术家/专辑/文件名的毫秒级实时搜索
- **播放列表导出导入**：支持导出为 `.m3u` 和 `.json` 格式，方便备份
- **音乐统计看板**：记录每首歌的播放次数和总听歌时长，展示Top10最爱歌曲

### 🎛 音频与硬核播放控制
- **十段均衡器 (EQ)**：基于 Web Audio API 的 `BiquadFilter` 实现，提供 8 种预设（Flat/Pop/Rock/Classical/Vocal/Bass/Electronic/Jazz），支持手动调节
- **播放速度与升降调**：0.5x~2.0x 变速播放，支持保持音调/允许变调切换
- **淡入淡出切歌**：可配置的 Crossfade（1-8秒），曲末自动渐弱渐强
- **睡眠定时器**：15/30/60 分钟倒计时，自动停止播放

### 📝 歌词增强
- **内嵌歌词解析**：自动读取 FLAC/MP3 文件内嵌的 USLT/SYLT 歌词标签
- **歌词时间轴微调**：+0.5s / -0.5s 按钮，动态修正 LRC 偏移

### 🖼️ 封面库独立
- 从播放列表中完全独立为专属功能模块（按 G 键或点击右上角🖼️按钮）
- 网格化展示，支持搜索专辑/艺术家，点击直接播放

### 📺 画中画迷你播放器
- 使用 Document Picture-in-Picture API，将播放器变为系统级悬浮窗
- 窗口置顶，包含专辑封面、歌名、控制按钮、单行歌词

### 📱 移动端与PWA
- **PWA 支持**：动态注入 manifest.json + Service Worker，可安装到桌面
- **移动端手势**：双击左侧快退10秒、双击右侧快进10秒、上下滑动调音量
- **沉浸模式左右长滑切歌**
- **移动端竖版视图全面优化**：按钮尺寸、间距、字号自适应

### ⚡ 性能优化
- **节能模式**：限制Canvas渲染帧率为30fps，降低设备发热
- 性能/全性能模式一键切换

### 🎵 Media Session 增强
- 实时同步 `PositionState`（进度条），锁屏/控制中心可拖动进度

### ⌨️ 新增快捷键
- `T` - 统计面板
- `G` - 封面库
- `Q` - 画中画

---

## v2.0.1 (2026-05-31)

### 🎛️ 交互优化
- **合并播放模式按钮**：将独立的「单曲循环」按钮合并到模式切换按钮中，现在是一个三段式按钮：`顺序 → 随机 → 单曲循环 → 顺序`，按 `M` / `R` / `S` 或手柄 `X` 键即可循环切换
- **修复右上角按钮被遮挡**：调整了标题栏的 z-index 层级，确保右上角「列表」「歌词」「设置」「载入音乐」按钮始终可点击

### 🎨 沉浸模式可视化重做
- **全新多层次视觉系统**：
  - **层次1 - 极光光晕**：动态色相变化的径向渐变背景
  - **层次2 - 中心发光核心**：随贝斯强度呼吸的发光体
  - **层次3 - 底部频谱弧线**：优雅的波形弧线（双线叠加+发光）
  - **层次4 - 两侧对称频谱柱**：渐变色的圆角频谱柱，顶部有光点
  - **层次5 - 顶部细线频谱**：微妙的高频指示线
  - **层次6 - 散布光点**：漂浮的"音符感"光点，跟随频谱闪烁
- **粒子系统增强**：粒子数量上限提升至 120，鼠标跟随生成更密集，点击空白区生成涟漪
- **颜色过渡更平滑**：取色模式和自由循环模式的色相过渡都更流畅

### 🖼️ 封面库独立
- **全新封面库视图**：从播放列表中独立出来，按封面/专辑自动聚合
  - 相同封面的歌曲归为一个专辑卡片
  - 显示封面缩略图、专辑名、歌曲数量
  - 无封面的歌曲统一归入「无封面」分组
  - 点击直接播放该分组第一首
  - 播放列表新增「全部」「收藏」「封面库」三个切换按钮

### ✨ UI/UX 细节优化
- **主界面频谱条**：改为彩色渐变，视觉更丰富
- **专辑封面滑动**：增加左右滑动动画效果
- **专辑信息增强**：解析并显示歌曲的专辑名称
- **文件信息面板**：新增专辑和时长显示
- **按钮交互反馈**：悬停和激活状态的视觉过渡更流畅
- **版本号更新**：v2.0.0 → v2.0.1

### 🐛 Bug 修复
- 修复右上角操作按钮被空白状态遮罩层遮挡的问题
- 修复播放模式按钮逻辑混乱（两个按钮控制同一状态）的问题

---

## v2.0.0 (2026-05-30)

### 🚀 技术优化
- **Web Worker 解析元数据**：在后台线程解析音乐标签，不阻塞 UI
- **IndexedDB 元数据缓存**：解析结果持久化，二次加载秒开
- **媒体会话增强**：完善 Media Session API，支持 seekto/seekbackward/seekforward
- **CUE 分轨支持**：解析 .cue 文件提取曲目信息和时间点
- **虚拟滚动支持**：大量音乐时使用虚拟滚动渲染
- **错误日志双存储**：IndexedDB + localStorage 记录播放异常

### 🎵 核心体验
- **播放错误容错**：解码失败自动跳下一首，列表中标红显示
- **播放列表管理**：右键菜单支持删除单曲、清空列表、查看文件信息
- **A-B 段落重复**：长按播放按钮进入，进度条点击设置 A/B 点
- **沉浸模式增强退出**：双击空白区、手势下滑、底部箭头三种方式

### 🖱️ 交互细节
- **进度条悬停预览**：显示该时间点的歌词片段
- **滑动切歌**：专辑封面左右滑动切换上下曲
- **播放队列拖拽排序**：HTML5 Drag & Drop 实时重排
- **歌词字体/对齐调节**：字号、行距、对齐方式滑块

### 🎨 视觉增强
- **10 套预设主题色**：赛博朋克、暖阳、极光、星夜、樱花等
- **专辑封面瀑布流**：封面墙视图
- **沉浸模式粒子互动**：鼠标/触摸粒子跟随散开，高潮爆炸
- **动态壁纸联动**：频谱实时影响背景光晕和粒子颜色

### ⌨️ 其他
- **快捷键大全面板**：按 `?` 键弹出，分类展示所有操作
- **手柄完全支持**：Xbox/PS 手柄全功能映射
