# Changelog

本文件记录准备推送到 GitHub / 准备打 tag 时的版本摘要。

## Unreleased

## v0.5.1 - 2026-05-18

### 修复

- 修复部分 Amazon 账号在 LWA Token 已通过、Listings 探测返回 `400 Invalid parameters provided` 时被误判为账号不可用的问题；现在会标记为“基础连接通过”，真实上架能力留到导入商品后的预览/提交阶段继续校验。
- 修复官方类目模板获取时，部分卖家账号调用带 `sellerId` 的 Product Type Definitions 返回 `InvalidInput` 导致模板生成失败的问题；系统会自动回退到站点级 schema。
- 修复账号列表和账号测试弹窗只显示“通过/失败”，无法区分“基础通过”和“完整上架探测通过”的状态提示。

### 验证

- 已验证 `CELLULAR_PHONE` 官方模板可正常生成。
- 已补充账号测试软通过和 Product Type Definitions 回退逻辑的自动化测试。

## v0.5.0 - 2026-04-27

本次重点：让**完全不懂技术的商家**也能自己装、自己配、自己上架，不再依赖售后远程协助。配套强化了价格 / 主图 / ASIN 重复 / 文件版本冲突等提交前的硬拦截，避免错单流到 Amazon。

### 新增 - 新手上手零门槛

- **AI Key 自助申请引导**：设置页 API 配置卡片顶部新增"还没有 Key？4 步注册自取"醒目卡片，把"注册账号 → 模型广场找模型 → APIkey 管理拿密钥 → 创建分组绑定模型 → 测试连接"5 个步骤分块写清楚，每一步配「打开页面 / 复制网址 / 找客服」按钮，整个流程不再需要客服解释。
- **客服微信兜底入口**：买 Key / 拿凭证任何一步卡住，都可以一键弹出客服二维码（图片放在 `web/static/img/vendor-qr.png`，售后自行替换；图片缺失会显示友好占位文案）。
- **亚马逊凭证 5 步图文教程**：亚马逊账号设置卡片新增"📖 不会拿凭证？看图文教程"按钮，弹出可勾选任务清单：登录 Seller Central → 拿 Seller ID / Merchant Token → 注册 SP-API 开发者应用拿 LWA Client ID / Secret → 自我授权拿 Refresh Token → 对照填写 4 个字段。每一步带有"在哪里找""注意事项""失败排查"，每勾完一步会自动变绿，结束后引导一键跳转添加账号弹窗。
- **字段术语鼠标悬停解释**：商品表列头、列设置面板、抽屉编辑区、变体家族编辑器、变体矩阵、新增账号弹窗等 30+ 处出现 `?` 小图标，鼠标悬停或键盘 Tab 聚焦即弹出中文解释卡片。覆盖 50+ 个字段，包括 `product_type / parentage_level / parent_sku / variation_theme / item_name / bullet_point_X / generic_keywords / main_image_url / fulfillment_channel / external_product_id_type / Seller ID / Marketplace ID / LWA Client ID / Refresh Token / 预览状态 / 上架状态 / 提交ID / 上架路线 / 最终状态 / 后台问题数` 等。术语字典集中在 `_js.txt` 的 `FIELD_TOOLTIPS_ZH`，需要更新文案直接改这一处即可。
- **公司网络代理配置**：`.env.example` 新增 `HTTP_PROXY` / `HTTPS_PROXY` / `NO_PROXY` 字段，方便公司内网/防火墙环境的客户填写。Doctor 自检会显示当前代理状态，外网检查全部失败且未配代理时主动提示"请联系 IT 拿代理填到 .env"，不再让客户对着 `TCP timeout` 发懵。
- **首次启动安全弹窗友好提示**：Mac `.command` 启动脚本头部新增大字提示，告诉客户首次打开遇到「无法验证开发者」时如何"右键 → 打开"过 Gatekeeper；Windows `.bat` 启动脚本同样补充了「Windows 已保护你的电脑」过 Defender 与杀毒软件加白名单的引导。

### 新增 - 提交前硬拦截（避免错单流到 Amazon）

- **价格成必填且强格式校验**：未填价格的子体不再放行，统一要求填写；新增 `_parse_numeric` 容错带美元符号 / 中英文逗号 / 全半角空格的写法，并设置 `MAX_SUBMISSION_PRICE = 10000` 上限，超出范围会在本地拦截并提示具体是哪一行。
- **主图缺失从警告升级为错误**：新建商品和子体没有主图 URL 时不允许预览/提交（父体或仅更新 offer 不强制），避免商品页空主图。
- **ASIN 重复检测**：导入文件中如果同一个 ASIN 被多个 SKU 占用，工作台会高亮提示并阻止跨行混用同一 ASIN。
- **Excel 文件版本冲突保护**：所有写回路径会比对客户端提交时的文件 mtime 与磁盘上的 mtime，发现期间被 Excel/WPS 改动会拒绝覆盖并提示"先关掉 Excel 再操作"，避免你这边正在保存、客户那边正在改导致的串数据。
- **Stage1 自动保存中间结果**：批量 AI 处理的每一段都会写入 `*.autosave.xlsx`，意外断电 / 关浏览器 / 杀进程后下次还能从断点续跑，不用整批重头来。

### 优化 - 售后/客服更省事

- **Doctor 报告网络段更细**：每条网络检查项细分 DNS / TCP / TLS / HTTP 4 步，失败时直接告诉客户是哪一层断；总结段会综合判断"全失败 + 未配代理 → 试试代理"还是"DNS 没问题但 TCP 不通 → 检查防火墙"。
- **设置页"获取 Key"卡片采用醒目橙色高对比设计**，与白色背景的输入框形成视觉阶梯，新手第一眼就能看到下一步该做什么。
- **凭证教程抽屉每完成一步会自动打钩变绿、半透明**，长流程不会让客户中途忘记自己做到哪一步。
- **客户向文档全程使用"商家""上架"等业务语言**，不出现 `endpoint` / `OAuth` / `marketplace_id` 等技术黑话，遇到必须出现的代号都用 `?` 提示卡解释。

### 修复

- 修复部分客户提交时把 `19.99 美元` 这种带单位的字符串原样发给 Amazon 导致 `INVALID_STANDARD_PRICE` 报错的问题（统一走 `_parse_numeric`）。
- 修复 v0.4.5 之前主图缺失只报 warning 但仍能继续提交，导致上架后图片为空的问题。
- 修复 stage1 批量 AI 跑到一半中断后，没自动保存导致已生成的内容白丢的问题。
- 删除上版本被误归档的根目录 `_js.txt`（与 `web/templates/_js.txt` 内容不同步的旧快照），仓库结构更干净。

### 升级建议

- 老客户升级建议直接覆盖整个解压目录；`.env`、`accounts.json`、`input/`、`output/`、`backups/`、`logs/` 不会被发行包覆盖，配置和历史数据完整保留。
- 升级后第一次启动会自动跑环境检测；如果 Doctor 报代理相关 WARN，按提示编辑 `.env` 加 `HTTP_PROXY` 即可。
- 上版本未填价格也能跑预览的"幸运用例"在本版本会本地拦截，请补齐价格后重试。

## v0.4.5 - 2026-04-28

### 新增
- Web 工作台新增首次使用引导和设置状态面板，按“AI 接口 -> Amazon 账号 -> 导入商品表 -> 一键自检”引导客户完成配置
- 新增轻量级 `/api/setup-status`，不触发外部 API 调用即可判断客户下一步该做什么
- 发行包新增 `客户先看这里.txt`，用客户能看懂的方式说明启动、配置、自检和处理顺序
- 发行包新增 `一键检测修复.command/.bat` Doctor，自动补齐基础运行文件并检查内置依赖、端口、浏览器、AI 中转域名和 Amazon 网络
- 发行包新增 ASCII 备用入口 `Start-Amazon-2.8`、`Doctor`、`Env-Check`、`Support-Bundle` 和 `Read-Me-First.txt`，避免中文文件名在部分电脑上乱码或被拦截
- 新增 `VERSION`、`release-manifest.json`、`dependency-inventory.json`、zip SHA256 和 `/api/version`，售后可精确确认客户运行版本
- Doctor 新增支持包导出，生成已脱敏的 `logs/support-bundle-*.zip`，包含诊断报告、日志、任务历史、版本/依赖清单、脱敏 `.env` 和 `accounts.json`
- 设置页新增配置网址导航、API 推荐一键填写和 Amazon 账号字段解释，客户可按提示填写 `https://api.kk666.best` 与默认模型
- 设置页新增 AI 简单/高级模式、推荐配置后端恢复接口、支持信息复制和支持包下载
- 工作台新增更醒目的“只需点这个”下一步按钮，并按“未配置 AI / 未配置 Amazon / 未导入文件 / 筛选为空”显示专门空状态
- 设置页拆分为 AI 接口、Amazon 账号、环境自检、高级设置四个标签，减少客户在一页里迷路
- Amazon 账号测试新增分项结果和最近测试记录，显示 LWA Token、站点映射、Listings API、权限分别是否通过
- 新增客户发版清单 `docs/RELEASE_CHECKLIST.md`，覆盖本地检查、CI、干净机器、客户流程和售后资料

### 优化
- 客户启动脚本增加“不关闭窗口”和“浏览器未打开时复制本地地址”的提示
- Doctor 网络检测拆分 DNS/TCP/TLS/HTTP，并识别云同步目录、临时目录、发行包缺失 `_internal` 等常见客户环境问题
- Web 后端增加 AI Base URL 与 Amazon 账号格式校验，客户模式下阻止旧域名或缺少 `{model}` 的错误配置
- 工作台风险按钮会根据 AI、Amazon 账号、Excel 导入状态直接锁定，并在 tooltip 告诉客户下一步
- Excel 写回改为“写临时文件 -> 校验可读 -> 原子替换”，被 Excel/WPS 占用时统一提示 `EXCEL_LOCKED`
- 源码离线部署检测到本地 wheelhouse 时跳过 pip 在线升级，并安装 release 构建依赖，避免无网络新电脑卡住
- GitHub Actions 增加前端 `node --check`，并通过 macOS/Windows 矩阵验证发行包根目录与打包程序自检
- `.env.example` 默认 AI Base URL 改为 `https://api.kk666.best`

## v0.4.4 - 2026-04-27

### 修复
- 将 `.env.example` 纳入仓库，确保 GitHub Actions 和新电脑源码部署都能生成默认配置

## v0.4.3 - 2026-04-27

### 修复
- 发行工作流改跑稳定非 Web 测试集，避免依赖本地模板缓存的 Web 测试阻塞客户包构建
- 发行工作流在 PyInstaller 构建后执行打包程序自身的 `--env-check` 烟测

## v0.4.2 - 2026-04-27

### 修复
- 发行版启动入口也强制使用 UTF-8 输出，避免 Windows 非 UTF-8 控制台启动提示乱码或报错

## v0.4.1 - 2026-04-27

### 修复
- GitHub Actions 发行构建安装 `pytest`，避免干净 runner 上测试步骤缺依赖
- 环境检测脚本输出强制使用 UTF-8，避免 Windows runner / 客户终端中文输出编码失败
- 发行版启动入口输出强制使用 UTF-8，避免客户直接运行 exe 时中文提示编码失败

## v0.4.0 - 2026-04-27

### 新增
- 新增客户交付发行包防呆能力：启动脚本先执行环境检测，通过后再启动 Web 工作台
- 新增 `tools/environment_check.py`，检测系统、依赖、目录权限、端口、AI 配置和 Amazon 账号配置
- 新增 `tools/bootstrap.py`，支持新电脑源码部署一键创建虚拟环境、安装依赖、生成配置和启动脚本
- 新增 `tools/build_dependency_bundle.py`，按平台整理离线 wheelhouse 依赖包
- GitHub Actions 同时构建应用发行包和离线依赖包，并在推送 `v*` tag 时上传到 GitHub Release
- 发行包新增 `环境检测.command` / `环境检测.bat`、客户快速开始说明和 `docs/客户部署说明.md`

### 优化
- 发行版端口 `5000` 被占用时自动尝试后续端口
- 发行包在根目录额外放置 README、客户说明、模板和 `.env.example`，避免客户翻找 `_internal`
- `.gitignore` 忽略客户输入、发行产物、离线依赖包和 macOS 元数据，避免误提交客户数据

### 修复
- 修复测试中替换 `load_dotenv` 后 `Config.reload()` 因参数不兼容失败的问题

### 测试
- 当前全量测试：`174 passed`
- 本地已验证 macOS arm64 发行包可执行 `--env-check`

## v0.3.0 - 2026-04-03

### 新增
- 搜索词合规校验：字节限制按站点区分 (US/EU=250, JP=500, IN=200)，空格和标点不计入
- 搜索词自动去重：生成后移除标题和五点中已有的单词
- 标题合规校验：禁用字符检测 (! $ ? _ { } ^ ~ ¬ ¦)、同词重复检测（最多2次）
- 标题自动修正：生成后自动删除禁用字符、去除多余重复词
- 五点描述合规校验：禁止价格、运费、退款承诺、公司信息、emoji
- 图片合规校验：生成后校验最长边 ≥ 1000px、主图白底采样校验
- 图片自动转 JPEG：Amazon 推荐格式
- 新增 4 个合规校验模块：search_term_utils / title_validation / bullet_validation / image_validation

### 修复
- 搜索词超限从 warning 提升为 error（Amazon 会拒绝整条搜索词，不是截断）
- 搜索词 Prompt 支持动态字节限制 `{byte_limit}`

### 测试
- 当前全量测试：`140 passed`（v0.2.0 为 105）
- 新增 35 个合规校验单元测试

## v0.2.0 - 2026-04-03

### 新增
- 新增缺项诊断链路：本地校验 + Amazon 预览 + 媒体连通性，按 SKU 汇总阻断项
- 新增变体家族编辑：批量更新同一家族多个 SKU、追加子体行
- 新增工作流状态自动重置：编辑字段后自动清除旧的诊断/预览/提交状态
- 新增 Gemini generateContent 协议支持，文本与图片 AI 配置完全独立
- 新增多站点语言感知，AI 提示词根据目标 marketplace 自动适配语言
- 新增 ASIN 回查：正式提交后自动补查 Listings API 获取 ASIN
- 新增 AI 目标受众、主题关键词、特殊功能等字段生成
- 新增图片流式下载与大小限制，防止恶意/超大文件撑爆内存
- 新增任务历史持久化到磁盘，应用重启后可恢复历史记录
- 新增凭证安全加固：accounts.json 和 .env 写入后自动收紧文件权限

### 安全修复
- 所有接受客户端文件路径的 API 端点统一通过 `_validate_file_path()` 做目录白名单校验
- 修复 `/api/submit` 异常处理中变量未定义导致的 `NameError` 崩溃
- `AmazonAuth.get_access_token()` 加双重检查锁，解决多线程并发刷新 token 的竞态条件
- `task_status['running']` 读取改为 `get_task_status()` 走锁，消除 TOCTOU 竞态

### 修复
- 修复 `generate_copy` 清理 markdown 代码块时无换行导致的 `IndexError`
- 修复 `_submit_batch` 全部 preflight 失败时 `failed` 统计漏算 `skipped` 数量
- 修复 `keywords.encode()` 对非字符串类型值的 `AttributeError`
- 清除 12 处 `request.json or {} or {}` 冗余表达式

### 优化
- 字段映射器新增 `bottoms_size` / `item_depth_width_height` / 变体维度校验
- Listings API 新增 `probe_connection`、`resolve_submission_asin`、token 过期自动重试
- 认证模块新增 `AmazonAuthError` / `AmazonTokenError` / `AmazonNetworkError` 异常体系
- Feed 处理报告解析增强，兼容多种 Amazon 返回格式
- Stage1 并发模型优化：文案/图片独立信号量 + 可取消
- `requirements.txt` 补充 `httpx` 依赖声明

### 测试
- 当前全量测试：`105 passed`
- 三轮代码审阅 + 回归验证全部通过

## v0.1.0 - 2026-03-28

### 新增
- 新增“官方字段模板”链路：类目推荐、模板生成、模板下载、模板上传自动诊断
- 新增模板经验缓存：Amazon preview 暴露过的真实缺字段，会沉淀到同类模板中
- 新增 `submit-task` 异步接口，Amazon 预览与正式提交进入任务中心

### 优化
- 工作台前移提交门禁：有模板阻断项的 SKU 不再误导性显示为可提交
- 模板诊断结果同步回写 `preview_status / preview_message / preview_time`
- 模板阻断项、缺项诊断、预览状态在工作台/抽屉内显示更一致
- 提交/预览进度改成阶段式中文提示

### 修复
- 修复模板字段与真实 Amazon 阻断字段不一致的问题，显式支持 `supplier_declared_dg_hz_regulation`
- 兼容旧别名 `hazmat_declaration`，统一映射到真实 Amazon 字段
- 修复单 SKU 改图输出路径冲突问题，生成文件改为唯一命名
- 修复模板诊断写回后的系统字段污染 Amazon 载荷问题

### 测试
- 当前全量测试：`66 passed`
- Web / 工作台回归测试：通过

## Tag 规范

每次准备推送并打 tag 时：

1. 先把 `Unreleased` 整理成版本段落，例如 `## v0.3.0 - 2026-03-28`
2. Git tag 注释与 GitHub release notes 直接使用该版本段落
3. 不允许只打 tag 不写更新说明
