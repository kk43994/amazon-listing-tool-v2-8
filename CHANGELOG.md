# Changelog

本文件记录准备推送到 GitHub / 准备打 tag 时的版本摘要。

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
