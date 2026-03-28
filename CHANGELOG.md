# Changelog

本文件记录准备推送到 GitHub / 准备打 tag 时的版本摘要。

## Unreleased

- 暂无

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
