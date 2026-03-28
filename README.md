# 亚马逊商品处理与自动提交工具

一个面向亚马逊上新流程的本地工作台，目标是把“采集数据 -> AI处理 -> 缺项诊断 -> Amazon 预览 -> 正式提交”串成一条可视化、可复查、可持续优化的闭环。

当前版本已经不是单纯的 Excel 处理脚本，而是包含以下三条主链路：

- 传统链路：导入 Excel -> AI 文案/图片 -> 校验/诊断 -> 提交
- 官方模板链路：自动推荐类目 -> 生成美国站官方字段模板 -> 上传采集结果 -> 自动诊断
- 任务中心链路：模板生成、模板诊断、Amazon 预览、正式提交全部进入任务中心追踪

## 当前核心能力

### 1. AI 内容与图片处理
- 标题、五点、描述、搜索词生成
- 主图/副图 AI 处理
- 单 SKU 与批量处理
- 处理结果回写 Excel
- 单 SKU 改图使用唯一输出路径，避免覆盖旧图片

### 2. Amazon 官方字段模板
- 根据链接 / 标题 / 关键词推荐美国站 `product_type`
- 按 Amazon schema 生成专属模板 Excel
- 模板包含：
  - 第一行：中文说明 / 必填提示 / 示例 / 默认值
  - 第二行：稳定字段名，便于第三方采集插件回填
- 支持单品模板与父子变体模板
- 模板会记住历史 Amazon preview 暴露出的真实缺字段，并在后续模板中提前提升显示

### 3. 自动诊断与提交流程
- 模板上传后自动执行：
  - 模板字段完整性检查
  - 本地规则校验
  - Amazon preview
- 诊断结果会回写 Excel，并同步显示到工作台
- 阻断项会前移到前端，不能提交的 SKU 不再误导用户
- 正式提交前会再次做门禁检查

### 4. 任务中心与进度
- 模板生成：任务化
- 模板诊断：任务化
- Amazon 预览：任务化
- 正式提交：任务化
- 支持阶段式进度展示，例如：
  - 读取商品文件
  - 本地字段校验
  - 调用 Amazon 预览
  - 提交门禁检查
  - 调用 Amazon 正式提交

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 配置环境变量
copy .env.example .env

# 3. 启动 Web 工作台
python web/app.py
```

默认访问：

```text
http://127.0.0.1:5000
```

## 推荐使用方式

### 方式 A：已有 Excel
1. 导入 Excel
2. 跑 AI 文案 / AI 生图
3. 点“缺项诊断”
4. 点“预览验证”
5. 通过后再正式提交

### 方式 B：官方模板先行
1. 在工作台顶部“官方字段模板”区域输入链接 / 标题 / 关键词
2. 识别类目并生成官方模板
3. 下载模板给采集插件或人工填充
4. 上传模板结果
5. 系统自动诊断并给出阻断项
6. 修完阻断项后再预览 / 提交

## 目录结构

```text
亚马逊2.8/
├── amazon/                  # Amazon SP-API 与字段映射
├── core/                    # AI、模板、媒体、Excel 核心逻辑
├── web/                     # Flask Web 工作台
├── config/                  # 字段配置
├── docs/                    # 补充文档
├── changelogs/              # 历史阶段性记录
├── input/                   # 输入文件
├── output/                  # 输出文件 / schema 缓存 / 模板缓存
├── stage1_pipeline.py       # AI 处理链路
├── stage2_pipeline.py       # SP-API 提交链路
└── test_*.py                # 自动化测试
```

## 主要依赖

- Python 3.10+
- Flask
- openpyxl
- Pillow
- requests
- boto3
- python-amazon-sp-api

## 当前已知限制

- 第一版官方模板链路固定以美国站为主
- 来源商品链接解析目前仍以“标题/关键词推荐 product_type”为主，不是完整 1688 抓取器
- 任务中心已支持阶段进度，但长期历史沉淀和断点恢复还可以继续增强
- 临时匿名图床脚本仍保留，仅供测试，不建议作为正式生产方案

## 发布与版本规则

后续每次准备推送 GitHub 时，统一执行以下规则：

1. 先更新 `README.md`
   - 补充当前可用能力
   - 标注新增主流程或关键变更

2. 再更新 `CHANGELOG.md`
   - 用 `Unreleased` 维护当前待发布内容
   - 推送前把本次更新整理成版本条目

3. 最后打 tag / 写 release notes
   - tag 描述必须写清楚“这次更新了什么”
   - release notes 内容以 `CHANGELOG.md` 对应条目为准

## 测试

建议常用回归命令：

```bash
python -m pytest -q
python -m pytest -q test_web.py test_web_flow.py
```

## 临时匿名图床

`upload_ai_images_temp.py` 仍保留，仅供少量测试：

```bash
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --in-place
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --overwrite
```

注意：这不是正式媒体托管方案。
