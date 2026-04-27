# 亚马逊商品处理与自动提交工具

一个面向亚马逊上新流程的本地工作台，目标是把“采集数据 -> AI处理 -> 缺项诊断 -> Amazon 预览 -> 正式提交”串成一条可视化、可复查、可持续优化的闭环。

当前版本已经不是单纯的 Excel 处理脚本，而是包含以下三条主链路：

- 传统链路：导入 Excel -> AI 文案/图片 -> 校验/诊断 -> 提交
- 官方模板链路：自动推荐类目 -> 生成美国站官方字段模板 -> 上传采集结果 -> 自动诊断
- 任务中心链路：模板生成、模板诊断、Amazon 预览、正式提交全部进入任务中心追踪

## 当前核心能力

### 1. AI 内容与图片处理
- 标题、五点、描述、搜索词、目标受众、主题关键词、特殊功能生成
- 支持 OpenAI 兼容接口与 Gemini generateContent 双协议
- 文本/图片 AI 配置完全独立，可分别使用不同模型和 endpoint
- 主图/副图 AI 处理，支持白底/生活场景/渐变等多种背景风格
- 单 SKU 与批量处理，并发模型支持文案/图片独立信号量
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
# 1. 创建虚拟环境
python3.11 -m venv .venv
source .venv/bin/activate

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置环境变量
cp .env.example .env

# 4. 启动 Web 工作台
python web/app.py
```

默认访问：

```text
http://127.0.0.1:5000
```

## 发行构建

当前仓库已经补上了面向终端用户的发行构建能力。

### 客户交付包

客户不需要安装 Python、pip 或 `requirements.txt`。推荐只交付 GitHub Release 里的可执行压缩包：

- macOS Apple Silicon：`AmazonListingTool-darwin-arm64.zip`
- Windows 64 位：`AmazonListingTool-windows-amd64.zip`

客户解压后双击 `启动亚马逊2.8.command` / `启动亚马逊2.8.bat` 即可。启动脚本会先运行环境检测，自动创建 `.env`、`accounts.json`、`input/`、`output/`、`logs/`，并在端口被占用时自动切换到后续端口。

如果客户反馈打不开，让客户双击 `环境检测.command` / `环境检测.bat`，把检测结果截图发给技术支持。详细交付说明见 `docs/客户部署说明.md`。

### 本地构建

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt -r requirements-release.txt
python tools/build_release.py
```

构建完成后会生成：

- `release/AmazonListingTool-<platform>/`
- `release/AmazonListingTool-<platform>.zip`

如需给内部售后准备源码离线依赖包：

```bash
python tools/build_dependency_bundle.py
```

构建完成后会生成 `release/AmazonListingTool-dependencies-<platform>.zip`。该包只用于源码部署，不需要发给普通客户。

### Windows 本地构建

在 Windows 上执行：

```bat
py -3.11 -m venv .venv
.venv\Scripts\pip install -r requirements.txt -r requirements-release.txt
tools\build_release.bat
```

### GitHub Actions 构建

仓库内置了 GitHub Actions 工作流：

- `.github/workflows/build-release.yml`

支持：

- `workflow_dispatch` 手动触发
- 推送 `v*` tag 时自动触发

工作流会分别构建：

- macOS arm64 发行包
- Windows amd64 发行包

并把 zip 包作为 Actions Artifact 上传。

### 新 Mac 初始化建议

1. 使用 Python 3.10+，当前推荐 `python3.11`
2. 首次拉仓后执行：

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

3. 准备两个敏感配置文件：
   - `.env`：AI 配置、运行时参数、CLI 用的 Amazon 凭证
   - `accounts.json`：Web 工作台默认使用的 Amazon 账号
4. 确保以下目录存在：
   - `input/`
   - `output/`
   - `logs/`

内部源码一键初始化也可以直接运行：

```bash
python3 tools/bootstrap.py
```

### Amazon 凭证说明

- Web 工作台的预览/提交链路默认读取 `accounts.json`
- `stage2_pipeline.py` / `main.py` 这类 CLI 链路仍会读取 `.env` 里的 `AMAZON_*`
- 如果你在设置页保存 SP-API 配置，系统会同时更新 `.env` 和 `accounts.json`

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

## 端到端测试建议

### 1. AI 文案 / AI 图片

- 使用 1688 或竞品采集得到的原始商品 Excel
- 先跑 Stage1，确认以下字段都已回写：
  - 标题
  - 五点
  - 描述
  - 搜索词
  - AI 主图 / 副图路径

### 2. 图片公网化

当前仓库保留了测试用临时图床流程：

```bash
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --in-place
```

用途：
- 把 Stage1 生成的本地 AI 图片上传到 `x0.at`
- 将公网 URL 回写到 Excel
- 用于后续 Amazon preview / submit 读取图片

说明：
- 这是测试方案，不是正式生产媒体托管方案
- 生产环境应切换到 S3 / CloudFront 一类稳定媒体存储

### 3. Amazon preview / submit

- 先跑 `/api/self-check`
- 再跑 Amazon `VALIDATION_PREVIEW`
- 修完 preview 暴露的问题后，再执行正式提交

建议把测试结果分成三层看：
- 本地字段校验是否通过
- Amazon preview 是否通过
- 正式提交是否返回 `submissionId` / `ASIN`

### 4. 测试数据口径

- 如果是“内部流程演练”，可以使用测试 UPC
- 如果目标是“真实新建商品长期保留”，必须改用真实 GTIN 或 GTIN exemption
- 使用测试 UPC 做正式提交时，允许通过系统链路验证，但后续可能因目录/品牌/条码一致性出问题

## 安全与稳定性

- 所有接受客户端文件路径的 API 端点均做目录白名单校验，防止路径穿越
- Amazon SP-API Token 刷新线程安全（双重检查锁）
- 敏感配置文件（`.env`、`accounts.json`）写入后自动收紧权限至 `0o600`
- `.gitignore` 已排除所有凭证文件

## 当前已知限制

- 第一版官方模板链路固定以美国站为主
- 来源商品链接解析目前仍以”标题/关键词推荐 product_type”为主，不是完整 1688 抓取器
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

## Amazon 内容合规校验

AI 生成内容后会自动做 Amazon 官方规范校验和修正：

| 字段 | 校验项 | 参考文档 |
|------|--------|----------|
| 标题 | 200字符限制、禁用字符、同词不超过2次 | [Product title requirements](https://sellercentral.amazon.com/help/hub/reference/external/GYTR6SYGFA5E3EQC) |
| 五点 | 500字符限制、禁止价格/运费/退款信息、禁止emoji | [Bullet point guidelines](https://sellercentral.amazon.com/help/hub/reference/external/GX5L8BF8GLMML6CX) |
| 搜索词 | 按站点字节限制(US:250/JP:500/IN:200)、空格不计入、自动去重 | [Generic Keyword Field Length](https://sellercentral.amazon.com/help/hub/reference/external/YTR72HN26BQ3TGN) |
| 图片 | 最长边≥1000px、主图白底校验、自动转JPEG | [Product image requirements](https://sellercentral.amazon.com/gp/help/external/G1881) |

## 测试

建议常用回归命令：

```bash
python -m pytest -q
python -m pytest -q test_web.py test_web_flow.py
ruff check .
```

## 临时匿名图床

`upload_ai_images_temp.py` 仍保留，仅供少量测试：

```bash
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --in-place
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --overwrite
```

注意：这不是正式媒体托管方案。
