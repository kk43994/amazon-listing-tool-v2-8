# 亚马逊商品处理与自动提交工具

## 项目概述
基于Xobi核心引擎的亚马逊商品内容处理与SP-API自动提交工具。

### 第一阶段：AI商品内容处理
- 商品图片AI背景替换（白底/场景）
- 标题/描述/Bullet Points AI生成
- Excel模板读写

### 第二阶段：自动化提交
- Amazon SP-API OAuth对接
- Excel→Listings字段映射
- 批量提交与状态追踪

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 配置API Key
cp .env.example .env
# 编辑 .env 填入API Key
# 如需把AI改图提交给Amazon，请配置 OUTPUT_IMAGE_PUBLIC_BASE 为公开图片URL前缀

# 3. 放入数据
# 将甲方Excel放入 input/ 目录

# 4. 运行第一阶段
python main.py --stage 1

# 5. 运行第二阶段 (需要SP-API权限)
python main.py --stage 2

# 临时方案：把 Stage1 生成的本地 AI 图片匿名上传到 x0.at，并写回 AI主图URL
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx
```

## 目录结构
```
亚马逊2.8/
├── config.py              # 配置
├── main.py                # 入口
├── core/                  # 核心引擎 (基于Xobi)
│   ├── ai_providers/      # AI供应商 (图片+文案)
│   ├── prompts/           # Prompt模板
│   ├── excel/             # Excel处理
│   └── url_utils.py       # 工具函数
├── amazon/                # Amazon SP-API模块
│   ├── auth.py            # OAuth认证
│   ├── listings.py        # Listings API
│   ├── feeds.py           # Feeds API (批量提交)
│   └── mapper.py          # 字段映射
├── input/                 # 输入 (甲方Excel+图片)
├── output/                # 输出 (处理后的Excel+图片)
└── logs/                  # 日志
```

## 技术栈
- Python 3.10+
- OpenAI SDK (兼容多供应商)
- openpyxl (Excel读写)
- Pillow (图片处理)
- python-sp-api (Amazon SP-API)

## 配置补充
- `OUTPUT_IMAGE_PUBLIC_BASE`: 可选。配置后会把本地 AI 主图路径转换为可公开访问的 URL，第二阶段提交时优先使用该 URL。
- `WEB_DEBUG`: 可选。控制 Flask Web 调试模式，默认关闭。

## 临时匿名图床
- `upload_ai_images_temp.py`: 用于临时测试。脚本会读取 Excel 中的 `AI主图路径`，上传到 `x0.at`，然后把返回的公网 URL 写回 `AI主图URL`。
- 示例:
```bash
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --in-place
python upload_ai_images_temp.py --input output/处理结果_xxx.xlsx --overwrite
```
- 注意: 这是匿名临时图床，只建议少量测试使用，不建议作为正式长期托管方案。
