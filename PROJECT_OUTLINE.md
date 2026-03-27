# 亚马逊商品上架自动化工具 — 项目大纲 V2.0

> 基于 Amazon SP-API 官方文档深度研究后重构
> 更新日期: 2026-02-09

---

## 一、项目定位

**核心价值：竞品采集→AI改写→一键上架，避免抄袭同时保留SEO价值**

### 数据来源
**甲方的Excel来自亚马逊其他店铺的竞品采集**，数据已经是英文的亚马逊Listing格式。

### 业务流程（完整闭环）

```
竞品店铺采集(英文listing Excel) 
  → AI智能改写(改标题+改卖点+改描述+换图片背景) 
    → 前后对比预览(竞品原文 vs AI改写) 
      → 人工确认/微调 
        → 选择目标亚马逊账号 
          → SP-API自动提交上架 
            → 状态追踪(成功/失败/警告)
```

### ⚠️ 关键注意
- **不是翻译！是英文→英文改写** — 保留核心关键词，但完全重写表述
- **必须避免内容重复** — Amazon会检测重复内容，可能导致listing被拒或封号
- **图片不能直接用** — 竞品图片有版权，必须用AI换背景/风格

---

## 二、SP-API 核心知识（官方文档研究成果）

### 2.1 上架API选择

| 方式 | API | 适用场景 | 限流 |
|------|-----|---------|------|
| **单个上架** | `putListingsItem` | 实时逐个提交，有即时反馈 | 5次/秒 |
| **批量上架** | `JSON_LISTINGS_FEED` + Feeds API | 超过1500件商品 | 5次/5分钟 |
| **局部更新** | `patchListingsItem` | 只改价格/库存/某个字段 | 5次/秒 |

**我们的方案：** 
- 默认用 `putListingsItem`（单个提交，即时反馈，适合中小批量）
- 超过1500件自动切换 `JSON_LISTINGS_FEED`（批量提交）

### 2.2 上架完整流程（SP-API官方推荐）

```
1. searchCatalogItems → 搜索亚马逊目录，看商品是否已存在(ASIN)
2. searchDefinitionsProductTypes → 确定产品类型(如LUGGAGE, DRINKING_CUP等)
3. getDefinitionsProductType → 获取该类型的JSON Schema(所有必填/选填字段)
4. 准备listing数据 → 按Schema构建attributes JSON
5. putListingsItem → 提交上架
6. 检查response → ACCEPTED/INVALID
7. 监听 LISTINGS_ITEM_STATUS_CHANGE 通知 → 异步获取最终结果
```

### 2.3 亚马逊商品字段分组（以LUGGAGE类型为例）

SP-API按 `propertyGroups` 分组，每个产品类型的字段不同，但通用分组如下：

#### 🔴 Product Identity（产品身份）— 必填
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `item_name` | 商品标题(200字符限制) | ✅ AI翻译+优化 |
| `brand` | 品牌名 | ❌ 原始数据 |
| `externally_assigned_product_identifier` | UPC/EAN/GTIN条码 | ❌ 原始数据 |
| `item_type_keyword` | 商品类型关键词 | ✅ AI推��� |
| `manufacturer` | 制造商 | ❌ 原始数据 |
| `model_number` | 型号 | ❌ 原始数据 |
| `merchant_suggested_asin` | 建议关联ASIN(匹配已有商品时) | ❌ API查询 |

#### 🟡 Product Details（商品详情）— 强烈推荐
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `product_description` | 商品描述(2000字符) | ✅ AI生成 |
| `bullet_point` | 卖点(最多5条,500字符/条) | ✅ AI生成 |
| `special_feature` | 特殊功能 | ✅ AI生成 |
| `color` | 颜色 | ❌ 原始数据 |
| `size` | 尺寸 | ❌ 原始数据 |
| `material` | 材质 | ❌ 原始数据 |
| `target_gender` | 目标性别 | ❌ 原始数据 |
| `age_range_description` | 适用年龄 | ❌ 原始数据 |
| `department` | 部门分类 | ✅ AI推荐 |

#### 🟢 Images（图片）— 必填至少主图
| 字段名 | 说明 | AI可处理? |
|--------|------|----------|
| `main_product_image_locator` | 主图URL(纯白背景,1000×1000+) | ✅ AI换背景 |
| `other_product_image_locator_1~8` | 副图(最多8张) | ✅ AI换背景 |
| `swatch_product_image_locator` | 色板图 | ❌ 原始数据 |

#### 🔵 Offer（销售条款）— 必填
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `purchasable_offer` | 价格(含currency+schedule) | ❌ 卖家设定 |
| `condition_type` | 商品状态(new_new/used等) | ❌ 卖家设定 |
| `fulfillment_availability` | 库存+配送方式(FBA/FBM) | ❌ 卖家设定 |
| `merchant_shipping_group` | 运费模板 | ❌ 卖家设定 |
| `list_price` | 标价 | ❌ 卖家设定 |

#### 🟣 Shipping（物流）— 推荐
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `item_dimensions` | 商品尺寸(长宽高) | ❌ 原始数据 |
| `item_package_dimensions` | 包装尺寸 | ❌ 原始数据 |
| `item_package_weight` | 包装重量 | ❌ 原始数据 |
| `item_weight` | 商品净重 | ❌ 原始数据 |

#### ⚪ Safety & Compliance（安全合规）— 部分必填
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `country_of_origin` | 原产国 | ❌ 原始数据 |
| `batteries_required` | 是否需要电池 | ❌ 原始数据 |
| `batteries_included` | 是否包含电池 | ❌ 原始数据 |
| `battery` | 电池信息(类型/容量) | ❌ 原始数据 |
| `supplier_declared_dg_hz_regulation` | 危险品声明 | ❌ 原始数据 |
| `california_proposition_65` | 加州65号提案合规 | ❌ 原始数据 |

#### 🟤 Variations（变体）— 有变体时必填
| 字段名 | 说明 | AI可生成? |
|--------|------|----------|
| `parentage_level` | 父/子层级 | ❌ 卖家设定 |
| `child_parent_sku_relationship` | 父子SKU关系 | ❌ 卖家设定 |
| `variation_theme` | 变体主题(颜色/尺寸等) | ❌ 卖家设定 |

### 2.4 关键概念

- **Product Type**: 亚马逊用产品类型决定需要哪些字段。不同类型(LUGGAGE vs SHIRT vs WIRELESS_ACCESSORY)所需字段完全不同
- **Product Type Definitions API**: 调用此API可动态获取任何产品类型的完整JSON Schema
- **requirements参数**: 
  - `LISTING` = 产品信息+销售条款(完整上架)
  - `LISTING_PRODUCT_ONLY` = 只有产品信息(先建商品不卖)
  - `LISTING_OFFER_ONLY` = 只有销售条款(跟卖已有ASIN)
- **marketplace_id**: 美国站 = `ATVPDKIKX0DER`
- **sellerId**: 卖家账号ID，决定上架到哪个账号

---

## 三、Excel字段重构

### 3.1 输入Excel（原始商品数据）

用户导入的Excel应包含所有原始商品信息：

| 分组 | 字段 | 必填 | 说明 |
|------|------|------|------|
| **基础** | SKU | ✅ | 卖家自定义编号 |
| **基础** | 商品名称(中文) | ✅ | 原始中文名 |
| **基础** | 品牌 | ✅ | Brand |
| **基础** | 产品类型 | ✅ | 亚马逊Product Type |
| **基础** | UPC/EAN | ⚠️ | 条形码(无则需申请豁免) |
| **图片** | 主图URL | ✅ | 原始主图 |
| **图片** | 副图1~8 URL | 📎 | 最多8张副图 |
| **文案** | 中文���述 | 📎 | 原始中文描述 |
| **文案** | 中文卖点1~5 | 📎 | 原始中文卖点 |
| **销售** | 售价(USD) | ✅ | 美元价格 |
| **销售** | 库存数量 | ✅ | 可售数量 |
| **销售** | 配送方式 | ✅ | FBA/FBM |
| **销售** | 商品状态 | ✅ | 全新/翻新/二手 |
| **物流** | 商品重量 | 📎 | 克/千克 |
| **物流** | 商品尺寸(长宽高) | 📎 | 厘米 |
| **物流** | 包装重量 | 📎 | 克/千克 |
| **物流** | 包装尺寸(长宽高) | 📎 | 厘米 |
| **属性** | 颜色 | 📎 | 原始值 |
| **属性** | 尺寸 | 📎 | 原始值 |
| **属性** | 材质 | 📎 | 原始值 |
| **属性** | 原产国 | 📎 | 如"China" |
| **属性** | 制造商 | 📎 | 制造商名称 |
| **属性** | 型号 | 📎 | Model Number |
| **合规** | 是否含电池 | 📎 | Yes/No |
| **合规** | 电池类型 | 📎 | 如含电池时填 |
| **变体** | 父SKU | 📎 | 有变体时填 |
| **变体** | 变体主题 | 📎 | Color/Size等 |
| **搜索** | 搜索关键词 | 📎 | 逗号分隔 |

### 3.2 输出Excel（AI处理后 + 前后对比）

| 列 | 说明 |
|----|------|
| SKU | 不变 |
| **原始标题(中文)** | 保留原始 |
| **→ AI标题(英文)** | AI翻译+SEO优化后 |
| **原始描述(中文)** | 保留原始 |
| **→ AI描述(英文)** | AI翻译+优化后 |
| **原始卖点1~5** | 保留原始 |
| **→ AI卖点1~5(英文)** | AI翻译+优化后 |
| **原始主图URL** | 保留原始 |
| **→ AI主图URL** | 换白底后的URL |
| **原始副图1~8** | 保留原始 |
| **→ AI副图1~8** | 换背景后 |
| **AI搜索关键词** | AI生成的Search Terms |
| 其他字段 | 原样保留 |
| **SP-API提交状态** | PENDING/ACCEPTED/INVALID/ERROR |
| **SP-API提交时间** | 时间戳 |
| **ASIN** | 上架成功后回填 |
| **Issues** | 上架失败原因 |

**关键设计：原始值和AI生成值并排对比，卖家一目了然**

---

## 四、功能模块重构

### 4.1 Stage 1: AI内容生成（已有,需增强）

```
输入: 原始Excel(中文商品数据)
处理:
  1. 读取每行商品数据
  2. AI翻译+优化标题(中→英, SEO优化, ≤200字符)
  3. AI生成5条Bullet Points(英文, 每条≤500字符)
  4. AI生成Product Description(英文, ≤2000字符)
  5. AI生成Search Terms(≤250字节,逗号分隔)
  6. AI图片背景替换(→纯白背景, 1000×1000+)
  7. 写入输出Excel(原始+AI结果并排)
输出: 带前后对比的Excel
```

**需要增强的点：**
- ✅ 标题长度控制(≤200字符)
- ✅ Bullet Points数量和长度控制
- ✅ Description长度控制
- ✅ Search Terms字节限制(250 bytes)
- 🆕 根据Product Type调整prompt(不同类型侧重点不同)
- 🆕 AI生成 `item_type_keyword` 推荐
- 🆕 图片输出尺寸强制1000×1000以上

### 4.2 Stage 2: SP-API上架（需重写）

```
输入: AI处理后的Excel + 账号配置
处理:
  1. 选择目标亚马逊账号(sellerId)
  2. 对每个商品:
     a. searchCatalogItems → 检查ASIN是否已存在
     b. getDefinitionsProductType → 获取Product Type Schema
     c. 构建attributes JSON(按Schema要求)
     d. putListingsItem → 提交上架
     e. 记录response(ACCEPTED/INVALID + issues)
  3. 回填Excel(状态+ASIN+错误信息)
输出: 更新后的Excel(含提交结果)
```

**关键改进：**
- 🆕 **多账号支持**: 可配置多个亚马逊卖家账号
- 🆕 **Product Type动态Schema**: 根据产品类型获取实际需要的字段
- 🆕 **字段映射验证**: 提交前验证所有required字段已填写
- 🆕 **ASIN匹配**: 先搜索已有商品，避免重复创建
- 🆕 **前后对比确认**: 提交前展示原始→AI生成的对比，人工确认
- 🆕 **错误处理**: 分类展示issues(ERROR/WARNING/INFO)
- 🆕 **批量模式**: >1500件自动切换JSON_LISTINGS_FEED

### 4.3 Web界面重构

```
首页Dashboard
├── 统计卡片(输入/输出/已上架/失败)
├── 快速���口(批量处理/单品编辑)
└── 最近活动

批量处理页 ← 核心页面
├── Step 1: 上传Excel(原始商品数据)
├── Step 2: 预览 & 选择处理选项
│   ├── ☑ AI标题翻译优化
│   ├── ☑ AI卖点生成
│   ├── ☑ AI描述生成
│   ├── ☑ AI图片换背景
│   └── ☑ AI搜索关键词
├── Step 3: AI处理(实时进度)
├── Step 4: 前后对比预览 ← 🆕 关键！
│   ├── 表格视图(原始 | AI生成 | 状态)
│   ├── 单品详情(点击展开对比)
│   ├── 编辑功能(不满意可手动修改)
│   └── 确认/全部通过
├── Step 5: 选择上架账号 ← 🆕
│   ├── 下拉选择已配置的亚马逊账号
│   ├── 显示marketplace信息
│   └── 字段完整性检查
└── Step 6: 提交上架 + 状态追踪 ← 🆕
    ├── 逐个提交进度
    ├── 成功/失败统计
    ├── 错误详情(可重试)
    └── 下载结果Excel

设置页
├── OpenAI API配置
├── 亚马逊账号管理 ← 🆕
│   ├── 添加账号(sellerId + credentials)
│   ├── 多账号列表
│   ├── 测试连接
│   └── 默认账号选择
└── 通用设置

单品编辑页 ← 🆕
├── 单个商品详细编辑
├── 实时AI生成预览
├── 图片上传+背景替换
└── 直接提交上架
```

---

## 五、亚马逊账号管理

### 5.1 每个账号需要的凭证

```json
{
  "accounts": [
    {
      "name": "主账号-美国站",
      "seller_id": "AXXXXXXXXXXXX",
      "marketplace_id": "ATVPDKIKX0DER",
      "marketplace_name": "Amazon US",
      "lwa_client_id": "amzn1.application-oa2-client.xxx",
      "lwa_client_secret": "xxx",
      "refresh_token": "Atzr|xxx",
      "is_default": true
    },
    {
      "name": "副账号-欧洲站",
      "seller_id": "AYYYYYYYYYYYY",
      "marketplace_id": "A1F83G8C2ARO7P",
      "marketplace_name": "Amazon UK",
      "lwa_client_id": "amzn1.application-oa2-client.yyy",
      "lwa_client_secret": "yyy",
      "refresh_token": "Atzr|yyy",
      "is_default": false
    }
  ]
}
```

### 5.2 SP-API认证流程

```
1. Login With Amazon (LWA) OAuth2
   POST https://api.amazon.com/auth/o2/token
   Body: grant_type=refresh_token & client_id & client_secret & refresh_token
   → 获取 access_token (有效���1小时)

2. 所有SP-API请求带 x-amz-access-token header
```

---

## 六、亚马逊各站点Marketplace ID

| 站点 | Marketplace ID | Endpoint |
|------|---------------|----------|
| 🇺🇸 美国 | ATVPDKIKX0DER | sellingpartnerapi-na.amazon.com |
| 🇨🇦 加拿大 | A2EUQ1WTGCTBG2 | sellingpartnerapi-na.amazon.com |
| 🇲🇽 墨西哥 | A1AM78C64UM0Y8 | sellingpartnerapi-na.amazon.com |
| 🇬🇧 英国 | A1F83G8C2ARO7P | sellingpartnerapi-eu.amazon.com |
| 🇩🇪 德国 | A1PA6795UKMFR9 | sellingpartnerapi-eu.amazon.com |
| 🇫🇷 法国 | A13V1IB3VIYZZH | sellingpartnerapi-eu.amazon.com |
| 🇮🇹 意大利 | APJ6JRA9NG5V4 | sellingpartnerapi-eu.amazon.com |
| 🇪🇸 西班牙 | A1RKKUPIHCS9HS | sellingpartnerapi-eu.amazon.com |
| 🇯🇵 日本 | A1VC38T7YXB528 | sellingpartnerapi-fe.amazon.com |
| 🇦🇺 澳大利亚 | A39IBJ37TRP1C6 | sellingpartnerapi-fe.amazon.com |

---

## 七、开发优先级

### Phase 1: 核心功能修正（本周）
1. ✅ Excel字段重构 — 按SP-API实际需求重新设计
2. ✅ AI生成增强 — 长度控制、Product Type感知
3. 🆕 前后对比功能 — 原始vs AI生成并排展示
4. 🆕 字段完整性检查 — 提交前验证required字段

### Phase 2: SP-API集成重写（下周）
1. 🆕 多账号管理 — 配置/切换/测试
2. 🆕 Product Type动态获取 — 调用getDefinitionsProductType
3. 🆕 putListingsItem正确调用 — 按Schema构建payload
4. 🆕 状态追踪 — 成功/失败/warnings回填Excel

### Phase 3: 体验优化（持续）
1. 🆕 单品编辑器 — 精细化编辑单个商品
2. 🆕 批量模式 — JSON_LISTINGS_FEED集成
3. 🆕 历史记录 — 上架记录查询
4. 🆕 通知系统 — LISTINGS_ITEM_STATUS_CHANGE webhook

---

## 八、与现有代码的差距分析

### 已有 ✅
- Excel读写(openpyxl)
- AI文案生成框架(OpenAI)
- AI图片处理框架
- Web界面基础(Flask)
- SP-API认证(LWA OAuth2)
- Feeds API基础

### 需要新增/重写 🆕
- **Excel字段映射**: 现有54列太多且乱，需按SP-API分组重构
- **前后对比**: 输出Excel需要原始+AI并排
- **Product Type动态Schema**: 现在是硬编码字段映射，应该动态获取
- **多账号管理**: 现在只支持单账号
- **ASIN匹配检查**: 上架前检查是否已存在
- **字段验证**: 按JSON Schema验证所有required字段
- **更好的错误处理**: 分类展示SP-API返回的issues
- **Web步骤流程**: 现在是一键处理，应该分步骤(上传→AI→对比→确认→上架)

---

## 九、技术架构

```
┌───────────────────��─────────────────────┐
│              Web界面 (Flask)             │
│  上传Excel → AI处理 → 对比预览 → 上架    │
└─────────┬───────────────────────────────┘
          │
┌─────────▼───────────────────────────────┐
│           业务逻辑层                      │
│  ┌──────────┐  ┌──────────┐  ┌────────┐ │
│  │ AI处理   │  │ 字段映射  │  │ 账号   │ │
│  │ Pipeline │  │ & 验证   │  │ 管理   │ │
│  └──────────┘  └──────────┘  └────────┘ │
└─────────┬───────────────────────────────┘
          │
┌─────────▼───────────────────────────────┐
│           外部API层                      │
│  ┌──────────┐  ┌───────────────────────┐ │
│  │ OpenAI   │  │ Amazon SP-API         │ │
│  │ (文案+图) │  │ ├ LWA Auth            │ │
│  │          │  │ ├ Listings Items API   │ │
│  │          │  │ ├ Product Type Defs    │ │
│  │          │  │ ├ Catalog Items API    │ │
│  │          │  │ └ Feeds API            │ │
│  └──────────┘  └───────────────────────┘ │
└─────────────────────────────────────────┘
```

---

## 十、关键注意事项

1. **Product Type不是固定的**: 每种商品类型需要的字段完全不同，必须动态获取Schema
2. **图片要求严格**: 主图必须纯白背景、≥1000×1000px、JPEG/PNG/GIF/TIFF
3. **UPC/EAN是关键**: 没有条码无法创建新商品(除非申请GTIN豁免)
4. **putListingsItem是替换操作**: 漏传的字段会被清空！必须传完整数据
5. **异步处理**: ACCEPTED只代表接受处理，不代表上架成功，需要监听通知
6. **Rate Limit**: 5次/秒，大批量时需要排队+重试机制
7. **Marketplace特定**: 不同站点的必填字段可能不同

---

*本文档基于Amazon SP-API官方文档研究整理*
*参考: https://developer-docs.amazon.com/sp-api/docs/listings-items-api*
