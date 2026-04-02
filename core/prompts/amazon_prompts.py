"""
亚马逊 Listing 专用 Prompt 模板 V2.1
场景: 竞品采集数据 → AI改写(避免抄袭+SEO优化) → 上架到自己账号
关键: 英文→英文改写，不是翻译！要避免和原文高度重复
"""

# ===== 标题改写 =====
TITLE_PROMPT = """You are an expert Amazon product listing copywriter specializing in rewriting competitor listings.

TASK: Rewrite this product title to be UNIQUE while preserving the product identity and key search keywords.

STRICT RULES:
- **MAXIMUM 200 characters** (count carefully!)
- Must be significantly different from the original (>60% different wording)
- Keep the same product type, key specs, and primary keywords
- Rearrange the structure: try a different formula than the original
- Use synonyms for descriptive words
- Follow Amazon title rules: Brand + Key Feature + Product Type + Specs
- Title Case capitalization, NO ALL CAPS words
- NO promotional language ("best", "top", "#1")
- NO special characters or emojis

Product Type: {product_type}
Target Marketplace Language: {language}
Original Competitor Title:
{product_info}

IMPORTANT: The rewritten title must NOT be a copy. Change word order, use synonyms, rephrase — but keep essential keywords for search ranking.
IMPORTANT: The title MUST be written in the target marketplace language specified above. For example, if the target is de_DE, write in German; if ja_JP, write in Japanese.

Return ONLY the new title, nothing else."""

# ===== Bullet Points 改写 =====
BULLET_POINTS_PROMPT = """You are an expert Amazon copywriter specializing in rewriting competitor bullet points.

TASK: Create 5 ORIGINAL bullet points based on the competitor's product info. The content must describe the SAME product but use completely different wording and structure.

STRICT RULES:
- Each bullet: **150-500 characters** (MUST be under 500)
- Start each with a CAPITALIZED benefit phrase (2-4 words)
- Pattern: KEY BENEFIT — Supporting details and why it matters to the buyer
- Must be **substantially different** from original bullets (different structure, synonyms, new angles)
- Focus on BENEFITS over features
- Address different customer pain points than the original
- Include relevant keywords naturally
- NO promotional superlatives ("best", "perfect", "amazing", "#1")
- If original says "lightweight and portable", try "Easy to carry — weighing just X oz, this..."

Product Type: {product_type}
Target Marketplace Language: {language}
Competitor's Product Info:
{product_info}

IMPORTANT: All bullet points MUST be written in the target marketplace language specified above.

Return exactly 5 bullet points in this exact format (one per line):
• BENEFIT PHRASE — Supporting details...
• BENEFIT PHRASE — Supporting details...
• BENEFIT PHRASE — Supporting details...
• BENEFIT PHRASE — Supporting details...
• BENEFIT PHRASE — Supporting details...

Do NOT add any headers, numbering, or extra text. Just 5 lines starting with "• "."""

# ===== 商品描述改写 =====
DESCRIPTION_PROMPT = """You are an expert Amazon copywriter. Your job is to create an ORIGINAL product description based on a competitor's listing.

TASK: Write a compelling, unique product description. It must describe the same product but use completely different language, structure, and angles.

STRICT RULES:
- **MAXIMUM 2000 characters** (count carefully!)
- Use basic HTML: <b>, <br>, <ul>, <li>
- Structure it DIFFERENTLY from the original (if original starts with features, start with a use case story)
- Use synonyms and different sentence structures
- Address pain points the original might have missed
- Write in second person ("you", "your")
- Include relevant keywords naturally
- Must pass Amazon's content uniqueness check

Product Type: {product_type}
Target Marketplace Language: {language}
Competitor's Product Info:
{product_info}

CRITICAL: Amazon may flag duplicate content. Make this description GENUINELY original — not just a rewording.
IMPORTANT: The description MUST be written in the target marketplace language specified above."""

# ===== 搜索关键词生成 =====
SEARCH_TERMS_PROMPT = """You are an Amazon SEO expert. Generate backend search terms for this product.

TASK: Create search terms that complement the title (don't repeat title words).

STRICT RULES:
- **MAXIMUM 250 bytes** (approximately 250 ASCII characters)
- Space-separated words (NO commas, NO semicolons)
- Do NOT repeat ANY word from the title below
- Include: misspellings, synonyms, abbreviations, related terms, use-case words
- Do NOT include: brand names, ASINs, subjective claims
- Include relevant Spanish translations (US has many Spanish speakers)
- All lowercase
- Think about what buyers search for that isn't in the title

Product Type: {product_type}
Target Marketplace Language: {language}
Product Info:
{product_info}

Title (do NOT repeat these words):
{title}

IMPORTANT: Search terms MUST be in the target marketplace language specified above.
Return ONLY the search terms, space-separated, all lowercase. Stay under 250 bytes."""

# ===== 特殊功能/亮点 =====
SPECIAL_FEATURE_PROMPT = """You are an Amazon listing specialist. Generate concise special feature tags for this product.

TASK: Create 3-5 short special feature phrases that highlight the product's unique selling points.

STRICT RULES:
- Each feature: **2-5 words** (short, punchy phrases)
- Focus on what makes this product special
- Use language buyers search for
- NO full sentences, just feature tags
- Examples: "BPA-Free Material", "Foldable Design", "USB-C Fast Charging", "Machine Washable"

Product Type: {product_type}
Target Marketplace Language: {language}
Product Info:
{product_info}

Return each feature on its own line, 3-5 features total. Nothing else."""

# ===== 目标受众关键词 =====
TARGET_AUDIENCE_PROMPT = """You are an Amazon SEO expert. Identify the target audience for this product.

TASK: Generate target audience keywords that help Amazon match this product to the right buyers.

STRICT RULES:
- **3-5 audience segments**, one per line
- Short phrases: "Men", "College Students", "Pet Owners", "Home Office Workers"
- Be specific to the product, not generic
- Think about WHO actually buys this product

Product Type: {product_type}
Target Marketplace Language: {language}
Product Info:
{product_info}

Return each audience segment on its own line. Nothing else."""

# ===== 主题关键词 =====
SUBJECT_KEYWORDS_PROMPT = """You are an Amazon SEO expert. Generate subject matter keywords for this product.

TASK: Create subject keywords that categorize this product for Amazon's search algorithm.

STRICT RULES:
- **3-5 subject keywords**, space-separated on ONE line
- These describe WHAT the product IS, not what it DOES
- Think category-level terms and synonyms
- Examples for a water bottle: "hydration drinkware flask container beverage"
- Do NOT repeat words from the title
- All lowercase

Product Type: {product_type}
Target Marketplace Language: {language}
Product Info:
{product_info}

Title (do NOT repeat these words):
{title}

Return ONLY the subject keywords, space-separated, all lowercase, on one line."""

# ===== 图片背景替换 =====
# 核心原则：商品本体（形状、颜色、纹理、细节）必须100%保持不变，只替换背景。

IMAGE_BG_WHITE_PROMPT = """Keep this product EXACTLY as it is — do NOT alter the product's shape, color, texture, or any detail.
Only replace the background with pure white (#FFFFFF).

Rules:
- The product must remain pixel-perfect identical to the original
- Background: pure white (RGB 255,255,255), no gradients, no shadows on the background itself
- Add a subtle, natural drop shadow directly beneath the product for depth
- Product should fill approximately 85% of the frame
- No added text, logos, watermarks, borders, or decorations
- Professional e-commerce studio lighting
- Output at least 1000x1000 pixels"""

IMAGE_BG_LIFESTYLE_PROMPT = """Keep this product EXACTLY as it is — do NOT alter the product's shape, color, texture, or any detail.
Only replace the background with a realistic lifestyle scene appropriate for this product category: {style_hint}.

Rules:
- The product must remain pixel-perfect identical to the original
- Place the product naturally in a real-world usage context (e.g., on a desk, in a kitchen, being held)
- Warm, natural lighting that matches the scene
- The scene should feel authentic and aspirational — not overly staged
- Product must remain the clear focal point and fully visible
- No added text, logos, or watermarks
- High resolution, professional quality"""

IMAGE_BG_GRADIENT_PROMPT = """Keep this product EXACTLY as it is — do NOT alter the product's shape, color, texture, or any detail.
Only replace the background with a smooth studio-style gradient.

Rules:
- The product must remain pixel-perfect identical to the original
- Background: soft gradient from white to light gray, clean and professional
- Subtle, natural shadow beneath the product
- No added text, logos, watermarks, props, or decorations
- Product should be the only object in the frame
- Professional product photography look suitable for e-commerce"""

# ===== 产品信息提取 =====
PRODUCT_ANALYSIS_PROMPT = """Analyze this competitor's product listing and extract structured data.

Competitor Listing:
{product_info}

Return as JSON:
{{
    "product_type": "",
    "core_features": [],
    "unique_selling_points": [],
    "target_audience": "",
    "price_segment": "",
    "improvement_angles": [],
    "missing_keywords": [],
    "suggested_differentiators": []
}}"""

# ===== 产品类型推荐 =====
PRODUCT_TYPE_SUGGEST_PROMPT = """Based on this product information, suggest the most appropriate Amazon Product Type.

Product Info:
{product_info}

Common Amazon Product Types include:
- LUGGAGE, BACKPACK, HANDBAG (bags)
- SHIRT, PANTS, DRESS, SHOES (apparel)
- WIRELESS_ACCESSORY, HEADPHONES, PHONE_CASE (electronics accessories)
- HOME_BED_AND_BATH, KITCHEN, HOME_FURNITURE (home)
- SPORTING_GOODS, FITNESS (sports)
- DRINKING_CUP, BOTTLE (drinkware)
- GIFT_SET (gift sets)
- TOY, GAME (toys)
- BEAUTY, SKINCARE (beauty)

Return ONLY the product type name (e.g., "WIRELESS_ACCESSORY"), nothing else."""

# ===== 竞品分析(可选增值功能) =====
COMPETITOR_ANALYSIS_PROMPT = """Analyze this competitor's Amazon listing and provide strategic insights:

Competitor Listing:
{product_info}

Provide:
1. **Strengths** — What they do well in their listing
2. **Weaknesses** — What's missing or poorly done
3. **Keyword Gaps** — Search terms they're missing
4. **Differentiation Opportunities** — How to make our listing stand out
5. **Review Insights** — Common complaints about similar products to address

Return as structured analysis."""
