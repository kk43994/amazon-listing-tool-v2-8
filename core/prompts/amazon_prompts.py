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
Original Competitor Title:
{product_info}

IMPORTANT: The rewritten title must NOT be a copy. Change word order, use synonyms, rephrase — but keep essential keywords for search ranking.

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
Competitor's Product Info:
{product_info}

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
Competitor's Product Info:
{product_info}

CRITICAL: Amazon may flag duplicate content. Make this description GENUINELY original — not just a rewording."""

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
Product Info:
{product_info}

Title (do NOT repeat these words):
{title}

Return ONLY the search terms, space-separated, all lowercase. Stay under 250 bytes."""

# ===== 图片背景替换 =====
IMAGE_BG_WHITE_PROMPT = """Remove the background from this product image and place the product on a pure white background (#FFFFFF).

Amazon main image requirements:
- Pure white background (RGB 255, 255, 255)
- Product fills 85% of the image frame
- No text, logos, watermarks, or additional graphics
- No borders or color blocks
- Product must be the only object
- Professional studio-quality lighting
- Soft, natural drop shadow
- Maintain original product details, texture, and colors
- Output size: at least 1000x1000 pixels (prefer 2000x2000)"""

IMAGE_BG_LIFESTYLE_PROMPT = """Place this product in a realistic lifestyle/usage scene.

Requirements for Amazon secondary images:
- Show the product being used naturally in context
- Warm, inviting lighting
- Relevant to the product's category: {style_hint}
- Authentic and aspirational setting
- High resolution, professional quality
- Product must be clearly visible and recognizable
- Must look DIFFERENT from competitor images"""

IMAGE_BG_GRADIENT_PROMPT = """Place this product on a clean studio-style gradient background.

Requirements:
- Background should transition softly from white to light gray
- Keep the product perfectly intact with realistic proportions
- Professional product photography lighting with subtle shadows
- No text, watermarks, props, or decorative graphics
- Product should remain the only focus of the image
- High resolution, suitable for e-commerce display"""

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
