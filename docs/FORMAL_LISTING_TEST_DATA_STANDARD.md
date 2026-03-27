# 正式新建测试数据最低标准

这份标准只解决一个问题:
拿什么样的数据去测“新建商品上架”，才不会把系统问题和客户脏数据混在一起。

## 1. 必须先确定商品标识模式

- `real_gtin`
  适用场景: 客户能提供真实的 UPC/EAN/GTIN，且能确认这是对外正式商品编码。
- `gtin_exemption`
  适用场景: 客户没有真实条码，但店铺对应品牌和类目已经拿到 Amazon 的 GTIN exemption。
- `internal_code`
  适用场景: 客户只有 ERP/内部编号。
  结论: 这种数据不能拿来做“正式新建商品”测试，只能做内部流程测试。

## 2. 能用于正式新建测试的最小字段

- `SKU`
  必须唯一，建议新建一个测试专用 SKU，不复用历史提交过的 SKU。
- `product_identity_mode`
  必填，只能是 `real_gtin` 或 `gtin_exemption`。
- `product_id` 和 `product_id_type`
  当 `product_identity_mode=real_gtin` 时必填。
  `UPC` 必须 12 位纯数字。
  `EAN` 必须 13 位纯数字。
  `GTIN` 必须是 8/12/13/14 位纯数字。
- `item_name`
  必填，长度建议不超过 200 字符。
- `brand_name`
  必填。必须和真实品牌、商标授权、Amazon 目录认知一致。
- `product_type`
  必填。必须是目标类目允许的 Amazon productType。
- `main_image_url`
  必填。必须是公网可访问的图片直链，不是本地路径，不是网页地址。
- `standard_price`
  正式提交时必填。
- `quantity`
  正式提交时建议填写。
- `condition_type`
  建议固定为 `new_new` 或系统可映射的新商品状态。

## 3. 通过这份标准仍然不代表一定能新建成功

以下问题不属于系统字段问题，而属于业务资质或目录归属问题:

- 条码已经被 Amazon 目录识别到现有 ASIN
- 当前账号无权在该品牌下新建
- 当前条码对应品牌与提交品牌不一致
- 当前类目需要额外合规字段或审批
- 当前店铺没有 GTIN exemption

## 4. 一条数据能不能拿来测“正式新建”

必须同时满足:

- 商品标识模式不是 `internal_code`
- 条码类型和条码长度一致
- 品牌、条码、商品本体三者一致
- 主图是公网图片直链
- 该账号对该品牌和类目具备创建权限
- 最好不是已经存在于 Amazon 目录中的成熟商品

## 5. 当前项目里的建议测试策略

- 测 AI 改写和 AI 改图:
  可以使用已有公开商品数据。
- 测 Amazon VALIDATION_PREVIEW:
  可以使用已有真实商品，但要接受它可能会命中已有 ASIN。
- 测“正式新建商品”:
  只用客户自己拥有品牌/授权/真实 GTIN 的数据，或者明确具备 GTIN exemption 的商品。

## 6. 当前最容易踩坑的错误

- 把 ERP 自动生成的码当成 UPC/EAN/GTIN
- 14 位码却标成 `UPC`
- 用现成大品牌成熟商品去测“新建”
- 把本地 AI 图片路径当成 Amazon 可提交图片
- 以为 Amazon 草稿箱能替代 SP-API 的真实校验逻辑
