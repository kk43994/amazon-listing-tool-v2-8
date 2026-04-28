# Release Checklist

用于每次发客户包前逐项确认，避免把半成品压缩包发给不会排查问题的客户。

## 1. 本地基础检查

- [ ] `python -m py_compile web/app.py release_entry.py tools/build_release.py tools/environment_check.py tools/verify_release.py`
- [ ] `node --check` 检查 `web/templates/_js.txt`
- [ ] `pytest -q` 全量通过
- [ ] `git diff --check` 无尾随空格/冲突标记

## 2. 打包检查

- [ ] 运行 `python tools/build_release.py`
- [ ] 运行 `python tools/verify_release.py release/AmazonListingTool-darwin-arm64 --system darwin`
- [ ] Windows CI 运行 `tools/verify_release.py release/AmazonListingTool-windows-amd64 --system windows`
- [ ] 包根目录必须能看到：`客户先看这里.txt`、`Read-Me-First.txt`、`一键检测修复`、`Doctor`、`环境检测`、`Env-Check`、`Support-Bundle`、`Open-Output`、`Open-Backups`、`.env.example`、模板 Excel、`release-manifest.json`、`dependency-inventory.json`
- [ ] Release 同时上传 `.zip.sha256`

## 3. 打包后程序自检

- [ ] 打包程序运行 `--env-check --quiet`
- [ ] 打包程序运行 `--doctor --quiet`
- [ ] 打包程序运行 `--smoke-test --quiet --port 0`
- [ ] 打包程序运行 `--support-bundle --quiet` 后能生成支持包 zip
- [ ] 离线依赖包 `tools/build_dependency_bundle.py` 构建成功，`vendor/wheelhouse` 不缺包

## 4. 干净机器人工验证

- [ ] macOS 干净机器/Finder 解压后，双击 `Start-Amazon-2.8.command` 或 `启动亚马逊2.8.command` 能打开网站
- [ ] Windows 干净机器/Explorer 解压后，双击 `Start-Amazon-2.8.bat` 或 `启动亚马逊2.8.bat` 能打开网站
- [ ] Windows Defender/杀软未删除 `_internal` 或 exe；若拦截，Doctor 能给出明确提示
- [ ] 不在压缩包内直接运行；不放 OneDrive/iCloud/Dropbox 同步目录；Doctor 能提示风险

## 5. 客户视角流程

- [ ] 首次打开能看到“傻瓜流程”和一个醒目的下一步按钮
- [ ] AI 设置默认简单模式，只需要填文字 Key/图片 Key；一键恢复推荐会写入 `https://api.kk666.best`
- [ ] Amazon 账号弹窗逐项说明 Seller ID、LWA Client ID、Secret、Refresh Token，并能显示分项测试结果
- [ ] 没有 AI、没有 Amazon 账号、没有文件时，工作台空状态都给出下一步按钮
- [ ] 示例 Excel 可以生成并导入
- [ ] 正式提交前必须有同账号、同站点、24 小时内的预览通过记录，并出现二次确认
- [ ] Excel 写回前会生成 `backups/` 备份；Excel/WPS 占用时提示 `EXCEL_LOCKED`

## 6. 售后资料

- [ ] `docs/客户部署说明.md` 与实际包内文件名一致
- [ ] `README.md` 与实际命令/入口一致
- [ ] `CHANGELOG.md` 已记录本次防呆、Doctor、打包变化
- [ ] 客户出问题时，只让客户点 `一键检测修复` 和 `导出支持包`，不要让客户手动找日志
