# 发版流程（维护者用）

> [!IMPORTANT]
> **特别说明（请先阅读）· 本项目仅供学习、研究与技术交流使用，不是商业产品，也不构成任何形式的服务承诺。**
>
> - 请只转换你**拥有合法权利**的文档，并遵守所在地法律法规。
> - 本工具不提供破解加密、绕过 DRM 或规避授权限制的功能。
> - 转换全程在本机完成：**不联网、不上传文件、不收集任何数据**。
> - 软件按「现状」（AS IS）提供，不担保转换结果完全无误；重要材料请**先备份并人工复核**。
> - 因使用或无法使用本软件产生的任何直接或间接损失，由你自行承担，开发者与贡献者不承担责任。
> - 与 Microsoft、金山办公（WPS）、Artifex 无任何隶属、赞助或背书关系，相关商标归各自所有者。
> - 完整条款见 [DISCLAIMER.md](DISCLAIMER.md)（共 11 节）；代码以 [MIT](LICENSE) 授权。

仓库只放源码和脚本，**exe / zip 等大文件一律作为 GitHub Release 附件上传**，不进 git。

> 发版时记得：Release 正文用 `.github/RELEASE_BODY.md`（CI 已固定使用），
> 不要删掉里面的「特别说明」块——这是对外文档必须带的三处声明之一。

## 1. 改版本号

- [ ] `app.py` 里的 `APP_VERSION = "x.y.z"`（安装包版本号、窗口标题都从这里取）
- [ ] `CHANGELOG.md` 顶部加一节 `## [x.y.z] — YYYY-MM-DD`

## 2. 自检 + 构建

```powershell
python selftest.py --quick                 # 全套检查见 README「自检」一节
powershell -ExecutionPolicy Bypass -File build.ps1
powershell -ExecutionPolicy Bypass -File installer\build_installer.ps1
```

## 3. 打包绿色版 + 校验值

```powershell
$v = (Select-String app.py -Pattern 'APP_VERSION\s*=\s*"([^"]+)"').Matches[0].Groups[1].Value
Compress-Archive -Path dist\PDF2Office -DestinationPath "PDF2Office-$v-Portable.zip" -Force
$setup = "installer\output\PDF2Office-$v-Setup.exe"
$files = @($setup, "PDF2Office-$v-Portable.zip")
$lines = foreach ($f in $files) { "$((Get-FileHash $f -Algorithm SHA256).Hash.ToLower())  $(Split-Path $f -Leaf)" }
$lines | Set-Content -Encoding ascii SHA256SUMS.txt
```

## 4. 提交 + 打 tag + 推送

```powershell
git add -A
git commit -m "release: x.y.z"
git tag -a vx.y.z -m "PDF2Office x.y.z"
git push origin main --tags
```

## 5. 发布 Release

推 `v*` tag 后，GitHub Actions（`.github/workflows/release.yml`）会自动构建并创建 Release。
**发版就走这一条路**：本地不构建发布包、也不上传附件。

> 为什么不本地上传：部分网络下本机到 `uploads.github.com` 的 TLS 会被重置，而
> `gh release upload --clobber` 是「先删后传」，一旦传失败 Release 就变成空的。
> 真需要手动兜底时，先用不同的附件名上传、确认成功后再删旧附件，不要直接 `--clobber`。
> 万一已经把 Release 弄空，重新跑一次 CI 即可恢复：`gh workflow run release.yml --ref vX.Y.Z`。

## 附件内容说明

| 附件 | 用途 |
| --- | --- |
| `PDF2Office-x.y.z-Setup.exe` | 安装包：开始菜单/桌面快捷方式 + 卸载项，默认装当前用户（无需管理员） |
| `PDF2Office-x.y.z-Portable.zip` | 绿色版：解压到任意文件夹（建议英文路径）双击 `PDF2Office.exe` |
| `SHA256SUMS.txt` | 两个附件的 SHA-256，方便用户校验 |

> 说明：安装包没有数字签名，用户首次运行可能看到 SmartScreen 提示，属正常现象。

> 完整免责声明见 [DISCLAIMER.md](DISCLAIMER.md)：本项目仅供学习与技术交流，非商业产品；因使用本软件产生的后果由使用者自行承担。
