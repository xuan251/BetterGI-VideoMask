# 给 AI 的安装指令 · VideoMask v1.6.28

> 用法：把本文件（或下面全部内容）直接发给你的 AI 助手，让它照做。
> 它需要的权限：读写本地文件、执行 PowerShell 命令。

---

## 你的任务

在用户的 Windows 电脑上安装一个 BetterGI 脚本（视频悬浮窗 VideoMask），
并且**一次性把权限和调度器配好**，不要留下"还要手动点几处"的尾巴。

**严格按 0→7 顺序执行。每完成一步先报告结果再继续；任何一步与预期不符就停下来问用户，不要猜。**

---

## 0. 前提

- 已装好 **BetterGI 0.50.0+**。
- 用户手上有解压好的 `VideoMask` 文件夹（含 `main.js`、`manifest.json`、`assets\`）。
  **只有 zip 没解压就先解压。**

---

## 1. 找到 BetterGI 安装目录

按优先级找，找到就停：

1. `D:\BetterGI`、`C:\BetterGI`、`D:\Program Files\BetterGI`、`%LOCALAPPDATA%\BetterGI`
2. 快捷方式目标：
   `(New-Object -ComObject WScript.Shell).CreateShortcut("$env:USERPROFILE\Desktop\BetterGI.lnk").TargetPath`
3. 直接问用户「右键桌面 BetterGI 图标 → 属性 → 起始位置」

**判定标准（必须同时满足）**：该目录下有 `BetterGI.exe` **且** 有 `User\JsScript\`。
不满足就继续找，别拿长得像的目录凑数。

---

## 2. 关掉 BetterGI，再复制文件

**必须先确认 BGI 进程已退出**（被占用时改动会被覆盖）。

```powershell
$bgi = "D:\BetterGI"                        # 换成第 1 步的真实路径
$src = "C:\path\to\VideoMask"               # 换成解压位置
$dst = Join-Path $bgi "User\JsScript\VideoMask"
if (Test-Path $dst) { Rename-Item $dst "VideoMask.bak.$(Get-Date -Format yyyyMMddHHmmss)" }
Copy-Item -Path $src -Destination $dst -Recurse -Force
Test-Path (Join-Path $dst "main.js")        # 必须返回 True
```

⚠️ 最常见的错：多套一层 → `VideoMask\VideoMask\main.js`。最后一行就是查这个。

---

## 3. 在调度器里建配置组、加入脚本、**同时授予 JS HTTP 权限**

这三件事其实是同一个文件：`<BGI>\User\ScriptGroup\<组名>.json`。

### 3.1 权限存在哪（重要）

每个 JS 脚本的「JS HTTP 权限」就是配置组里这条 project 的 **`allowJsHTTPHash`** 字段：

```json
{
  "name": "视频悬浮窗",
  "folderName": "VideoMask",
  "jsScriptSettingsObject": null,
  "index": 1,
  "type": "Javascript",
  "status": "Enabled",
  "schedule": "Daily",
  "runNum": 1,
  "allowJsNotification": true,
  "allowJsHTTPHash": "https://*.bilibili.com/*|https://bilibili.com/*|https://*.biliapi.net/*|https://*.hdslb.com/*|https://*.bilivideo.com/*|https://*.akamaized.net/*"
}
```

- `allowJsHTTPHash` = `manifest.json` 的 `http_allowed_urls` **按原顺序用 `|` 连起来**，一个字符都不能差。
- 空字符串 = 没授权。**BGI 的规则是：hash 与 manifest 不一致就要人工重新授权一次**——
  这就是为什么"没预置权限的人要重启两次才能索引到内容"。所以这一步必须一次写对。
- **不要手抄**，从 manifest 现算：

```powershell
$man  = Get-Content (Join-Path $dst "manifest.json") -Raw | ConvertFrom-Json
$hash = ($man.http_allowed_urls -join "|")
$dispName = if ($man.name) { $man.name } else { "视频悬浮窗" }
"hash = $hash"
```

### 3.2 写进配置组

```powershell
$sgDir = Join-Path $bgi "User\ScriptGroup"
$files = @(Get-ChildItem $sgDir -Filter *.json -ErrorAction SilentlyContinue)
```

**情况 A：已有配置组** → 取第一个（或用户指定的），在 `projects` 里：

- 已有 `folderName -eq "VideoMask"` 的条目 → 只更新 `allowJsHTTPHash`，并把 `status` 设为 `"Enabled"`。
- 没有 → 追加一条，`index` 取现有最大值 +1，其余字段照上面模板。

**情况 B：一个配置组都没有** → **不要凭空编造 `config` 结构**（那里有几十个
pathingConfig / autoFightConfig 字段，写错会让 BGI 读配置失败）。改为告诉用户：
「请在 BetterGI → 全自动 → 调度器 → 新建配置组，名字随意，建好告诉我」，
等用户建完再回来执行情况 A。

参考实现（情况 A）：

```powershell
$f = $files[0].FullName
Copy-Item $f "$f.bak" -Force
$g = Get-Content $f -Raw | ConvertFrom-Json
if (-not $g.projects) { $g.projects = @() }
$hit = $g.projects | Where-Object { $_.folderName -eq "VideoMask" }
if ($hit) {
  $hit.allowJsHTTPHash = $hash
  $hit.status = "Enabled"
} else {
  $maxIdx = ($g.projects | Measure-Object -Property index -Maximum).Maximum
  if (-not $maxIdx) { $maxIdx = 0 }
  $g.projects += [pscustomobject]@{
    name = $dispName; folderName = "VideoMask"; jsScriptSettingsObject = $null
    index = [int]$maxIdx + 1; type = "Javascript"; status = "Enabled"
    schedule = "Daily"; runNum = 1; allowJsNotification = $true; allowJsHTTPHash = $hash
  }
}
[System.IO.File]::WriteAllText($f, ($g | ConvertTo-Json -Depth 100),
  (New-Object System.Text.UTF8Encoding($false)))
```

写完**必须回读校验**：重新解析该文件，确认存在 `folderName = "VideoMask"` 且
`allowJsHTTPHash` 与 `$hash` 完全相等。不等就还原 `.bak` 并告诉用户。

> 这个文件由 BGI 自己维护，它开/关时会重写。所以**只能在 BGI 关闭时改**，改完再启动。

---

## 4. 打开「后台运行」（决定"自动过剧情会不会失效"）

浮窗会和 BGI 的自动剧情抢游戏前台焦点；没开后台运行时，游戏一失焦，自动剧情就点不到游戏。

改 `<BGI>\User\config.json` 里的 **`autoSkipConfig.runBackgroundEnabled`** → `true`。

**要求：定点字符串替换，不要把整个 JSON 读出来再序列化回去**（会改日期格式、打乱键序）。

```powershell
$p = Join-Path $bgi "User\config.json"
Copy-Item $p "$p.bak" -Force
$txt = [System.IO.File]::ReadAllText($p)
if ($txt -match '"runBackgroundEnabled"\s*:\s*(true|false)') {
    $txt = $txt -replace '"runBackgroundEnabled"\s*:\s*false', '"runBackgroundEnabled": true'
    [System.IO.File]::WriteAllText($p, $txt, (New-Object System.Text.UTF8Encoding($false)))
    "已打开后台运行"
} else {
    "没找到 runBackgroundEnabled，跳过并提示用户手动开"
}
```

⚠️ `commonConfig.mainBackgroundEnabled` 名字很像，其实是**主界面背景图开关**，别动它。

---

## 5. 校验清单

逐项检查并报告，全过才算装好：

| 检查 | 期望 |
|------|------|
| `<BGI>\User\JsScript\VideoMask\main.js` | 直接存在（不是多套一层） |
| `main.js` 内 `SCRIPT_VERSION` | `"1.6.28"` |
| `manifest.json` 的 `version` | `1.6.28` |
| `settings.json` 里 `marksOn` 的 `default` | `true`（打怪点暂停默认开） |
| `settings.json` 里 `hkPlayPause` 的 `default` | `"~"` |
| `assets\player.html` | 存在且含 `打开节目单` |
| 配置组 json | 有 `folderName = "VideoMask"`，`allowJsHTTPHash` 与 manifest 完全一致，`status = "Enabled"` |
| `config.json` | `autoSkipConfig.runBackgroundEnabled = true` |
| `main.js` 语法 | 有 node 就跑 `node --check main.js`，不报错 |

---

## 6. 只剩一步要人工

**打开 BetterGI → 刷新脚本列表 → 运行「视频悬浮窗」**（也可以在 全自动 → 调度器 里点「运行」）。

不需要再去点「JS HTTP 权限」了 —— 第 3 步已经写进配置组。

如果界面上那个权限下拉仍显示「未启用」：说明 hash 没对上。回第 3 步重算，
**不要**让用户手动点一下了事（点了以后脚本一更新又会掉）。

---

## 7. 出问题怎么查

日志：`<BGI>\log\better-genshin-impact*.log`（取最新）

| 搜什么 | 含义 | 处理 |
|--------|------|------|
| `VideoMask v1.6.28 启动` | 脚本起来了 | 没有 = 路径层级错或脚本列表没刷新 |
| `JS HTTP` / `不允许请求此 URL` | 权限没生效 | 回第 3 步核对 hash |
| `取直链失败` + `412` | B站风控 | 让用户在设置里填 SESSDATA |
| `取视频地址超时` | 网络 | 点画面上的「重试」 |
| `视频起播超时` | mp4 地址失效/慢 | 点「重新获取」 |

用户若反馈"过剧情还是不动"：确认 BGI 设置里的「后台运行」是开的，然后**重启 BGI**（配置改动有时要重启才生效）。

---

## 禁止事项

- ❌ 不要 `rm -rf` / 递归删除。覆盖旧版用**改名备份**。
- ❌ 不要动 `User\` 下与本次无关的文件。
- ❌ 不要改 `manifest.json` 的 `http_allowed_urls`（改错会导致接口全被拦）。
- ❌ 不要整份重写 `config.json`，只改那一个布尔值。
- ❌ 不要凭空编造配置组的 `config` 段（见 3.2 情况 B）。
- ❌ 不要替用户填 SESSDATA。

---

## 附：这是什么脚本

把 B 站原神攻略视频浮在游戏窗口上。暂停点来自 UP 主全集 566 集（4950 分钟）字幕统计：
说「镇守宝箱」后 82% 会叫你打怪 → 停；说「这里有个宝箱」只有 14% → 不停。
共 1526 个点（合并相近点后 1461），分「打怪点」和「看完再停（挑战/boss）」两类。
默认播放/暂停热键是 `~`（Esc 下面那颗），因为笔记本的 F9 常被亮度键占用。
