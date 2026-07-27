# VDIReportGeneratePart — 项目说明

> 本文件是给 AI agent 看的项目快照。下次调用时**先 Read 这份文档**，即可在不重新探索仓库的情况下接手任何改动任务。最后更新：2026-07-27。

## 1. 项目目标

BMW 内部 IT 支持用的两条独立 Excel 报表自动化流水线，跑在 Windows 工作站上、读写 BMW 内网 UNC 共享（`\\china.bmw.corp\winfs\...`），只能在能访问该内网的 Windows 机器上运行，**不能在 dev sandbox 跑**。

| 入口脚本 | 干什么 |
|---|---|
| `run.py` | 合并 VDI VMware 报表（Z3 + Z4）→ 与 AD Nameanddepartment 报表 join → 输出 `VDI_final.xlsx` + `VDI_AD_final.xlsx` |
| `run_asset.py` | 读 Checkout PC List 资产报表 → 用 Selenium 自动驱动 `wass.bmwgroup.net` Inventory 向导，生成 VDI 笔记本 last-login 报表并下载 .xlsx |

两条流水线**互相独立**，可以单独跑。

## 2. 仓库结构

```
VDIReportGeneratePart/
├── run.py                       # VDI+AD 合并入口
├── run_asset.py                 # asset->wass 入口
├── requirements.txt
├── .gitignore                   # *.xlsx 已忽略，输出不入库
└── vdi_report/
    ├── __init__.py              # 暴露 VDIReportAnalyzer
    ├── config.py                # 所有路径/列名/常量（单一真相源）
    ├── analyzer.py              # VDIReportAnalyzer：Step1 merge + Step2 join
    ├── vdi_merger.py            # Step1：Z3+Z4 合并
    ├── ad_lookup.py             # Step2：AD join、explode 多用户
    ├── paths.py                 # 共享上找最新 VDI/AD 文件夹
    ├── asset_pipeline.py        # asset->wass 编排
    ├── asset_reader.py          # 读 Checkout PC List，分块
    └── wass_automation.py       # Selenium 驱动 wass 向导（最复杂，bug 多在这里）
```

## 3. config.py 关键常量

| 常量 | 含义 |
|---|---|
| `VDI_REPORT_ROOT` | VDI VMware 报表根（`...VDI auto report\VDI VMware report`） |
| `AD_REPORT_ROOT` | AD NameanddepartmentDailyReportExport 根 |
| `ASSET_REPORT_ROOT` | Checkout PC List 资产报表根（`...VDI auto report\Asset report`） |
| `ASSET_FILE_PREFIX` | `"Checkout PC List"` |
| `WASS_URL` | `https://wass.bmwgroup.net` |
| `WASS_IMPORT_CHUNK_SIZE` | `4999`（wass 单次导入上限 5000，留 1 余量） |
| `WASS_RESULT_ELEMENTS` | `["Last Logon", "Last User", "Last User (Account)", "Last User (Email)"]` — 注意空格，必须与 wass picker 中 `<div class="EmText EmTextClick">` 文本完全一致 |
| `WASS_BROWSER` | `"edge"`（默认，Windows 自带） |
| `WASS_REPORT_TIMEOUT_SECONDS` | `1800` |
| `Z3_PREFIX` / `Z4_PREFIX` | `"Z3"` / `"Z4"` |

所有路径调整、列名变更、元素名变更都在 `config.py` 改，**其他文件不需要动**。

## 4. 流水线 A：VDI + AD 合并（`run.py`）

```
find_latest_vdi_folder(VDI_REPORT_ROOT)
   → 找到 Z3*.xlsx + Z4*.xlsx
   → merge_vdi_reports：concat 两帧，保留 [Id, IPv4 Address, Assigned Users]
   → 加 __source__ 列（Z3/Z4）
   → 写 VDI_final.xlsx

find_latest_ad_report(AD_REPORT_ROOT)
   → 读 AD 报表，保留 [Account, DepartmentCode, Name, EmailAddress]
   → explode_assigned_users：把 "user1;user2" 拆成多行
   → join_vdi_with_ad：VDI.AssignedUsers ≈ AD.Account
   → 写 VDI_AD_final.xlsx
```

CLI：`python run.py [-v] [--vdi-root X] [--ad-root X] [--output-dir X] [--no-explode] [--keep-source]`

## 5. 流水线 B：asset → wass（`run_asset.py`）

### 5.1 读资产 + 分块

`asset_reader.find_latest_asset_folder` 从今天往回找 `YYYY-MM-DD` 子文件夹（最多 30 天），里面要有 `Checkout PC List MMDD.xlsx`。`load_asset_names` 读 `Name` 列、去重（保留顺序，避免 4999 配额被重复名浪费）。`chunk_names` 按 4999 切块；`chunk_report_name`：第 0 块用基础名 `VDI-laptop-lastlogin-YYYY-MM-DD`，后续块加 `-02`、`-03`...

### 5.2 wass 向导 10 步（`wass_automation.WassReportAutomator.create_report`）

| # | 动作 | locator |
|---|---|---|
| 1 | 打开 `wass.bmwgroup.net`，等 Inventory tab 可见（首次需手动 SSO） | `#MainNaviBottom_2` |
| 2 | Inventory → Create Report → New Report | XPath `_ci("create report")` / `_ci("new report")` |
| 3 | 填报告名 | `#Name` |
| 4 | 点 Import 图标 → 点 "Import computer list" 链接 | `#ImportList` / `a.EmTextLink[href*='Host_Import']` |
| 5 | 把 4999 个名字用 **JS 一次性** 写入 `#ListResult` textarea（不用 send_keys，否则逐字符发送会卡死） | `#ListResult` |
| 6 | 等 Matrix42 弹确认框 → blindly 点 OK/Yes/Confirm | `_dismiss_import_confirmation` |
| 7 | 等 15s 让服务器处理 → 点 Step1 Next（onclick 含 `Inv_Rep_Step2.aspx`） | `_click_button_by_onclick` |
| 8 | 点 `#Output` 打开 Result elements picker → 展开 Login 组 → 用 `CheckBoxToggle('Choice_NNN')` 勾选 4 个元素 → 点 OK（onclick 含 `ClosePopup('Level3')`） | `#Output` |
| 9 | Step2 Next（`Inv_Rep_Step3.aspx`）→ Generate OK（`Save_Report.aspx`） | `_click_button_by_onclick` |
| 10 | 轮询刷新按钮直到 "Report successfully created" → 右键行 → Result → Excel Download → 点下载链接 → 等 .xlsx 落到 `download_dir` | `img[alt='Update']` / `result_menu_item` / `excel_download_button` / `download_link` |

多块时 `asset_pipeline._merge_reports` 把所有下载的 .xlsx 合并成 `VDI-laptop-lastlogin-<date>_merged.xlsx`。

CLI：`python run_asset.py [-v] [--browser edge|chrome] [--headless] [--chunk-size N] [--dry-run] [--asset-root X] [--download-dir X]`

`--dry-run` 只做读+分块，不开浏览器，用来验证资产文件解析。

## 6. wass_automation.py 关键设计点

### 6.1 Matrix42 Button widget 匹配

wass 的按钮不是标准 `<button>`，而是 `<div id="Button<guid>" class="ui-button..."><span class="ui-button-text">Label</span></div>`，onclick 通过紧随其后的 `<script>` 里的 `EmControls.Events.SetOnClick("btnId", {"OnClick":"..."})` 动态绑定。**同一页有多个同名按钮**（如多处 "Next"/"OK"），所以必须用 onclick 内容来消歧。

`_click_button_by_onclick(name, label, onclick_contains)`：
- XPath 找所有 `<div id^="Button">` + 内部 `<span class="ui-button-text">` 文本等于 label 的候选
- 对每个候选，JS 遍历整个 `document.getElementsByTagName('script')` 找到提到该 btn_id 的 script 文本
- 检查 `onclick_contains` 是否是该 script 文本的子串 + 按钮可见 → 点击

### 6.2 ⚠️ 已修复的关键 bug：`\uXXXX` 转义

**症状**：勾完 4 个 result elements 后点 OK 卡住，报 `button 'result_elements_ok_button' (label='OK', onclick~'ClosePopup('Level3')') not found`，候选显示 `script_has_target=False`。

**根因**：Matrix42 把 onclick handler 存在 JSON 字符串里，单引号被转义成 `\u0027`。`<script>.textContent` 返回**原始源码**，6 个字面字符 `\ u 0 0 2 7`，浏览器只在**执行**时才解释成 `'`。所以代码查字面量 `ClosePopup('Level3')` 永远匹配不上。

**修复**（已合入）：
1. 新增 `_decode_js_unicode(s)`：用 `re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1),16)), s)` 把 `\uXXXX` 解码成真字符
2. `_click_button_by_onclick` 在子串匹配前先 `script_norm = _decode_js_unicode(script)`，用 `script_norm` 比对
3. 新增 `_click_any_visible_button(name, label, timeout)` 作为更稳的 fallback：只挑可见的、span 文本精确等于 label 的 Matrix42 Button div，替换原来太松的 `_click("ok_button")` / `_click("next_button")`

`select_result_elements` 和 `finish_wizard` 的 fallback 已切到新 helper。

### 6.3 ⚠️ 已修复的关键 bug：右键菜单 "Result" 点不到

**症状**：报告跑完后右键行，菜单弹出（Execute/Result/Copy/Edit/Delete，Result 是第 2 项），但 "Result" 没被点到。

**根因**：wass 用 jQuery contextMenu 插件，菜单是 `<ul class="context-menu-list"><li class="context-menu-item"><span>Result</span></li></ul>`。旧 locator `text()='result'` 会匹配页面上**任何**文本是 "Result" 的元素——包括报表列表的 "Result" 列头。列头在 DOM 顺序上排在菜单 `<span>` 前面，所以 `_click` 点到列头，菜单项没触发。

**修复**（已合入）：
1. `result_menu_item` locator 改为限定到 `ul.context-menu-list` 内的 `li.context-menu-item`
2. 新增 `_click_context_menu_item(name, label, timeout)`：scope 到 `.context-menu-item`，等 **presence**（不是 clickable——jQuery contextMenu 的 `<li>` 常过不了 Selenium 的 clickable 检查），JS click 触发插件委托 handler
3. `download_result` 先试新 helper，失败再退回 `_click("result_menu_item")`

### 6.4 ⚠️ 已修复的关键 bug：报告列表里多条同名报告，找错了行

**症状**：报表列表累积了用户所有历史报告，按名字匹配会返回多条（重跑、或 `-02`/`-03` chunk 后缀子串匹配基础名），旧代码取第一个匹配——不一定是刚生成的那条。

**根因**：旧 `download_result` 用 `//tr[td[contains(.,'报告名')]]` 取第一个匹配。`contains` 会把 `VDI-laptop-lastlogin-2026-07-27` 和 `-02` 后缀的都匹配上，且不保证顺序。

**wass 报告行结构**（从用户 HTML 提取）：
```html
<tr cr="true" cm="Menu1" cp="#id#|32806||" class="TrMarked">
  <td><div class="EmIcon" style="...Report_Success.png"></div></td>  <!-- 图标 -->
  <td>32806</td>              <!-- 报告 ID（单调递增）-->
  <td>name test</td>          <!-- 报告名 -->
  <td>Private</td>            <!-- scope -->
  <td>Roc Xie, FG-CN-60</td>  <!-- owner -->
  <td>27.07.2026, 13:09</td>  <!-- 创建时间（德式 DD.MM.YYYY, HH:MM）-->
</tr>
```

**修复**（已合入）：新增 `_find_report_row(report_name, timeout)`：
- 找所有 `tr[@cr='true' and @cm]`（带右键菜单的报告行）
- 名字 cell 文本**精确等于** report_name（不是子串），避免 `-02` 误匹配
- 从 `cp` 属性 `#id#|32806||` 解析报告 ID，fallback 到第 2 个 `<td>` 文本
- 挑 ID **最大**的那条（ID 单调递增 = 最新生成的）
- `download_result` 替换旧的 `row_xpaths` 循环

### 6.5 ⚠️ 已修复的关键 bug：右键报告行 StaleElementReferenceException

**症状**：报告状态刷新到 "Report successfully created" 后，`download_result` 找到了行（`picked id=32807`），但 `ActionChains.context_click(row)` 抛 `StaleElementReferenceException`，JS fallback 也抛同样的异常，程序终止。

**根因**：wass 报表列表用 jQuery/Kendo 渲染。`wait_for_completion` 每轮点 refresh 按钮，最后状态变 "Report successfully created" 后表格 DOM 被重建——`_find_report_row` 拿到的 `<tr>` 引用在重建瞬间失效。`ActionChains.context_click(row)` 拿着失效的 element 去操作就抛 stale；fallback `execute_script(..., row)` 也把同一个失效 element 当 `arguments[0]` 传过去，同样抛 stale。

**修复**（已合入）：
1. `wait_for_completion` 检测到完成后多等 2s，让表格 DOM 重建完
2. 新增 `_open_context_menu_for_report(name, max_attempts=5)`：retry loop 里 `_find_report_row` → 立刻 `context_click`，stale 就 sleep + 重新 find。还检查 `ul.context-menu-list` 是否真的弹出来了，没弹出就重试
3. 新增 `_wait_for_context_menu(timeout)`：等 jQuery contextMenu `<ul>` 可见
4. `download_result` 替换原来一次性 `context_click + JS fallback` 为调用新方法
5. import 加 `StaleElementReferenceException`

### 6.6 ⚠️ 已修复的关键 bug：同名报告状态读错（误判完成）

**症状**：`wait_for_completion` 1 秒内就报 "Report successfully created" 然后立刻去下载——但新报告可能还在跑。用户确认 wass 列表里有多条同名报告（重跑产生）。

**根因**（两层）：
1. 旧 `_read_status` 用 `//*[contains(.,'报告名')]//*[contains(.,'Report successfully created')]` 子串匹配，匹配到旧报告容器
2. 即使按 ID 找最新行，icon fallback 还有"行内文本含 successfully created"判断——`row.text` 会包含后代所有可见文本，可能吸入邻近元素的 "successfully created" 提示，导致新报告行也被误判完成

**修复**（已合入）：
1. `wait_for_completion` 提交前先记录当前最大 ID（`_max_report_id`），之后只信任 ID **严格大于**该值的新行——彻底忽略提交前已存在的同名报告
2. `_read_status` 加 `min_id` 参数，过滤掉旧报告行
3. `_read_row_status` 只信任 icon：`Report_Success`/`success` → DONE，其他 icon / 无 icon → RUNNING。**去掉所有文本 fallback**，避免吸入邻近元素的状态文本

### 6.7 ⚠️ 已修复的关键 bug：Excel Download 点击无反应

**症状**：右键报告 → Result 后，侧边栏 "Excel Download" 点击没反应，没切换到下载页面。

**根因**：Matrix42 把 onclick 绑在 `<td>` 上（通过 `EmControls.Events.SetOnClickMulti(new Array(td1_id, td2_id), {...})`），不是内部的 `<div class="EmText">`。旧 locator `_ci("excel download")` 优先匹配到 `<div>`，Selenium click 落在 div 上，但 Matrix42 的 handler 可能用 `event.target` 判断，事件没冒泡触发 td 的 handler → LoadContent 导航没执行。

**wass Excel Download HTML 结构**：
```html
<td id="cef5e254-..." class="EmText" style="cursor:pointer;">
  <div class="EmText">Excel Download</div>
</td>
<script>
EmControls.Events.SetOnClickMulti(new Array("td1_id", "td2_id"),
  {"OnClick":"javascript:HighlightElement(\u0027Result\u0027, ...);
   LoadContent(\u0027Level2_Form\u0027, \u0027StatusContentResult\u0027,
   \u0027Inv_Rep_Status_Download.aspx?id=32811\u0027);"});
</script>
```

**修复**（已合入）：新增 `_click_excel_download(timeout)`：
- 直接定位 `<td>`（带 `EmText` class 且含 "Excel Download" 文本），不点 `<div>`
- **用 `ActionChains.move_to_element(td).click()` 真实点击**（`isTrusted=true`）—— 旧版用 `td.click()` / JS `.click()` 产生 `isTrusted=false` 的 click 事件，Matrix42 handler 拒绝执行，右侧 "Download result file" 面板不出现
- 尝试 label td 和 icon td 两个 cell（`#Icon_Result` 父 td）
- **点击后验证右侧面板是否真的出现**（找 `.xlsx` 链接或 "Download result file" / "file was successfully created" / "file is being created" 文本），没出现就换下一个 cell
- 终极 fallback：从页面 `<script>` 正则提取完整 `OnClick` handler 字符串，`_decode_js_unicode` 解码 `\uXXXX`，strip `javascript:` 前缀，`eval` 执行整个 handler（含 `HighlightElement` + `LoadContent`），绕过 Matrix42 事件绑定
- `download_result` 先试新 helper，失败再退回 `_click("excel_download_button")`

### 6.8 ⚠️ 已修复的关键 bug：download link 点击不触发下载

**症状**：Excel Download 已成功跳转，右侧 "Download result file" 链接（`<a class="EmTextLink" href="...xlsx">`）已展示，但点击后没真正下载文件。

**根因**：和 Excel Download 同样的 `isTrusted` 问题——Matrix42 的 `EmTextLink` onclick handler 可能检查 `event.isTrusted`，Selenium `element.click()` / JS `.click()` 产生 `isTrusted=false` 的 click，handler 拒绝执行，浏览器不开始下载。

**修复**（已合入）：新增 `_click_download_link(timeout)`：
- 找 `a.EmTextLink[href*='.xlsx']`，`scrollIntoView` + `ActionChains.move_to_element(link).click()` 真实点击（`isTrusted=true`）
- **点击后验证文件是否真的开始下载**：检查 `download_dir` 里 8s 内是否出现新 `.xlsx` 或 `.crdownload`
- 没下载就换下一个 link 重试（已试过的 href 用 `tried_hrefs` 跳过）
- **终极 fallback**：`self._d.get(href)` 直接让浏览器导航到 xlsx URL，绕过 Matrix42 事件绑定，强制下载
- `download_result` 替换原来一次性 `_click("download_link")` 为调用新方法

### 6.9 ⚠️ 已修复的关键 bug：下载文件不在指定位置 / 弹 Save as 对话框

**症状**：用户反馈"下载文件的位置地址一定放在我指定的地方。可以选择 Save as，指定位置"——即 Edge 浏览器下载文件时可能弹出 Save as 对话框，文件不落到 `download_dir`，Selenium 卡住。

**根因**：BMW 企业 Edge 组策略覆盖了 Selenium 在 prefs 里设置的 `download.prompt_for_download: False` / `download.default_directory`，强制弹 Save as 对话框。

**修复**（已合入）三层保障：
1. **启动时 CDP 强制设置下载目录**：`_set_download_dir_via_cdp()` 调用 `driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": ...})`，CDP 命令受企业策略尊重，能把下载钉死在 `download_dir`
2. **`_click_download_link` 加 requests + cookies HTTP 下载 fallback**（新增 `_http_download`）：
   - 从 `<a href="...xlsx">` 提取 URL
   - 用 Selenium 的 session cookies + Referer + User-Agent 流式下载到 `download_dir`
   - **完全绕过浏览器下载 UI**，不弹任何对话框，文件 100% 落在指定位置
   - 401 时 fallback 到 `requests_ntlm.HttpNtlmAuth("", "")`（BMW 内网常用 Windows 集成认证）
3. **`_click_download_link` 三层策略**：ActionChains 真实点击 + 验证文件 → requests HTTP 下载 → `self._d.get(href)` 浏览器导航（依赖 CDP 钉的下载目录）
4. `requirements.txt` 加 `requests` + `requests_ntlm`

### 6.10 其他坑

- `import_computer_list` 用 JS 写 textarea value + dispatch input/change 事件，**不能** `send_keys`（4999 行会逐字符发送，Edge 会卡死）
- `import_computer_list` 末尾**不**点 OK——那个 OK 会关掉整个 New Report 弹窗，让用户回到报告名页。正确流程是提交 import link 后直接点 Step1 Next
- `_dismiss_import_confirmation` blindly 点 OK/Yes/Confirm 之一，Matrix42 消息框 OK 的 onclick 是 `CloseMessageBox()`
- 勾选 checkbox 用 `CheckBoxToggle('Choice_NNN')` 直接调 Matrix42 自带函数，不靠 DOM click 落点（picker 弹窗里下方 row 的 click 经常不触发）
- 整个文档可能渲染在 `<iframe>` 里，`_find_clickable` 先在 top document 找，找不到再遍历所有 iframe（一层深）
- 右键行用 `ActionChains.context_click(row)`，失败回退到 JS dispatch contextmenu 事件

### 6.11 部署注意事项

- 必须是能访问 `wass.bmwgroup.net` 的 BMW 内网 Windows 机器
- 首次跑要手动在浏览器完成 SSO 登录，`ensure_logged_in` 等 `login_timeout`（默认 120s）
- `--headless` 经常让企业 SSO 失败，默认不开
- Edge/Chrome driver 通过 `webdriver-manager` 自动下载缓存到 `~/.wdm/`
- 下载目录默认是当天资产 `YYYY-MM-DD` 文件夹

## 7. 依赖

`requirements.txt`：`pandas`、`openpyxl`、`xlrd`、`python-dateutil`、`selenium>=4.10`、`webdriver-manager`。Windows 上用 `.venv`。

## 8. 常见问题排查速查

| 症状 | 看哪里 |
|---|---|
| `Latest asset folder: ...` 不对 | `config.ASSET_REPORT_ROOT` / `ASSET_FILE_PREFIX` |
| `Column 'Name' not found` | 资产 xlsx 表头变了 → `config.ASSET_EXPECTED_COLUMNS` |
| 卡在 Inventory tab | SSO 没登 / `inventory_tab` locator |
| 卡在某个 Next/OK | `_click_button_by_onclick` 的 `onclick_contains` 字符串；候选诊断日志会列出 `script_has_target` |
| `script_has_target=False` 但按钮明显在 | 检查 JS 转义（`\uXXXX`），用 `_decode_js_unicode` 解码后再比 |
| 勾选的元素名不对 | `config.WASS_RESULT_ELEMENTS`，注意 `Last User (Account)` 空格 |
| 下载的 xlsx 找不到 | `download_dir` 权限 / `.crdownload` 未结束 / `download_link` locator |
| 报告一直 running 不结束 | `WASS_REPORT_TIMEOUT_SECONDS` / wass 后端慢，正常 |
| 报告 1 秒就 "successfully created" 但实际没跑完 | 同名报告状态读错 → `_read_status` 已改按最新 ID 行的 icon 判断（6.6 节） |
| 右键行抛 StaleElementReferenceException | 表格 DOM 重建导致 row 引用失效 → `_open_context_menu_for_report` retry（6.5 节） |

## 9. 开发约定

- 单一真相源：路径、列名、元素名都在 `config.py`，其他模块只 import 不重复定义
- locator 全部用文本/XPath（`_ci()` helper），不依赖易变的 id/class
- 所有 Selenium 操作都有 fallback：normal click 失败 → JS click；onclick 匹配失败 → 任意可见同名按钮；top document 找不到 → 遍历 iframe
- 日志用 `logging.getLogger(__name__)`，`run*.py` 配置 `%(asctime)s %(levelname)-7s %(name)s` 格式
- 不写测试（项目跑在内网，sandbox 无法验证）；改完用 `python -m py_compile` 验证语法
