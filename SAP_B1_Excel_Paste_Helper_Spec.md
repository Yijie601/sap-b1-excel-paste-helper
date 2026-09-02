# SAP B1 Excel Paste Helper — Product & Technical Specification

## 1. 项目目标

开发一个轻量、快速、稳定的 Windows 桌面工具，用来把 Excel 中已经整理好的 AP Invoice 数据快速填入 **SAP Business One — AP Invoice**。

这个工具：

- 不做 OCR
- 不需要 AI
- 不直接读取 Excel 文件
- 只读取 Windows Clipboard
- 不自动点击 SAP 的 **Add / Update**
- 不自动切换窗口
- 不依赖 Python
- 一次 F8 自动完成五个动作，Header paste 之间固定异步等待 **0.8 秒**

目标体验：

```text
Excel
Ctrl+C
↓
用户自己切到 SAP AP Invoice
↓
按 1 次 F8 / Logitech 鼠标侧键
↓
自动依次填写 4 个 Header + 整块 Items
↓
停止
↓
用户人工检查
↓
用户自己按 Add
```

## 2. 推荐技术栈

建议：

```text
C#
.NET 8
WinForms
Win32 API / SendInput
```

发布目标：

```text
Windows x64
Self-contained
Single-file EXE
```

例如：

```bash
dotnet publish -c Release -r win-x64 --self-contained true
```

最终希望公司电脑：

- 不需要 Python
- 不需要另外安装 .NET Runtime
- 直接运行 EXE
- 尽量不需要 Administrator

如果 SAP Business One 本身以 Administrator 权限运行，则 Helper 可能也需要相同权限。

## 3. 用户日常操作流程

正常 workflow：

```text
1. 用户在 Excel 选择同一张 invoice 的 B:N
2. Ctrl+C
3. 用户自己 Alt+Tab / 点击切换到 SAP
4. 确保 SAP AP Invoice 已打开
5. 按一次 F8
6. Helper 按 Supplier → Posting Date → Supplier Ref. → Remarks 自动粘贴，每次间隔 0.8 秒
7. Helper 最后一次粘贴全部 E:N Items
8. 用户检查
9. 用户自己按 Add
```

支持把 Logitech Options+ 鼠标侧键设置为：

```text
F8
```

这样实际操作可以变成：

```text
Ctrl+C
↓
切 SAP
↓
鼠标侧键
↓
完成
```

## 4. Excel Clipboard 数据结构

用户复制范围固定为：

```text
B:N
```

不要包含 Excel Header。

对应：

| Excel Column | 内容 |
|---|---|
| B | Supplier Name |
| C | Document Date |
| D | Document Number |
| E | SAP Code / Item No. |
| F | Outlet |
| G | QTY |
| H | Total |
| I | Unit Price |
| J | VAT Code |
| K | Department / Blank |
| L | Discount |
| M | UoM / Blank |
| N | Whse |

Clipboard 中 Excel 数据一般为：

```text
Column = TAB (\t)
Row = CRLF (\r\n)
```

程序必须：

- 保留空白 cell
- 不能 Trim 掉中间空 column
- 不能因为某些 cell 为空就改变 column 数量

正常：

```text
B:N = 13 columns
```

如果不足 13 columns：

```text
Invalid Excel selection.
Please copy columns B:N.
```

## 5. Invoice 判断逻辑

一次 Clipboard 只能包含一张 Invoice。

所有 rows 的以下字段必须相同：

```text
Supplier Name
Document Date
Document Number
```

例如：

```text
Row 1: LIM SOON POH TRADING / 13-08-2026 / 260813/162
Row 2: LIM SOON POH TRADING / 13-08-2026 / 260813/162
Row 3: LIM SOON POH TRADING / 13-08-2026 / 260814/001
```

必须停止：

```text
Multiple invoices detected.
```

不要继续操作 SAP。

## 6. Header 数据

Header 只取第一行：

```text
Supplier Name = Column B
Document Date = Column C
Document No   = Column D
```

例如：

```text
Supplier Name: LIM SOON POH TRADING
Date:          13-08-2026
Document No:   260813/162
```

## 7. Supplier Mapping

Excel 只有 Supplier Name，但 SAP 要 BP / Supplier Code。

例如：

```text
LIM SOON POH TRADING → VL1002
```

程序需要 Supplier Mapping。

第一版可以使用：

```text
supplier_mapping.csv
```

例如：

```csv
Supplier Name,SAP Code
LIM SOON POH TRADING,VL1002
```

规则：

- Trim whitespace
- Ignore case
- 不做 fuzzy match
- 不使用 AI 猜 Supplier
- 找不到时直接停止

错误：

```text
Supplier mapping not found:
XXXXX
```

后续软件 UI 可增加：

```text
Supplier Mapping Manager
```

支持：

- 新增
- 编辑
- 删除
- 搜索 Supplier
- Import / Export CSV

## 8. Date Conversion

Excel 可能是：

```text
13-08-2026
13/08/2026
13.08.2026
2026-08-13
```

统一转换成 SAP：

```text
13.08.26
```

支持：

```text
dd-MM-yyyy
dd/MM/yyyy
dd.MM.yyyy
yyyy-MM-dd
```

输出：

```text
dd.MM.yy
```

## 9. SAP Header Mapping

### Supplier

Excel：

```text
LIM SOON POH TRADING
```

直接使用 Excel Supplier Name，不做 mapping：

```text
LIM SOON POH TRADING
```

填写到 SAP：

```text
Supplier
```

输入 Supplier 后不发送 TAB。下一步直接点击 Posting Date 并输入日期，让离开 Supplier
字段的动作触发 SAP commit；等待 SAP 完成 Supplier/date processing 后，再点击 Supplier Ref.。

SAP 之后会加载：

```text
Supplier Name
Contact Person
Currency
Item Matrix availability
etc.
```

### Supplier Ref. No.

点击已校准的 Supplier Ref. 坐标后粘贴：

```text
Document Number
```

例如：

```text
260813/162
```

### Posting Date

填写：

```text
Document Date
```

例如：

```text
13.08.26
```

### Document Date

不再单独点击或填写。输入 Posting Date 后由 SAP 自动更新 Document Date。

### Remarks

填写：

```text
Document Number
```

例如：

```text
260813/162
```

## 10. SAP Item Matrix

Excel：

```text
E:N
```

已经按 SAP Matrix 顺序整理。

SAP Matrix 对应：

```text
Item No.
Outlets/Location
Quantity
Total (LC)
Unit Price
VAT Code
Department
Discount %
UoM Name
Whse
```

程序不要逐 cell 输入。

应该重新构造 TAB-separated 数据：

```text
ItemNo\tOutlet\tQty\tTotal\tUnitPrice\tVAT\tDepartment\tDiscount\tUoM\tWhse
```

多行：

```text
row1\r\n
row2\r\n
row3
```

然后：

```text
Click first Item No. cell
Ctrl+V
```

SAP Business One 已经确认支持：

```text
multi-row
multi-cell
Excel-style paste
```

因此无论：

```text
1 row
2 rows
10 rows
30 rows
```

都只进行一次 Clipboard Paste。

## 11. 已确认的 SAP UI 行为

未点击字段时，Windows UI Automation 通常只能看到：

```text
Name: AP Invoice
ControlType: Window
ClassName: TMMDIChildClass
FrameworkId: Win32
```

点击实际 input 后，可以检测到：

```text
ClassName: TMEditTextClass
```

已经确认以下位置点击后会激活真正输入控件：

```text
Supplier
Supplier Ref.
Posting Date
Remarks
First Item No.
```

因此正式版：

- 不需要 AI Vision
- 不需要 OCR
- 不需要截图识别字段

## 12. Desktop Coordinate Mode

根据公司电脑的实际兼容性测试，正式 runtime 使用用户捕获的绝对 Windows desktop coordinates。

正常 F8 不依赖：

```text
SAP process name
A/P Invoice window title
TMMDIChildClass
TMEditTextClass
AP Invoice window rectangle
```

要求用户让 SAP 保持最大化，或保持与 Calibration 相同的桌面位置。移动 SAP、改变缩放、分辨率或显示器排列后需要重新 Calibration。

## 13. Absolute Desktop Coordinates

每一台电脑都必须实际捕获：

```text
Supplier
Supplier Ref. No.
Posting Date
Remarks
First Item No.
```

不能使用内置 default values。配置必须带 absolute-coordinate format version；旧 relative coordinates 一律视为未校准。

## 14. Calibration / 自定义位置

这是正式版的重要功能。

位置不能长期写死。

程序需要：

```text
Calibration Mode
```

### Calibration UX

Settings：

```text
SAP Calibration

Supplier          [Capture]
Supplier Ref.     [Capture]
Posting Date      [Capture]
Remarks           [Capture]
First Item No.    [Capture]

[Test Calibration]
[Save]
```

点击：

```text
Capture Supplier
```

程序提示：

```text
Click the Supplier field in SAP.
```

用户点击 SAP 对应位置。

程序直接保存：

```text
Mouse Desktop X/Y
```

例如：

```text
Supplier
X = 249
Y = 62
```

保存：

```json
{
  "coordinateVersion": 2,
  "supplier": { "x": 249, "y": 62 },
  "supplierRef": { "x": 249, "y": 97 },
  "postingDate": { "x": 1799, "y": 80 },
  "remarks": { "x": 249, "y": 817 },
  "itemNo": { "x": 536, "y": 283 }
}
```

建议保存到：

```text
calibration.json
```

## 15. Calibration 的关键原则

Calibration 是：

```text
一次性 / 偶尔设置
```

不是每次 F8 都重新扫描。

正常 runtime：

```text
读取 calibration.json
↓
读取 absolute desktop X/Y
↓
直接 Click / SendInput
```

正常 F8 流程绝对不要：

```text
OCR
AI Vision
Screenshot recognition
Repeated UI tree scanning
寻找文字标签
```

这些会增加：

- latency
- instability
- CPU usage
- false detection

Calibration 只负责把位置保存好。

## 16. Test Calibration

增加：

```text
Test Calibration
```

作用：

只依次把鼠标移动到：

```text
Supplier
Supplier Ref.
Posting Date
Remarks
Item No.
```

不要输入任何内容。

用户肉眼确认：

```text
位置是否正确
```

如果 SAP resolution / layout 改变，可以重新 Capture。

## 17. Global Hotkey

默认：

```text
F8
```

要求：

- Helper 可以最小化到 tray
- 软件不在 foreground 也能收到 F8

推荐使用：

```text
RegisterHotKey
```

Settings 可以修改：

```text
F8
Ctrl+Alt+F8
Ctrl+Shift+F8
etc.
```

Logitech Options+：

```text
Mouse Side Button → F8
```

## 18. 不自动切换窗口

最终需求已经确认：

```text
Helper 不负责切到 SAP
```

用户自己：

```text
Excel Ctrl+C
↓
Alt+Tab / Click SAP
↓
F8
```

F8 后不会侦测 SAP process、窗口标题或内部控件。Helper 会按校准保存的
Windows desktop absolute coordinates 直接点击与粘贴，因此用户必须先把 SAP
A/P Invoice 放在校准时相同的屏幕、位置与大小，再按 F8。

如果 SAP 不在预期位置，用户应停止操作并重新校准；Helper 不会猜测或切换窗口。

不要自己：

```text
Alt+Tab
SetForegroundWindow to guessed window
Search Excel then switch SAP
```

## 19. 单次 F8 执行顺序

```text
1. Excel Ctrl+C 后读取并验证 B:N，解析第一行 B:D，同时建立所有行的 E:N block
2. F8：点击 Supplier，粘贴第一行 B，然后异步等待 0.8 秒
3. 点击 Posting Date，粘贴转换后的第一行 C，然后异步等待 0.8 秒
4. 点击 Supplier Ref.，粘贴第一行 D，然后异步等待 0.8 秒
5. 点击 Remarks，粘贴第一行 D，然后异步等待 0.8 秒
6. 点击 First Item No.，一次粘贴所有已复制行的 E:N block
7. 每个动作都把原始 B:N 恢复为 plain text；不恢复可能阻塞 UI 的完整 Excel OLE formats
8. Items 完成后停止
```

绝对不要：

```text
20. Add
```

## 20. 输入实现

正式版优先：

```text
Win32 SendInput
```

需要：

```text
SetCursorPos
SendInput mouse click
SendInput keyboard
RegisterHotKey
```

Clipboard：

```text
System.Windows.Forms.Clipboard
```

不建议把：

```text
SendKeys.SendWait
```

作为长期核心实现。

Prototype 可以使用，但 production 建议使用 Win32 SendInput。

## 21. 性能要求

正式 EXE 必须明显比 PowerShell Prototype 快。

主要原则：

```text
不要每个 field 固定 Sleep 300–1000ms
```

普通 field：

```text
Posting Date
Remarks
```

应该：

```text
Click
20–50 ms
Paste/Input
继续
```

目标每个普通 field：

```text
约 30–100 ms
```

## 22. Supplier 等待策略

公司 SAP 对 synthetic Tab 不稳定，因此 Helper 不发送 Tab，而是在一次 F8 内按坐标点击每个栏位，并使用固定 0.8 秒 paste guard。

```text
Fill Supplier
↓
Wait 0.8s
↓
Fill Posting Date
↓
Wait 0.8s
↓
Fill Supplier Ref.
```

这样不需要 SAP process、控件或 busy-state detection；所有等待均使用异步 Task.Delay，不阻塞 Helper UI。

## 23. Item Paste 性能

Items 必须保持：

```text
E:N entire block
↓
ONE Ctrl+V
```

不要：

```text
row by row
cell by cell
```

Clipboard string 构造应该是毫秒级。

例如：

```text
10 rows
30 rows
50 rows
```

程序本身只做一次 paste。

之后的等待主要取决于：

```text
SAP validation / calculation
```

而不是 Helper。

## 24. 性能目标

公司电脑兼容性优先，参考目标：

```text
Clipboard parse       < 10ms
Read calibration      < 5ms
Each header action    one click + one paste
Between pastes        fixed asynchronous 0.8s guard
Final item action     single E:N paste operation
```

正常：

```text
2–5 item rows
```

目标：

```text
一次 F8
→
自动完成五个 paste actions

约 4–7 秒，另加 SAP Items 计算时间
```

如果 item rows 很多，SAP 自己计算时间可以另外计算。

Logging 中需要记录：

```text
Automation Duration
```

方便以后优化。

## 25. Clipboard Restore

用户原本复制的是：

```text
B:N full invoice
```

Helper 过程中会临时把：

```text
Supplier Code
Document Ref
Date
E:N Item Table
```

放入 Clipboard。

完成以后必须恢复：

```text
original B:N clipboard
```

即使错误发生，也尽可能 restore。

## 26. Safety

### Multiple Invoice

如果：

```text
Supplier / Date / Document No.
```

任何 row 不一样：

```text
STOP
```

### Unknown Supplier

```text
STOP
```

不要猜。

### SAP Not Foreground

```text
STOP
```

### AP Invoice Not Found

```text
STOP
```

### Wrong Column Count

如果不是至少：

```text
13 columns
```

提示：

```text
Invalid Excel selection.
Please copy columns B:N.
```

### Header Row Detection

如果第一行包含：

```text
Supplier Name
Document Date
Document Number
```

提示：

```text
Excel header detected.
Please copy data rows only.
```

### Add / Update

软件：

```text
NEVER click Add
NEVER click Update
```

由用户人工完成。

## 27. 软件 UI

不要复杂。

Main Window：

```text
SAP B1 Excel Helper

Status
● Ready

Hotkey
F8

SAP
Detected / Not detected

Supplier mappings
136

[Supplier Mapping]
[Calibration]
[Settings]
[View Log]
```

支持：

```text
Minimize to Tray
Start with Windows (optional)
```

Tray：

```text
SAP Helper
```

右键：

```text
Ready
Run Now
Supplier Mapping
Calibration
Settings
Open Log
Exit
```

## 28. 成功提示

成功时不要弹 modal dialog。

使用：

```text
Windows Toast
Tray Balloon
Small Overlay
Sound
```

例如：

```text
✓ 260813/162
2 item rows pasted
1.12s
```

约：

```text
1–2 秒后消失
```

错误时才使用 Modal Dialog。

## 29. Logging

记录：

```text
Timestamp
Supplier Name
SAP Code
Document Number
Row Count
Status
Error
Duration
```

例如：

```text
2026-08-28 10:30:12
SUCCESS
Supplier: LIM SOON POH TRADING
Code: VL1002
Ref: 260813/162
Rows: 2
Duration: 1.14s
```

## 30. 推荐项目结构

```text
SapB1ExcelHelper/
│
├── Program.cs
├── MainForm.cs
│
├── Services/
│   ├── ClipboardService.cs
│   ├── ExcelClipboardParser.cs
│   ├── SapAutomationService.cs
│   ├── HotkeyService.cs
│   ├── CalibrationService.cs
│   └── SupplierMappingService.cs
│
├── Models/
│   ├── InvoiceClipboardData.cs
│   ├── InvoiceItem.cs
│   ├── SupplierMapping.cs
│   └── SapCalibration.cs
│
├── Config/
│   ├── suppliers.json
│   └── calibration.json
│
└── Logs/
```

## 31. Data Models

### InvoiceClipboardData

```csharp
public class InvoiceClipboardData
{
    public string SupplierName { get; set; }
    public string SupplierCode { get; set; }
    public DateTime DocumentDate { get; set; }
    public string DocumentNumber { get; set; }

    public List<InvoiceItem> Items { get; set; }
}
```

### InvoiceItem

```csharp
public class InvoiceItem
{
    public string ItemNo { get; set; }
    public string Outlet { get; set; }
    public string Qty { get; set; }
    public string Total { get; set; }
    public string UnitPrice { get; set; }
    public string VatCode { get; set; }
    public string Department { get; set; }
    public string Discount { get; set; }
    public string Uom { get; set; }
    public string Warehouse { get; set; }
}
```

### SapCalibration

```csharp
public class SapPoint
{
    public int X { get; set; }
    public int Y { get; set; }
}

public class SapCalibration
{
    public SapPoint Supplier { get; set; }
    public SapPoint PostingDate { get; set; }
    public SapPoint Remarks { get; set; }
    public SapPoint ItemNo { get; set; }
}
```

## 32. MVP 第一阶段

第一阶段只完成：

```text
✅ Windows EXE
✅ Self-contained x64
✅ Tray app
✅ Global F8
✅ Clipboard B:N parser
✅ Preserve blank cells
✅ Header detection
✅ Same invoice validation
✅ Direct Supplier Name paste
✅ Date conversion
✅ Absolute desktop coordinate targeting
✅ calibration.json
✅ Supplier
✅ Supplier Ref. via calibrated direct click
✅ Posting Date
✅ Document Date auto-update by SAP
✅ Remarks
✅ E:N entire block paste
✅ Restore clipboard
✅ Success notification
✅ Error handling
✅ Logging
✅ Duration measurement
✅ NEVER click Add
```

## 33. 第二阶段

完成 MVP 稳定后增加：

```text
Calibration UI
Supplier Mapping UI
Editable Hotkeys
Start with Windows
Test Calibration
Import / Export Mapping
Better SAP ready detection
Performance telemetry
Optional operation history
```

## 34. 核心 UX 原则

软件存在感要非常低。

正常成功流程：

```text
Ctrl+C
↓
切 SAP
↓
F8 / Logitech Side Button × 1
↓
0.8 秒间隔自动完成五个动作
↓
Done
```

不要：

```text
Open Window
Next
Confirm
OK
Next
Paste
Confirm
```

日常成功流程：

```text
0 dialogs
```

## 35. 核心性能原则

Calibration 是：

```text
setup-time
```

而不是：

```text
runtime detection
```

正常 F8 runtime 不允许为了寻找 field 而反复执行：

```text
OCR
AI Vision
Screenshot Matching
Repeated UI Automation Tree Scanning
```

正常 runtime 应该：

```text
Read saved absolute desktop coordinates
+
Direct Win32 SendInput
```

目标是可靠的单次 F8 Header 顺序输入；Items 始终保持一次性 E:N 整块粘贴。

## 36. 给 Codex 的最终开发要求

> Build this as a production-quality Windows utility rather than another PowerShell prototype.
>
> Prioritize speed, reliability, minimal user interaction, and safe failure.
>
> Calibration must be a one-time setup workflow: capture each SAP field by having the user click it and save the exact absolute Windows desktop coordinates.
>
> Runtime automation must NOT use OCR, computer vision, screenshots, or repeated UI-tree scanning to locate fields.
>
> Use the stored absolute desktop coordinates and direct Win32 SendInput during normal F8 execution. Do not require SAP process, title, child-window, or control-class detection.
>
> Prefer reliable field focus and Supplier commit timing on company PCs. Keep delays bounded and paste the item matrix only once.
>
> Paste the entire E:N item matrix as one tab-delimited clipboard block.
>
> Never automate the final SAP Add or Update action.
