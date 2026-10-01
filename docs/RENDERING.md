# 把排版好的 `.docx` 渲染成 PDF / 图片

本技能只负责**改格式**，渲染交给文字处理软件。三种路径，按推荐排序。

## ① LibreOffice（最省事）

```bash
soffice --headless --convert-to pdf --outdir . 论文_DUT.docx
```

不受 macOS 沙箱与 TCC 影响，Agent 会话内、终端里、CI 上都能跑。
缺字体时中文可能变方框，装上 SimSun / 黑体 / Times New Roman 即可。

## ② Microsoft Word（本机已装）

```bash
./scripts/convert_with_word.sh 论文_DUT.docx          # 输出同名 .pdf
./scripts/convert_with_word.sh 论文_DUT.docx 出.pdf   # 指定输出
```

**必须在系统「终端」里跑，不能在 Agent 沙箱里跑**，原因见下。

### 为什么沙箱里会失败

在 DSH 等带文件沙箱的 Agent 会话里执行 Word 自动化，会看到：

```
execution error: “Microsoft Word”遇到一个错误：发生权限违例。 (-10004)
```

实测结论（在提权沙箱下逐项对照得到）：

| 操作 | 沙箱内 | 提权后 |
|---|---|---|
| `osascript -e 'return 1+1'`（纯计算） | ✅ | ✅ |
| `tell application "Word" to return 1`（只寻址） | ✅ | ✅ |
| `tell application "Finder" to return name of home`（真正发事件） | ❌ -10004 | ✅ |
| `tell application "System Events" to ...` | ❌ -10004 | ✅ |
| 写 `$HOME`（工作区外） | ❌ Operation not permitted | ✅ |
| `ps -e`（进程表） | ❌ 只返回 1 行 | ✅ |

结论：**`osascript` 本身能启动、能寻址目标 App，但「把 AppleEvent 送出去」被沙箱拦掉了。**
拦截发生在送达之前，所以 **macOS 根本不会弹出「自动化」授权框**——
在沙箱里等授权框是等不到的，重试多少次都一样。

在「终端」里跑同样的命令，事件能正常送达，授权框也会正常出现；
点「好」之后该权限被记住，之后就能用了。

### 如果仍然失败

按顺序做这两步：

1. **系统设置 → 隐私与安全性 → 自动化**
   找到运行脚本的程序（终端 / iTerm），勾选其下的「Microsoft Word」。
   若列表里没有这一项，说明系统从未记录过授权请求，先做第 2 步。
2. **系统设置 → 隐私与安全性 → 完全磁盘访问权限**
   把「终端」加进去并打开开关，然后 **⌘Q 完全退出终端再重开**（不是关窗口）。
   原因：Word 自己的沙箱容器 `~/Library/Group Containers/UBF8T346G9.Office`
   若不可写，Word **保存 PDF 时会直接返回 -10004**，这与自动化授权是两码事。
   可用 `touch "$HOME/Library/Group Containers/UBF8T346G9.Office/.w"` 自测。

### AppleScript 写法上的两个坑

`scripts/convert_with_word.sh` 已经处理好，这里记录备查：

1. **`save as` 是 Word 的「顶层命令」，不是 `document` 的方法。**
   而 AppleScript 对 Word 的术语解析经常失败，报

   ```
   “document "x.docx"”不理解“save as”信息。 (-1708)
   ```

   即使语法完全正确也会报。解法是外面套一层 `using terms from application "Microsoft Word"`：

   ```applescript
   using terms from application "Microsoft Word"
       tell application "Microsoft Word"
           set d to open POSIX file "/abs/in.docx"
           save as (item 1 of d) file name "/abs/out.pdf" file format format PDF
           close document 1 saving no
       end tell
   end using terms from
   ```

2. **`open` 返回的是 document 的 list**，直接 `set d to open …` 然后当单个文档用会报
   `变量“d”没有定义 (-2753)`。要么取 `item 1 of`，要么用 `document 1`。

查 Word 支持哪些命令/枚举，直接读它的 sdef：

```bash
grep -o '<command name="[^"]*"' "/Applications/Microsoft Word.app/Contents/Resources/Word.sdef"
awk '/<enumeration name="WdSaveFormat"/,/<\/enumeration>/' "/Applications/Microsoft Word.app/Contents/Resources/Word.sdef"
```

## ③ 从 PDF 出图

```python
import pymupdf                      # pip install pymupdf
d = pymupdf.open('论文_DUT.pdf')
d[0].get_pixmap(dpi=150).save('封面.png')
```

`scripts/make_demo_real.py` 就是这么把 README 的效果图拼出来的。
