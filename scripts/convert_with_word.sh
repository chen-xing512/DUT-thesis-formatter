#!/bin/bash
# ============================================================================
#  convert_with_word.sh — 用本机 Microsoft Word 把 .docx 转成 PDF
#
#  为什么需要这个脚本（在 DSH 沙箱外运行）：
#    Agent 会话跑在沙箱里时，向 Finder / Word / System Events 发送 Apple Event
#    会被拦掉，osascript 返回 -10004「权限违例」，而且在沙箱里**不会弹出**
#    macOS 的自动化授权对话框。审计日志被 SIP 保护、也读不到，无法在会话内自证。
#    在终端里直接跑这个脚本，AppleEvent 能正常送达，授权框也会正常出现。
#
#  用法：
#     chmod +x convert_with_word.sh
#     ./convert_with_word.sh 论文.docx              # 输出到同目录
#     ./convert_with_word.sh 论文.docx 出.pdf        # 指定输出
# ============================================================================
set -uo pipefail

die() { printf '\033[31m错误：%s\033[0m\n' "$1" >&2; exit 1; }
ok()  { printf '\033[32m%s\033[0m\n' "$1"; }

[ $# -ge 1 ] || die "用法: $0 <输入.docx> [输出.pdf]"

IN="$1"
OUT="${2:-${IN%.*}.pdf}"

if [ ! -f "$IN" ]; then
  # 相对路径按当前目录补全，避免 Word 因为拿不到绝对路径而报错
  [ -f "$PWD/$IN" ] && IN="$PWD/$IN" || die "找不到输入文件：$IN"
fi
IN="$(cd "$(dirname "$IN")" && pwd)/$(basename "$IN")"
case "$OUT" in
  /*) : ;;
  *) OUT="$(cd "$(dirname "$OUT")" 2>/dev/null && pwd || pwd)/$(basename "$OUT")" ;;
esac
[ -f "$IN" ] || die "找不到输入文件：$IN"

echo "输入：$IN"
echo "输出：$OUT"
echo

# --- 1) Word 是否安装 -------------------------------------------------------
[ -d "/Applications/Microsoft Word.app" ] || die "未安装 Microsoft Word"

# --- 2) 预热：先发一个最小 AppleEvent，触发/暴露授权状态 ---------------------
echo "[1/3] 检查自动化授权（这一步可能弹出系统授权框，请点「好」）..."
if ! osascript -e 'tell application "Microsoft Word" to return version' >/dev/null 2>&1; then
  cat >&2 <<'EOF'
无法向 Microsoft Word 发送 Apple Event（通常是 -10004 权限违例）。

请按顺序排查：
  a) 打开「系统设置 → 隐私与安全性 → 自动化」，
     找到运行本脚本的程序（Terminal / iTerm / 启动脚本的宿主 App），
     勾选其下的「Microsoft Word」。若列表里没有该项，说明系统从未记录过授权请求，
     先执行下面 b)。
  b) 打开「系统设置 → 隐私与安全性 → 完全磁盘访问权限」，
     把「终端」(Terminal.app 或 iTerm) 加进去并打开开关，然后**完全退出并重开终端**。
     Word 自身的沙箱容器（~/Library/Group Containers/UBF8T346G9.Office）不可写时，
     Word 保存 PDF 也会直接返回 -10004。
  c) 如果仍然失败，在本脚本目录下执行下面这条最小测试命令，把输出贴出来：
       osascript -e 'tell application "Microsoft Word" to return version'
EOF
  exit 2
fi
ok "  AppleEvent 可达 Word"

# --- 3) 转换 ---------------------------------------------------------------
# 三个坑（都实测踩过）：
#   ① ``set theDoc to open inPath`` 在本机 Word 16 上**不返回文档对象**，
#      于是下一句报 -2753「变量 theDoc 没有定义」；必须用
#      ``open inPath`` 后再取 ``active document``。
#   ② 目标文件若已在 Word 里打开，``active document`` 可能是**旧的**在内存副本：
#      Word 按路径缓存文档，``open`` 一个已打开的文件不会重新读盘，于是你会
#      导出上一次的旧内容，而且**毫无提示**（实测踩过：改了 docx 再导出，
#      PDF 里的封面还是改之前的样子）。所以下面先按文件名把同名文档关掉。
#   ③ **导出格式必须写枚举名 ``format PDF``，不要写数字常量。**
#      曾经写成 15（以为是 wdFormatPDF），实际 15 = wdFormatXMLTemplateMacroEnabled，
#      于是生成的文件扩展名是 .pdf、内容却是 Word XML 模板；Word 不报错、
#      脚本不报错，只有 ``file(1)`` 才露馅。wdFormatPDF 的数字其实是 17。
#      所以下面第 [3/3] 步会**校验产物真的是 PDF**，不是就报错退出。
echo "[2/3] 让 Word 打开文档并另存为 PDF（Word 会短暂出现在前台）..."
BASE="$(basename "$IN")"
RC=0
osascript <<OSA
set inPath to POSIX file "$IN"
set outPath to "$OUT"
set fname to "$BASE"
tell application "Microsoft Word"
	activate
	-- 同名文档若已打开，先关掉（不保存），否则会导出内存里的旧内容
	repeat with d in documents
		try
			if name of d is fname then close d saving no
		end try
	end repeat
	open inPath
	set theDoc to active document
	save as theDoc file name outPath file format format PDF
	close theDoc saving no
end tell
return "saved"
OSA
RC=$?

echo "[3/3] 校验产物..."
if [ "${RC:-1}" -ne 0 ]; then
  die "Word 报错（退出码 ${RC:-?}）。常见原因：输出路径无写权限、文档被占用、或上面的授权未生效。"
fi
if [ ! -s "$OUT" ]; then
  die "Word 未报错，但没有生成 $OUT。请确认输出目录可写。"
fi
if command -v file >/dev/null; then
  file "$OUT"
  if ! file -b "$OUT" | grep -qi 'PDF'; then
    die "导出的不是 PDF（file 显示：$(file -b "$OUT")）。
    请检查 AppleScript 里的 save as 参数——必须写枚举名 \`format PDF\`；
    写数字常量（如 15）会得到一个「扩展名是 .pdf 的 Word 模板」。"
  fi
fi
ok "完成：$OUT  ($(wc -c < "$OUT" | tr -d ' ') 字节)"
