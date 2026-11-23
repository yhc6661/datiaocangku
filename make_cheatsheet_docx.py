"""生成《FSTSP 云服务器常用命令速查》DOCX，再导出 PDF。"""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

INK = RGBColor(0x1F, 0x29, 0x37)
GRAY = RGBColor(0x6B, 0x72, 0x80)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
CODE_FONT = "Consolas"
UI_FONT = "微软雅黑"


# --------------------------------------------------------------------------- #
# 基础工具
# --------------------------------------------------------------------------- #
def set_font(run, name=UI_FONT, size=None, bold=None, color=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def shade(cell_or_par, fill):
    el = cell_or_par._tc if hasattr(cell_or_par, "_tc") else cell_or_par._p
    pr = el.get_or_add_tcPr() if hasattr(el, "get_or_add_tcPr") else el.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pr.append(shd)


def cell_borders(cell, color="E2E8F0", size=4):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(size))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)
        borders.append(e)
    tcpr.append(borders)


def heading(doc, text, number=None, brk=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3.5)
    p.paragraph_format.keep_with_next = True
    if brk:
        p.paragraph_format.page_break_before = True
    if number:
        r = p.add_run(f"{number}  ")
        set_font(r, size=13, bold=True, color=BLUE)
    r = p.add_run(text)
    set_font(r, size=13, bold=True, color=INK)
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "3")
    bottom.set(qn("w:color"), "DBEAFE")
    pbdr.append(bottom)
    p._p.get_or_add_pPr().append(pbdr)
    return p


def set_cell_margins(t, top=30, bottom=30, left=85, right=85):
    tblpr = t._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for edge, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblpr.append(mar)


def keep_together(t):
    """禁止表格跨页：行内不拆行 + 除最后一行外都 keep_with_next。"""
    for row in t.rows:
        trpr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trpr.append(cant)
    for row in t.rows[:-1]:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True


def table(doc, rows, widths=(3.9, 12.1), headers=("用途", "命令"), mono=True):
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    header = t.rows[0].cells
    for i, title in enumerate(headers):
        cell = header[i]
        shade(cell, "1D4ED8")
        cell_borders(cell, "1D4ED8")
        cell.width = Cm(widths[i])
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title)
        set_font(r, size=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for idx, (label, cmd) in enumerate(rows):
        cells = t.add_row().cells
        fill = "FFFFFF" if idx % 2 == 0 else "F8FAFC"
        for i in range(2):
            cells[i].width = Cm(widths[i])
            shade(cells[i], fill)
            cell_borders(cells[i])
            cells[i].paragraphs[0].paragraph_format.space_before = Pt(2)
            cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)
        r = cells[0].paragraphs[0].add_run(label)
        set_font(r, size=9, color=INK)
        r = cells[1].paragraphs[0].add_run(cmd)
        if mono:
            set_font(r, name=CODE_FONT, size=8.0, color=RGBColor(0x0F, 0x17, 0x2A))
        else:
            set_font(r, size=8.8, color=RGBColor(0x1F, 0x29, 0x37))
    set_cell_margins(t)
    keep_together(t)
    return t


def subhead(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, size=10.5, bold=True, color=INK)
    return p


def note(doc, text, kind="info"):
    fill = "EFF6FF" if kind == "info" else "FEF3C7"
    bar = "1D4ED8" if kind == "info" else "D97706"
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.rows[0].cells[0]
    cell.width = Cm(16)
    shade(cell, fill)
    tcpr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge, sz, col in (("top", 0, fill), ("bottom", 0, fill), ("right", 0, fill), ("left", 18, bar)):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single" if sz else "nil")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), col)
        borders.append(e)
    tcpr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    set_font(r, size=8.5, color=RGBColor(0x33, 0x41, 0x55))
    set_cell_margins(t)
    keep_together(t)
    return t


def code_box(doc, lines):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.rows[0].cells[0]
    cell.width = Cm(16)
    shade(cell, "F1F5F9")
    cell_borders(cell, "E2E8F0")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    for i, line in enumerate(lines):
        r = p.add_run(line)
        set_font(r, name=CODE_FONT, size=8.5, color=RGBColor(0x0F, 0x17, 0x2A))
        if i < len(lines) - 1:
            r.add_break()
    set_cell_margins(t)
    keep_together(t)
    return t


def footer(doc):
    section = doc.sections[0]
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("FSTSP 云服务器速查  ·  第 ")
    set_font(r, size=8, color=GRAY)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "16")
    rpr.append(sz)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = "1"
    run.append(t)
    fld.append(run)
    p._p.append(fld)
    r = p.add_run(" 页")
    set_font(r, size=8, color=GRAY)


# --------------------------------------------------------------------------- #
# 文档
# --------------------------------------------------------------------------- #
def build(path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(1.4)
    sec.left_margin = sec.right_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = UI_FONT
    normal.font.size = Pt(9.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), UI_FONT)

    # 标题
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("FSTSP 云服务器")
    set_font(r, size=21, bold=True, color=INK)
    r = p.add_run("  常用命令速查")
    set_font(r, size=21, bold=True, color=BLUE)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(
        "腾讯云轻量 4 核 4G · Ubuntu 22.04 · 项目目录 /opt/fstsp · 更新于 2026-11-23"
    )
    set_font(r, size=9, color=GRAY)

    # 0. 三条基本认知
    heading(doc, "先分清「在哪台机器上」", "0")
    note(
        doc,
        "提示符 PS C:\\...> = 你的 Windows；fstsp@VM-0-14-ubuntu:~$ = 云服务器。"
        "同一句话意思完全不同，敲命令前先看提示符。",
        "info",
    )
    table(
        doc,
        [
            ("本机路径", "C:\\Users\\yanghaichen\\Documents\\deepseek-harness\\default-workspace\\fstsp-server"),
            ("宿主机路径", "/opt/fstsp/work"),
            ("容器内路径", "/app/work      ← 与上一行是同一个目录"),
            ("容器里的 Python", "3.11 + OR-Tools + numpy + pandas"),
            ("宿主机的 Python", "3.10，没装任何第三方库"),
        ],
        widths=(4.0, 12.0),
        headers=("项目", "路径 / 说明"),
    )

    # 1. 连接
    heading(doc, "连接与退出", "1")
    table(
        doc,
        [
            ("连服务器", "ssh fstsp@106.52.92.189"),
            ("连上执行一条就退出", 'ssh fstsp@106.52.92.189 "hostname"'),
            ("在服务器里退出", "exit"),
            ("看自己在哪台机器", "hostname          # 输出 VM-0-14-ubuntu 即服务器"),
        ],
    )
    note(doc, "不需要密码：本机 ~/.ssh/id_ed25519 已写入服务器，ssh / scp 通用。", "info")

    # 2. 传文件
    heading(doc, "上传与下载文件（在本机 PowerShell 执行）", "2")
    table(
        doc,
        [
            ("上传一个文件", "scp .\\test.py fstsp@106.52.92.189:/opt/fstsp/work/"),
            ("上传多个文件", "scp .\\a.py .\\b.py fstsp@106.52.92.189:/opt/fstsp/work/"),
            ("上传整个目录", "scp -r .\\data fstsp@106.52.92.189:/opt/fstsp/work/"),
            ("取回结果", "scp -r fstsp@106.52.92.189:/opt/fstsp/results .\\cloud-results"),
        ],
    )
    note(doc, "目标路径结尾的 / 表示「放进这个目录」；漏掉可能被当成重命名。scp 会自己连服务器，不用先 ssh。", "warn")

    # 3. 目录与文件
    heading(doc, "看目录、看文件", "3")
    table(
        doc,
        [
            ("我在哪", "pwd"),
            ("当前目录有什么", "ls -lh"),
            ("含隐藏文件 / 按时间排", "ls -la        ls -lht"),
            ("树状看两层结构", "tree -L 2 /opt/fstsp"),
            ("进目录 / 回家 / 上一层", "cd /opt/fstsp     cd ~     cd .."),
            ("看文件内容", "cat instances/demo_01.json"),
            ("实时看日志", "tail -f logs/demo_01.log     # Ctrl+C 停止"),
            ("编辑文件", "nano work/test.py            # Ctrl+O 存，Ctrl+X 退"),
        ],
    )

    # 4. 增删改
    heading(doc, "新建、复制、移动、删除", "4")
    table(
        doc,
        [
            ("建目录（含父目录）", "mkdir -p work/exp01"),
            ("建空文件", "touch work/a.py"),
            ("复制文件 / 目录", "cp a.py b.py        cp -r dir1 dir2"),
            ("移动或改名", "mv work/a.py work/run_01.py"),
            ("删文件", "rm work/test.py"),
            ("删目录", "rm -r work/exp01"),
            ("删前逐个确认", "rm -i work/*.py"),
        ],
    )
    note(doc, "没有回收站。通配符先 ls 看命中什么再删；rm 后面多一个空格可能删掉系统目录。", "warn")

    # 5. 跑 Python
    heading(doc, "运行 Python 程序", "5")
    table(
        doc,
        [
            ("容器里跑（推荐）", "cd /opt/fstsp && docker compose exec -T solver python /app/work/test.py"),
            ("指定工作目录跑", "docker compose exec -T -w /app/work solver python test.py"),
            ("跑项目求解器", "docker compose exec -T solver python -m fstsp.solve -i /app/instances/demo_01.json -t 60"),
            ("宿主机跑（仅标准库）", "python3 /opt/fstsp/work/test.py"),
            ("环境自检", "docker compose exec -T solver python /app/work/demo_env.py"),
        ],
    )
    note(doc, "改过程序要重建镜像才生效：docker compose build && docker compose up -d", "info")

    # 6. 后台跑
    heading(doc, "后台运行（断线不停）", "6")
    table(
        doc,
        [
            ("项目封装好的提交方式", "cd /opt/fstsp && bash scripts/run_job.sh demo_01 120"),
            ("批量跑全部算例", "bash scripts/run_batch.sh 120 3      # 每个 120 秒，3 路并行"),
            ("新建会话", "tmux new -s run"),
            ("离开但任务继续", "Ctrl+B 松开后按 D"),
            ("回到会话", "tmux attach -t run"),
            ("看有哪些会话 / 结束", "tmux ls          exit"),
        ],
    )

    # 7. Docker
    heading(doc, "Docker 常用四条", "7")
    table(
        doc,
        [
            ("容器在跑吗", "docker ps"),
            ("有哪些镜像", "docker images"),
            ("改代码后重建", "docker compose build && docker compose up -d"),
            ("看容器日志", "docker compose logs --tail 50"),
        ],
    )

    # 8. 资源
    heading(doc, "查看服务器资源", "8")
    table(
        doc,
        [
            ("CPU / 内存实时", "htop            # q 退出"),
            ("磁盘剩余", "df -h"),
            ("内存", "free -h"),
            ("各目录占用", "du -sh /opt/fstsp/*"),
        ],
    )

    # 9. 坑
    heading(doc, "新手最常踩的六个坑", "9")
    table(
        doc,
        [
            ("漏空格：cd/opt/fstsp", "写成 cd /opt/fstsp，命令名和参数之间必须有空格"),
            ("通配符直接删", "先 ls *.py 看清命中什么，再 rm"),
            ("窗口被鼠标点中卡住", "控制台进入选中状态，按 Esc 恢复"),
            ("在服务器里敲 scp", "先 exit 回本机再传文件"),
            ("PowerShell 抄 rm -rf", "本机用 Remove-Item -Recurse，参数不通用"),
            ("容器里找 /opt/fstsp", "容器内路径是 /app，两个文件系统是分开的"),
        ],
        headers=("现象", "正确做法"),
        mono=False,
    )

    # 10. 八条
    heading(doc, "一天里真正会敲的 8 条", "10")
    code_box(
        doc,
        [
            "# —— 本机 PowerShell ——",
            "ssh fstsp@106.52.92.189",
            "scp .\\test.py fstsp@106.52.92.189:/opt/fstsp/work/",
            "scp -r fstsp@106.52.92.189:/opt/fstsp/results .\\cloud-results",
            "",
            "# —— 云服务器 bash ——",
            "cd /opt/fstsp",
            "ls -lh work/",
            "docker compose exec -T solver python /app/work/test.py",
            "bash scripts/run_job.sh demo_01 120",
            "tail -f logs/*.log",
        ],
    )

    subhead(doc, "路径对照（同一个文件的两种叫法）")
    table(
        doc,
        [
            ("宿主机", "/opt/fstsp/work"),
            ("容器内", "/app/work        ← 与上一行是同一个目录"),
            ("宿主机", "/opt/fstsp/instances        容器内 /app/instances"),
            ("宿主机", "/opt/fstsp/results          容器内 /app/results"),
        ],
        headers=("视角", "路径"),
        mono=True,
    )

    subhead(doc, "三条红线")
    note(doc, "① rm 后面不要多敲空格：rm -rf /opt/fstsp/work / 会连根目录一起删。", "warn")
    note(doc, "② 没有回收站：重要结果先 scp 回本地，再考虑删。", "warn")
    note(doc, "③ 密码、私钥不要贴进聊天窗口或 Git 仓库。", "warn")

    subhead(doc, "出问题时按这个顺序查")
    table(
        doc,
        [
            ("命令报错看不懂", "docker compose logs --tail 50     或   tail -50 logs/xxx.log"),
            ("容器没在跑", "docker compose up -d"),
            ("连不上服务器", "等 30 秒重试；再看腾讯云控制台防火墙的 22 端口"),
            ("忘了自己在哪台机器", "hostname     # 输出 VM-0-14-ubuntu 即服务器"),
        ],
        headers=("现象", "先看这里"),
        mono=False,
    )

    footer(doc)
    doc.save(path)
    return path


if __name__ == "__main__":
    build("docs/server-cheatsheet.docx")
    print("saved docs/server-cheatsheet.docx")
