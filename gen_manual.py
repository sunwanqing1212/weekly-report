# -*- coding: utf-8 -*-
# 生成《分地市周报报表工具 使用说明书.docx》
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# 全局默认字体：中文微软雅黑
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')


def set_cn(run, size=None, bold=None, color=None):
    run.font.name = 'Calibri'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = RGBColor(*color)


def h(text, level):
    p = doc.add_heading('', level=level)
    r = p.add_run(text)
    set_cn(r, size={0: 20, 1: 15, 2: 12}.get(level, 12), bold=True,
           color=(0x1F, 0x3B, 0x63) if level <= 1 else (0x2E, 0x5A, 0x88))
    return p


def para(text, bold=False, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    r = p.add_run(text)
    set_cn(r, bold=bold)
    return p


def bullets(items):
    for it in items:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(it)
        set_cn(r)


def numbered(items):
    for it in items:
        p = doc.add_paragraph(style='List Number')
        r = p.add_run(it)
        set_cn(r)


def code(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.74)
    r = p.add_run(text)
    r.font.name = 'Consolas'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Consolas')
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    return p


# ================= 封面/标题 =================
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run('分地市周报报表工具\n使用说明书')
set_cn(r, size=22, bold=True, color=(0x1F, 0x3B, 0x63))
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run('适用目录：C:\\Users\\sunwq\\Documents\\20260812    版本：2026-08-13')
set_cn(r, size=9, color=(0x88, 0x88, 0x88))

# ================= 1 工具用途 =================
h('一、工具用途', 1)
para('本工具依据每周的两份广电周报（xlsx 文件），自动完成分地市用户数与用户流量的周对比分析，输出两份成果文件：')
bullets([
    'xlsx 报表：含 Excel 原生组合图表（柱状图：上周/本周；折线图：增幅，副坐标轴）及明细附表（上周值、本周值、增量、增幅、占比，含全省汇总行）；另含"期末总用户数"工作表，列出前后两周最后一天的全省总用户数（单位：万）及增减额。图表引用表内单元格，修改数据后图表自动更新，可直接编辑。',
    'html 报表：与 xlsx 内容一致的网页版图表（ECharts 交互式图表，离线可用，双击浏览器打开），便于快速查看和演示；末尾附表三为"期末总用户数（两周最后一天）"。',
])
para('统计口径：用户数（单位：万）与流量（单位：T）均取各周 7 天的日均值进行对比（增量 = 本周日均 − 上周日均，增幅 = 增量 ÷ 上周日均，占比 = 该地市本周日均 ÷ 全省本周日均）。图表按增幅升序排列，增幅最低的地市在最左侧、最高的在最右侧。期末总用户数取两周各自最后一天（期末）的全省总用户数，增减额 = 本周最后一天 − 上周最后一天。流量附表另含月均DOU列：月均DOU(GB) = 本周日均流量 ÷ 本周日均用户数 × 当月天数（1T 按 1024GB 折算，月天数取本周最后一天所在月份）。')

# ================= 2 每周操作步骤 =================
h('二、每周操作步骤（推荐）', 1)
numbered([
    '将新一周的周报 xlsx 文件放入本目录（文件名保持 gd_YYYY_MM_DD_....xlsx 格式，旧周报无需删除）。',
    '如已打开上次生成的 xlsx 报表，请先关闭（Excel 会锁定文件，导致无法写入）。',
    '双击运行「生成报表.bat」，等待窗口提示"生成完成"。',
    '在目录中打开新生成的文件，文件名形如：分地市周报对比_用户数与流量_2026-08-06_第32周.xlsx（.html），其中日期为最新统计截止日，周数为 ISO 8601 年内周数。每周生成新文件名，不会互相覆盖。',
])
para('工具会自动识别目录中最新的两份周报作为"上周"和"本周"，图表中的周期标签（如"上周(07-31~08-06日均)"）也会根据周报表头日期自动生成，全程无需修改任何代码。', bold=True)

# ================= 3 手动分步执行 =================
h('三、手动分步执行（可选）', 1)
para('与双击 bat 效果完全相同，适用于需要逐步排查问题的场景。在目录下打开命令行，依次执行：')
code('node dump_data.js')
code('tools\\python\\python.exe gen_xlsx.py')
code('node gen_report.js')
para('说明：第一步从两份周报读取数据并计算，结果写入 report_data.json；第二步生成 xlsx 报表；第三步生成 html 报表。')

# ================= 4 目录文件说明 =================
h('四、目录内文件说明', 1)
tbl = doc.add_table(rows=1, cols=2)
tbl.style = 'Light Grid Accent 1'
hdr = tbl.rows[0].cells
for i, txt in enumerate(['文件 / 目录', '作用']):
    p = hdr[i].paragraphs[0]
    r = p.add_run(txt)
    set_cn(r, bold=True)
rows = [
    ('gd_2026_07_31_....xlsx 等', '每周原始周报（数据源），按文件名日期自动识别最新两份'),
    ('生成报表.bat', '一键生成入口，双击运行'),
    ('report_data.js', '数据读取与计算、图表配置（自动识别文件与日期标签）'),
    ('dump_data.js', '将计算结果导出为 report_data.json'),
    ('gen_xlsx.py', '生成含原生图表的 xlsx 报表'),
    ('gen_report.js', '生成 html 报表'),
    ('report_data.json', '中间数据文件，自动生成，无需关心'),
    ('tools\\', '隔离的便携 Python 运行环境（不影响系统），勿删'),
    ('node_modules\\', 'Node 依赖（xlsx、echarts 等），勿删'),
    ('分地市周报对比_..._第N周.xlsx / .html', '生成的成果文件，每周一份'),
]
for name, desc in rows:
    cells = tbl.add_row().cells
    r = cells[0].paragraphs[0].add_run(name)
    set_cn(r, size=9.5)
    r = cells[1].paragraphs[0].add_run(desc)
    set_cn(r, size=9.5)
tbl.columns[0].width = Cm(6.5)
tbl.columns[1].width = Cm(9.5)

# ================= 5 常见问题 =================
h('五、常见问题', 1)
h('1. 运行 bat 提示"无法写入 … xlsx"', 2)
para('生成的 xlsx 正在 Excel 中打开，文件被锁定。关闭 Excel 中该文件后重新运行 bat 即可。')
h('2. 运行 bat 提示"目录中周报文件不足两份"', 2)
para('目录中至少需要两份文件名形如 gd_YYYY_MM_DD_....xlsx 的周报。请检查新周报文件名是否规范（勿带 ~$ 前缀的临时文件，该文件为 Excel 锁定文件，关闭对应 Excel 文档后会自动消失）。')
h('3. 想对比指定的某两周', 2)
para('工具始终取目录中文件名日期最新的两份周报。如需指定对比周，将其余周报文件暂时移到其他文件夹即可。')
h('4. "第 N 周"口径不同', 2)
para('文件名中的周数采用 ISO 8601 周（周一为一周开始）。若公司采用自然周（1 月 1 日起算）等其他口径，需修改 report_data.js 中的 isoWeek 函数，可联系维护人员调整。')
h('5. 想修改图表样式、指标口径', 2)
para('图表配置集中在 report_data.js 的 chartOption 函数（html 图表）和 gen_xlsx.py 的 add_sheet 函数（xlsx 图表）中；统计口径（日均、增幅、占比公式）在 report_data.js 的 buildRows 函数中。修改后重新运行 bat 即可生效。')
h('6. 图表复制到 Word（A4）后文字变形、标签重叠', 2)
para('xlsx 中的图表已按 A4 纵向版心宽度（约 14.8cm）设计，直接粘贴到 Word 无需缩放。粘贴时建议选择"保留源格式"或"粘贴为图片"；如需调整大小，请等比缩放，不要单独拉伸宽度或高度，以免文字变形。')

# ================= 6 环境依赖 =================
h('六、环境依赖', 1)
bullets([
    'Node.js：使用系统已安装的 Node（v24），无需额外配置。',
    'Python：使用本目录 tools\\ 下的便携版 Python 3.12（含 xlsxwriter、python-docx），与系统隔离，不影响其他程序；请勿删除 tools 目录。',
    'ECharts 图表库已内嵌于 html 文件中，html 报表离线可用。',
    'xlsx 报表需要 Microsoft Excel（或 WPS）打开查看和编辑。',
])

doc.save('分地市周报报表工具-使用说明书.docx')
print('DOCX written: 分地市周报报表工具-使用说明书.docx')
