# -*- coding: utf-8 -*-
# 生成含 Excel 原生组合图表（柱状图：上周/本周；折线图：增幅，副坐标轴）的 xlsx
# 图表引用表内单元格，数据可直接编辑，图表随数据自动更新
import json
import math
import xlsxwriter

with open('report_data.json', encoding='utf-8') as f:
    data = json.load(f)

OUT = data['meta']['outBase'] + '.xlsx'
wb = xlsxwriter.Workbook(OUT)

fmt_title = wb.add_format({'bold': True, 'font_size': 12, 'font_name': '微软雅黑'})
fmt_head = wb.add_format({'bold': True, 'bg_color': '#F2F5F9', 'border': 1,
                          'align': 'center', 'font_name': '微软雅黑'})
fmt_city = wb.add_format({'border': 1, 'font_name': '微软雅黑'})
fmt_num = wb.add_format({'border': 1, 'num_format': '#,##0', 'font_name': '微软雅黑'})
fmt_num1 = wb.add_format({'border': 1, 'num_format': '#,##0.0', 'font_name': '微软雅黑'})
fmt_num2 = wb.add_format({'border': 1, 'num_format': '#,##0.00', 'font_name': '微软雅黑'})
fmt_pct_up = wb.add_format({'border': 1, 'num_format': '+0.00"%";-0.00"%"',
                            'font_color': '#C0392B', 'font_name': '微软雅黑'})
fmt_pct_dn = wb.add_format({'border': 1, 'num_format': '+0.00"%";-0.00"%"',
                            'font_color': '#27AE60', 'font_name': '微软雅黑'})
fmt_share = wb.add_format({'border': 1, 'num_format': '0.00"%"', 'font_name': '微软雅黑'})
fmt_total = wb.add_format({'border': 1, 'bold': True, 'num_format': '#,##0.00', 'font_name': '微软雅黑'})
fmt_total_int = wb.add_format({'border': 1, 'bold': True, 'num_format': '#,##0', 'font_name': '微软雅黑'})
fmt_total_1 = wb.add_format({'border': 1, 'bold': True, 'num_format': '#,##0.0', 'font_name': '微软雅黑'})
fmt_delta_up = wb.add_format({'border': 1, 'num_format': '+#,##0.00;-#,##0.00',
                              'font_color': '#C0392B', 'font_name': '微软雅黑'})
fmt_delta_dn = wb.add_format({'border': 1, 'num_format': '+#,##0.00;-#,##0.00',
                              'font_color': '#27AE60', 'font_name': '微软雅黑'})

TABLE_ROW = 19  # 表格标题所在行（0 起始），图表占上方约 18 行


def add_sheet(name, chart_title, rows, unit, total, prev_label, curr_label, extra=None, pct_col=None):
    ws = wb.add_worksheet(name)
    ws.set_column('A:A', 10)
    ws.set_column('B:D', 14)
    ws.set_column('E:F', 10)
    if extra:
        ws.set_column('G:G', 13)

    # ---------- 明细表（图表数据源） ----------
    title_text = f'分地市{name}明细（单位：{unit}；增量/增幅为本周相对上周，占比为该地市占全省比例）'
    if extra:
        title_text += f'；{extra[0]} = 本周日均流量 ÷ 本周日均用户数 × 当月天数（1T=1024GB）'
    ws.write(TABLE_ROW, 0, title_text, fmt_title)
    headers = ['地市', f'上周({unit})', f'本周({unit})', f'增量({unit})', '增幅', '占比(本周)']
    if extra:
        headers.append(extra[0])
    for c, h in enumerate(headers):
        ws.write(TABLE_ROW + 1, c, h, fmt_head)

    first_data = TABLE_ROW + 2          # 首行数据（0 起始）
    # 用户数：2位小数；流量：流量取整数、增量保留1位小数
    if unit == '万':
        fmtv, fmtd, fmt_tv, fmt_td = fmt_num2, fmt_num2, fmt_total, fmt_total
    else:
        fmtv, fmtd, fmt_tv, fmt_td = fmt_num, fmt_num1, fmt_total_int, fmt_total_1
    for i, r in enumerate(rows):
        row = first_data + i
        ws.write(row, 0, r['city'], fmt_city)
        ws.write_number(row, 1, r['prev'], fmtv)
        ws.write_number(row, 2, r['curr'], fmtv)
        ws.write_number(row, 3, r['delta'], fmtd)
        ws.write_number(row, 4, r['pct'], fmt_pct_up if r['delta'] >= 0 else fmt_pct_dn)
        ws.write_number(row, 5, r['share'], fmt_share)
        if extra:
            ws.write_number(row, 6, r[extra[1]], fmt_num2)

    last_data = first_data + len(rows) - 1
    t_delta = round(total['curr'] - total['prev'], 2 if unit == '万' else 1)
    t_pct = round((total['curr'] - total['prev']) / total['prev'] * 100, 2)
    t_prev = total['prev'] if unit == '万' else round(total['prev'])
    t_curr = total['curr'] if unit == '万' else round(total['curr'])
    trow = last_data + 1
    ws.write(trow, 0, '全省', fmt_head)
    ws.write_number(trow, 1, t_prev, fmt_tv)
    ws.write_number(trow, 2, t_curr, fmt_tv)
    ws.write_number(trow, 3, t_delta, fmt_td)
    ws.write_number(trow, 4, t_pct, fmt_total)
    ws.write_number(trow, 5, 100, fmt_total)
    if extra:
        ws.write_number(trow, 6, total[extra[1]], fmt_total)

    # ---------- 原生组合图表 ----------
    # 图表宽度按 A4 纵向版心（约 14.8cm）设计，复制到 Word 后无需缩放，
    # 文字不被压缩变形；上周柱不显示数据标签，避免标签拥挤重叠
    lab_fmt = '#,##0.0' if unit == '万' else '#,##0'  # 数据标签：用户数保留1位小数
    col = wb.add_chart({'type': 'column'})
    col.add_series({
        'name': prev_label,
        'categories': [name, first_data, 0, last_data, 0],
        'values': [name, first_data, 1, last_data, 1],
        'fill': {'color': '#9DB8D2'},
        'gap': 60, 'overlap': -10,
    })
    col.add_series({
        'name': curr_label,
        'categories': [name, first_data, 0, last_data, 0],
        'values': [name, first_data, 2, last_data, 2],
        'fill': {'color': '#4E79A7'},
        'data_labels': {'value': True, 'num_format': lab_fmt, 'font': {'size': 9, 'name': '微软雅黑'}},
    })

    line = wb.add_chart({'type': 'line'})
    line.add_series({
        'name': '增幅',
        'categories': [name, first_data, 0, last_data, 0],
        # 数值引用隐藏工作表中的 ROUND(...,1) 公式列：逐点标签不继承系列 num_format 时
        # 也能保证只显示 1 位小数，且随明细表编辑自动更新
        'values': f"='基准线数据'!${pct_col}$1:${pct_col}${len(rows)}",
        'y2_axis': True,
        'line': {'color': '#E15759', 'width': 1.75},
        'marker': {'type': 'circle', 'size': 6, 'border': {'color': '#E15759'}, 'fill': {'color': 'white'}},
        'data_labels': {'value': True, 'num_format': '0.0"%";-0.0"%"', 'position': 'above',
                        'font': {'size': 9, 'color': '#C0392B', 'name': '微软雅黑'},
                        # 逐点着色：与表格一致，正增幅红色、负增幅绿色；
                        # 每个点都显式指定 position，否则部分查看器不继承系列级定位，标签会落回折线上
                        'custom': [{'position': 'above',
                                    'font': {'size': 9, 'name': '微软雅黑',
                                             'color': '#C0392B' if r['pct'] >= 0 else '#27AE60'}}
                                   for r in rows]},
    })
    # 增幅 0 基准线：Excel 无原生基准线，用隐藏工作表中的常数 0 虚线折线实现
    # （不用隐藏列：隐藏列依赖 plotVisOnly 行为，Excel 兼容性无保障）
    line.add_series({
        'name': '0基准线',
        'categories': [name, first_data, 0, last_data, 0],
        'values': f"='基准线数据'!$A$1:$A${len(rows)}",
        'y2_axis': True,
        'line': {'color': '#999999', 'width': 1, 'dash_type': 'dash'},
        'marker': {'type': 'none'},
    })
    # 两轴留白：柱形压低至约 70% 高度，增幅折线整体抬到上部空白区，
    # 使折线数据标签位于折线上方且不遮挡柱形与柱标签
    max_bar = max(max(r['prev'], r['curr']) for r in rows)
    y1_max = math.ceil(max_bar * 1.4)
    pcts = [r['pct'] for r in rows]
    p_min, p_max = min(pcts), max(pcts)
    span = max(p_max - p_min, 1)
    y2_min = round(p_min - span * 2.7, 1)
    y2_max = round(p_max + span * 0.65, 1)
    # 组合图中折线系列实际使用折线图自身的副轴，col 与 line 需设置相同副轴参数。
    # 注意：副轴不能用 visible=False（轴被删除后 Excel 会让折线回退到主轴缩放，
    # 折线会压平在横轴上），改为保留轴、只隐藏刻度文字与轴线
    y2_opts = {'min': y2_min, 'max': y2_max,
               'name': '增幅(%)', 'name_font': {'size': 10, 'name': '微软雅黑'},
               'label_position': 'none', 'line': {'none': True},
               'major_gridlines': {'visible': False}}
    line.set_y2_axis(y2_opts)
    col.combine(line)

    col.set_title({'name': chart_title, 'name_font': {'size': 12, 'bold': True, 'name': '微软雅黑'}})
    col.set_x_axis({'num_font': {'size': 10, 'rotation': -45, 'name': '微软雅黑'}})
    col.set_y_axis({'name': unit, 'num_format': '#,##0', 'max': y1_max,
                    'num_font': {'size': 10, 'name': '微软雅黑'},
                    'major_gridlines': {'visible': False}})
    col.set_y2_axis(y2_opts)
    # delete_series：图例中隐藏“0基准线”条目（系列顺序：上周0、本周1、增幅2、基准线3）
    col.set_legend({'position': 'bottom', 'font': {'size': 10, 'name': '微软雅黑'},
                    'delete_series': [3]})
    col.set_size({'width': 560, 'height': 350})
    col.set_style(10)

    ws.insert_chart(0, 0, col)


add_sheet('周日均用户数', '广电分地市用户数（周日均）及周增幅（按增幅升序）',
          data['userRows'], '万',
          {'prev': data['totals']['prevU'], 'curr': data['totals']['currU']},
          data['meta']['prevLabel'], data['meta']['currLabel'],
          pct_col='B')

add_sheet('日均流量', '广电分地市流量（日均）及周增幅（按增幅升序）',
          data['trafRows'], 'T',
          {'prev': data['totals']['prevT'], 'curr': data['totals']['currT'],
           'dou': data['totals']['douT']},
          data['meta']['prevLabel'], data['meta']['currLabel'],
          extra=('月均DOU(GB)', 'dou'), pct_col='C')

# ---------- 隐藏工作表：增幅 0 基准线（A列）与增幅1位小数数据源（B/C列，公式引用明细表E列） ----------
zws = wb.add_worksheet('基准线数据')
n_city = len(data['trafRows'])
for i in range(n_city):
    zws.write_number(i, 0, 0)
    for col_idx, sheet_name in ((1, '周日均用户数'), (2, '日均流量')):
        excel_row = TABLE_ROW + 3 + i  # 明细表首行数据的 Excel 行号（TABLE_ROW+2 起，0 起始转 1 起始）
        zws.write_formula(i, col_idx, f"=ROUND('{sheet_name}'!E{excel_row},1)")
zws.hide()

# ---------- 期末总用户数：前后两周最后一天全省总用户数（万）及增减额 ----------
ld = data['lastDayUsers']
ws = wb.add_worksheet('期末总用户数')
ws.set_column('A:A', 16)
ws.set_column('B:B', 12)
ws.set_column('C:C', 14)
ws.write(0, 0, '前后两周最后一天总用户数（单位：万）', fmt_title)
for c, h in enumerate(['项目', '日期', '总用户数(万)']):
    ws.write(1, c, h, fmt_head)
ws.write(2, 0, '上周最后一天', fmt_city)
ws.write(2, 1, ld['prevDate'], fmt_city)
ws.write_number(2, 2, ld['prev'], fmt_num2)
ws.write(3, 0, '本周最后一天', fmt_city)
ws.write(3, 1, ld['currDate'], fmt_city)
ws.write_number(3, 2, ld['curr'], fmt_num2)
ws.write(4, 0, '增减额', fmt_head)
ws.write(4, 1, '—', fmt_city)
ws.write_number(4, 2, ld['delta'], fmt_delta_up if ld['delta'] >= 0 else fmt_delta_dn)

try:
    wb.close()
except Exception:
    raise SystemExit(f'错误：无法写入 {OUT}。文件很可能正在 Excel 中打开，请先关闭该文件后重新运行。')
print('XLSX written:', OUT)
