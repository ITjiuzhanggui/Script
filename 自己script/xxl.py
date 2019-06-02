# coding: utf-8
import xlsxwriter

# 创建EXcel文件
workbook = xlsxwriter.Workbook('demo.xlsx')

# 创建一个工作表对象
worksheet = workbook.add_worksheet('2019_WW12')
worksheet2 = workbook.add_worksheet('2019_WW13')


def get_merge_range(*args, **kwargs):
    def write_(*args, **kwargs):
        yield worksheet.merge_range(*args, **kwargs)

    [item for item in write_(*args, **kwargs)]
    # for i in write(*args,**kwargs):
    #     pass


# worksheet.set_row
def get_sheet_set_row(*args, **kwargs):
    def sheet_(*args, **kwargs):
        yield worksheet.set_row(*args, **kwargs)

    [item for item in sheet_(*args, **kwargs)]


cell_format = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
})

cell_2_format = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
})

cell_3_format = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'gray',
})

cell_4_format = workbook.add_format({
    'fg_color': 'green',
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 25,
    'border': 1,
})

cell_5_format = workbook.add_format({
    'bg_color':'green',
    'pattern': True,
    'align': 'center',
    'valign': 'vcenter',
})

# 设置宽度
worksheet.set_column('A:A', 18)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:C', 10)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 25)
worksheet.set_column('G:G', 20)
worksheet.set_column('H:H', 15)
worksheet.set_column('I:I', 15)
worksheet.set_column('J:J', 25)
worksheet.set_column('K:K', 10)
worksheet.set_column('L:L', 20)
worksheet.set_column('M:M', 15)
worksheet.set_column('N:N', 15)

# 设置行高
# get_sheet_set_row('A:A', 35)
# get_sheet_set_row('B:B', 35)
# get_sheet_set_row('C:C', 35)
# get_sheet_set_row('D:D', 35)

# 合并单元格---P1
get_merge_range('A1:A5', 'OLINUX-5025', cell_format)
get_merge_range('B1:B5', 'ST_PERF_Single_80_geometry_APL_I', cell_format)

get_merge_range('A7:A11', 'OLINUX-5026', cell_format)
get_merge_range('B7:B11', 'ST_PERF_Single_30_geometry_APL_I', cell_format)

get_merge_range('A13:A17', 'OLINUX-5028', cell_format)
get_merge_range('B13:B17', 'ST_PERF_Multi_80_geometry_APL_I', cell_format)

get_merge_range('A19:A23', 'OLINUX-5031', cell_format)
get_merge_range('B19:B23', 'ST_PERF_Tex_80_APL_I', cell_format)

get_merge_range('A25:A29', 'OLINUX-5032', cell_format)
get_merge_range('B25:B29', 'ST_PERF_Long_80_APL_I', cell_format)

get_merge_range('A31:A35', 'OLINUX-5046', cell_format)
get_merge_range('B31:B35', 'ST_PERF_Single_80_shader_APL_I', cell_format)

get_merge_range('A37:A41', 'OLINUX-5047', cell_format)
get_merge_range('B37:B41', 'ST_PERF_Multi_80_shader_APL_I', cell_format)

get_merge_range('A43:A47', 'OLINUX-5048', cell_format)
get_merge_range('B43:B47', 'ST_PERF_Tex_flat_APL_I', cell_format)

# 合并单元格---P2
get_merge_range('A49:A53', 'OLINUX-5036', cell_format)
get_merge_range('B49:B53', 'ST_PERF_Long_80_APL_I_fps', cell_format)

get_merge_range('A55:A59', 'OLINUX-5037', cell_format)
get_merge_range('B55:B59', 'ST_PERF_Multi_30_geometry_APL_I_fps', cell_format)

get_merge_range('A61:A65', 'OLINUX-5038', cell_format)
get_merge_range('B61:B65', 'ST_PERF_Multi_30_shader_APL_I_fps', cell_format)

get_merge_range('A67:A71', 'OLINUX-5039', cell_format)
get_merge_range('B67:B71', 'ST_PERF_Multi_80_geometry_APL_I_fps', cell_format)

get_merge_range('A73:A77', 'OLINUX-5040', cell_format)
get_merge_range('B73:B77', 'ST_PERF_Multi_80_shader_APL_I_fps', cell_format)

get_merge_range('A79:A83', 'OLINUX-5041', cell_format)
get_merge_range('B79:B83', 'ST_PERF_Single_30_geometry_APL_I_fps', cell_format)

get_merge_range('A85:A89', 'OLINUX-5042', cell_format)
get_merge_range('B85:B89', 'ST_PERF_Single_30_shader_APL_I_fps', cell_format)

get_merge_range('A91:A95', 'OLINUX-5043', cell_format)
get_merge_range('B91:B95', 'ST_PERF_Single_80_geometry_APL_I_fps', cell_format)

get_merge_range('A97:A101', 'OLINUX-5044', cell_format)
get_merge_range('B97:B101', 'ST_PERF_Single_80_shader_APL_I_fps', cell_format)

get_merge_range('A103:A107', 'OLINUX-5045', cell_format)
get_merge_range('B103:B107', 'ST_PERF_Tex_80_APL_I_fps', cell_format)

# worksheet.write_string('C1', 'P1/P1', cell_format)
list_P1_C = []
for i in range(1, 6):
    list_P1_C.append("C%s" % i)
    list_P1_C.append("C%d" % (i + 6))
    list_P1_C.append("C%d" % (i + 12))
    list_P1_C.append("C%d" % (i + 18))
    list_P1_C.append("C%d" % (i + 24))
    list_P1_C.append("C%d" % (i + 30))
    list_P1_C.append("C%d" % (i + 36))
    list_P1_C.append("C%d" % (i + 42))

for item in list_P1_C:
    worksheet.write_string(item, 'P1/P1', cell_format)

list_P2_C = []
for j in range(49, 54):
    list_P2_C.append("C%s" % j)
    list_P2_C.append("C%s" % (j + 6))
    list_P2_C.append("C%s" % (j + 12))
    list_P2_C.append("C%s" % (j + 18))
    list_P2_C.append("C%s" % (j + 24))
    list_P2_C.append("C%s" % (j + 30))
    list_P2_C.append("C%s" % (j + 36))
    list_P2_C.append("C%s" % (j + 42))
    list_P2_C.append("C%s" % (j + 48))
    list_P2_C.append("C%s" % (j + 54))

for item in list_P2_C:
    worksheet.write_string(item, 'P2/P2', cell_format)

list_P1_F = []
for i in range(1, 6):
    list_P1_F.append("F%s" % i)
    list_P1_F.append("F%d" % (i + 6))
    list_P1_F.append("F%d" % (i + 12))
    list_P1_F.append("F%d" % (i + 18))
    list_P1_F.append("F%d" % (i + 24))
    list_P1_F.append("F%d" % (i + 30))
    list_P1_F.append("F%d" % (i + 36))
    list_P1_F.append("F%d" % (i + 42))

for item in list_P1_F:
    worksheet.write_string(item, 'P1/P1', cell_format)
    
list_P2_F = []
for j in range(49, 54):
    list_P2_F.append("F%s" % j)
    list_P2_F.append("F%s" % (j + 6))
    list_P2_F.append("F%s" % (j + 12))
    list_P2_F.append("F%s" % (j + 18))
    list_P2_F.append("F%s" % (j + 24))
    list_P2_F.append("F%s" % (j + 30))
    list_P2_F.append("F%s" % (j + 36))
    list_P2_F.append("F%s" % (j + 42))
    list_P2_F.append("F%s" % (j + 48))
    list_P2_F.append("F%s" % (j + 54))

for item in list_P2_F:
    worksheet.write_string(item, 'P2/P2', cell_format)
# 单元格分割行---P1
get_merge_range('A6:N6', None, cell_3_format)
get_sheet_set_row(5, 5, cell_3_format)  # A6行高5像素

get_merge_range('A12:N12', None, cell_3_format)
get_sheet_set_row(11, 5, cell_3_format)  # A12行高5像素

get_merge_range('A18:N18', None, cell_3_format)
get_sheet_set_row(17, 5, cell_3_format)  # A18行高5像素

get_merge_range('A24:N24', None, cell_3_format)
get_sheet_set_row(23, 5, cell_3_format)  # A24行高5像素

get_merge_range('A30:N30', None, cell_3_format)
get_sheet_set_row(29, 5, cell_3_format)  # A30行高5像素

get_merge_range('A36:N36', None, cell_3_format)
get_sheet_set_row(35, 5, cell_3_format)  # A36行高5像素

get_merge_range('A42:N42', None, cell_3_format)
get_sheet_set_row(41, 5, cell_3_format)  # A42行高5像素

get_merge_range('A48:N48', 'P2', cell_4_format)
get_sheet_set_row(47, 20)  # A48行高为35像素

# 单元格分割行---P2V
get_merge_range('A54:N54', None, cell_3_format)
get_sheet_set_row(53, 5, cell_3_format)  # A54行高5像素

get_merge_range('A60:N60', None, cell_3_format)
get_sheet_set_row(59, 5, cell_3_format)  # A60行高5像素

get_merge_range('A66:N66', None, cell_3_format)
get_sheet_set_row(65, 5, cell_3_format)  # A66行高5像素

get_merge_range('A72:N72', None, cell_3_format)
get_sheet_set_row(71, 5, cell_3_format)  # A72行高5像素

get_merge_range('A78:N78', None, cell_3_format)
get_sheet_set_row(77, 5, cell_3_format)  # A78行高5像素

get_merge_range('A84:N84', None, cell_3_format)
get_sheet_set_row(83, 5, cell_3_format)  # A84行高5像素

get_merge_range('A90:N90', None, cell_3_format)
get_sheet_set_row(89, 5, cell_3_format)  # A90行高5像素

get_merge_range('A96:N96', None, cell_3_format)
get_sheet_set_row(95, 5, cell_3_format)  # A96行高5像素

get_merge_range('A102:N102', None, cell_3_format)
get_sheet_set_row(101, 5, cell_3_format)  # A102行高5像素

get_merge_range('A108:N108', None, cell_3_format)
get_sheet_set_row(107, 5, cell_3_format)  # A108行高5像素

# 合并单元格4998-4999
get_merge_range(
    'A109:D188',
    'Per Domain Graphics & Media Prioritization (multiple domains) [APL-I]  5%(-)' + '\n' +
    'Per Domain Graphics & Media SLA with QoS (multiple domains) [APL-I] - QoS in SOS  80%(+)' + '\n' +
    'Per Domain Graphics & Media SLA with QoS (multiple domains) [APL-I] - QoS in UOS  90%(+)' + '\n',
    cell_2_format)

# 4998-4999 F列
get_merge_range('F109:I123', '', cell_2_format)
get_merge_range('E109:E123', None, cell_5_format)
get_merge_range('E124:N124', None, cell_3_format)

get_merge_range('F125:I139', '', cell_2_format)
get_merge_range('E125:E139', None,  cell_5_format)
get_merge_range('E140:N140', None, cell_3_format)

get_merge_range('F141:I155', '', cell_2_format)
get_merge_range('E141:E155', None,  cell_5_format)
get_merge_range('E156:N156', None, cell_3_format)

get_merge_range('F157:I171', '', cell_2_format)
get_merge_range('E157:E171', None,  cell_5_format)
get_merge_range('E172:N172', None, cell_3_format)

get_merge_range('F173:I187', '', cell_2_format)
get_merge_range('E173:E187', None,  cell_5_format)
get_merge_range('E188:N188', None, cell_3_format)

# 4998-4999 N列
get_merge_range('K109:N123', '', cell_2_format)
get_merge_range('J109:J123', '', cell_5_format)

get_merge_range('K125:N139', '', cell_2_format)
get_merge_range('J125:J139', '', cell_5_format)

get_merge_range('K141:N155', '', cell_2_format)
get_merge_range('J141:J155', '', cell_5_format)

get_merge_range('K157:N171', '', cell_2_format)
get_merge_range('J157:J171', '', cell_5_format)

get_merge_range('K173:N187', '', cell_2_format)
get_merge_range('J173:J187', '', cell_5_format)

# 隐藏第二行
# worksheet2.set_row(1, None, None, {'hidden': True})
# worksheet.write('A19', 'World', bold)cell_3_format
# 设置单行高度为40像素
# get_sheet_set_row(0, 40, cell_format=bold)
# get_sheet_set_row("A0:A4", None, None, {'hidden': True})
# worksheet.write(3, 0, 35.9999)
# worksheet.write(4, 0, '=SUM(A3:A4)')
# worksheet.insert_image('B5', 'fasfsfsfsaf')
workbook.close()
