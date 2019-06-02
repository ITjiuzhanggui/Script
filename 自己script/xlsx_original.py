# coding: utf-8
import xlsxwriter

# 创建EXcel文件
workbook = xlsxwriter.Workbook('demo.xlsx')

# 创建一个工作表对象
worksheet = workbook.add_worksheet('2019_WW12')
worksheet2 = workbook.add_worksheet('2019_WW13')

cell_format = workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
})


cell_2_format = workbook.add_format({
                            'bold': True,
                            'border': 1,
                            'align': 'center',
                            'valign': 'vcenter'
})


cell_3_format = workbook.add_format({
                            'bold': True,
                            'align': 'center',
                            'valign': 'vcenter',
                            'fg_color': 'gray'
})


cell_4_format = workbook.add_format({
                            'fg_color': 'green',
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_size': 25,
})

# 设置宽度
worksheet.set_column('A:A', 15)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:C', 10)
worksheet.set_column('D:D', 25)

# 设置行高
# worksheet.set_row('A:A', 35)
# worksheet.set_row('B:B', 35)
# worksheet.set_row('C:C', 35)
# worksheet.set_row('D:D', 35)

# 合并单元格---P1
worksheet.merge_range('A1:A5', 'OLINUX-5025', cell_format)
worksheet.merge_range('B1:B5', 'ST_PERF_Single_80_geometry_APL_I', cell_format)

worksheet.merge_range('A7:A11', 'OLINUX-5026', cell_format)
worksheet.merge_range('B7:B11', 'ST_PERF_Single_30_geometry_APL_I', cell_format)

worksheet.merge_range('A13:A17', 'OLINUX-5028', cell_format)
worksheet.merge_range('B13:B17', 'ST_PERF_Multi_80_geometry_APL_I', cell_format)

worksheet.merge_range('A19:A23', 'OLINUX-5031', cell_format)
worksheet.merge_range('B19:B23', 'ST_PERF_Tex_80_APL_I', cell_format)

worksheet.merge_range('A25:A29', 'OLINUX-5032', cell_format)
worksheet.merge_range('B25:B29', 'ST_PERF_Long_80_APL_I', cell_format)

worksheet.merge_range('A31:A35', 'OLINUX-5046', cell_format)
worksheet.merge_range('B31:B35', 'ST_PERF_Single_80_shader_APL_I', cell_format)

worksheet.merge_range('A37:A41', 'OLINUX-5047', cell_format)
worksheet.merge_range('B37:B41', 'ST_PERF_Multi_80_shader_APL_I', cell_format)

worksheet.merge_range('A43:A47', 'OLINUX-5048', cell_format)
worksheet.merge_range('B43:B47', 'ST_PERF_Tex_flat_APL_I', cell_format)

# 合并单元格---P2
worksheet.merge_range('A49:A53', 'OLINUX-5036', cell_format)
worksheet.merge_range('B49:B53', 'ST_PERF_Long_80_APL_I_fps', cell_format)

worksheet.merge_range('A55:A59', 'OLINUX-5037', cell_format)
worksheet.merge_range('B55:B59', 'ST_PERF_Multi_30_geometry_APL_I_fps', cell_format)

worksheet.merge_range('A61:A65', 'OLINUX-5038', cell_format)
worksheet.merge_range('B61:B65', 'ST_PERF_Multi_30_shader_APL_I_fps', cell_format)

worksheet.merge_range('A67:A71', 'OLINUX-5039', cell_format)
worksheet.merge_range('B67:B71', 'ST_PERF_Multi_80_geometry_APL_I_fps', cell_format)

worksheet.merge_range('A73:A77', 'OLINUX-5040', cell_format)
worksheet.merge_range('B73:B77', 'ST_PERF_Multi_80_shader_APL_I_fps', cell_format)

worksheet.merge_range('A79:A83', 'OLINUX-5041', cell_format)
worksheet.merge_range('B79:B83', 'ST_PERF_Single_30_geometry_APL_I_fps', cell_format)

worksheet.merge_range('A85:A89', 'OLINUX-5042', cell_format)
worksheet.merge_range('B85:B89', 'ST_PERF_Single_30_shader_APL_I_fps', cell_format)

worksheet.merge_range('A91:A95', 'OLINUX-5043', cell_format)
worksheet.merge_range('B91:B95', 'ST_PERF_Single_80_geometry_APL_I_fps', cell_format)

worksheet.merge_range('A97:A101', 'OLINUX-5044', cell_format)
worksheet.merge_range('B97:B101', 'ST_PERF_Single_80_shader_APL_I_fps', cell_format)

worksheet.merge_range('A103:A107', 'OLINUX-5045', cell_format)
worksheet.merge_range('B103:B107', 'ST_PERF_Tex_80_APL_I_fps', cell_format)

worksheet.write_string('C1', 'P1/P1', cell_format)
worksheet.write_string('C2', 'P1/P1', cell_format)
worksheet.write_string('C3', 'P1/P1', cell_format)
worksheet.write_string('C4', 'P1/P1', cell_format)
worksheet.write_string('C5', 'P1/P1', cell_format)


# 单元格分割行---P1
worksheet.merge_range('A6:Z6', None, cell_3_format)
worksheet.set_row(5, 5, cell_3_format)  # A6行高5像素

worksheet.merge_range('A12:Z12', None, cell_3_format)
worksheet.set_row(11, 5, cell_3_format)  # A12行高5像素

worksheet.merge_range('A18:Z18', None, cell_3_format)
worksheet.set_row(17, 5, cell_3_format)  # A18行高5像素

worksheet.merge_range('A24:Z24', None, cell_3_format)
worksheet.set_row(23, 5, cell_3_format)  # A24行高5像素

worksheet.merge_range('A30:Z30', None, cell_3_format)
worksheet.set_row(29, 5, cell_3_format)  # A30行高5像素

worksheet.merge_range('A36:Z36', None, cell_3_format)
worksheet.set_row(35, 5, cell_3_format)  # A36行高5像素

worksheet.merge_range('A42:Z42', None, cell_3_format)
worksheet.set_row(41, 5, cell_3_format)  # A42行高5像素

worksheet.merge_range('A48:AQ48', 'P2', cell_4_format)
worksheet.set_row(47, 20)  # A48行高为35像素

# 单元格分割行---P2
worksheet.merge_range('A54:Z54', None, cell_3_format)
worksheet.set_row(53, 5, cell_3_format)  # A54行高5像素

worksheet.merge_range('A60:Z60', None, cell_3_format)
worksheet.set_row(59, 5, cell_3_format)  # A60行高5像素

worksheet.merge_range('A66:Z66', None, cell_3_format)
worksheet.set_row(65, 5, cell_3_format)  # A66行高5像素

worksheet.merge_range('A72:Z72', None, cell_3_format)
worksheet.set_row(71, 5, cell_3_format)  # A72行高5像素

worksheet.merge_range('A78:Z78', None, cell_3_format)
worksheet.set_row(77, 5, cell_3_format)  # A78行高5像素

worksheet.merge_range('A84:Z84', None, cell_3_format)
worksheet.set_row(83, 5, cell_3_format)  # A84行高5像素

worksheet.merge_range('A90:Z90', None, cell_3_format)
worksheet.set_row(89, 5, cell_3_format)  # A90行高5像素

worksheet.merge_range('A96:Z96', None, cell_3_format)
worksheet.set_row(95, 5, cell_3_format)  # A96行高5像素

worksheet.merge_range('A102:Z102', None, cell_3_format)
worksheet.set_row(101, 5, cell_3_format)  # A102行高5像素

worksheet.merge_range('A108:Z108', None, cell_3_format)
worksheet.set_row(107, 5, cell_3_format)  # A108行高5像素


# 合并单元格4998-4999
worksheet.merge_range(
                    'A109:E188',
                    'Per Domain Graphics & Media Prioritization (multiple domains) [APL-I]  5%(-)' + '\n' +
                    'Per Domain Graphics & Media SLA with QoS (multiple domains) [APL-I] - QoS in SOS  80%(+)' + '\n' +
                    'Per Domain Graphics & Media SLA with QoS (multiple domains) [APL-I] - QoS in UOS  90%(+)' + '\n',
                    cell_2_format)

# 4998-4999 F列
worksheet.merge_range('F109:L123', '', cell_2_format)
worksheet.merge_range('F124:S124', None, cell_3_format)

worksheet.merge_range('F125:L139', '', cell_2_format)
worksheet.merge_range('F140:S140', None, cell_3_format)

worksheet.merge_range('F141:L155', '', cell_2_format)
worksheet.merge_range('F156:S156', None, cell_3_format)

worksheet.merge_range('F157:L171', '', cell_2_format)
worksheet.merge_range('F172:S172', None, cell_3_format)

worksheet.merge_range('F173:L187', '', cell_2_format)
worksheet.merge_range('F188:S188', None, cell_3_format)


# 4998-4999 M列
worksheet.merge_range('M109:S123', '', cell_2_format)

worksheet.merge_range('M125:S139', '', cell_2_format)

worksheet.merge_range('M141:S155', '', cell_2_format)

worksheet.merge_range('M157:S171', '', cell_2_format)

worksheet.merge_range('M173:S187', '', cell_2_format)

# 隐藏第二行
# worksheet2.set_row(1, None, None, {'hidden': True})
# worksheet.write('A19', 'World', bold)cell_3_format
#

# 设置单行高度为40像素
# worksheet.set_row(0, 40, cell_format=bold)
# worksheet.set_row("A0:A4", None, None, {'hidden': True})
# worksheet.write(2, 1, 32)
# worksheet.write(1, 1, 32)
# worksheet.write(3, 0, 35.9999)
# worksheet.write(4, 0, '=SUM(A3:A4)')


# worksheet.insert_image('B5', 'fasfsfsfsaf')
workbook.close()
