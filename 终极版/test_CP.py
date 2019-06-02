import xlsxwriter

class JsonFileXlsx(object):
    def __init__(self, workbook, worksheet):
        self.workbook = xlsxwriter.Workbook('deml____.xlsx')
        self.worksheet = self.workbook.add_worksheet('2019_WW13')
        self.list_P2_C = []
        self.data = {
            'A:A': 18, 'B:B': 40, 'C:C': 10, 'D:D': 20, \
            'E:E': 25, 'G:G': 20, 'H:H': 15, \
            'I:I': 15, 'J:J': 25, \
            'K:K': 10, 'L:L': 20, 'M:M': 15, 'N:N': 15,
        }
        self.cell_format = self.workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
        })
        self.cell_2_format = self.workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })
        self.cell_3_format = self.workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'gray',
        })
        self.cell_4_format = self.workbook.add_format({
            'fg_color': 'green',
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 25,
            'border': 1,
        })
        self.cell_5_format = self.workbook.add_format({
            'bg_color': 'green',
            'pattern': True,
            'align': 'center',
            'valign': 'vcenter',
        })
        self.cells = [
            ('A1:A5', 'OLINUX-5025', self.cell_format),
            ('B1:B5', 'ST_PERF_Single_80_geometry_APL_I', self.cell_format),
            ('A7:A11', 'OLINUX-5026', self.cell_format),
            ('B7:B11', 'ST_PERF_Single_30_geometry_APL_I', self.cell_format),
            ('A13:A17', 'OLINUX-5028', self.cell_format),
            ('B13:B17', 'ST_PERF_Multi_80_geometry_APL_I', self.cell_format),
            ('A19:A23', 'OLINUX-5031', self.cell_format),
            ('B19:B23', 'ST_PERF_Tex_80_APL_I', self.cell_format),
            ('A25:A29', 'OLINUX-5032', self.cell_format),
            ('B25:B29', 'ST_PERF_Long_80_APL_I', self.cell_format),
            ('A31:A35', 'OLINUX-5046', self.cell_format),
            ('B31:B35', 'ST_PERF_Single_80_shader_APL_I', self.cell_format),
            ('A37:A41', 'OLINUX-5047', self.cell_format),
            ('B37:B41', 'ST_PERF_Multi_80_shader_APL_I', self.cell_format),
            ('A43:A47', 'OLINUX-5048', self.cell_format),
            ('B43:B47', 'ST_PERF_Tex_flat_APL_I', self.cell_format),
            ('A49:A53', 'OLINUX-5036', self.cell_format),
            ('B49:B53', 'ST_PERF_Long_80_APL_I_fps', self.cell_format),
            ('A55:A59', 'OLINUX-5037', self.cell_format),
            ('B55:B59', 'ST_PERF_Multi_30_geometry_APL_I_fps', self.cell_format),
            ('A61:A65', 'OLINUX-5038', self.cell_format),
            ('B61:B65', 'ST_PERF_Multi_30_shader_APL_I_fps', self.cell_format),
            ('A67:A71', 'OLINUX-5039', self.cell_format),
            ('B67:B71', 'ST_PERF_Multi_80_geometry_APL_I_fps', self.cell_format),
            ('A73:A77', 'OLINUX-5040', self.cell_format),
            ('B73:B77', 'ST_PERF_Multi_80_shader_APL_I_fps', self.cell_format),
            ('A79:A83', 'OLINUX-5041', self.cell_format),
            ('B79:B83', 'ST_PERF_Single_30_geometry_APL_I_fps', self.cell_format),
            ('A85:A89', 'OLINUX-5042', self.cell_format),
            ('B85:B89', 'ST_PERF_Single_30_shader_APL_I_fps', self.cell_format),
            ('A91:A95', 'OLINUX-5043', self.cell_format),
            ('B91:B95', 'ST_PERF_Single_80_geometry_APL_I_fps', self.cell_format),
            ('A97:A101', 'OLINUX-5044', self.cell_format),
            ('B97:B101', 'ST_PERF_Single_80_shader_APL_I_fps', self.cell_format),
            ('A103:A107', 'OLINUX-5045', self.cell_format),
            ('B103:B107', 'ST_PERF_Tex_80_APL_I_fps', self.cell_format),
            ('A6:N6', None, self.cell_3_format),
            ('A12:N12', None, self.cell_3_format),
            ('A18:N18', None, self.cell_3_format),
            ('F109:I123', '', self.cell_2_format),
            ('E109:E123', None, self.cell_5_format),
            ('E124:N124', None, self.cell_3_format),
            ('F125:I139', '', self.cell_2_format),
            ('E125:E139', None, self.cell_5_format),
            ('E140:N140', None, self.cell_3_format),
            ('F141:I155', '', self.cell_2_format),
            ('E141:E155', None, self.cell_5_format),
            ('E156:N156', None, self.cell_3_format),
            ('F157:I171', '', self.cell_2_format),
            ('E157:E171', None, self.cell_5_format),
            ('E172:N172', None, self.cell_3_format),
            ('F173:I187', '', self.cell_2_format),
            ('E173:E187', None, self.cell_5_format),
            ('E188:N188', None, self.cell_3_format),
            ('K109:N123', '', self.cell_2_format),
            ('J109:J123', '', self.cell_5_format),
            ('K125:N139', '', self.cell_2_format),
            ('J125:J139', '', self.cell_5_format),
            ('K141:N155', '', self.cell_2_format),
            ('J141:J155', '', self.cell_5_format),
            ('K157:N171', '', self.cell_2_format),
            ('J157:J171', '', self.cell_5_format),
            ('K173:N187', '', self.cell_2_format),
            ('J173:J187', '', self.cell_5_format),
            (
                'A109:D188',
                'Per Domain Graphics & Media Prioritization (multiple domains), [APL-I]  5%(-),' + '\n' +
                'Per Domain Graphics & Media SLA with QoS (multiple domains), [APL-I] - QoS in SOS  80%(+),' + '\n' +
                'Per Domain Graphics & Media SLA with QoS (multiple domains), [APL-I] - QoS in UOS  90%(+),' + '\n',
                self.cell_2_format),
            ('A24:N24', None, self.cell_3_format),
            ('A30:N30', None, self.cell_3_format),
            ('A36:N36', None, self.cell_3_format),
            ('A42:N42', None, self.cell_3_format),
            ('A48:N48', 'P2', self.cell_4_format),
            ('A54:N54', None, self.cell_3_format),
            ('A60:N60', None, self.cell_3_format),
            ('A66:N66', None, self.cell_3_format),
            ('A72:N72', None, self.cell_3_format),
            ('A78:N78', None, self.cell_3_format),
            ('A84:N84', None, self.cell_3_format),
            ('A90:N90', None, self.cell_3_format),
            ('A96:N96', None, self.cell_3_format),
            ('A108:N108', None, self.cell_3_format)
        ]
        self.sellow=[
            (5, 5, self.cell_3_format),  # A6行高5像素
            (11, 5, self.cell_3_format),  # A12行高5像素
            (17, 5, self.cell_3_format),  # A18行高5像素
            (23, 5, self.cell_3_format),  # A24行高5像素
            (29, 5, self.cell_3_format),  # A30行高5像素
            (35, 5, self.cell_3_format),  # A36行高5像素
            (41, 5, self.cell_3_format),  # A42行高5像素
            (47, 20),  # A48行高为35像素
            (53, 5, self.cell_3_format),  # A54行高5像素
            (59, 5, self.cell_3_format),  # A60行高5像素
            (65, 5, self.cell_3_format),  # A66行高5像素
            (71, 5, self.cell_3_format),  # A72行高5像素
            (77, 5, self.cell_3_format),  # A78行高5像素
            (83, 5, self.cell_3_format),  # A84行高5像素
            (89, 5, self.cell_3_format),  # A90行高5像素
            (95, 5, self.cell_3_format),  # A96行高5像素
            (101, 5, self.cell_3_format),  # A102行高5像素
            (107, 5, self.cell_3_format),  # A108行高5像素
        ]

        for item in self.cells:
            self.get_merge_range(*item)
        for item in self.sellow:
            self.get_sheet_set_row(*item)

    def get_merge_range(self, *args, **kwargs):
        def write_(*args, **kwargs):
            yield self.worksheet.merge_range(*args, **kwargs)

        [item for item in write_(*args, **kwargs)]

    def get_sheet_set_row(self, *args, **kwargs):
        def sheet_(*args, **kwargs):
            yield self.worksheet.set_row(*args, **kwargs)

        [item for item in sheet_(*args, **kwargs)]

    def P1P2(self):
        for i in range(1, 6):
            self.list_P2_C.append("C%d" % (i + 48))
            self.list_P2_C.append("C%d" % (i + 48 + 6))
            self.list_P2_C.append("C%d" % (i + 48 + 12))
            self.list_P2_C.append("C%d" % (i + 48 + 18))
            self.list_P2_C.append("C%d" % (i + 48 + 24))
            self.list_P2_C.append("C%d" % (i + 48 + 30))
            self.list_P2_C.append("C%d" % (i + 48 + 36))
            self.list_P2_C.append("C%d" % (i + 48 + 42))
            self.list_P2_C.append("C%d" % (i + 48 + 48))
            self.list_P2_C.append("C%d" % (i + 48 + 54))
        for item in self.list_P2_C:
            num = int(item[1:])
            item = item.replace("%d" % num, str(num - 48))
            yield self.worksheet.write_string(item, 'P1/P1', self.cell_format)
            item1 = item.replace("C", "F")
            yield self.worksheet.write_string(item1, 'P1/P1', self.cell_format)
        for item in self.list_P2_C:
            yield self.worksheet.write_string(item, 'P2/P2', self.cell_format)
            item1 = item.replace("C", "F")
            yield self.worksheet.write_string(item1, 'P2/P2', self.cell_format)

    def run(self):
        for item in self.P1P2():
            pass
        # 设置宽度
        for k, v in self.data.items():
            self.worksheet.set_column(k, v)
        self.workbook.close()


if __name__ == '__main__':
    import time
    gtime =time.time()
    j = JsonFileXlsx('workbook', 'worksheet')
    j.run()
    print(time.time()- gtime)