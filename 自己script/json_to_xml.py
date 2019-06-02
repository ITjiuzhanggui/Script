class JsonfileToXlsx():
    def __init__(self, column, row):
        import xlsxwriter
        import os
        self.workbook = xlsxwriter.Workbook('demo.xlsx')
        self.jsonforfile = [i for i in os.listdir('.') if \
                             i in os.listdir('.') and \
                                 i.endswith('json') or i.startswith('1+')]
        print(self.jsonforfile)
        self.total_p1 = 0
        self.total_p2 = 0

        self.row = row
        self.column = column
        self.rowdiffer = 1
        self.init_all = True

    def init_xls(self,worksheet,num=0,type='DRAM'):
        if self.init_all:
            self.init_all = False
            for item in self.init_xls(worksheet,4,type='DCPMM'):
                pass
        yield worksheet.write(0, num+0, type)
        yield worksheet.write(1, num+0, 'Query ID')
        yield worksheet.write(1, num+1, 'R1')
        yield worksheet.write(1, num+2, 'R2')
        yield worksheet.write(11, num+0, 'GEOMEAN')
        yield worksheet.write(13, num+0, 'PERF GAIN (DCPMM .VA DRAM)')
        yield worksheet.write(14, num+0, 'RATE')
        yield worksheet.write(13, num+1, 'R1')
        yield worksheet.write(13, num+2, 'R2')


    def get_date(self):
        import json
        for jsonfilename in self.jsonforfile:
            if jsonfilename.startswith('1+') and self.init_all:
                with open(jsonfilename, 'r') as f:
                    data = json.load(f)
                    yield (data, jsonfilename,'1+')
            else:
                with open(jsonfilename, 'r') as f:
                    data = json.load(f)
                    yield (data, jsonfilename, None)


    def save_to_xlsx(self):
        for json_date, name, if_all in self.get_date():
            worksheet = self.workbook.add_worksheet(name.split('-')[0])
            if if_all:
                for item in self.init_xls(worksheet):
                    pass
                rnum = self.row
                cnum = self.column
                dnum = self.rowdiffer
                for k, v in json_date.items():
                    yield worksheet.write(rnum, cnum, k)
                    yield worksheet.write(
                        rnum, cnum + dnum, str(v[0][1])
                    )
                    self.total_p1 += v[0][1]
                    yield worksheet.write(
                        rnum, cnum + dnum + 1, str(v[1][1])
                    )
                    self.total_p2 += v[1][1]
                    rnum += 1
                yield worksheet.write(
                    rnum, cnum + dnum, self.total_p1
                )
                yield worksheet.write(
                    rnum, cnum + dnum + 1, self.total_p2
                )
            else:
                for item in self.init_xls(worksheet):
                    pass
                rnum = self.row
                cnum = self.column+4
                dnum = self.rowdiffer
                for k, v in json_date.items():
                    yield worksheet.write(rnum, cnum, k)
                    yield worksheet.write(
                        rnum, cnum + dnum, str(v[0][1])
                    )
                    self.total_p1 += v[0][1]
                    yield worksheet.write(
                        rnum, cnum + dnum + 1, str(v[1][1])
                    )
                    self.total_p2 += v[1][1]
                    rnum += 1
                yield worksheet.write(
                    rnum, cnum + dnum, self.total_p1
                )
                yield worksheet.write(
                    rnum, cnum + dnum + 1, self.total_p2
                )

    def run(self):
        for i in self.save_to_xlsx():
            pass
        self.workbook.close()

if __name__ == '__main__':
    import time
    c =time.time()
    a = JsonfileToXlsx(0,2)
    a.run()
    print(time.time()-c)