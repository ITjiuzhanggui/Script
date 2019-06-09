import pandas as pd

pd.set_option("expand_frame_repr", False)


def calc_rate(default_golang_list, clear_golang_list):
    rate_list = []
    if len(default_golang_list) != len(clear_golang_list):
        print("数据不对应！")
        return None

    tmp_result_list = []
    for i in range(len(default_golang_list)):
        print(clear_golang_list[i], "    ", default_golang_list[i])
        tmp_ret = float(clear_golang_list[i]) / float(default_golang_list[i])
        # print(tmp_ret)
        result = 1 / tmp_ret - 1
        # print(result)
        tmp_result_list.append(result)

    for i in range(len(tmp_result_list)):
        ret = tmp_result_list[i]
        ret = str(ret * 100)[:4] + "%"
        rate_list.append(ret)
    return rate_list


def GolangExcel():
    df_json = pd.read_json(r"C:\Users\xinhuizx\python_Code\MQ_script\data_LOG.json")
    print(df_json)
    clearlinux_version_dict = df_json.loc["clearlinux_version"].loc["status_Clr"]
    clearlinux_version = clearlinux_version_dict["clear_linux"]

    default_dict = df_json.loc["golang"].loc["default"]
    clear_dict = df_json.loc["golang"].loc["clear"]

    x_list = ["BenchmarkBuild", "BenchmarkGarbage", "BenchmarkHTTP", "BenchmarkJSON"]
    x_col = pd.Series(x_list)

    default_golang_list = [default_dict["BenchmarkBuild"], default_dict["BenchmarkGarbage"],
                           default_dict["BenchmarkHTTP"], default_dict["BenchmarkJSON"]]

    default_col = pd.Series(default_golang_list)
    clear_golang_list = [clear_dict["BenchmarkBuild"], clear_dict["BenchmarkGarbage"],
                         clear_dict["BenchmarkHTTP"], clear_dict["BenchmarkJSON"]]

    rate_col = calc_rate(default_golang_list, clear_golang_list)
    clear_col = pd.Series(clear_golang_list)

    data_frame = {"X": x_col, "Default docker": default_col, "Clear docker": clear_col, "Rate": rate_col}
    data_frame2 = {"X2": x_col, "Default docker2": default_col, "Clear docker2": clear_col, "Rate2": rate_col}
    df_excel = pd.DataFrame(data_frame)
    df_excel2 = pd.DataFrame(data_frame2)

    # 写入此文件
    writer = pd.ExcelWriter(r"C:\Users\xinhuizx\python_Code\MQ_script\MQ_tset.xlsx")
    df_excel.to_excel(writer, sheet_name="golang", index=False, startrow=0)
    df_excel2.to_excel(writer, sheet_name="golang", index=False, startrow=9)
    writer.save()
    print("写入成功！")


if __name__ == "__main__":
    GolangExcel()

# col 为列名 相当于Key
# x_list 相当于值
# rate_col 算法的值
