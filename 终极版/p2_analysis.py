import os
import sys
import re


def find_guest_file(base_dir, guest_total_files, key):
    date = "20181202"
    image_dir = base_dir % date

    # os.system("find . -name '*.Guest.out > '")
    ret = os.system(
        "stdbuf -oL find {image_dir} -name '*{key}' > {file}".format(image_dir=image_dir, file=guest_total_files, key=key))

    if ret != 0:
        print("find command failed.")
        return


def filter_p2_file(guest_total_files):
    case_name_list = ["OLINUX-5036", "OLINUX-5037", "OLINUX-5038", "OLINUX-5039",
                      "OLINUX-5040", "OLINUX-5041", "OLINUX-5042", "OLINUX-5043",
                      "OLINUX-5044", "OLINUX-5045"]

    with open(guest_total_files, 'r') as f:
        tmp_files = f.readlines()

    if tmp_files is None:
        print("Error: guest file is None")
    guest_files = []
    for filename in tmp_files:
        for case_name in case_name_list:
            if case_name == filename.split("/")[-2]:
                # print(filename.split("/")[-2])
                guest_files.append(filename)
    # print(guest_files)

    return guest_files


def read_file(guest_files_dict):
    ret_dict = {}
    for case_name, files in guest_files_dict.items():
        ret_dict[case_name] = []
        # tmp_list = []
        for filename in files:
            tmp_dict = {}
            fps = calc_avg_fps(filename)
            tmp_dict["filename"] = filename
            tmp_dict["fps"] = fps
            ret_dict[case_name].append(tmp_dict)

    for case_name, files in ret_dict.items():
        print(case_name, ":", len(files))
    return ret_dict


def save_file(ret_dict, filename):
    if ret_dict is None:
        print("Error: ret dict is None")
        sys.exit(-1)

    fd = open(filename, "w")
    for case_name, values in ret_dict.items():
        fd.write(case_name+":\n")
        for value in values:
            fd.write("\t" + value["filename"] + "\n\t\t" + "fps:\t" + value["fps"] + "\n")

        fd.write("\n")

    fd.close()
    print("save file ok!")


def calc_avg_fps(filename):
    if filename is None:
        print("Error: filename is not exists")
        sys.exit(-1)

    fps_sum = 0.0
    num = 1
    tmp_lines = []

    with open(filename.split("\n")[0]) as f:
        lines = f.readlines()

    for line in lines:
        fps = re.search(r"seconds: (\d+\.\d+)", line)
        if fps:
            tmp_lines.append(line)

    for line in tmp_lines[:29]:
        fps = re.search(r"seconds: (\d+\.\d+)", line)
        if fps:
            fps_sum += float(fps.group(1))
            num += 1

    # print("num:", num)

    avg_fps = fps_sum / num
    # print("the avg fps of %d: %0.3f" % (num, avg_fps))
    return "%.3f" % avg_fps


def merge_case(guest_files):
    result_dict = {}
    for i in range(0, len(guest_files)):
        case_name = guest_files[i].split("/")[-2]
        result_dict[case_name] = []

    for filename in guest_files:
        case_name = filename.split("/")[-2]
        if case_name in result_dict.keys():
            result_dict[case_name].append(filename)
    # print(result_dict)
    return result_dict


def main():
    base_dir = "/home/zxh/rawbuild/result/%s"
    guest_total_files = "p2_guest_total_files.txt"
    fps_file = "/home/zxh/rawbuild/result/20181202/fps_result.txt"
    key = ".Guest.out"
    find_guest_file(base_dir, guest_total_files, key)
    guest_files = filter_p2_file(guest_total_files)
    print("length of guest files", len(guest_files))
    result_dict = merge_case(guest_files)
    ret_dict = read_file(result_dict)
    save_file(ret_dict, fps_file)


if __name__ == '__main__':
    main()
