# -*-coding: utf-8 -*-

import paramiko
import time
import re
from scp import SCPClient


class RemoteHelper(object):
    def __init__(self, hostname, port, username, password):
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = password

        self.client = None
        self.scpclient = None
        # 保存命令执行的结果的文件
        self.filename = "result.txt"

    def upload_file(self):

        if self.client is None:
            print("upload_file():ssh连接没有成功")
            return

        try:
            scpclient = SCPClient(self.client.get_transport(), socket_timeout=15.0)
            self.scpclient = scpclient
            # 上传文件到远程
            scpclient.put("/home/zds/0905/sos_workload/daimler_ic", "/root", recursive=True)
            # 把文件从远程复制下来
            # time.sleep(60*3.5)
            # scpclient.get("/root/daimler_ic/test.txt", local_path="/home/zds/Test_text")

        except Exception as e:
            print("upload_file():", e)
            return
        print("上传文件成功")

    def remote_pull(self):

        if self.scpclient is None:
            print('remote_pull():ssh连接没有成功')
            return

        try:
            # 下拉文件到本地
            self.scpclient.get("/root/daimler_ic/result.txt", local_path="/home/zds/Downloads/")
        except Exception as e:
            print("remote_pull():", e)
            return
        print('下拉成功')

    def connect_remote_host(self):
        # 创建SSH对象
        client = paramiko.SSHClient()
        # 添加需要连接的IP
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            client.connect(hostname=self.hostname, port=self.port, username=self.username, password=self.password)
        except Exception as e:
            # 连接失败
            print("connect_remote_host():", e)
            return None,
        self.client = client

    def exec_shell_cmd(self, cmds):
        # 使用send来把命令发送到远程执行
        # 在SSH server端创建一个交互式的shell，且可以按自己的需求配置伪终端，
        # 可以在invoke_shell()函数中添加参数配置。
        ssh = self.client.invoke_shell()
        # 打开一个文件，
        fd = open(file=self.filename, mode='w')
        for cmd in cmds:
            print('cmd:', cmd)
            # 如果命令是以.开头，那么说明执行的是./daimler_ic-wayland -x 2560 -y 960 -m 61 -fullscreen命令
            # 这时就需要多给他一些时间来执行
            if cmd.startswith("stdbuf"):
                # 用send函数发送cmd到SSH server
                # self.client.exec_command(cmd)
                ssh.send(cmd)
                # 每次能接收数据的大小
                output = ssh.recv(4096)
                print(len(output))
                print("output:", output.decode())
                fd.write(output.decode())
                # 暂停一会，让命令跑一会
                # 如果要计算2分钟的
                time.sleep(60 * 1.7)
                # contorl -c 命令
                ssh.send(chr(3))
            else:
                # 执行普通的命令，而且这些命令的输出结果都不要经

                ssh.send(cmd)
                output = ssh.recv(4096)
                print("output:", output.decode())
                time.sleep(1)

        print("==================命令执行完成=====================")
        # 关闭文件
        # fd.close()
        # 关闭远程连接
        # self.close()
    #
    def close(self):
        self.client.close()

    # def calc(self):
    #     print('calc latecy...')
    #     # 读取文件进行解析
    #     with open(self.filename, 'r')as load_f:
    #         # with open('sos.10.15.4.txt', 'r')as load_f:
    #         lines = load_f.readlines()
    #         p_start = []
    #         p_end = []
    #         sum_up = 0.0
    #         sum_down = 0.0
    #         avg_up = 0.0
    #         avg_down = 0.0
    #
    #         for line in lines:
    #             try:
    #                 restup_up = re.search(r'[0]{1}/0', line)
    #                 if restup_up:
    #                     num = re.search(r'\d{2}\.\d{3}', line)
    #                     if num:
    #                         p_start.append(num.group())
    #
    #             except Exception as e:
    #                 print(e)
    #                 print('Real number read above ‘reset’ failed')
    #
    #         for line in lines:
    #             try:
    #                 reset_dowm = re.search(r'[1-9][0-9]/0', line)
    #                 if reset_dowm:
    #                     num = re.search(r'\d{2}\.\d{3}', line)
    #                     if num:
    #                         p_end.append(num.group())
    #
    #             except Exception as e:
    #                 print(e)
    #                 print('Real number read below ‘reset’ failed')
    #
    #         del p_start[0]
    #         # print(p_start)
    #         # print(p_end)
    #
    #         for i in p_start:
    #             sum_up = sum_up + float(i)
    #
    #         # for i in p_end:
    #         #     sum_down = sum_down + float(i)
    #
    #         for i in range(0, 10):
    #             sum_dowm = sum_down + float(p_end[i])
    #
    #         print('The sum above "reset" is:%f' % sum_up)
    #         print('The sum under "reset" is:%f' % sum_down)
    #
    #         avg_up = sum_up / len(p_start)
    #         print('The average value above "reset" is:%f' % avg_up)
    #
    #         avg_down = sum_down / len(p_end)
    #         print('The average value under "reset" is:%f' % avg_down)

    # def calc(self):
    #     print('calc latecy...')
    #     with open(self.filename, 'r')as load_f:
    #         lines = load_f.readlines()
    #         p_start = []
    #         p_end = []
    #         sum_up = 0.0
    #         sum_down = 0.0
    #         avg_up = 0.0
    #         avg_down = 0.0
    #
    #         for line in lines:
    #
    #             try:
    #                 restup_up = re.search(r'[0]{1}/0', line)
    #                 if restup_up:
    #                     num = re.search(r'\d{2}\.\d{3}', line)
    #                     # print(restup_up)
    #                     if num:
    #                         p_start.append(num.group())
    #
    #             except Exception as e:
    #                 # print(e)
    #                 pass
    #         # for line in lines:
    #         #     try:
    #         #         reset_dowm = re.search(r'[1-9][0-9]/0', line).group()
    #         #         if reset_dowm:
    #         #             num = re.search(r'\d{2}\.\d{3}', line).group()
    #         #             # print(reset_dowm)
    #         #             p_end.append(num)
    #         #
    #         #     except Exception as e:
    #         #         pass
    #         for line in lines:
    #             i = 0
    #             if '{}/0'.format(i) in line:
    #                 # 1/0-9/0
    #                 reset_dowm = re.search(r'[1-9]/0', line)
    #                 print(reset_dowm)
    #                 i += 1
    #
    #             else:
    #                 reset_dowm = re.search(r'[1-9][0-9]/0', line)
    #
    #             if reset_dowm:
    #                 print(reset_dowm)
    #                 num = re.search(r'\d{2}\.\d{3}', line)
    #                 if num:
    #                     p_end.append(num.group())
    #                     print(p_end)
    #                     print(len(p_end))
    #
    #
    #         del p_start[0]
    #         # print(p_start)
    #         print(p_end)
    #         # print(p_start)
    #
    #         for i in p_start:
    #             sum_up = sum_up + float(i)
    #
    #         # for i in p_end:
    #         for i in range(0, 10):  # sum_down = sum_down + float(i)
    #             print(p_end[i])
    #             sum_down = sum_down + float(p_end[i])
    #
    #
    #         print('The sum above "reset" is:%f' % sum_up)
    #         print('The sum under "reset" is:%f' % sum_down)
    #
    #         avg_up = sum_up / len(p_start)
    #         print('The average value above "reset" is:%f' % avg_up)
    #
    #         # avg_down = sum_down / len(p_end)
    #         avg_down = sum_down / len(p_end[0:10])
    #
    #         print('The average value under "reset" is:%f' % avg_down)

    def calc_reset_up(self):
        with open(self.filename, 'r') as f:
            lines = f.readlines()
            reset_up_list = []
            reset_down_list = []
            # 匹配reset state 之上的数据
            for line in lines:
                reset_up = re.search(r'0/0', line)
                if reset_up:
                    reset_up_num = re.search(r'\d{2}\.\d{3}', line)
                    if reset_up_num:
                        reset_up_list.append(reset_up_num.group())
                        # print(reset_up_num.group())

            if reset_up_list is None:
                print("reset_up_list是空的")
                return
            # 删除reset_up_list中的第一个元素
            del reset_up_list[0]

            reset_up_sum = 0.0
            for num in reset_up_list:
                reset_up_sum += float(num)
        avg_reset_up = reset_up_sum / len(reset_up_list)
        print('The average value of above "reset" is:', avg_reset_up)
        print(reset_up_list)
        return avg_reset_up

    def calc_reset_down(self):
        with open(self.filename, 'r') as f:
            lines = f.readlines()

            # print(lines)
        reset_state = "Reset stats"
        reset_state_index = 0
        index = 0
        for line in lines:
            if reset_state in line:
                # 找到Reset stats的位置
                reset_state_index = index
            index += 1
        i = 1
        # 获取Reset stats之后的所有数据
        reset_down_lines = lines[reset_state_index+1:]
        reset_down_list = []
        for line in reset_down_lines:
            if i < 10:
                reset_down = re.search(r'[1-9]/0', line)
            else:
                reset_down = re.search(r'[1-9][0-9]/0', line)
            i += 1
            if reset_down:
                reset_down_num = re.search(r'\d{2}\.\d{3}', line)
                if reset_down_num:
                    reset_down_list.append(reset_down_num.group())

        if len(reset_down_list) < 10:
            print("reset_down_list的数据不足10条")
            print(reset_down_list)
            return

        reset_down_sum = 0.0
        for i in range(0, 10):
            reset_down_sum += float(reset_down_list[i])
        avg_reset_down = reset_down_sum / len(reset_down_list[0:10])
        print('The average value of down "reset" is:', avg_reset_down)
        return avg_reset_down


def main():
    hostname = "10.239.85.38"
    port = 22
    username = "root"
    password = "zhangxinhuicf"

    remoter = RemoteHelper(hostname, port, username, password)
    remoter.connect_remote_host()
    if remoter.client is None:
        print("连接失败！！！")

    print("连接成功")
    remoter.upload_file()
    # 利用send函数发送cmd到SSH server，添加'\n'做回车来执行shell命令。
    # 注意不同的情况，如果执行完telnet命令后，telnet的换行符是\r\n
    cmds = ["cd /root/daimler_ic\n",
            "swupd autoupdate --disable\n",
            "echo -1023 > /sys/module/i915/parameters/gvt_workload_priority\n",
            "echo 'PermitRootLogin yes' > /etc/ssh/sshd_config\n",
            "systemctl enable ias\n",
            "systemctl start ias\n",
            "export XDG_RUNTIME_DIR=/run/ias\n",
            "stdbuf -oL ./daimler_ic-wayland -x 2560 -y 960 -m 61 -fullscreen > /root/daimler_ic/result.txt\n"]
    # 在远程执行命令
    remoter.exec_shell_cmd(cmds)
    # 下拉远程数据
    remoter.remote_pull()
    # 计算结果xl
    # remoter.calc()
    remoter.calc_reset_up()
    remoter.calc_reset_down()


if __name__ == '__main__':
    main()
