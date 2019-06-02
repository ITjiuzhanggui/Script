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

            self.scpclient.get("/root/daimler_ic/result.txt", local_path="/home/zds/Downloads/")
        except Exception as e:
            print("remote_pull():", e)
            return
        print('下拉成功')

    def connect_remote_host(self):

        client = paramiko.SSHClient()

        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            client.connect(hostname=self.hostname, port=self.port, username=self.username, password=self.password)
        except Exception as e:

            print("connect_remote_host():", e)
            return None,
        self.client = client

    def exec_shell_cmd(self, cmds):

        ssh = self.client.invoke_shell()

        fd = open(file=self.filename, mode='w')
        for cmd in cmds:
            print('cmd:', cmd)

            if cmd.startswith("stdbuf"):

                ssh.send(cmd)

                output = ssh.recv(4096)
                print(len(output))
                print("output:", output.decode())
                fd.write(output.decode())

                time.sleep(60 * 1.7)

                ssh.send(chr(3))
            else:

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



    def calc_reset_up(self):
        with open(self.filename, 'r') as f:
            lines = f.readlines()
            reset_up_list = []
            reset_down_list = []

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

                reset_state_index = index
            index += 1
        i = 1

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

    cmds = ["cd /root/daimler_ic\n",
            "swupd autoupdate --disable\n",
            "echo -1023 > /sys/module/i915/parameters/gvt_workload_priority\n",
            "echo 'PermitRootLogin yes' > /etc/ssh/sshd_config\n",
            "systemctl enable ias\n",
            "systemctl start ias\n",
            "export XDG_RUNTIME_DIR=/run/ias\n",
            "stdbuf -oL ./daimler_ic-wayland -x 2560 -y 960 -m 61 -fullscreen > /root/daimler_ic/result.txt\n"]

    remoter.exec_shell_cmd(cmds)

    remoter.remote_pull()

    # remoter.calc()
    remoter.calc_reset_up()
    remoter.calc_reset_down()


if __name__ == '__main__':
    main()
