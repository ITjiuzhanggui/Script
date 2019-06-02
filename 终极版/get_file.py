


# shell 命令
# 创建文件 , 精确到分钟
'''
date=`date +%Y%m%d%H%M`

touch data
'''


# python

import os
# 获取当前文件的 路径
import time
import safe_linux

curpath = os.path.dirname(os.path.realpath(__file__))


# 当前路径下创建 文件夹(logs)
logs_dir = os.path.join(curpath, "logs")
os.makedirs(logs_dir, exist_ok=True)

cmd = ["make","update"]
rc, outstr, errstr = safe_linux.OSUtil.safe_popen(cmd)

logs = os.path.join(logs_dir, "%s"%int(time.time()))
with open(logs, 'w') as f:
    f.write(outstr)