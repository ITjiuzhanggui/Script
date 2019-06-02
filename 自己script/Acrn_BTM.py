#!/usr/bin/python2.7
# -*- coding: utf-8 -*-

import os
import sys
import re
import time
import json
import argparse
import csv
from threading import Thread
from collections import OrderedDict
reload(sys)
sys.setdefaultencoding("ISO-8859-1")

class Acrnboot:
    def __init__(self, config, dir):
        self.config = json.load(open(config, 'r'), object_pairs_hook=OrderedDict)
        self.logdir = dir
        self.result = OrderedDict()
        self.passset = 0
        self.serline = 0

    def getSerlog(self, sernum):
        serialcmd = self.config["capture"]["thread"][0]% (sernum, acrn_pwd, acrn_user, os.path.join(self.logdir, "{0}.log".format(sernum.replace('/dev/', ''))))
        time.sleep(2)
        os.system(serialcmd)

    def modifyUsb3(self):
        ttyUSB3_stage = []
        ttyUSB3_log = os.path.join(self.logdir, 'ttyUSB3.log')
        if os.path.exists(ttyUSB3_log):
            for i in range(len(open(ttyUSB3_log).readlines())):
                if open(ttyUSB3_log).readlines()[i].find("Starting Kernel") != -1:
                    ttyUSB3_stage.extend( open(ttyUSB3_log).readlines()[i+1:])
        if len(ttyUSB3_stage):
            fo = open(os.path.join(self.logdir, 'ttyUSB3_.log'), 'w')
            for j in ttyUSB3_stage:
                fo.writelines(j)
            fo.close()

    def parseLog(self):
        logResult = {}
        for i in self.config["log"].keys():
            tmpResult = []
            for j in self.config["log"][i]["tag"]:
                log_name = os.path.join(self.logdir, self.config["log"][i]["config"]["log_name"])
                if os.path.exists(log_name) and os.path.getsize(log_name) != 0:
                    for k in open(log_name).readlines():
                        if k and re.search(j["filter"], k):
                            time_pattern_config = self.config["log"][i]["config"]['filter_time_pattern']
                            tmptime = re.findall(time_pattern_config, k)
                            if not filter(lambda x:x[1] == j["name"], tmpResult):
                                logtime = 0
                                if time_pattern_config.find('ms') != -1:
                                    tmptime = re.findall('\d+', tmptime[0])
                                    if len(tmptime) == 1:
                                        logtime = int(tmptime[0])
                                    if len(tmptime) == 2:
                                        logtime = int(tmptime[0])*10**3+int(tmptime[1])
                                else:
                                    if time_pattern_config == '\d+.\d+':
                                        logtime = float(tmptime[0])*10**3
                                    if time_pattern_config == '\d+:\d+:\d+.\d+':
                                        logtime = (float(tmptime[0].split(':')[0])*3600+float(tmptime[0].split(':')[1])*60+float(tmptime[0].split(':')[2]))*10**3
                                tmpResult.append((logtime, j["name"]))
            if self.config["log"][i]["config"]["parent_axis"] == "main":
                tmpResult.append((000, "init"))
            if self.config["log"][i]["config"].has_key("insert_data"):
                insert_data = self.config['log'][i]["config"]["insert_data"]
                if not filter(lambda x: x[1] == insert_data[1], tmpResult):
                    tmpResult.append(tuple(insert_data))
            logResult[i] = sorted(tmpResult, key=lambda x:x[0])
        return logResult

    def raiseResult(self):
        total = []
        Result = []
        logResult = self.parseLog()
        for i in self.config["log"].keys():
            offsettime = 0
            tmpLoglist = []
            parent_axis = self.config["log"][i]["config"]["parent_axis"]
            if not parent_axis == "main":
                if logResult.has_key(i) and logResult.has_key(parent_axis):
                    for j in logResult[i]:
                        for k in logResult[parent_axis]:
                            if j[1] == k[1]:
                                offsettime = float(k[0] - j[0])
                    if offsettime == 0:
                        print "\033[1;31mWarning:'{0} & {1} have no same log' \033[0m".format(i, parent_axis)
                        continue
                    for line in logResult[i]:
                        logTime = float("%.3f" % (line[0] + offsettime))
                        tmpTuple = logTime, line[1]
                        tmpLoglist.append(tmpTuple)
                    logResult[parent_axis].extend(tmpLoglist)
                    total = logResult[parent_axis]
        try:
            for i in total:
                Result.append((float("%.3f" % (i[0]*(0.001))),i[1].encode('utf-8')))
            self.result["total"] = sorted(set(Result), key=lambda x:x[0])
            for j in self.config["stage"]:
                data = self.generate_list(Result, j)
                self.result[j.encode('utf-8')] = data
            print "--*-- Total --*--"
            for resdata in self.result["total"]:
                time.sleep(0.1)
                print resdata
            try:
                end_tag = self.result["step_android"]["detail"][-1]
                print "Result:{0}s ({1})".format(end_tag[0], end_tag[1])
            except Exception, e:
                print "\033[1;31mGenerate Result Error:%s' \033[0m"%e
            time.sleep(1)
            self.writeCsv(self.result)
            return self.result
        except Exception, e:
            print "append data error:%s" % e

    def generate_list(self, result, data):
        display = []
        stage = []
        content = {}
        stage_list = self.config["stage"][data]
        try:
            for m in stage_list:
                for n in m["durations"]:
                    for q in result:
                        if n == q[1]:
                            display.append(q)
            display = sorted(set(display), key=lambda x: x[0])
        except Exception, e:
            print "Generate display_{0} data Error:{1}".format(data, e)

        try:
            for lines in stage_list:
                stage_time = -1
                for line in display:
                    if lines["start"] == line[1]:
                        stage_time = line[0]
                    if stage_time > -1 and lines["end"] == line[1] and (
                        ("%.3f" % (line[0] - stage_time), lines["name"]) not in stage):
                        stage.append(("%.3f" % (line[0] - stage_time), lines["name"]))
            content["detail"] = display
            content["stage"] = stage
            content["data"] = data
            return content
        except Exception, e:
            print "Generate stages failed:%s" % e

    def writeCsv(self,datas):
        try:
            with open(os.path.join(self.logdir,"report.csv"), 'wb') as csvFile:
                writer = csv.writer(csvFile, 'excel-tab')
                writer.writerow(["Total:"])
                for key,value in datas.items():
                    if key == 'total':
                        for entry in datas["total"]:
                            writer.writerow(entry)
                    else:
                        writer.writerow("")
                        writer.writerow(["%s:" % key])
                        for stages in value["stage"]:
                            writer.writerow(stages)
                        writer.writerow("")
                        writer.writerow(["%s:" % (key+"_detail")])
                        for stages in value["detail"]:
                            writer.writerow(stages)
        except Exception, e:
            print e

    def parseMain(self):
        raw_input("************ PLEASE MAKE SURE THE DUT IS POWER OFF & PRESS THE 'ENTER KEY' TO CONTINUE... ************")
        if os.popen("adb devices").read().split('\n')[1] == '':
            t1 = Thread(target=self.getSerlog, args=('/dev/ttyUSB2',))
            t1.start()
            time.sleep(2)
            # cmd = 'echo %s| sudo -S python ./src/serial_com.py -p /dev/ttyUSB2 -c n1#' % linux_pwd
            # if os.system(cmd) == 0:
            t2 = Thread(target=self.getSerlog, args=('/dev/ttyUSB3',))
            t2.start()
            time.sleep(5)
            print "************ \033[1;31m WAIT FOR DUT POWER ON('PLEASE PRESS THE IGNITION BUTTON')...\033[0m ************"
            t1.join()
            t2.join()
            time.sleep(2)
            print "************ CAPTURE ADB LOGS ************"
            for i in self.config["capture"]["commands"]:
                if i.find('adb') != -1:
                    if i.find('%s') != -1:
                        i = i.replace('%s', self.logdir)
                    try:
                        os.system(i)
                    except Exception, e:
                        print e
                if i.find('sleep') != -1:
                    i = int(i.replace('sleep ', ''))
                    time.sleep(i)
            time.sleep(2)
            print "************ PARSE LOGS ************"
            time.sleep(2)
            self.modifyUsb3()
            self.raiseResult()

        else:
            time.sleep(2)
            self.parseMain()
def usage():
    parser = argparse.ArgumentParser(description="Test boot time for acrn")
    parser.add_argument('-u', '--user', dest='acrn_username',type=str,default='',help='Username for acrn_sos')
    parser.add_argument('-p', '--pwd', dest='acrn_password',type=str,default='',help='password for acrn_sos')
    parser.add_argument('-P', '--Pwd', dest='linux_password',type=str,default='',help='Password for linux host')
    parser.add_argument('-c', '--config', dest='config_name', type=str, help='config name in the config folder')
    return parser.parse_args()

if __name__ == "__main__":
    linux_pwd = ''
    acrn_user = ''
    acrn_pwd = ''
    config_name = ''
    args = usage()
    if args.linux_password:
        linux_pwd = args.linux_password
    else:
        print "Must input linux host passwd,useage: '-P linux_pwd'"
        sys.exit()
    if args.acrn_password:
        acrn_pwd = args.acrn_password
    else:
        print "Must input acrn_sos passwd,useage: '-p pwd'"
        sys.exit()
    if args.acrn_username:
        acrn_user = args.acrn_username
    else:
        print "Must input acrn_sos username,useage: '-u user'"
        sys.exit()
    if args.config_name:
        config_name = args.config_name
    else:
        print "Must input config name in the config folder,usage: '-c config_name'"
        sys.exit()
    os.system("echo %s|sudo -S chmod 777 /dev/ttyUSB3" % linux_pwd)
    os.system("echo %s|sudo -S chmod 777 /dev/ttyUSB2" % linux_pwd)
    logdir = os.path.join(os.path.abspath(__file__).strip(os.path.abspath(__file__).split('/')[-1]), time.strftime("%Y%m%d_%H%M%S"))
    if not os.path.exists(logdir):
        os.makedirs(logdir)
    print "************ BREAKDOWN DATA PATH: %s ************"%logdir
    config = os.path.join(os.path.abspath('./'), 'config', config_name)
    Acrnboot(config, logdir).parseMain()








