#!/usr/bin/env python
# conding=utf-8

'a test module'

import os
import re


def show():
    pass


def calc():
    print('calc latecy...')
    with open('sos.10.15.4.txt', 'r')as load_f:
        lines = load_f.readlines()
        p_start = []
        p_end = []
        sum_up = 0.0
        sum_down = 0.0
        avg_up = 0.0
        avg_down = 0.0

        for line in lines:
            try:
                restup_up = re.search(r'[0]{1}/0', line).group()
                if restup_up:
                    num = re.search(r'\d{2}\.\d{3}', line).group()
                    # print(restup_up)
                    p_start.append(num)

            except Exception as e:
                pass

        for line in lines:
            try:
                reset_dowm = re.search(r'[1-9]{1,2}/0', line).group()
                if reset_dowm:
                    num = re.search(r'\d{2}\.\d{3}', line).group()
                    p_end.append(num)

            except Exception as e:
                pass

        del p_start[0]
        print(p_start)
        print(p_end)

        for i in p_start:

            sum_up = sum_up + float(i)

        for i in p_end:
            sum_down = sum_down + float(i)

        print('The sum above "reset" is:%f' % sum_up)
        print('The sum under "reset" is:%f' % sum_down)

        avg_up = sum_up / len(p_start)
        print('The average value above "reset" is:%f' % avg_up)

        avg_down = sum_down / len(p_end)
        print('The average value under "reset" is:%f' % avg_down)


def kill_process_by_name():
    with open ('sos.10.15.4.txt','r')as kill_p:
        lines = kill_p.readlines
        for line in lines:
            try:
               proc_top = re.search(r'10/0]').group()
               if
            except Exception as e:
                pass


def main():
    pass


if __name__ == '__main__':
    calc()
