#!/usr/bin/python
#coding=utf-8
#python 2.7
#删除当前目录下所有 ~ 结尾的文件
#运行python run.py



import string
from string import Template
import os
import sys
import stat
import shutil
import re

class TemplateIgnoreInvalid(Template):
    # override pattern to make sure `invalid` never matches
    pattern = r"""
    %(delim)s(?:
      (?P<escaped>%(delim)s) |   # Escape sequence of two delimiters
      (?P<named>%(id)s)      |   # delimiter and a Python identifier
      {(?P<braced>%(id)s)}   |   # delimiter and a braced identifier
      (?P<invalid>^$)            # never matches (the regex is not multilined)
    )
    """ % dict(delim=re.escape(Template.delimiter), id=Template.idpattern)

#@rootdir 目录
#@filterfunc 筛选文件参数
#@func 文件具体处理函数
def walk_func(rootdir, func, filterfunc):
    for parent, _, filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字

        for filename in filenames:                        #输出文件信息
            if filterfunc(filename):
                func(os.path.join(parent,filename))


def filter_template(filename):
    if filename.endswith('~'):
        return True
    else:
        return False

def handle_file(filename):
    os.remove(filename)



def get_local_ip(ifname):
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    inet = fcntl.ioctl(s.fileno(), 0x8915, struct.pack('256s', ifname[:15]))
    ret = socket.inet_ntoa(inet[20:24])
    return ret



def construct_cfg_file():
    walk_func("./", handle_file, filter_template)


if __name__ == '__main__':
    #跳转目录
    os.chdir("../game_platform")
    print "start construct cfg"
    construct_cfg_file()
    print "construct cfg succ"
