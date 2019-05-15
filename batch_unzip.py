# -*- coding: utf-8 -*-
"""
Created on Sun Jun 11 16:43:17 2017

@author: manna
"""

import os
import time
import threading
import subprocess
from loggingutil import getlogging

logger = getlogging().getLogger('batch_unzip.py')

def find(name, *types):
    for root, dirs, files in os.walk(name):
        for f in files:  
            if os.path.isfile(os.path.join(root, f)) and os.path.splitext(f)[1][1:] in types:
                print(os.path.join(root, f))


def unzip(path, *types):

    for root, dirs, files in os.walk(path):
        for f in files:
            logger.debug('aaaaaaaaaaa')
            if os.path.isfile(os.path.join(root, f)) and os.path.splitext(f)[1][1:] in types:
                # print os.path.join(root,f)
                unzip_command = u'7z.exe x -y {} "-o{}\\*" -r'.format(os.path.join(root, f), unzip_dir)
                # unzip_command= u'"unzip.exe"  {} -d {}*'.format(os.path.join(root,f), unzip_dir)
                logger.debug(unzip_command)
                # error
                logger.debug(os.system(unzip_command))

                if os.system(unzip_command):
                    logger.debug('中文233333333333333322')
                    if subprocess.call(unzip_command, shell=True):
                        print(u'successful!!!')


if __name__ == '__main__':  
    work_dir = u"data"
    unzip_dir = u'unzip'   # 创建解压缩文件存储路径
    t1 = time.time()
    # 加入线程，搜索D盘 以.sql、.zip结尾的文件
    t = threading.Thread(target=unzip, args=(work_dir, "7z", "zip"))
    t.start()  
    t.join()  
    # 计算执行时间
    logger.info(time.time()-t1)
