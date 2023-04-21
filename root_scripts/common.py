# -*- coding: utf-8 -*-
import os
import sys
import logging 
import time
import subprocess
import re
from datetime import datetime
libpath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if not libpath in sys.path:
    sys.path.append(libpath)
print(libpath)
import random
#from mini_test.configs import AppConfig
from functools import wraps
import tkinter as tk
from tkinter import filedialog   #文件对话框模块
import configparser
import time
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By   #By类：定位元素

def timethis(func):
    """
    Decorator that reports the execution time.
    """
    @wraps(func)    #在代码运行期间动态增加功能的方式，称之为“装饰器”（Decorator）
    def wrapper(*args, **kwargs):   #*args表示任何多个无名参数，它是一个tuple元祖  **kwargs表示关键字参数，它是一个dict字典
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        return result
    return wrapper

class UIParser():
    @staticmethod
    def nest(self,func):
        def wrapper(*args,**kwargs):
            func(args)
        return wrapper
    @staticmethod
    def run(obj,params,exceptfunc = None):
        device = obj #if isinstance(obj,Device) else obj.device
        #isinstance(obj, Device)如果参数obj是Device的实例，或者obj是Device类的子类的一个实例， 返回True。如果obj不是一个给定类型的的对象， 则返回结果总是False  >>> isinstance(1.0, float) true
        def param_parser(param):
            if isinstance(param, dict):
                for k,v in param.items():
                    if v == None:
                        param.pop(k)
            else:
                for v in param:
                    if v == None:
                        param.remove(v)
       
        def error(param):
            if param.has_key("assert") and param['assert'] == False:                       
                return False
            else:
                print ("%s error!"%param)
                exceptfunc() if (exceptfunc) else None
                return True 
                                
        def listfoo(param):
            resault = True
            if isinstance(param["content"],list):
                for content in param["content"]:
                    param_tmp = param
                    param_tmp["content"]=content
                    resault = resault and listfoo(param)
            elif param["id"] == "meta":
                resault = resault and getattr(obj,param["content"])(*param["action"]["param"] if param.has_key("action") and param["action"].has_key("param") else [])
            else:
                print (param)
                if param_parser(param["id"])=={}:
                    return True
                select = device(**{param["id"]:param["content"]})
                #action=select.wait.exists(timeout = 5000) if ((not param.has_key("wait")) or ((param.has_key("wait") and param["wait"]))) else select.wait.exists(timeout = int(param["wait"]))
                action=select.wait.exists(timeout = 5000) if not param.has_key("wait") else select.wait.exists(timeout = int(param["wait"]))
                if action and not (param.has_key("action") and param["action"]==None):
                    getattr(select,"click")(None) if not param.has_key("action") else getattr(select,param["action"]["type"])(*param["action"]["param"] if param["action"].has_key("param") else [])
                    time.sleep(param["action"]["delay"] if param.has_key("aciton") and param["action"].has_key("delay") else 0)
                resault = resault and action
            return resault
        
        def dictfoo(param):
            resault = True    
            if not param.has_key("id"):
                return False   
            if isinstance(param["id"],list):
                for content in param["id"]:
                    param_tmp = param
                    param_tmp["id"]=content
                    if param["id"].has_key("action"):
                        param_tmp["action"]=param["action"]
                    resault = resault and listfoo(param)
            elif param["id"].has_key("meta"):
                resault = resault and getattr(obj,param["id"]["meta"])(*param["action"]["param"] if param.has_key("action") and param["action"].has_key("param") else [])
            else:
                if param_parser(param["id"])=={}:
                    return True
                select = device(**param["id"])
                #action=select.wait.exists(timeout = 5000) if ((not param.has_key("wait")) or ((param.has_key("wait") and param["wait"]))) else select.wait.exists(timeout = int(param["wait"]))    
                action=select.wait.exists(timeout = 5000) if not param.has_key("wait") else select.wait.exists(timeout = int(param["wait"]))
                if action and not (param.has_key("action") and param["action"]==None):
                    getattr(select,"click")(None) if not param.has_key("action") else getattr(select,param["action"]["type"])(*param["action"]["param"] if param["action"].has_key("param") else [])
                    time.sleep(param["action"]["delay"] if param.has_key("aciton") and param["action"].has_key("delay") else 0)
                resault = resault and action
            return resault
        
        for param in params:
            if isinstance(param,list):
                UIParser.run(obj,param)
            else:
                if param.has_key("id") and isinstance(param["id"],dict):
                    if not dictfoo(param):
                        if (error(param)):
                            return False
                elif param.has_key("id") and param.has_key("content") and not isinstance(param["id"],dict):
                    if not listfoo(param):
                        if (error(param)):
                            return False
        return True

def create_folder():
    """Create folder to save pic & log.     
    Return a folder path or None
    Exception: OSError
    """
    log_path = os.environ.get("LOG_PATH")
    if log_path is None:
        log_path =  sys.path[0][sys.path[0].find(':')+1:] + '\\results'
    if not os.path.exists(log_path):
        # logger.debug("log_path not exsit,create it")
        #logger.debug(u"log_path不存在".encode('utf8'))
        os.makedirs(log_path)
    if not os.path.exists(log_path):
        return None
    return log_path

def createlogger(name):
    """Create a logger named specified name with the level set in config file.
    return a logger
    """
#    config = GetConfigs("common")    #构造一个config类
    lev_key = 'DEBUG'   #用getstr函数获取到LOG_FITER的值，进而控制log打印的内容
    lev_dict = {"DEBUG": logging.DEBUG, "INFO": logging.INFO,
                "WARNING": logging.WARNING, "ERROR": logging.ERROR,
                "CRITICAL": logging.CRITICAL}
    logger = logging.getLogger(name)    #初始化一个日志对象
    logger.setLevel(lev_dict[lev_key])   #设置日志等级，一旦设置了日志等级，则调用比等级低的日志记录函数则不会输出
                                #日志级别大小关系为：CRITICAL > ERROR > WARNING > INFO > DEBUG > NOTSET
    ch = logging.StreamHandler()    #日志输出到流，可以是sys.stderr、sys.stdout或者文件 就是打印到屏幕
    t = time.strftime("%Y_%m_%d", time.localtime())
    file_name = 'results/log_' + t + '.log'
    # if file_name != '':
    fh = logging.FileHandler(file_name, encoding='utf-8')   #日志输出到文件,在windows下面，新文件的默认编码是gbk，这样的话，python解释器会用gbk编码去解析我们的网络数据流txt，然而txt此时已经是decode过的unicode编码，这样的话就会导致解析不了，出现上述问题。 解决的办法就是，改变目标文件的编码
    #.%(msecs)03d,'%y-%m-%d %H:%M:%S'
    formatter = logging.Formatter(
        '%(asctime)s: [%(levelname)s] [%(name)s] [%(funcName)s] %(message)s'
        )    #生成log格式定义，依次为log创建时间，保留3位的毫秒时间，输出的等级标签，logger的名称，调用的函数名，log的提示信息
    ch.setFormatter(formatter)   #设置输出流的格式为上面定义的格式
    logger.addHandler(ch)   #把设置好的ch添加到logger中
    # if file_name != '':
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger

logger = createlogger("COMMON")    #定义完函数直接生成一个logger对象

def log_traceback(traceback):
    """print traceback information with the log style.
    """
    str_list = traceback.split("\n")
    for string in str_list:
        logger.warning(string)

def random_name(index_num):
    """#为联系人随机名字"""
    numseed = "0123456789"
    logger.debug('Create a random name.')
    sa = []
    for i in range(5):
        sa.append(random.choice(numseed))
    stamp = ''.join(sa)
    strname = 'Autotest%02d_' %(index_num+1) +stamp
    return strname

class Common(object):     #用于初始化设备
    """Provide common functions for all scripts."""  
    def __init__(self, mod='POP_TEST', timeout = 5000):
        self.timeout = timeout
        self.logger = createlogger(mod)
        self.log_path = create_folder()
        sys.stdout = output_print()
        #self.config = GetConfigs("common")   #初始化GetConfigs类
        #self.appconfig = AppConfig("appinfo")  #自动初始化类AppConfig，同时初始化了Configs类，预先找到配置文件路劲"configure","%s.ini" 名字为appinfo

    # def createlogger(self, name):
    #     """Create a logger named specified name with the level set in config file.
    #     return a logger
    #     """
    #     #    config = GetConfigs("common")    #构造一个config类
    #     lev_key = 'DEBUG'  # 用getstr函数获取到LOG_FITER的值，进而控制log打印的内容
    #     lev_dict = {"DEBUG": logging.DEBUG, "INFO": logging.INFO,
    #                 "WARNING": logging.WARNING, "ERROR": logging.ERROR,
    #                 "CRITICAL": logging.CRITICAL}
    #     logger = logging.getLogger(name)  # 初始化一个日志对象
    #     logger.setLevel(lev_dict[lev_key])  # 设置日志等级，一旦设置了日志等级，则调用比等级低的日志记录函数则不会输出
    #     # 日志级别大小关系为：CRITICAL > ERROR > WARNING > INFO > DEBUG > NOTSET
    #     ch = logging.StreamHandler()  # 日志输出到流，可以是sys.stderr、sys.stdout或者文件 就是打印到屏幕
    #     t = time.strftime("%Y_%m_%d", time.localtime())
    #     file_name = 'results/log_' + t + '.log'
    #     # if file_name != '':
    #     fh = logging.FileHandler(file_name)  # 日志输出到文件
    #     # .%(msecs)03d,'%y-%m-%d %H:%M:%S'
    #     formatter = logging.Formatter(
    #         '%(asctime)s: [%(levelname)s] [%(name)s] [%(funcName)s] %(message)s'
    #     )  # 生成log格式定义，依次为log创建时间，保留3位的毫秒时间，输出的等级标签，logger的名称，调用的函数名，log的提示信息
    #     ch.setFormatter(formatter)  # 设置输出流的格式为上面定义的格式
    #     logger.addHandler(ch)  # 把设置好的ch添加到logger中
    #     # if file_name != '':
    #     fh.setFormatter(formatter)
    #     logger.addHandler(fh)
    #     return logger

    def select_file(self):
        root=tk.Tk()
        root.withdraw()
        filepath=filedialog.askopenfilename()  #获取文件名
        return filepath

    def iselementexist_by_xpath(self, browser, element):
        flag=True
        try:
            browser.find_element_by_xpath(element)
            return flag
        except:
            flag=False
            return flag

    def iselementexist_by_classname(self, browser, element):
        flag=True
        try:
            browser.find_element_by_class_name(element)
            return flag
        except:
            flag=False
            return flag

    def iselementexist_by_id(self, browser, element):
        flag=True
        try:
            browser.find_element_by_id(element)
            return flag
        except:
            flag=False
            return flag

    def iselementexist_by_name(self, browser, element):
        flag=True
        try:
            browser.find_element_by_name(element)
            return flag
        except:
            flag=False
            return flag

    def wait_until_not_exist_by_id(self, browser, element, timeout=10):
        """显示等待ID名的控件消失
        关于显示等待：WebDriverWait   WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'wrapper')))
            presence_of_element_located可以替换为
            title_is
            title_contains
            presence_of_element_located
            visibility_of_element_located
            visibility_of
            presence_of_all_elements_located
            text_to_be_present_in_element
            text_to_be_present_in_element_value
            frame_to_be_available_and_switch_to_it
            invisibility_of_element_located
            element_to_be_clickable - it is Displayed and Enabled.
            staleness_of
            element_to_be_selected
            element_located_to_be_selected
            element_selection_state_to_be
            element_located_selection_state_to_be
            alert_is_present
            By.ID中的ID可替换为'CLASS_NAME', 'CSS_SELECTOR', 'ID', 'LINK_TEXT', 'NAME', 'PARTIAL_LINK_TEXT', 'TAG_NAME', 'XPATH"""
        st = time.time()
        try:
            result = WebDriverWait(browser, timeout, 0.5).until_not(EC.presence_of_element_located((By.ID, element)))
            if result:
                et = time.time()
                ut = round(et - st, 2)
                print("wait element disappear %s seconds." % ut)
                return True
            return False
        except TimeoutException:
            return False

    def wait_until_exist_by_id(self, browser, element, timeout=10):
        """显示等待ID名的控件出现"""
        st = time.time()
        try:
            result = WebDriverWait(browser, timeout, 0.5).until(EC.presence_of_element_located((By.ID, element)))
            if result:
                et = time.time()
                ut = round(et-st, 2)
                print("wait element %s seconds." % ut)
                return True
            return False
        except TimeoutException:
            return False

    def wait_until_not_exist_by_classname(self, browser, element, timeout=10):
        """显示等待class名的控件消失"""
        try:
            result = WebDriverWait(browser, timeout, 0.5).until_not(EC.presence_of_element_located((By.CLASS_NAME, element)), message = "%s is not disappeared"%element)
            if result:
                return True
            return False
        except TimeoutException:
            return False
        except:
            self.logger.warning('selenium.common.exceptions.WebDriverException: Message: target frame detached')
            return True

    def wait_until_exist_by_classname(self, browser, element, timeout=10):
        """显示等待class名的控件出现"""
        st = time.time()
        try:
            result = WebDriverWait(browser, timeout, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME, element)))
            if result:
                et = time.time()
                ut = round(et - st, 2)
                print("wait element %s seconds." % ut)
                return True
            return False
        except TimeoutException:
            return False

    def wait_until_not_exist_by_xpath(self, browser, element, timeout=10):
        """显示等待xpath名的控件消失"""
        try:
            result = WebDriverWait(browser, timeout, 0.5).until_not(EC.presence_of_element_located((By.XPATH, element)))
            if result:
                return True
            return False
        except TimeoutException:
            return False

    def wait_until_exist_by_xpath(self, browser, element, timeout=10):
        """显示等待xpath名的控件出现"""
        st = time.time()
        try:
            result = WebDriverWait(browser, timeout, 0.5).until(EC.presence_of_element_located((By.XPATH, element)))
            if result:
                et = time.time()
                ut = round(et - st, 2)
                print("wait element %s seconds." % ut)
                return True
            return False
        except TimeoutException:
            return False

    def scroll_to_element_by_class(self, browser, element):
        save_b = browser.find_element_by_class_name(element)
        browser.execute_script("arguments[0].scrollIntoView();", save_b)
        self.logger.info("Scroll the view to %s" % element)

    def scroll_to_element_by_id(self, browser, element):
        save_b = browser.find_element_by_id(element)
        browser.execute_script("arguments[0].scrollIntoView();", save_b)
        self.delay(1)
        self.logger.info("Scroll the view to %s" % element)

    def scroll_to_head(self, browser):
        """滚动回页面头部"""
        browser.execute_script("window.scrollTo(0,0);")
        self.delay(1)
        self.logger.info("Scroll the view to the head")

    def check_time(self, set_time='22'):
        """判断当前时间,当到达匹配时间返回True"""
        t = time.strftime("%H", time.localtime())
        print("It's %s o'clock." %t)
        if t == set_time or t == '0' + set_time:
            return True
        return False

    def get_current_date(self):
        """输出当前年月日"""
        t = time.strftime("%Y-%m-%d", time.localtime())
        return t

    def get_file_path(self):
        t = time.strftime("%Y_%m_%d", time.localtime())
        filename = 'log_' + t + '.log'
        return filename

    def read_config(self, *args):
        """根据需要的信息，获取配置文件中的站点对应的税率，折扣名+对应折扣，积分配置，返回对应的字典"""
        path=os.path.dirname(os.path.abspath(__file__))
        config_path=path+'/config.ini'
        print(config_path)
        config=configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
        need_return=[]
        sections= config.sections()  #拿到所有的section
        if len(config)>0:
            for arg in args:
                arg = arg.lower()
                if arg in sections:
                    self.logger.info("Find %s items, output!" % arg)
                    site_list = config.items(arg)
                    site_dict = dict(site_list)
                    need_return.append(site_dict)
                else:
                    self.logger.debug("Can not find section named:%s" % arg)
            return_len=len(need_return)
            self.logger.info("Totally find %s dicts, return them!"%return_len)
            if return_len == 1:
                return need_return[0]
            elif return_len == 2:
                return need_return[0], need_return[1]
            elif return_len == 3:
                return need_return[0], need_return[1], need_return[2]
        else:
            self.logger.warning("配置文件为空!")
            return False

    def screenshot(self, browser, name=''):
        path = os.path.dirname(os.path.dirname(__file__)) + '/results/'
        nowtime=time.strftime("%Y_%m_%d_%H_%M_%S")
        file_name=path+nowtime+"_"+name+'.png'
        browser.get_screenshot_as_file(file_name)
        self.logger.debug("Save screenshot:%s" % file_name)

    def get_current_page_width_and_height(self, browser):
        """返回当前页面的宽和高"""
        height = browser.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
        width = browser.execute_script("return document.documentElement.scrollWidth")
        return width, height

    def delay(self, waittime=1):
        time.sleep(waittime)

    def get_value_in_string(self, string_value=''):
        """获取字符串中纯数字内容"""
        result = ''.join(filter(str.isdigit, string_value))
        return int(result)

class output_print(object):
    def __init__(self, filename="result.log"):
        t=time.strftime("%Y_%m_%d", time.localtime())
        filename='log_' + t + '.log'
        self.terminal = sys.stdout
        if not os.path.isfile('results/'+filename):
            self.log = open('results/'+filename, "w+", encoding='utf-8')
            print("Create log file:results/%s" % filename)
        else:
            self.log = open('results/'+filename, "a", encoding='utf-8')
            print("Open log file:results/%s" % filename)

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()  #实时保存

    def flush(self):
        pass



             
if __name__ == "__main__":
    a = Common("web_test")
#    a.start_app("Sound Recorder")
                
                