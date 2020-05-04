#!/usr/bin/env
# coding=utf-8
import time
import pandas as pd
# 爬虫模块
from selenium import webdriver
from selenium.webdriver.chrome.options import Options  # 导入chrome选项
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# OS模块
import os
import shutil
# ui模块
import tkinter as tk
from tkinter import filedialog

start=time.perf_counter()
print("->")
print("->开始执行，请等待")
print("->请选择要处理的EXCEL文件")
print("->")

application_window = tk.Tk()
# 设置文件对话框会显示的文件类型
my_filetypes = [('all files', '.*'), ('text files', '.txt')]
# 请求选择文件
filename = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="请选择你要处理的委托明细",
                                    filetypes=my_filetypes)
foldname=filename.split("/")[-1]
foldname=foldname.split(".")[0]
local='C:\\'+foldname
local1='C:\\'+foldname+'\\ScreenShot'
if os.path.exists(local) is False:
    os.mkdir(local)
if os.path.exists(local1) is False:
    os.mkdir(local1)
else:
    shutil.rmtree(local1, True)
    os.mkdir(local1)

end = time.perf_counter()
duration = round(end - start, 3)
print("->初始化目录，耗时:", duration)
print("->")

# 读表,拆分列
IPdata = pd.read_excel(filename,index_col=0)
df = IPdata[u'新委托来源'].str.split(";",expand=True)
IPdata['站点编号'] = df[0].apply(lambda x: x.replace("S:","") if 'S:' in x else None)
IPdata['手机号码'] = df[1].apply(lambda x: x.replace("MPN:","") if 'MPN:' in x else None)
IPdata['手机IP'] = df[2].apply(lambda x: x.replace("MIP:","") if 'MIP:' in x else None)
IPdata['手机串号'] = df[3].apply(lambda x: x.replace("IMEI:","") if 'IMEI:' in x else None)
IPdata['IP'] = df[1].apply(lambda x: x.replace("IIP:","")  if 'IIP:' in x else None)
IPdata['MAC'] = df[2].apply(lambda x: x.replace("MAC:","")  if 'MAC:' in x else None)
IPdata['硬盘序列号'] = df[3].apply(lambda x: x.replace("HD:","")  if 'HD:' in x else None)
IPdata['IP归属地']=None

end=time.perf_counter()
duration=round(end-start,3)
print("->分列完成，耗时:",duration)

#汇聚取唯一值
frame=IPdata[['手机号码', '手机IP', '手机串号','IP','MAC','硬盘序列号','IP归属地']]
frame=frame.drop_duplicates(subset=['手机号码', '手机IP', '手机串号','IP','MAC','硬盘序列号','IP归属地'])

#合并IP列，生成cache表
cache=pd.DataFrame()
cache['ALLIP']=frame['手机IP'].apply(lambda x:x if x is not None else '')+frame['IP'].apply(lambda x:x if x is not None else '')
cache=cache.drop_duplicates(subset=['ALLIP'])
cache['IP归属地']=None
cache.reset_index(drop=True, inplace=True)

end = time.perf_counter()
duration = round(end-start,3)
print("->聚合完成，耗时:",duration)
print("->")
print("->准备启动GOOGLE浏览器")
print("->")
#开始爬虫
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--log-level=3")  # fatal
chrome_options.add_argument('disable-infobars')
chrome_options.add_argument('--headless') # 设置chrome浏览器无头模式
chrome_options.add_argument('--disable-gpu') # 如果不加这个选项，有时定位会出现问题
# chrome_options.add_argument('--user-data-dir=C:\\Users\\wulf\\AppData\\Local\\Google\\Chrome\\User Data\\Default') #设置成用户自己的数据目录
chrome_options.add_experimental_option( "prefs",{'profile.managed_default_content_settings.javascript': 2})
chrome_path = 'chromedriver.exe'
driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
driver.get("https://www.baidu.com/s?wd=169.235.24.133")
WebDriverWait(driver, 20, 0.2).until(EC.title_contains('169.235.24.133'))
end = time.perf_counter()
duration = round(end-start,3)
print("->开始百度，耗时:",duration)
print("->")
for i in cache.index:
    col1 = cache.loc[i,'ALLIP']
    if  len(col1)>6 and  col1.count('.') == 3:
        inputwd = driver.find_element_by_name("wd")  # 搜索输入文本框的name属性值
        but = driver.find_element_by_xpath('//input[@type="submit"]')  # 搜索提交按钮
        inputwd.clear()  # 清楚文本框里的内容
        inputwd.send_keys(col1)  # 输入IP关键词‘ ’
        but.send_keys(Keys.RETURN)  # 输入回车键  but.click()  #点击按钮
        try:
            WebDriverWait(driver, 20, 0.5).until(EC.title_contains(col1))
            data1 = driver.find_element_by_class_name("op-ip-detail").text
            data1 = data1.replace(col1, '')
            col2 = data1.replace("IP地址: ", "")
            cache.loc[i,u'IP归属地'] = col2
            result = driver.find_element_by_xpath("//span[@class='c-gap-right']/../../../../../../../..")
            driver.execute_script('arguments[0].scrollIntoViewIfNeeded();', result)  # 将可视区域滑动到元素所在区域
            result.screenshot(local+"\\ScreenShot\\"+col1+".png")
            print("-" + col1 + " | " + col2)
        finally:
            pass
    else:
        print("检测到异常IP格式") #IP长度和字符不符合常规
        pass
driver.quit()

end=time.perf_counter()
duration=round(end-start,3)
print("->")
print("->解析完成，耗时:",duration)
print("->开始匹配")

#UPDATE归属地数据
a = IPdata.set_index('手机IP')
b = frame.set_index('手机IP')
c = cache.set_index('ALLIP')
a.update(c)
b.update(c)
IPdata['IP归属地'] = a[u'IP归属地'].values.astype(str)
frame['IP归属地'] = b[u'IP归属地'].values.astype(str)
a = IPdata.set_index('IP')
b = frame.set_index('IP')
c = cache.set_index('ALLIP')
a.update(c)
b.update(c)
IPdata['IP归属地'] = a[u'IP归属地'].values.astype(str)
frame['IP归属地'] = b[u'IP归属地'].values.astype(str)

print("->")
print("->匹配完成，耗时:",duration)
print("->开始存盘")

#数据表存盘
IPw = pd.ExcelWriter(local+'\\委托明细.xlsx',index=False)
IPdata.to_excel(IPw,sheet_name='委托明细')
frame.to_excel(IPw,sheet_name='汇总结果')
cache.to_excel(IPw,sheet_name='解析结果')
IPw.save()
IPw.close()
os.system("start explorer "+local)

end=time.perf_counter()
duration=round(end-start,3)
print("\n->已完成，程序耗时:",duration)