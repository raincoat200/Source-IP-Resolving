# Source IP Resolving(金仕达柜台委托来源IP解析)
实现功能：

1. 处理委托明细，分列，汇总唯一性。支持识别普通柜台、两融柜台。
2. 兼容网上委托、手机委托。（异常委托来源来不及过滤处理，抱歉。）
3. 支持国内外IP自动百度解析地址。
4. 支持解析结果截屏上传合规系统。
5. 程序生成文件以xls文件名自动归类。
    
    
#初始化程序运行环境：   

##1. python 3.8.2 代码解析器
[链接](https://www.python.org/downloads/release/python-382/)

> * 64bit选Windows x86-64 executable installer[下载](https://www.python.org/ftp/python/3.8.2/python-3.8.2-amd64.exe)
>* [32bit](https://www.python.org/ftp/python/3.8.2/python-3.8.2.exe)
>

**安装提示：**   
* 给所有用户安装   
* 添加系统PATH变量 
* [参考教程](https://www.liaoxuefeng.com/wiki/1016959663602400/1016959856222624)
-----------------------------------------------------------------------------------------
##2. pip环境
 
**cmd窗口下运行pip更新**   
    `python -m pip install --upgrade pip`   
    `pip install --upgrade -i https://pypi.douban.com/simple moudle_name` 

**导入三方库**
`
    pip install wheel pandas xlrd openpyxl selenium
`

-----------------------------------------------------------------------------------------

##3. google chrome爬虫浏览器 

本人使用的是最新版，版本 81.0.4044.113（正式版本） （64 位），对应附件的chromedriver.exe 
chromedriver的版本一定要与Chrome的版本一致，不然就不起作用。 
如果版本不一致，建议参考教程链接下载对应版本 
[教程](https://www.cnblogs.com/lfri/p/10542797.html) [下载链接](http://chromedriver.storage.googleapis.com/index.html)

-----------------------------------------------------------------------------------------


##4. 开始运行代码吧！
> bat为快捷运行窗口    
> py文件为源代码  
> demo.xls为金仕达（普通或两融）导出的委托明细文件，不需要做任何处理。    

双击bat批处理程序 -> 打开你要处理的xls -> 一切如你所见。
