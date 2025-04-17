# wechat_articles_spider
For WeChat Official Accounts, a small-scale crawler program that doesn't use packet capture tools.
一个小规模并且不需要使用抓包工具的微信公众号爬虫程序，基本做到开袋即食
截止2025/4/17，包含了三种功能，分别是
1、对一系列公众号名爬取最近两天内更新的所有文章的标题、URL、更新时间，测试样例见Crawl_WeChat_Official_Account_History.py
2、对某个公众号名、爬取其所有历史文章的标题、URL、更新时间，测试样例见Crawl_WeChat_Official_Account_Search_byname.py
3、对某个公众号名、爬取其所有历史文章并从文章中查找由用户自定义的关键词（支持三个）同时支持为三个词语赋予不同的权重，最后每篇文章
得分 = 词语出现次数·权重 ，最后将所有文章按得分降序输出，测试样例见Crawl_All_sort_by_keyword.py

使用教学（新手向）

使用前准备
安装依赖库
在vs code 等编译器终端粘贴这段代码并回车
pip install selenium webdriver-manager pandas wechatarticles
需要下载谷歌浏览器，如有则跳过这一步（ https://www.google.com/chrome/）
注册一个微信公众号账户，注册链接https://mp.weixin.qq.com/ （进入后注册一个新公众号账户，切记不要用官方的账号！！！，因为账号可能被封禁）
文件格式

一、功能一使用教学
程序需要一个名为accounts.xlsx的Excel文件，包含要爬取的公众号名称，格式如下（本文件夹已经包括）
nickname
机器之心
新华网
...

如果文件不存在，程序会自动创建一个示例文件。
如后期需要对查找的公众号进行增删，直接在这个Excel文件中nickname这一列添加新的名称（或删除 注：删除时要选择下方单元格上移，不要留空行）
使用方法
在vscode Crawl_WeChat_Official_Account_History.py 之后代码会执行以下过程
1.检查是否有保存的token和cookie，有效则直接使用（类似于检查身份验证，在有效期内则直接登录，否则跳转至chrome浏览器，需要微信扫码验证登录）
2.读取程序自动读取accounts.xlsx中的公众号列表
3.对每个公众号爬取当天发布的文章
4.将结果保存到以当天日期命名的Excel文件中
输出结果
程序会生成以下文件:
1.weixin_credentials.py: 保存token和cookie信息供下次运行使用（不要删除）
2.X月X号wechat_articles.xlsx: 包含所有爬取文章的Excel文件，包括以下信息: 
首行统计信息(需要爬取的公众号总数、当日有更新的数量、当日未更新的数量)
公众号名称(nickname)
文章标题(title)
文章链接(link)
发布时间(publish_time)

二、功能二使用教学
直接在代码中修改
# 指定公众号名称
    target_account = "机器之心"  # 修改为您想爬取的公众号名称
然后设置希望爬取数量（建议小于100，减少封号风险）
# 设置最大爬取文章数量
    max_articles = 10
    
运行Crawl_WeChat_Official_Account_Search_byname.py
运行过程中的问题，具体见 一、功能一使用教学 部分
输出结果包含
公众号名称(nickname)
文章标题(title)
文章链接(link)
发布时间(publish_time)

三、功能三使用教学
基本类似于二了，在代码内部修改要爬取的公众号名称、希望爬取的关键词、权重
.....
运行Crawl_All_sort_by_keyword.py
最后将所有文章具体信息按得分降序输出


四、注意事项
1.运行的时候关闭VPN，否则可能导致访问超时
2.爬取频率受到严格控制，使用随机延迟(5-15秒)减少被检测风险
3.如果公众号最近没有发文，不会在结果中显示该公众号的任何文章
4.程序运行过程可能较慢，请耐心等待
5.如遇到验证码或其他安全检查，请手动处理
6.不要短时间内多次运行程序


可能的问题与解决方案

1.极个别数公众号延迟比较严重，可能没法获取最新的文章（当天只能爬取到昨天的文章），目前仅发现“同伙管”公众号有这个问题，建议手动检查
2.无法登录: 确保VPN已关闭，再次尝试。仍无法登录。删除weixin_credentials.py再次运行程序，本次会扫码登录，若还是无法登录，可能是账号被封禁，更换账号登录
3.无法获取文章: 检查公众号名称是否正确，凭证是否有效
4.被封禁: 重新注册一个微信公众号
5.没有找到当天文章: 这可能是因为公众号最近确实没有发文，请检查公众号最新状态
