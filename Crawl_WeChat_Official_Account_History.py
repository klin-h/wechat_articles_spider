#!/usr/bin/env python
# -*- coding: utf-8 -*-

from wechat_mp_crawler import WechatArticleManager, read_accounts_from_excel

def main():
    """
    从Excel中读取公众号列表，并爬取这些公众号的最近文章
    """
    # 初始化管理器
    manager = WechatArticleManager()
    
    # 从Excel文件读取公众号列表
    accounts_file = "accounts.xlsx"
    account_list = read_accounts_from_excel(accounts_file)
    
    if not account_list:
        print(f"从 {accounts_file} 中未读取到有效的公众号名称，程序终止")
        return
    
    print(f"待爬取的公众号列表: {', '.join(account_list)}")
    
    # 设置参数
    articles_per_account = 15  # 每个公众号获取的文章数量(最大值)
    days = 2  # 获取最近几天的文章 (默认为2，即今天和昨天)
    
    # 执行爬取操作
    success, articles = manager.crawl_multiple_accounts(
        nickname_list=account_list,
        articles_per_account=articles_per_account,
        days=days
    )
    
    if success:
        print(f"成功完成爬取！获取到 {len(articles)} 篇文章")
    else:
        print("爬取过程中出现问题，请检查日志")


if __name__ == "__main__":
    main()
