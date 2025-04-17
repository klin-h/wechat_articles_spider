#!/usr/bin/env python
# -*- coding: utf-8 -*-

from wechat_mp_crawler import WechatArticleManager

def main():
    """
    爬取指定公众号的历史文章（最多100篇）
    """
    # 初始化管理器
    manager = WechatArticleManager()
    
    # 指定公众号名称
    target_account = "机器之心"  # 修改为您想爬取的公众号名称
    
    # 设置最大爬取文章数量
    max_articles = 10
    
    # 可选：指定输出文件名
    # output_file = "xxx_history.xlsx"
    
    print(f"准备爬取公众号 '{target_account}' 的历史文章（最多 {max_articles} 篇）")
    
    # 执行爬取操作
    success, articles = manager.crawl_account_history(
        nickname=target_account,
        max_articles=max_articles,
        # output_file=output_file  # 可选参数
    )
    
    if success:
        print(f"成功完成爬取！获取到 {len(articles)} 篇文章")
        
        # 打印前5篇文章信息
        print("\n前5篇文章:")
        for i, article in enumerate(articles[:5]):
            if i < len(articles):
                print(f"{i+1}. {article['title']} - 发布于 {article.get('publish_date', '未知日期')}")
                print(f"   链接: {article['link']}")
                print()
    else:
        print("爬取过程中出现问题，请检查日志")


if __name__ == "__main__":
    main()
