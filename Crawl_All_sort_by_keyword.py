#!/usr/bin/env python
# -*- coding: utf-8 -*-

from wechat_mp_crawler_with_head import WechatArticleManager

def main():
    """
    在指定公众号的文章中搜索关键词，并显示浏览器窗口
    """
    # 初始化管理器，设置headless=False显示浏览器窗口
    manager = WechatArticleManager(headless=False)
    
    # 指定公众号名称
    target_account = "同济大学"  # 修改为您想爬取的公众号名称
    
    # 设置最大爬取文章数量
    max_articles = 10
    
    # 设置关键词和权重
    keywords = ["人工智能", "数据科学", "程序设计"]  # 最多3个关键词
    weights = [1.5, 1.2, 1.0]  # 对应的权重，数量应与关键词相同
    
    print(f"准备在公众号 '{target_account}' 的文章中搜索关键词: {keywords}")
    print(f"关键词权重: {weights}")
    print("注意：浏览器窗口将会显示，请在出现二维码时完成扫码")
    
    # 执行搜索操作
    success, articles = manager.search_keywords_in_account(
        nickname=target_account,
        keywords=keywords,
        weights=weights,
        max_articles=max_articles
    )
    
    if success:
        print(f"成功完成搜索！共处理 {len(articles)} 篇文章")
        
        # 打印排名前5的文章
        print("\n排名前5的文章:")
        for i, article in enumerate(articles[:5]):
            if i < len(articles):
                print(f"{i+1}. {article['title']} - 分数: {article.get('keyword_score', 0)}")
                print(f"   关键词分布: {article.get('keyword_counts', {})}")
                print()
    else:
        print("搜索过程中出现问题，请检查日志")


if __name__ == "__main__":
    main()
