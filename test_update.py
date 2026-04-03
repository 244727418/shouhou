#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GitHub版本检测功能
"""

import requests

# 配置信息
CURRENT_VERSION = "1.4"
GITHUB_API_URL = "https://api.github.com/repos/244727418/shouhou/releases/latest"

def test_update_check():
    """测试更新检测功能"""
    print("=== GitHub版本检测测试 ===")
    print(f"当前版本: {CURRENT_VERSION}")
    print(f"API地址: {GITHUB_API_URL}")
    print()
    
    try:
        # 发送GitHub API请求
        headers = {
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'Test-Update-Checker'
        }
        
        print("正在连接GitHub API...")
        response = requests.get(GITHUB_API_URL, headers=headers, timeout=10)
        
        print(f"状态码: {response.status_code}")
        
        if response.status_code == 200:
            release_data = response.json()
            
            # 获取最新版本号（去掉v前缀）
            latest_version = release_data.get('tag_name', '').lstrip('v')
            
            print(f"GitHub最新版本: {latest_version}")
            print(f"Release标题: {release_data.get('name', '无标题')}")
            
            # 检查是否有可执行文件
            assets = release_data.get('assets', [])
            print(f"附件数量: {len(assets)}")
            
            for asset in assets:
                name = asset.get('name', '')
                print(f"  - {name}")
            
            # 比较版本号
            if latest_version:
                if latest_version > CURRENT_VERSION:
                    print("✅ 检测到新版本！")
                    print(f"   当前版本: {CURRENT_VERSION}")
                    print(f"   最新版本: {latest_version}")
                elif latest_version == CURRENT_VERSION:
                    print("ℹ️ 当前已经是最新版本")
                else:
                    print("❓ 版本号异常")
            else:
                print("❌ 无法获取版本号")
                
        elif response.status_code == 404:
            print("❌ GitHub Release不存在")
            print("   请访问 https://github.com/244727418/shouhou/releases")
            print("   创建Tag为 v1.4.1 的Release")
            
        else:
            print(f"❌ API请求失败: {response.status_code}")
            print(f"   响应内容: {response.text[:200]}")
            
    except Exception as e:
        print(f"❌ 测试失败: {type(e).__name__}: {e}")

if __name__ == "__main__":
    test_update_check()