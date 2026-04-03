#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细调试GitHub版本检测问题
"""

import requests
import json

def debug_update_check():
    """详细调试更新检测功能"""
    print("=== GitHub版本检测详细调试 ===")
    
    # 测试多个可能的API地址
    api_urls = [
        "https://api.github.com/repos/244727418/shouhou/releases/latest",
        "https://api.github.com/repos/244727418/shouhou/releases",
        "https://api.github.com/repos/244727418/shouhou"
    ]
    
    for api_url in api_urls:
        print(f"\n🔍 测试API: {api_url}")
        
        try:
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': 'Debug-Update-Checker'
            }
            
            response = requests.get(api_url, headers=headers, timeout=10)
            print(f"   状态码: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                
                if 'releases' in api_url:
                    # releases列表
                    releases = data
                    print(f"   Release数量: {len(releases)}")
                    for i, release in enumerate(releases[:3]):  # 只显示前3个
                        tag = release.get('tag_name', '无标签')
                        name = release.get('name', '无标题')
                        print(f"   [{i+1}] Tag: {tag}, 标题: {name}")
                elif 'latest' in api_url:
                    # 最新release
                    tag = data.get('tag_name', '无标签')
                    name = data.get('name', '无标题')
                    print(f"   最新Release: {name} (Tag: {tag})")
                    
                    assets = data.get('assets', []) 
                    print(f"   附件数量: {len(assets)}")
                    for asset in assets:
                        print(f"     - {asset.get('name')}")
                else:
                    # 仓库信息
                    print(f"   仓库名: {data.get('name')}")
                    print(f"   描述: {data.get('description')}")
                    
            elif response.status_code == 404:
                print("   ❌ 404 - 资源不存在")
                print(f"   响应: {response.text[:200]}")
            else:
                print(f"   ❌ 错误: {response.status_code}")
                print(f"   响应: {response.text[:200]}")
                
        except Exception as e:
            print(f"   ❌ 异常: {type(e).__name__}: {e}")
    
    # 测试直接访问仓库页面
    print(f"\n🌐 仓库页面: https://github.com/244727418/shouhou")
    print(f"🌐 Release页面: https://github.com/244727418/shouhou/releases")
    
    # 检查网络连接
    print("\n🔧 网络连接测试:")
    try:
        test_response = requests.get("https://api.github.com", timeout=5)
        print(f"   GitHub API可访问: {test_response.status_code == 200}")
    except:
        print("   ❌ 无法连接到GitHub API")

if __name__ == "__main__":
    debug_update_check()