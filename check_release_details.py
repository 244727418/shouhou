#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查GitHub Release状态
"""

import requests
import time

def check_release_details():
    """详细检查Release状态"""
    print("=== GitHub Release详细状态检查 ===")
    
    # 测试多个API端点
    endpoints = [
        ("最新Release", "https://api.github.com/repos/244727418/shouhou/releases/latest"),
        ("所有Release", "https://api.github.com/repos/244727418/shouhou/releases"),
        ("仓库信息", "https://api.github.com/repos/244727418/shouhou"),
    ]
    
    for name, url in endpoints:
        print(f"\n🔍 {name}:")
        print(f"   地址: {url}")
        
        try:
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': 'Release-Checker',
                'Cache-Control': 'no-cache'  # 禁用缓存
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            print(f"   状态码: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                
                if 'latest' in url:
                    print(f"   ✅ 成功获取最新Release")
                    print(f"      Tag: {data.get('tag_name')}")
                    print(f"      标题: {data.get('name')}")
                    print(f"      创建时间: {data.get('created_at')}")
                    
                    assets = data.get('assets', [])
                    print(f"      附件数量: {len(assets)}")
                    for asset in assets:
                        print(f"        - {asset.get('name')} ({asset.get('size', 0)} bytes)")
                        
                elif 'releases' in url and 'latest' not in url:
                    releases = data
                    print(f"   Release总数: {len(releases)}")
                    
                    for i, release in enumerate(releases):
                        tag = release.get('tag_name')
                        name = release.get('name')
                        draft = release.get('draft', False)
                        prerelease = release.get('prerelease', False)
                        
                        status = ""
                        if draft:
                            status = " [草稿]"
                        elif prerelease:
                            status = " [预发布]"
                            
                        print(f"   [{i+1}] Tag: {tag}, 标题: {name}{status}")
                        
                        # 检查是否是语义化版本
                        if tag and tag.startswith('v') and tag[1:].replace('.', '').isdigit():
                            print(f"        ✅ 有效的语义化版本")
                        else:
                            print(f"        ⚠️ 非标准版本格式")
                            
                else:
                    print(f"   仓库: {data.get('name')}")
                    print(f"   描述: {data.get('description')}")
                    print(f"   最后更新: {data.get('updated_at')}")
                    
            elif response.status_code == 404:
                print(f"   ❌ 404 - 资源不存在")
                
                # 如果是latest返回404，但releases列表有数据，说明API有问题
                if 'latest' in url:
                    print(f"   ⚠️ 这可能是因为:")
                    print(f"      - GitHub API缓存延迟")
                    print(f"      - Release被标记为草稿/预发布")
                    print(f"      - Release格式问题")
                    
            else:
                print(f"   ❌ 错误: {response.status_code}")
                print(f"      响应: {response.text[:200]}")
                
        except Exception as e:
            print(f"   ❌ 异常: {type(e).__name__}: {e}")
    
    # 检查是否是草稿或预发布状态
    print("\n🔧 检查Release状态:")
    try:
        releases_url = "https://api.github.com/repos/244727418/shouhou/releases"
        response = requests.get(releases_url, headers={'User-Agent': 'Checker'}, timeout=10)
        
        if response.status_code == 200:
            releases = response.json()
            if releases:
                latest_release = releases[0]  # 第一个应该是最新的
                
                draft = latest_release.get('draft', False)
                prerelease = latest_release.get('prerelease', False)
                
                print(f"   最新Release状态:")
                print(f"     - 草稿: {draft}")
                print(f"     - 预发布: {prerelease}")
                print(f"     - 发布时间: {latest_release.get('published_at', '未发布')}")
                
                if draft:
                    print("   ❗ Release被标记为草稿，需要发布")
                if prerelease:
                    print("   ❗ Release被标记为预发布")
                if not latest_release.get('published_at'):
                    print("   ❗ Release未正式发布")
                    
    except Exception as e:
        print(f"   检查失败: {e}")

if __name__ == "__main__":
    check_release_details()