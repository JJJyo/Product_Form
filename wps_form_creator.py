#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS表单生成器
基于reptiles项目学习到的WPS表单API实现
"""

import json
import requests
import random
import string
from datetime import datetime

# WPS Cookie
WPS_SID = "V02S3ypjutra8juOvM7K29C8SXiZooM00acecadb0049df50f3"

# API基础地址
API_BASE = "https://f-api.wps.cn/ksform/api"


def generate_random_string(length=6):
    """生成随机字符串"""
    characters = string.ascii_lowercase + string.digits
    return ''.join(random.choice(characters) for _ in range(length))


def get_headers():
    """获取请求头"""
    return {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://f.wps.cn/',
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
    }


def get_cookies():
    """获取Cookie"""
    return {
        'wps_sid': WPS_SID,
    }


def create_form(title, subtitle=""):
    """
    创建WPS表单
    
    Args:
        title: 表单标题
        subtitle: 表单副标题
    
    Returns:
        表单ID和分享链接
    """
    # 构建表单数据
    form_data = {
        "name": title,
        "type": "Draft",
        "kind": "ksform",
        "tableId": "",
        "shareId": "",
        "property": {
            "subNameAlign": "left",
        },
        "setting": {
            "baseSetting": {
                "checkLogin": False,
                "checkOnce": False,
                "checkPhone": False,
                "canModify": False,
                "isAnonymous": False,
                "isNotify": False,
                "startTime": 0,
                "stopTime": 0,
                "fillLimit": 0,
                "fillType": "anyone",
            },
            "extendSetting": {
                "answerQrcodeConfig": {
                    "isAnswerQrcodeOpen": True,
                    "type": "show",
                }
            }
        },
        "questionMap": {},
        "questionLayouts": [],
        "questionLogicMap": {},
        "stockMap": {},
    }
    
    # 创建表单
    url = f"{API_BASE}/v3/draft"
    response = requests.post(
        url,
        headers=get_headers(),
        cookies=get_cookies(),
        json=form_data
    )
    
    print(f"Create form status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
    
    if response.status_code == 200:
        result = response.json()
        return result.get("data", {})
    else:
        raise Exception(f"创建表单失败: {response.text}")


def add_text_question(form_id, title, required=False):
    """
    添加文本问题
    
    Args:
        form_id: 表单ID
        title: 问题标题
        required: 是否必填
    """
    qid = generate_random_string(6)
    
    question_data = {
        "baseInfo": {
            "note": "",
            "serialNumberType": "continuous",
            "new": True,
            "isReadonly": False,
            "need": required,
            "isProtectPrivacy": False,
            "noteFiles": [],
            "isShowNote": False,
            "isNoRepeat": False,
            "delete": False,
            "isHidden": False
        },
        "title": title,
        "type": "input",
        "qid": qid,
        "version": 0
    }
    
    # 添加问题到表单
    url = f"{API_BASE}/v3/draft/{form_id}/questions"
    response = requests.post(
        url,
        headers=get_headers(),
        cookies=get_cookies(),
        json=question_data
    )
    
    print(f"Add question status: {response.status_code}")
    return qid


def add_image_question(form_id, title, image_url):
    """
    添加图片问题
    
    Args:
        form_id: 表单ID
        title: 问题标题
        image_url: 图片URL
    """
    qid = generate_random_string(6)
    
    # 先上传图片获取fileSid
    # TODO: 实现图片上传逻辑
    
    question_data = {
        "baseInfo": {
            "note": "",
            "serialNumberType": "continuous",
            "new": True,
            "isReadonly": False,
            "need": False,
            "isProtectPrivacy": False,
            "noteFiles": [],
            "isShowNote": False,
            "isNoRepeat": False,
            "delete": False,
            "isHidden": False
        },
        "title": title,
        "type": "newImage",
        "qid": qid,
        "version": 0,
        "subItems": [
            {
                "itemId": generate_random_string(6),
                "sort": 1,
                "text": "请上传图片"
            }
        ],
        "fileSizeLimit": 10,
        "fileAcceptType": ["image"]
    }
    
    url = f"{API_BASE}/v3/draft/{form_id}/questions"
    response = requests.post(
        url,
        headers=get_headers(),
        cookies=get_cookies(),
        json=question_data
    )
    
    print(f"Add image question status: {response.status_code}")
    return qid


def publish_form(form_id):
    """
    发布表单
    
    Args:
        form_id: 表单ID
    
    Returns:
        表单分享链接
    """
    url = f"{API_BASE}/v3/draft/{form_id}/publish"
    response = requests.post(
        url,
        headers=get_headers(),
        cookies=get_cookies()
    )
    
    print(f"Publish form status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
    
    if response.status_code == 200:
        result = response.json()
        return result.get("data", {})
    else:
        raise Exception(f"发布表单失败: {response.text}")


def create_product_form(title, products):
    """
    创建商品信息收集表单
    
    Args:
        title: 表单标题
        products: 商品列表
    
    Returns:
        表单链接
    """
    # 创建表单
    form_data = create_form(title)
    form_id = form_data.get("id")
    
    print(f"表单创建成功，ID: {form_id}")
    
    # 添加商品名称问题
    add_text_question(form_id, "商品名称", required=True)
    
    # 添加价格问题
    add_text_question(form_id, "商品价格", required=True)
    
    # 添加定金问题
    add_text_question(form_id, "定金", required=False)
    
    # 添加备注问题
    add_text_question(form_id, "备注", required=False)
    
    # 添加图片问题
    add_image_question(form_id, "商品图片", "")
    
    # 发布表单
    publish_data = publish_form(form_id)
    
    return publish_data.get("shareUrl", "")


if __name__ == "__main__":
    # 测试创建表单
    try:
        share_url = create_product_form("商品信息收集表", [])
        print(f"\n表单创建成功！")
        print(f"分享链接: {share_url}")
    except Exception as e:
        print(f"创建表单失败: {e}")
