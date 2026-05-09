#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全自动商品信息收集表生成器 - 金山表单版

功能说明:
    1. 从腾讯文档在线表格中读取商品信息
    2. 使用金山表单API自动创建收集表
    3. 自动添加表单题目（商品名称、价格、图片等）
    4. 生成表单链接，供他人填写

前置要求:
    腾讯文档: 需要 access_token
    金山文档: 需要 app_id, app_key, 以及 OAuth2.0 access_token

使用方法:
    1. 填写配置信息
    2. 运行: python wps_form_generator.py
"""

import json
import hashlib
import requests
from datetime import datetime, timezone
from typing import List, Dict, Any, Optional
import uuid


# ==================== 用户配置区域 ====================

# 腾讯文档配置（读取源数据）
TENCENT_CONFIG = {
    "client_id": "9f95052132d446609c1c0f029271736c",
    "file_id": "DS2hSRUNLQUJyTXZF",
    "sheet_id": "BB08J2",
    "range": "A1:H70",
    "access_token": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbHQiOiI5Zjk1MDUyMTMyZDQ0NjYwOWMxYzBmMDI5MjcxNzM2YyIsInR5cCI6MSwiZXhwIjoxNzgwODQyNTU0Ljg3OTAyOSwiaWF0IjoxNzc4MjUwNTU0Ljg3OTAyOSwic3ViIjoiMzk1Mjc0ZTA0YWEwNDJkNzhlODMxOTYzZDRmNjNhMTkifQ.XWqTy4rewCgve35x4z8pfPFrN81tp3dVkVzERhnhImg",
    "open_id": "395274e04aa042d78e831963d4f63a19",
}

# 金山文档配置（创建表单）
KDOCS_CONFIG = {
    "app_id": "AK20260508QYUEYQ",
    "app_key": "cfa558741f245981a3c372264c09eda5",
    # 需要先通过OAuth2.0获取access_token
    # 授权链接: https://developer.kdocs.cn/h5/auth?app_id=AK20260508QYUEYQ&scope=user_basic,create_form&redirect_uri=YOUR_REDIRECT_URI
    "access_token": "",  # 填写获取到的access_token
}

# 表单配置
FORM_CONFIG = {
    "title": "商品信息收集表",
    "subtitle": "请填写商品相关信息",
}

# 字段映射配置
FIELD_MAPPING = {
    "商品名称": ["商品名", "名称", "产品名称", "商品", "品名", "name", "title"],
    "商品价格": ["价格", "单价", "售价", "price", "cost"],
    "商品图片": ["图片", "商品图", "image", "photo", "pic"],
    "商品规格": ["规格", "型号", "配置", "spec", "model"],
    "商品库存": ["库存", "数量", "stock", "quantity"],
    "供应商": ["供应商", "厂家", "品牌", "厂商", "supplier", "brand"],
    "发售日期": ["发售日", "上市日期", "日期", "date", "time"],
    "备注": ["备注", "说明", "描述", "备注信息", "desc"],
}

# =====================================================


class TencentDocsAPI:
    """腾讯文档开放平台 API 封装"""

    BASE_URL = "https://docs.qq.com/openapi"

    def __init__(self, client_id: str, access_token: str = "", open_id: str = ""):
        self.client_id = client_id
        self.access_token = access_token
        self.open_id = open_id

    def _get_headers(self, content_type: str = "application/json") -> Dict:
        headers = {
            "Access-Token": self.access_token,
            "Open-Id": self.open_id,
            "Client-Id": self.client_id,
        }
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def get_sheet_data(self, file_id: str, sheet_id: str, range_str: str) -> List[List[str]]:
        """获取指定范围内的表格数据"""
        if not self.access_token:
            raise Exception("请先获取Access Token")

        path = f"/spreadsheet/v3/files/{file_id}/{sheet_id}/{range_str}"
        url = f"{self.BASE_URL}{path}"

        response = requests.get(url, headers=self._get_headers("application/json"))
        result = response.json()

        if "ret" in result and result.get("ret") != 0:
            raise Exception(f"获取表格数据失败: {result.get('msg', '未知错误')}")

        if "gridData" in result:
            grid_data = result.get("gridData", {})
        else:
            grid_data = result.get("data", {}).get("gridData", {})
        rows = grid_data.get("rows", [])

        data = []
        for row in rows:
            row_data = []
            for cell in row.get("values", []):
                cell_value = cell.get("cellValue", {})
                text = cell_value.get("text", "")
                row_data.append(text)
            data.append(row_data)

        return data


class KDocsFormAPI:
    """金山文档表单 API 封装"""

    BASE_URL = "https://developer.kdocs.cn"

    def __init__(self, app_id: str, app_key: str, access_token: str = ""):
        self.app_id = app_id
        self.app_key = app_key.encode("utf-8")
        self.access_token = access_token

    def _generate_signature(self, content_md5: str, content_type: str, date: str) -> str:
        """生成 WPS-2 签名"""
        string_to_sign = self.app_key + content_md5.encode("utf-8") + content_type.encode("utf-8") + date.encode("utf-8")
        signature = hashlib.sha1(string_to_sign).hexdigest()
        return signature

    def _make_request(self, method: str, path: str, body: Optional[Dict] = None, query_params: Optional[Dict] = None) -> Dict:
        """发送带签名的请求"""
        date = datetime.now(timezone.utc).strftime("%a, %d %b %Y %H:%M:%S GMT")

        if body is not None:
            body_json = json.dumps(body, ensure_ascii=False, separators=(',', ':'))
            body_bytes = body_json.encode("utf-8")
            content_md5 = hashlib.md5(body_bytes).hexdigest()
        else:
            uri = path
            if query_params:
                uri += "?" + requests.compat.urlencode(query_params)
            content_md5 = hashlib.md5(uri.encode("utf-8")).hexdigest()
            body_bytes = None

        content_type = "application/json"
        signature = self._generate_signature(content_md5, content_type, date)
        authorization = f"WPS-2:{self.app_id}:{signature}"

        headers = {
            "Date": date,
            "Content-Md5": content_md5,
            "Content-Type": content_type,
            "Authorization": authorization,
        }

        if self.access_token:
            headers["X-Access-Token"] = self.access_token

        url = f"{self.BASE_URL}{path}"
        if query_params:
            url += "?" + requests.compat.urlencode(query_params)

        if method.upper() == "GET":
            response = requests.get(url, headers=headers)
        elif method.upper() == "POST":
            response = requests.post(url, headers=headers, data=body_bytes)
        elif method.upper() == "PUT":
            response = requests.put(url, headers=headers, data=body_bytes)
        elif method.upper() == "DELETE":
            response = requests.delete(url, headers=headers)
        else:
            raise ValueError(f"不支持的HTTP方法: {method}")

        return response.json()

    def create_form(self, title: str, subtitle: str = "", questions: Dict = None, config: Dict = None) -> Dict:
        """
        创建表单

        Args:
            title: 表单标题
            subtitle: 表单副标题
            questions: 表单题目
            config: 表单配置
        """
        if not self.access_token:
            raise Exception("请先获取Access Token")

        body = {
            "title": title,
            "subtitle": subtitle,
            "questions": questions or {},
            "config": config or {
                "check_login": False,
                "check_once": False,
                "expire": 0,
            },
            "answer_items": {
                "closing_remark_text": "感谢您的填写！"
            }
        }

        # 添加题目顺序
        if questions:
            body["question_sections"] = [
                {
                    "id": "section_1",
                    "questions": [{"qid": qid} for qid in questions.keys()]
                }
            ]

        result = self._make_request("POST", "/api/v1/openapi/personal/forms", body=body)

        if result.get("code") == 0:
            return result.get("data", {})
        else:
            raise Exception(f"创建表单失败: {result}")


def analyze_headers(headers: List[str], field_mapping: Dict) -> Dict[int, str]:
    """分析表头，匹配字段映射"""
    column_mapping = {}
    used_standard_fields = set()

    for col_idx, header in enumerate(headers):
        header_lower = header.lower().strip()
        matched = False

        for standard_field, keywords in field_mapping.items():
            if standard_field in used_standard_fields:
                continue

            for keyword in keywords:
                if keyword.lower() in header_lower or header_lower in keyword.lower():
                    column_mapping[col_idx] = standard_field
                    used_standard_fields.add(standard_field)
                    matched = True
                    break

            if matched:
                break

        if not matched and header.strip():
            column_mapping[col_idx] = header.strip()

    return column_mapping


def create_form_questions(column_mapping: Dict[int, str], headers: List[str]) -> Dict:
    """
    根据列映射创建表单题目

    Args:
        column_mapping: 列索引到字段名的映射
        headers: 原始表头

    Returns:
        表单题目定义
    """
    questions = {}

    # 字段类型映射到表单题目类型
    field_type_map = {
        "商品名称": "input",
        "商品价格": "floatInput",
        "商品图片": "newImage",
        "商品规格": "input",
        "商品库存": "floatInput",
        "供应商": "input",
        "发售日期": "date",
        "备注": "multiInput",
    }

    for col_idx, field_name in sorted(column_mapping.items()):
        qid = f"q_{col_idx}"
        q_type = field_type_map.get(field_name, "input")

        question = {
            "id": qid,
            "title": field_name,
            "type": q_type,
            "need": False,
            "hidden": False,
            "delete": False,
        }

        # 图片题特殊配置
        if q_type in ["newImage", "newMultiImage"]:
            question["subItems"] = [
                {"itemId": f"{qid}_item", "sort": 1, "text": "请上传图片"}
            ]
            question["fileSizeLimit"] = 10
            question["fileAcceptType"] = ["image"]

        questions[qid] = question

    return questions


def auto_generate_form(tencent_config: Dict, kdocs_config: Dict, form_config: Dict):
    """
    主流程：自动读取商品信息并生成金山表单
    """
    print("=" * 70)
    print("全自动商品信息收集表生成器 - 金山表单版")
    print("=" * 70)

    # 步骤1: 初始化腾讯文档API
    print("\n[步骤1] 连接腾讯文档...")
    tencent_api = TencentDocsAPI(
        client_id=tencent_config["client_id"],
        access_token=tencent_config["access_token"],
        open_id=tencent_config["open_id"]
    )
    print("  ✓ 腾讯文档API连接成功")

    # 步骤2: 读取源表格数据
    print("\n[步骤2] 读取源表格数据...")
    try:
        data = tencent_api.get_sheet_data(
            file_id=tencent_config["file_id"],
            sheet_id=tencent_config["sheet_id"],
            range_str=tencent_config["range"]
        )
        print(f"  ✓ 成功读取 {len(data)} 行数据")
    except Exception as e:
        print(f"  ✗ 读取失败: {e}")
        return

    # 步骤3: 分析表头
    print("\n[步骤3] 分析表头结构...")
    if not data:
        print("  ✗ 没有数据")
        return

    headers = data[0]
    print(f"  原始表头: {headers}")

    column_mapping = analyze_headers(headers, FIELD_MAPPING)
    print(f"  识别到的字段:")
    for col_idx, field_name in sorted(column_mapping.items()):
        print(f"    列{col_idx + 1}: {headers[col_idx]} -> {field_name}")

    # 步骤4: 检查金山文档配置
    print("\n[步骤4] 检查金山文档配置...")
    if not kdocs_config.get("access_token"):
        print("  ⚠ 缺少金山文档Access Token")
        print("\n  请按以下步骤获取Access Token:")
        print(f"  1. 访问授权链接:")
        print(f"     https://developer.kdocs.cn/h5/auth")
        print(f"     ?app_id={kdocs_config['app_id']}")
        print(f"     &scope=user_basic,create_form")
        print(f"     &redirect_uri=YOUR_REDIRECT_URI")
        print(f"  2. 用户授权后，从回调URL获取code参数")
        print(f"  3. 用code换取access_token:")
        print(f"     GET https://developer.kdocs.cn/api/v1/oauth2/access_token")
        print(f"     ?code={{code}}&app_id={kdocs_config['app_id']}&app_key={kdocs_config['app_key']}")
        return

    # 步骤5: 初始化金山表单API
    print("\n[步骤5] 连接金山表单API...")
    kdocs_api = KDocsFormAPI(
        app_id=kdocs_config["app_id"],
        app_key=kdocs_config["app_key"],
        access_token=kdocs_config["access_token"]
    )
    print("  ✓ 金山表单API连接成功")

    # 步骤6: 创建表单题目
    print("\n[步骤6] 创建表单题目...")
    questions = create_form_questions(column_mapping, headers)
    print(f"  ✓ 生成 {len(questions)} 个题目:")
    for qid, q in questions.items():
        print(f"    - {q['title']} ({q['type']})")

    # 步骤7: 创建表单
    print("\n[步骤7] 创建金山表单...")
    try:
        form_data = kdocs_api.create_form(
            title=form_config["title"],
            subtitle=form_config["subtitle"],
            questions=questions,
            config={
                "check_login": False,
                "check_once": False,
                "expire": 0,
            }
        )
        form_id = form_data.get("id", {}).get("super_token")
        print(f"  ✓ 表单创建成功")
        print(f"    表单Token: {form_id}")
    except Exception as e:
        print(f"  ✗ 创建表单失败: {e}")
        return

    # 步骤8: 输出结果
    print("\n" + "=" * 70)
    print("✓ 全自动收集表生成完成!")
    print("=" * 70)
    print(f"\n  📋 表单标题: {form_config['title']}")
    print(f"  🔗 表单链接: https://f.wps.cn/{form_id}")
    print(f"\n  📊 数据汇总:")
    print(f"     - 读取源数据: {len(data) - 1} 条商品记录")
    print(f"     - 生成题目数: {len(questions)} 个")
    print(f"\n  🚀 使用方法:")
    print("     1. 点击上方链接进入表单填写页面")
    print("     2. 直接分享此链接给他人填写")
    print("     3. 填写结果会自动汇总")
    print("=" * 70)


def demo_mode():
    """演示模式"""
    print("=" * 70)
    print("演示模式: 展示全自动表单生成流程")
    print("=" * 70)

    demo_data = [
        ["商品名称", "价格", "图片链接", "规格", "库存", "供应商"],
        ["iPhone 15 Pro", "8999", "https://example.com/iphone.jpg", "256GB 黑色", "100", "Apple"],
        ["MacBook Air M3", "11999", "https://example.com/mac.jpg", "16GB+512GB", "50", "Apple"],
    ]

    print("\n[模拟数据]")
    for row in demo_data:
        print(f"  {' | '.join(row)}")

    headers = demo_data[0]
    print(f"\n[表头分析]")
    print(f"  原始表头: {headers}")

    column_mapping = analyze_headers(headers, FIELD_MAPPING)
    print(f"\n[字段匹配结果]")
    for col_idx, field_name in sorted(column_mapping.items()):
        print(f"  列{col_idx + 1}: '{headers[col_idx]}' -> '{field_name}'")

    print(f"\n[生成的表单题目]")
    questions = create_form_questions(column_mapping, headers)
    for qid, q in questions.items():
        print(f"  - {q['title']} (类型: {q['type']})")

    print("\n" + "=" * 70)
    print("演示完成! 配置金山文档Access Token后即可运行正式模式")
    print("=" * 70)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--demo":
        demo_mode()
    else:
        if TENCENT_CONFIG["client_id"].startswith("YOUR_"):
            print("⚠ 请先配置API密钥信息!")
            print("\n使用方法:")
            print("1. 查看演示模式: python wps_form_generator.py --demo")
            print("2. 编辑代码顶部的 TENCENT_CONFIG 和 KDOCS_CONFIG")
            print("3. 运行: python wps_form_generator.py")
        else:
            auto_generate_form(TENCENT_CONFIG, KDOCS_CONFIG, FORM_CONFIG)
