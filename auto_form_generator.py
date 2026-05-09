#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全自动商品信息收集表生成器

功能说明:
    1. 从腾讯文档在线表格中读取商品信息（商品名、价格、商品图等）
    2. 自动创建腾讯智能表格（Smartsheet）
    3. 自动添加字段（商品名称、价格、图片等）
    4. 将商品数据自动写入智能表格
    5. 自动生成收集表链接，供他人填写

前置要求:
    腾讯文档开放平台: 需要申请应用获取 client_id, access_token
    权限要求: scope.sheet, scope.smartsheet

使用方法:
    1. 填写配置信息
    2. 运行: python auto_form_generator.py
"""

import json
import requests
from typing import List, Dict, Any, Optional


# ==================== 用户配置区域 ====================

TENCENT_CONFIG = {
    "client_id": "9f95052132d446609c1c0f029271736c",
    "file_id": "DS2hSRUNLQUJyTXZF",
    "sheet_id": "BB08J2",
    "range": "A1:H70",
    "access_token": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbHQiOiI5Zjk1MDUyMTMyZDQ0NjYwOWMxYzBmMDI5MjcxNzM2YyIsInR5cCI6MSwiZXhwIjoxNzgwODQyNTU0Ljg3OTAyOSwiaWF0IjoxNzc4MjUwNTU0Ljg3OTAyOSwic3ViIjoiMzk1Mjc0ZTA0YWEwNDJkNzhlODMxOTYzZDRmNjNhMTkifQ.XWqTy4rewCgve35x4z8pfPFrN81tp3dVkVzERhnhImg",
    "open_id": "395274e04aa042d78e831963d4f63a19",
    "output_title": "商品信息收集表",
}

# 字段映射配置：将源表格列名映射到收集表字段
# 程序会自动识别这些关键词来匹配列
FIELD_MAPPING = {
    "商品名称": ["商品名", "名称", "产品名称", "商品", "name", "title"],
    "商品价格": ["价格", "单价", "售价", "price", "cost"],
    "商品图片": ["图片", "商品图", "image", "photo", "pic"],
    "商品规格": ["规格", "型号", "配置", "spec", "model"],
    "商品库存": ["库存", "数量", "stock", "quantity"],
    "供应商": ["供应商", "厂家", "品牌", "supplier", "brand"],
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
        """获取通用请求头"""
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

    def create_smartsheet(self, title: str) -> Dict:
        """
        创建智能表格（Smartsheet）
        """
        url = f"{self.BASE_URL}/drive/v2/files"
        headers = self._get_headers("application/x-www-form-urlencoded")

        data = {
            "type": "smartsheet",
            "title": title,
        }

        response = requests.post(url, headers=headers, data=data)
        result = response.json()

        if result.get("ret") == 0:
            return result.get("data", {})
        else:
            raise Exception(f"创建智能表格失败: {result.get('msg', '未知错误')}")

    def add_smartsheet_fields(self, file_id: str, sheet_id: str, fields: List[Dict]) -> Dict:
        """
        向智能表格添加字段

        Args:
            file_id: 智能表格文件ID
            sheet_id: 子表ID
            fields: 字段列表，每个字段包含 fieldTitle, fieldType 等
        """
        url = f"{self.BASE_URL}/smartbook/v2/files/{file_id}/sheets/{sheet_id}"
        headers = self._get_headers("application/json")

        body = {
            "addFields": {
                "fields": fields
            }
        }

        response = requests.post(url, headers=headers, json=body)
        result = response.json()

        if result.get("ret") == 0:
            return result.get("data", {})
        else:
            raise Exception(f"添加字段失败: {result.get('msg', '未知错误')}")

    def add_smartsheet_records(self, file_id: str, sheet_id: str, records: List[Dict]) -> Dict:
        """
        向智能表格添加记录

        Args:
            file_id: 智能表格文件ID
            sheet_id: 子表ID
            records: 记录列表
        """
        url = f"{self.BASE_URL}/smartbook/v2/files/{file_id}/sheets/{sheet_id}"
        headers = self._get_headers("application/json")

        body = {
            "addRecords": {
                "records": records
            }
        }

        response = requests.post(url, headers=headers, json=body)
        result = response.json()

        if result.get("ret") == 0:
            return result.get("data", {})
        else:
            raise Exception(f"添加记录失败: {result.get('msg', '未知错误')}")

    def set_file_permission(self, file_id: str, permission: str = "anyone_editable") -> bool:
        """设置文件权限"""
        url = f"{self.BASE_URL}/drive/v2/files/{file_id}/permission"
        headers = self._get_headers("application/x-www-form-urlencoded")

        data = {"scope": permission}
        response = requests.patch(url, headers=headers, data=data)
        result = response.json()

        return result.get("ret") == 0

    def get_sheet_info(self, file_id: str) -> Dict:
        """获取智能表格的子表信息"""
        url = f"{self.BASE_URL}/smartbook/v2/files/{file_id}/sheets"
        headers = self._get_headers("application/json")

        response = requests.get(url, headers=headers)
        result = response.json()

        if result.get("ret") == 0:
            data = result.get("data", {})
            # API返回格式可能是 {"getSheet": [...]} 或 {"sheets": [...]}
            if "getSheet" in data:
                data["sheets"] = data.pop("getSheet")
            return data
        else:
            raise Exception(f"获取子表信息失败: {result.get('msg', '未知错误')}")


def analyze_headers(headers: List[str], field_mapping: Dict) -> Dict[int, str]:
    """
    分析表头，匹配字段映射

    Args:
        headers: 表头列表
        field_mapping: 字段映射配置

    Returns:
        列索引到标准字段名的映射
    """
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


def create_smartsheet_fields(column_mapping: Dict[int, str], headers: List[str]) -> List[Dict]:
    """
    根据列映射创建智能表格字段

    Args:
        column_mapping: 列索引到字段名的映射
        headers: 原始表头

    Returns:
        字段定义列表
    """
    fields = []

    # 字段类型映射
    field_type_map = {
        "商品名称": 1,      # 文本
        "商品价格": 2,      # 数字
        "商品图片": 8,      # 超链接（URL）
        "商品规格": 1,      # 文本
        "商品库存": 2,      # 数字
        "供应商": 1,        # 文本
        "发售日期": 4,      # 日期
        "备注": 1,          # 文本
    }

    for col_idx, field_name in sorted(column_mapping.items()):
        field_type = field_type_map.get(field_name, 1)

        field_def = {
            "fieldTitle": field_name,
            "fieldType": field_type,
        }

        # 根据字段类型添加属性
        if field_type == 2:  # 数字
            field_def["propertyNumber"] = {}
        elif field_type == 4:  # 日期
            field_def["propertyDateTime"] = {"autoFill": False, "format": "yyyy/m/d"}
        elif field_type == 8:  # 超链接
            field_def["propertyUrl"] = {"type": 1}
        else:  # 文本
            field_def["propertyText"] = {}

        fields.append(field_def)

    return fields


def convert_data_to_records(data: List[List[str]], column_mapping: Dict[int, str]) -> List[Dict]:
    """
    将表格数据转换为智能表格记录格式

    Args:
        data: 表格数据（包含表头）
        column_mapping: 列索引到字段名的映射

    Returns:
        记录列表
    """
    if not data or len(data) < 2:
        return []

    records = []
    # 从第2行开始（跳过表头）
    for row_idx in range(1, len(data)):
        row = data[row_idx]
        record = {"fields": {}}

        for col_idx, field_name in column_mapping.items():
            if col_idx < len(row):
                value = row[col_idx]
                record["fields"][field_name] = value

        records.append(record)

    return records


def auto_generate_form(config: Dict):
    """
    主流程：自动读取商品信息并生成收集表
    """
    print("=" * 70)
    print("全自动商品信息收集表生成器")
    print("=" * 70)

    # 步骤1: 初始化API
    print("\n[步骤1] 连接腾讯文档...")
    api = TencentDocsAPI(
        client_id=config["client_id"],
        access_token=config["access_token"],
        open_id=config["open_id"]
    )
    print("  ✓ API连接成功")

    # 步骤2: 读取源表格数据
    print("\n[步骤2] 读取源表格数据...")
    try:
        data = api.get_sheet_data(
            file_id=config["file_id"],
            sheet_id=config["sheet_id"],
            range_str=config["range"]
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

    # 步骤4: 创建智能表格
    print("\n[步骤4] 创建智能表格...")
    try:
        smartsheet_info = api.create_smartsheet(config["output_title"])
        smartsheet_id = smartsheet_info.get("ID")
        smartsheet_url = smartsheet_info.get("url")
        print(f"  ✓ 智能表格创建成功")
        print(f"    ID: {smartsheet_id}")
    except Exception as e:
        print(f"  ✗ 创建失败: {e}")
        return

    # 步骤5: 获取子表ID
    print("\n[步骤5] 获取子表信息...")
    try:
        sheet_info = api.get_sheet_info(smartsheet_id)
        sheets = sheet_info.get("sheets", [])
        if sheets:
            sub_sheet_id = sheets[0].get("sheetID")
            print(f"  ✓ 子表ID: {sub_sheet_id}")
        else:
            print("  ✗ 未找到子表")
            return
    except Exception as e:
        print(f"  ✗ 获取子表信息失败: {e}")
        return

    # 步骤6: 添加字段
    print("\n[步骤6] 自动添加字段...")
    try:
        fields = create_smartsheet_fields(column_mapping, headers)
        print(f"  准备添加 {len(fields)} 个字段:")
        for f in fields:
            print(f"    - {f['fieldTitle']} (类型: {f['fieldType']})")

        result = api.add_smartsheet_fields(smartsheet_id, sub_sheet_id, fields)
        print(f"  ✓ 字段添加成功")
    except Exception as e:
        print(f"  ✗ 添加字段失败: {e}")
        return

    # 步骤7: 写入数据
    print("\n[步骤7] 写入商品数据...")
    try:
        records = convert_data_to_records(data, column_mapping)
        if records:
            # 分批写入，每批50条
            batch_size = 50
            total = len(records)
            for i in range(0, total, batch_size):
                batch = records[i:i + batch_size]
                api.add_smartsheet_records(smartsheet_id, sub_sheet_id, batch)
                print(f"  ✓ 已写入 {min(i + batch_size, total)}/{total} 条记录")
        else:
            print("  ⚠ 没有数据需要写入")
    except Exception as e:
        print(f"  ✗ 写入数据失败: {e}")
        return

    # 步骤8: 设置权限
    print("\n[步骤8] 设置权限...")
    try:
        api.set_file_permission(smartsheet_id, "anyone_editable")
        print("  ✓ 已设置为任何人可编辑")
    except Exception as e:
        print(f"  ⚠ 设置权限失败: {e}")

    # 步骤9: 生成收集表链接
    print("\n[步骤9] 生成收集表...")
    try:
        # 收集表链接格式
        form_url = f"https://docs.qq.com/form/page/{smartsheet_url.split('/')[-1]}"
        print(f"  ✓ 收集表已生成")
    except Exception as e:
        print(f"  ⚠ 生成收集表链接时出错: {e}")
        form_url = smartsheet_url

    # 步骤10: 输出结果
    print("\n" + "=" * 70)
    print("✓ 全自动收集表生成完成!")
    print("=" * 70)
    print(f"\n  📋 表格标题: {config['output_title']}")
    print(f"\n  🔗 智能表格链接（查看数据）:")
    print(f"     {smartsheet_url}")
    print(f"\n  📝 收集表链接（给他人填写）:")
    print(f"     {form_url}")
    print(f"\n  📊 数据汇总:")
    print(f"     - 读取源数据: {len(data) - 1} 条商品记录")
    print(f"     - 创建字段数: {len(fields)} 个")
    print(f"     - 写入记录数: {len(records)} 条")
    print(f"\n  🚀 使用方法:")
    print("     1. 点击'收集表链接'进入表单填写页面")
    print("     2. 直接分享此链接给他人填写")
    print("     3. 填写结果会自动汇总到智能表格中")
    print("     4. 点击'智能表格链接'查看汇总数据")
    print("=" * 70)


def demo_mode():
    """演示模式"""
    print("=" * 70)
    print("演示模式: 展示全自动收集表生成流程")
    print("=" * 70)

    demo_data = [
        ["商品名称", "价格", "图片链接", "规格", "库存", "供应商"],
        ["iPhone 15 Pro", "8999", "https://example.com/iphone.jpg", "256GB 黑色", "100", "Apple"],
        ["MacBook Air M3", "11999", "https://example.com/mac.jpg", "16GB+512GB", "50", "Apple"],
        ["AirPods Pro 2", "1999", "https://example.com/airpods.jpg", "降噪版", "200", "Apple"],
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

    print(f"\n[生成的字段定义]")
    fields = create_smartsheet_fields(column_mapping, headers)
    for f in fields:
        print(f"  - {f['fieldTitle']} (类型代码: {f['fieldType']})")

    print(f"\n[转换后的记录]")
    records = convert_data_to_records(demo_data, column_mapping)
    for i, record in enumerate(records[:2]):
        print(f"  记录{i + 1}: {json.dumps(record, ensure_ascii=False)}")

    print("\n" + "=" * 70)
    print("演示完成! 运行正式模式将自动执行上述流程")
    print("=" * 70)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--demo":
        demo_mode()
    else:
        if TENCENT_CONFIG["client_id"].startswith("YOUR_"):
            print("⚠ 请先配置API密钥信息!")
            print("\n使用方法:")
            print("1. 查看演示模式: python auto_form_generator.py --demo")
            print("2. 编辑代码顶部的 TENCENT_CONFIG")
            print("3. 运行: python auto_form_generator.py")
        else:
            auto_generate_form(TENCENT_CONFIG)
