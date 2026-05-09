#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
腾讯文档商品信息读取并生成金山文档(WPS在线文档)

功能说明:
    1. 从腾讯文档在线表格中读取商品信息
    2. 在金山文档中创建新的在线表格
    3. 将商品信息写入金山文档

前置要求:
    1. 腾讯文档开放平台: 需要申请应用获取 client_id, client_secret
    2. 金山文档开放平台: 需要申请应用获取 app_id, app_key

使用方法:
    1. 填写配置信息（见下方 CONFIG 区域）
    2. 运行: python tencent_to_kdocs.py
"""

import json
import hashlib
import hmac
import base64
import requests
from datetime import datetime, timezone
from urllib.parse import urlencode
from typing import List, Dict, Any, Optional


# ==================== 用户配置区域 ====================

# 腾讯文档配置
TENCENT_CONFIG = {
    "client_id": "9f95052132d446609c1c0f029271736c",
    "client_secret": "",
    "file_id": "DS2hSRUNLQUJyTXZF",
    "sheet_id": "BB08J2",
    "range": "A1:H70",
    "access_token": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbHQiOiI5Zjk1MDUyMTMyZDQ0NjYwOWMxYzBmMDI5MjcxNzM2YyIsInR5cCI6MSwiZXhwIjoxNzgwODQyNTU0Ljg3OTAyOSwiaWF0IjoxNzc4MjUwNTU0Ljg3OTAyOSwic3ViIjoiMzk1Mjc0ZTA0YWEwNDJkNzhlODMxOTYzZDRmNjNhMTkifQ.XWqTy4rewCgve35x4z8pfPFrN81tp3dVkVzERhnhImg",
    "open_id": "395274e04aa042d78e831963d4f63a19",
    "form_title": "商品信息收集表",
}

# =====================================================


class TencentDocsAPI:
    """腾讯文档开放平台 API 封装"""

    BASE_URL = "https://docs.qq.com/openapi"

    def __init__(self, client_id: str, client_secret: str = "", access_token: str = "", open_id: str = ""):
        self.client_id = client_id
        self.client_secret = client_secret
        self.access_token = access_token
        self.open_id = open_id

    def set_access_token(self, access_token: str, open_id: str = ""):
        """直接设置已有的 Access Token"""
        self.access_token = access_token
        self.open_id = open_id

    def get_access_token(self, authorization_code: str = "") -> str:
        """
        通过 OAuth2.0 授权码获取 Access Token
        """
        url = f"{self.BASE_URL}/oauth/v2/token"
        data = {
            "grant_type": "authorization_code",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "code": authorization_code,
            "redirect_uri": "https://your-app.com/callback"
        }

        response = requests.post(url, data=data)
        result = response.json()

        if "access_token" in result:
            self.access_token = result["access_token"]
            self.open_id = result.get("open_id", "")
            return self.access_token
        else:
            raise Exception(f"获取AccessToken失败: {result}")

    def get_sheet_data(self, file_id: str, sheet_id: str, range_str: str) -> List[List[str]]:
        """
        获取指定范围内的表格数据

        Args:
            file_id: 在线表格唯一标识
            sheet_id: 工作表ID（空字符串表示默认工作表）
            range_str: 查询范围，如 "A1:D10"

        Returns:
            二维列表，每个元素是单元格的文本值
        """
        if not self.access_token:
            raise Exception("请先获取Access Token")

        # 构建请求URL
        if sheet_id:
            path = f"/spreadsheet/v3/files/{file_id}/{sheet_id}/{range_str}"
        else:
            # 如果没有sheet_id，先获取sheet列表
            sheet_id = self._get_first_sheet_id(file_id)
            path = f"/spreadsheet/v3/files/{file_id}/{sheet_id}/{range_str}"

        url = f"{self.BASE_URL}{path}"

        headers = {
            "Access-Token": self.access_token,
            "Open-Id": self.open_id,
            "Client-Id": self.client_id,
            "Accept": "application/json"
        }

        response = requests.get(url, headers=headers)
        result = response.json()

        # v3 API返回格式: {gridData: {...}} 或 {ret, msg, data: {gridData: {...}}}
        if "ret" in result and result.get("ret") != 0:
            raise Exception(f"获取表格数据失败: {result.get('msg', '未知错误')}")

        # 解析返回的数据
        if "gridData" in result:
            grid_data = result.get("gridData", {})
        else:
            grid_data = result.get("data", {}).get("gridData", {})
        rows = grid_data.get("rows", [])

        # 转换为二维列表
        data = []
        for row in rows:
            row_data = []
            for cell in row.get("values", []):
                cell_value = cell.get("cellValue", {})
                # 提取文本内容
                text = cell_value.get("text", "")
                row_data.append(text)
            data.append(row_data)

        return data

    def _get_first_sheet_id(self, file_id: str) -> str:
        """获取第一个工作表的ID"""
        url = f"{self.BASE_URL}/sheetbook/v2/{file_id}/sheets-info"
        headers = {
            "Access-Token": self.access_token,
            "Open-Id": self.open_id,
            "Client-Id": self.client_id,
        }

        response = requests.get(url, headers=headers)
        result = response.json()

        if result.get("ret") == 0:
            sheets = result.get("data", {}).get("sheets", [])
            if sheets:
                return sheets[0].get("sheetId", "")

        raise Exception("无法获取工作表ID")

    def create_form(self, title: str) -> Dict:
        """
        创建在线收集表

        Args:
            title: 收集表标题

        Returns:
            创建的收集表信息，包含ID和URL
        """
        if not self.access_token:
            raise Exception("请先获取Access Token")

        url = f"{self.BASE_URL}/drive/v2/files"
        headers = {
            "Access-Token": self.access_token,
            "Open-Id": self.open_id,
            "Client-Id": self.client_id,
            "Content-Type": "application/x-www-form-urlencoded",
        }

        data = {
            "type": "form",
            "title": title,
        }

        response = requests.post(url, headers=headers, data=data)
        result = response.json()

        if result.get("ret") == 0:
            return result.get("data", {})
        else:
            raise Exception(f"创建收集表失败: {result.get('msg', '未知错误')}")

    def set_form_permission(self, form_id: str, permission: str = "anyone_editable") -> bool:
        """
        设置收集表权限

        Args:
            form_id: 收集表ID
            permission: 权限类型，默认 anyone_editable（任何人可编辑）

        Returns:
            是否成功
        """
        if not self.access_token:
            raise Exception("请先获取Access Token")

        url = f"{self.BASE_URL}/drive/v2/files/{form_id}/permission"
        headers = {
            "Access-Token": self.access_token,
            "Open-Id": self.open_id,
            "Client-Id": self.client_id,
            "Content-Type": "application/x-www-form-urlencoded",
        }

        data = {
            "scope": permission,
        }

        response = requests.patch(url, headers=headers, data=data)
        result = response.json()

        return result.get("ret") == 0


class KDocsAPI:
    """金山文档开放平台 API 封装"""

    BASE_URL = "https://developer.kdocs.cn"

    def __init__(self, app_id: str, app_key: str, access_token: str = ""):
        self.app_id = app_id
        self.app_key = app_key.encode("utf-8")
        self.access_token = access_token

    def set_access_token(self, access_token: str):
        """设置访问令牌"""
        self.access_token = access_token

    def get_access_token(self, code: str) -> str:
        """
        通过OAuth2.0授权码获取access_token
        """
        url = f"{self.BASE_URL}/api/v1/oauth2/access_token"
        params = {
            "code": code,
            "app_id": self.app_id,
            "app_key": self.app_key.decode("utf-8")
        }
        headers = {
            "Content-Type": "application/json"
        }

        response = requests.get(url, params=params, headers=headers)
        result = response.json()

        if result.get("code") == 0:
            self.access_token = result["data"]["access_token"]
            return self.access_token
        else:
            raise Exception(f"获取access_token失败: {result}")

    def _generate_signature(self, content_md5: str, content_type: str, date: str) -> str:
        """
        生成 WPS-2 签名

        签名算法: sha1(app_key + Content-Md5 + Content-Type + DATE)
        """
        string_to_sign = self.app_key + content_md5.encode("utf-8") + content_type.encode("utf-8") + date.encode("utf-8")
        signature = hashlib.sha1(string_to_sign).hexdigest()
        return signature

    def _make_request(self, method: str, path: str, body: Optional[Dict] = None, query_params: Optional[Dict] = None) -> Dict:
        """发送带签名的请求"""
        # 生成日期头
        date = datetime.now(timezone.utc).strftime("%a, %d %b %Y %H:%M:%S GMT")

        # 准备请求体
        if body is not None:
            body_json = json.dumps(body, ensure_ascii=False, separators=(',', ':'))
            body_bytes = body_json.encode("utf-8")
            content_md5 = hashlib.md5(body_bytes).hexdigest()
        else:
            # GET请求使用URI计算MD5
            uri = path
            if query_params:
                uri += "?" + urlencode(query_params)
            content_md5 = hashlib.md5(uri.encode("utf-8")).hexdigest()
            body_bytes = None

        content_type = "application/json"

        # 生成签名
        signature = self._generate_signature(content_md5, content_type, date)
        authorization = f"WPS-2:{self.app_id}:{signature}"

        # 构建请求头
        headers = {
            "Date": date,
            "Content-Md5": content_md5,
            "Content-Type": content_type,
            "Authorization": authorization,
        }

        # 如果有access_token，添加到请求头
        if self.access_token:
            headers["X-Access-Token"] = self.access_token

        # 构建完整URL
        url = f"{self.BASE_URL}{path}"
        if query_params:
            url += "?" + urlencode(query_params)

        # 发送请求
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

        result = response.json()
        return result

    def create_spreadsheet(self, filename: str, parent_token: str = "", creator: str = "") -> str:
        """
        创建空白在线表格

        Args:
            filename: 文档名称（需包含扩展名，如 .ksheet）
            parent_token: 父目录ID
            creator: 创建者ID

        Returns:
            新创建文档的 file_token
        """
        body = {
            "filename": filename,
        }
        if parent_token:
            body["parent_token"] = parent_token
        if creator:
            body["creator"] = creator

        result = self._make_request("POST", "/api/v1/openapi/appspace/files", body=body)

        if result.get("code") == 0:
            file_token = result.get("data", {}).get("file_token")
            print(f"✓ 成功创建金山文档: {filename}")
            print(f"  文档ID: {file_token}")
            return file_token
        else:
            raise Exception(f"创建文档失败: {result}")

    def update_cells(self, file_token: str, sheet_idx: int, data: List[List[str]]) -> bool:
        """
        更新指定工作表的单元格数据

        Args:
            file_token: 文档ID
            sheet_idx: 工作表索引（从0开始）
            data: 二维列表，要写入的数据

        Returns:
            是否成功
        """
        if not data:
            print("⚠ 没有数据需要写入")
            return False

        # 构建 ranges 数据
        ranges = []
        for row_idx, row in enumerate(data):
            for col_idx, value in enumerate(row):
                ranges.append({
                    "row_from": row_idx,
                    "row_to": row_idx,
                    "col_from": col_idx,
                    "col_to": col_idx,
                    "formula": str(value),
                    "op_type": "formula"
                })

        body = {
            "ranges": ranges
        }

        path = f"/api/v1/openapi/ksheet/{file_token}/sheets/{sheet_idx}/cells"
        result = self._make_request("POST", path, body=body)

        if result.get("code") == 0:
            print(f"✓ 成功写入 {len(data)} 行 {max(len(r) for r in data)} 列数据")
            return True
        else:
            print(f"✗ 写入数据失败: {result}")
            return False

    def get_sheets_info(self, file_token: str) -> List[Dict]:
        """获取表格中所有工作表的信息"""
        path = f"/api/v1/openapi/ksheet/{file_token}/sheets"
        result = self._make_request("GET", path)

        if result.get("code") == 0:
            return result.get("data", {}).get("sheets", [])
        else:
            raise Exception(f"获取工作表信息失败: {result}")


def transfer_data(tencent_config: Dict, kdocs_config: Dict, auth_code: str = ""):
    """
    主流程：从腾讯文档读取数据并写入金山文档

    Args:
        tencent_config: 腾讯文档配置
        kdocs_config: 金山文档配置
        auth_code: 腾讯文档OAuth授权码（可选，如配置中已有access_token则不需要）
    """
    print("=" * 60)
    print("开始数据迁移: 腾讯文档 → 金山文档")
    print("=" * 60)

    # 步骤1: 初始化腾讯文档API并获取数据
    print("\n[步骤1] 连接腾讯文档...")

    # 优先使用配置中已有的access_token
    existing_token = tencent_config.get("access_token", "")
    existing_open_id = tencent_config.get("open_id", "")
    if existing_token:
        tencent_api = TencentDocsAPI(
            client_id=tencent_config["client_id"],
            access_token=existing_token,
            open_id=existing_open_id
        )
        print("  使用已提供的 Access Token")
    else:
        tencent_api = TencentDocsAPI(
            client_id=tencent_config["client_id"],
            client_secret=tencent_config["client_secret"]
        )
        # 获取Access Token（需要用户授权）
        if auth_code:
            tencent_api.get_access_token(auth_code)
        else:
            print("⚠ 请提供腾讯文档OAuth授权码或access_token")
            return

    # 读取表格数据
    print(f"  正在读取表格数据 (范围: {tencent_config['range']})...")
    try:
        data = tencent_api.get_sheet_data(
            file_id=tencent_config["file_id"],
            sheet_id=tencent_config.get("sheet_id", ""),
            range_str=tencent_config["range"]
        )
        print(f"✓ 成功读取 {len(data)} 行数据")
    except Exception as e:
        print(f"✗ 读取腾讯文档失败: {e}")
        return

    # 步骤2: 创建腾讯收集表
    print("\n[步骤2] 创建腾讯收集表...")
    try:
        form_info = tencent_api.create_form(
            title=tencent_config.get("form_title", "商品信息收集表")
        )
        form_id = form_info.get("ID")
        form_url = form_info.get("url")
        print(f"✓ 成功创建收集表: {form_info.get('title')}")
        print(f"  收集表ID: {form_id}")
    except Exception as e:
        print(f"✗ 创建收集表失败: {e}")
        return

    # 步骤3: 设置收集表权限（任何人可填写）
    print("\n[步骤3] 设置收集表权限...")
    try:
        success = tencent_api.set_form_permission(form_id, "anyone_editable")
        if success:
            print("✓ 已设置为任何人可填写")
        else:
            print("⚠ 权限设置可能未生效")
    except Exception as e:
        print(f"⚠ 设置权限时出错: {e}")

    # 步骤4: 输出结果
    print("\n" + "=" * 60)
    print("✓ 收集表创建完成!")
    print("=" * 60)
    print(f"\n  收集表标题: {form_info.get('title')}")
    print(f"  收集表链接: {form_url}")
    print(f"\n  使用方法:")
    print("  1. 点击上方链接进入收集表")
    print("  2. 在腾讯文档中编辑收集表，添加需要收集的字段")
    print("  3. 发布后即可分享给他人填写")
    print("  4. 所有填写结果会自动汇总到关联表格中")
    print("=" * 60)


def demo_mode():
    """
    演示模式：使用模拟数据展示程序流程
    无需真实API密钥即可查看程序逻辑
    """
    print("=" * 60)
    print("演示模式: 展示程序执行流程")
    print("=" * 60)

    # 模拟商品数据
    demo_data = [
        ["商品编号", "商品名称", "规格", "单价", "库存", "供应商"],
        ["P001", "iPhone 15 Pro", "256GB 黑色", "8999", "100", "Apple官方"],
        ["P002", "MacBook Air M3", "16GB+512GB", "11999", "50", "Apple官方"],
        ["P003", "AirPods Pro 2", "降噪版", "1999", "200", "Apple官方"],
        ["P004", "小米14 Ultra", "16GB+1TB", "6499", "80", "小米科技"],
        ["P005", "华为Mate 60 Pro", "12GB+512GB", "6999", "60", "华为技术"],
    ]

    print("\n[模拟数据] 从腾讯文档读取的商品信息:")
    print("-" * 60)
    for row in demo_data:
        print("  |  " + "  |  ".join(row) + "  |")
    print("-" * 60)

    print("\n[步骤说明]")
    print("1. 程序会调用腾讯文档API读取指定表格的数据")
    print("2. 在金山文档中创建新的在线表格")
    print("3. 将读取的数据写入新创建的金山文档")

    print("\n[配置要求]")
    print("• 腾讯文档: 需要 client_id, client_secret, 以及用户授权")
    print("• 金山文档: 需要 app_id, app_key")
    print("\n请修改代码顶部的配置信息后运行正式模式")


if __name__ == "__main__":
    import sys

    # 检查是否有 --demo 参数
    if len(sys.argv) > 1 and sys.argv[1] == "--demo":
        demo_mode()
    else:
        # 检查配置是否已填写
        if TENCENT_CONFIG["client_id"].startswith("YOUR_"):
            print("⚠ 请先配置API密钥信息!")
            print("\n使用方法:")
            print("1. 查看演示模式: python tencent_to_kdocs.py --demo")
            print("2. 编辑代码顶部的 TENCENT_CONFIG")
            print("3. 运行: python tencent_to_kdocs.py")
            print("\n申请API密钥:")
            print("• 腾讯文档: https://docs.qq.com/open/")
        else:
            # 正式运行（只使用腾讯文档配置）
            transfer_data(TENCENT_CONFIG, {})
