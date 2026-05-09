# 商品订购系统 - Code Wiki

## 目录

1. [项目概述](#1-项目概述)
2. [项目架构](#2-项目架构)
3. [文件结构](#3-文件结构)
4. [模块详解](#4-模块详解)
   - 4.1 [前端展示层](#41-前端展示层)
   - 4.2 [数据处理层](#42-数据处理层)
   - 4.3 [API集成层](#43-api集成层)
5. [关键类与函数](#5-关键类与函数)
6. [数据流](#6-数据流)
7. [依赖关系](#7-依赖关系)
8. [配置说明](#8-配置说明)
9. [运行方式](#9-运行方式)
10. [部署说明](#10-部署说明)

---

## 1. 项目概述

**项目名称**: 商品订购系统  
**项目类型**: 静态网站 + Python自动化工具集  
**主要功能**: 
- 商品信息展示与订购（前端页面）
- 订单数据本地存储与管理
- Excel商品数据自动处理与转换
- 腾讯文档/金山文档API集成（表单生成）

**目标用户**: 面向需要收集商品订购信息的商家/个人，支持通过GitHub Pages免费部署。

---

## 2. 项目架构

```
┌─────────────────────────────────────────────────────────────┐
│                        用户交互层                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │  index.html  │  │admin_x7k9... │  │   浏览器本地存储   │  │
│  │  商品订购页   │  │  订单管理后台 │  │  localStorage    │  │
│  └──────────────┘  └──────────────┘  └──────────────────┘  │
├─────────────────────────────────────────────────────────────┤
│                        数据处理层                            │
│  ┌──────────────────────────────────────────────────────┐  │
│  │              update_form.py                          │  │
│  │  Excel读取 → 图片提取 → 价格处理 → JSON生成 → Git推送  │  │
│  └──────────────────────────────────────────────────────┘  │
├─────────────────────────────────────────────────────────────┤
│                        API集成层                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │auto_form_... │  │wps_form_...  │  │tencent_to_kdocs  │  │
│  │腾讯智能表格   │  │ 金山表单API  │  │  腾讯→金山迁移    │  │
│  └──────────────┘  └──────────────┘  └──────────────────┘  │
│  ┌──────────────┐  ┌──────────────┐                        │
│  │wps_form_c... │  │  (其他工具)   │                        │
│  │ WPS表单Cookie│  │              │                        │
│  └──────────────┘  └──────────────┘                        │
├─────────────────────────────────────────────────────────────┤
│                        数据存储层                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │product_data  │  │products_with_│  │    images/       │  │
│  │   .json     │  │  images.json  │  │   商品图片资源    │  │
│  └──────────────┘  └──────────────┘  └──────────────────┘  │
│  ┌──────────────┐  ┌──────────────┐                        │
│  │ 20260510截止 │  │   README.md   │                        │
│  │   .xlsx     │  │   项目说明    │                        │
│  └──────────────┘  └──────────────┘                        │
└─────────────────────────────────────────────────────────────┘
```

### 架构特点

- **前后端分离**: 前端为纯静态HTML/CSS/JS，无后端服务器依赖
- **本地存储**: 订单数据使用浏览器 `localStorage` 持久化
- **自动化流水线**: Python脚本实现Excel→JSON→GitHub的自动化发布
- **多平台API适配**: 支持腾讯文档、金山文档(WPS)等多个平台

---

## 3. 文件结构

```
product_form/
│
├── index.html                    # 商品订购主页面（用户端）
├── admin_x7k9m2p4q8n1.html       # 订单管理后台页面
├── README.md                     # 项目说明文档
│
├── update_form.py                # Excel数据自动处理脚本（核心工具）
│
├── auto_form_generator.py        # 腾讯文档智能表格自动生成
├── wps_form_generator.py         # 金山表单自动生成（OAuth版）
├── wps_form_creator.py           # WPS表单创建（Cookie版）
├── tencent_to_kdocs.py           # 腾讯文档→金山文档数据迁移
│
├── product_data.json             # 原始商品数据（表格格式）
├── products_with_images.json     # 处理后商品数据（带图片路径）
│
├── images/                       # 商品图片资源目录
│   ├── product_2.png
│   ├── product_3.png
│   └── ... (共约70张)
│
├── 20260510截止.xlsx             # 商品信息源Excel文件
│
└── content.md                    # 内容文档（漫画脚本，非项目相关）
```

---

## 4. 模块详解

### 4.1 前端展示层

#### 4.1.1 index.html - 商品订购页面

**职责**: 面向订购用户的商品展示与下单界面

**主要功能**:
- 加载并展示 `products_with_images.json` 中的商品列表
- 支持商品数量输入与实时价格计算
- 收集用户信息（姓名、手机号、微信号、备注）
- 订单确认弹窗与提交
- 订单数据保存至 `localStorage`
- 生成可复制的订单文本并自动复制到剪贴板

**关键DOM结构**:
```html
.container
├── .header              # 页面标题
├── .user-info           # 用户信息表单
├── .products-table      # 商品列表表格
├── .total-section       # 订单汇总区域
├── .submit-btn          # 提交按钮
└── #confirmModal        # 订单确认弹窗
```

**核心JavaScript函数**:

| 函数名 | 说明 |
|--------|------|
| `loadProducts()` | 异步加载商品JSON数据 |
| `renderProducts()` | 渲染商品列表到表格 |
| `updateQuantity(id, qty, price)` | 更新商品数量 |
| `updateSummary()` | 更新订单汇总统计 |
| `submitOrder()` | 打开订单确认弹窗 |
| `confirmSubmit()` | 确认提交订单 |
| `generateOrderText(order)` | 生成可读的订单文本 |

**数据存储格式**:
```javascript
// localStorage key: 'product_orders_v1'
{
  user: { name, phone, wechat, remark },
  items: [{ id, name, price, quantity, subtotal }],
  totalPrice: number,
  submitTime: string
}
```

#### 4.1.2 admin_x7k9m2p4q8n1.html - 订单管理后台

**职责**: 商家管理订单的后台界面

**主要功能**:
- 从 `localStorage` 读取所有订单
- 统计面板：订单总数、商品总数、订单总金额
- 订单列表展示与详情查看
- 单条订单删除
- 批量清空所有订单
- 导出订单为CSV文件

**核心JavaScript函数**:

| 函数名 | 说明 |
|--------|------|
| `loadOrders()` | 加载订单数据 |
| `displayOrders(orders)` | 渲染订单列表 |
| `updateStats(orders)` | 更新统计面板 |
| `viewOrderDetail(index)` | 展开/收起订单详情 |
| `deleteOrder(index)` | 删除指定订单 |
| `clearAllOrders()` | 清空所有订单 |
| `exportAllOrders()` | 导出CSV文件 |

**注意**: 由于使用 `localStorage`，订单数据仅在当前浏览器中存储，不同设备/浏览器间不共享。

---

### 4.2 数据处理层

#### 4.2.1 update_form.py - Excel数据处理核心

**职责**: 将Excel商品数据转换为前端可用的JSON格式，并自动推送到GitHub

**主要流程**:
```
读取Excel → 提取图片 → 处理价格(+10) → 生成JSON → Git推送
```

**函数清单**:

| 函数名 | 参数 | 返回值 | 说明 |
|--------|------|--------|------|
| `extract_price(price_str)` | `str` | `str` | 提取价格数字并+10，支持范围价格取平均 |
| `extract_images_from_excel(excel_path, output_dir)` | `str, str` | `set` | 从Excel提取图片，返回有图片的行号集合 |
| `process_excel(excel_path)` | `str` | `list` | 主处理函数：读取Excel、提取数据、生成JSON |
| `push_to_github()` | 无 | `bool` | 自动执行git add/commit/push |
| `main()` | 无 | 无 | 入口函数，支持命令行参数指定Excel文件 |

**Excel数据映射**:

| Excel列 | JSON字段 | 说明 |
|---------|----------|------|
| A (第1列) | `发售日` | 发售日期 |
| B (第2列) | `截止日` | 截止日期 |
| C (第3列) | `厂商` | 制造商 |
| D (第4列) | `品名` | 商品名称 |
| E (第5列) | `价格` | 处理后的价格（+10） |
| F (第6列) | `定金` | 定金金额 |
| G (第7列) | `下单店铺` | 原"备注"字段 |
| 图片 | `图片` | 提取的图片路径 |

**价格处理逻辑**:
- 移除货币符号（¥, $）和逗号
- 范围价格（如 `30-40`）取平均值后+10
- 普通价格直接+10
- 转换失败则保留原值

**使用方式**:
```bash
# 自动查找目录中的Excel文件
python update_form.py

# 指定Excel文件
python update_form.py 20260510截止.xlsx
```

---

### 4.3 API集成层

#### 4.3.1 auto_form_generator.py - 腾讯文档智能表格

**职责**: 通过腾讯文档API自动创建智能表格（Smartsheet）收集表

**核心类**: `TencentDocsAPI`

| 方法 | 说明 |
|------|------|
| `get_sheet_data(file_id, sheet_id, range_str)` | 读取腾讯文档表格数据 |
| `create_smartsheet(title)` | 创建智能表格 |
| `add_smartsheet_fields(file_id, sheet_id, fields)` | 添加字段定义 |
| `add_smartsheet_records(file_id, sheet_id, records)` | 批量写入数据 |
| `set_file_permission(file_id, permission)` | 设置文件权限 |
| `get_sheet_info(file_id)` | 获取子表信息 |

**字段映射配置** (`FIELD_MAPPING`):

| 标准字段 | 匹配关键词 |
|----------|-----------|
| 商品名称 | 商品名、名称、产品名称、name、title |
| 商品价格 | 价格、单价、售价、price、cost |
| 商品图片 | 图片、商品图、image、photo、pic |
| 商品规格 | 规格、型号、配置、spec、model |
| 商品库存 | 库存、数量、stock、quantity |
| 供应商 | 供应商、厂家、品牌、supplier、brand |
| 发售日期 | 发售日、上市日期、date、time |
| 备注 | 备注、说明、描述、desc |

**字段类型映射**:

| 字段 | 类型代码 | 类型说明 |
|------|----------|----------|
| 商品名称 | 1 | 文本 |
| 商品价格 | 2 | 数字 |
| 商品图片 | 8 | 超链接(URL) |
| 发售日期 | 4 | 日期 |

**使用方式**:
```bash
# 演示模式（无需API密钥）
python auto_form_generator.py --demo

# 正式运行（需配置TENCENT_CONFIG）
python auto_form_generator.py
```

#### 4.3.2 wps_form_generator.py - 金山表单生成（OAuth版）

**职责**: 通过金山文档开放平台API创建表单

**核心类**:

| 类 | 说明 |
|----|------|
| `TencentDocsAPI` | 腾讯文档API封装（读取源数据） |
| `KDocsFormAPI` | 金山表单API封装（创建表单） |

**金山表单API签名算法**:
```
签名 = sha1(app_key + Content-Md5 + Content-Type + DATE)
Authorization = "WPS-2:" + app_id + ":" + 签名
```

**字段类型映射**:

| 标准字段 | 表单类型 |
|----------|----------|
| 商品名称 | input (文本输入) |
| 商品价格 | floatInput (数字输入) |
| 商品图片 | newImage (图片上传) |
| 发售日期 | date (日期选择) |
| 备注 | multiInput (多行文本) |

#### 4.3.3 wps_form_creator.py - WPS表单创建（Cookie版）

**职责**: 基于WPS Cookie直接调用内部API创建表单

**特点**:
- 使用 `wps_sid` Cookie进行身份验证
- 直接调用 `f-api.wps.cn` 内部接口
- 无需OAuth授权流程

**核心函数**:

| 函数 | 说明 |
|------|------|
| `create_form(title, subtitle)` | 创建草稿表单 |
| `add_text_question(form_id, title, required)` | 添加文本问题 |
| `add_image_question(form_id, title, image_url)` | 添加图片问题 |
| `publish_form(form_id)` | 发布表单 |
| `create_product_form(title, products)` | 完整流程：创建→添加问题→发布 |

#### 4.3.4 tencent_to_kdocs.py - 数据迁移工具

**职责**: 从腾讯文档读取数据，创建腾讯收集表

**核心类**:

| 类 | 说明 |
|----|------|
| `TencentDocsAPI` | 腾讯文档API（含OAuth2.0支持） |
| `KDocsAPI` | 金山文档API（含WPS-2签名） |

**主要方法**:

| 方法 | 说明 |
|------|------|
| `TencentDocsAPI.get_access_token(code)` | OAuth2.0授权码换Token |
| `TencentDocsAPI.create_form(title)` | 创建收集表 |
| `TencentDocsAPI.set_form_permission(id, scope)` | 设置权限为任何人可编辑 |
| `KDocsAPI.create_spreadsheet(filename)` | 创建金山在线表格 |
| `KDocsAPI.update_cells(token, idx, data)` | 写入单元格数据 |

---

## 5. 关键类与函数

### 5.1 Python模块依赖关系

```
update_form.py
├── openpyxl          # Excel读写
├── json              # JSON序列化
├── os, sys, shutil   # 系统操作
├── subprocess        # Git命令执行
└── pathlib           # 路径处理

auto_form_generator.py
├── requests          # HTTP请求
├── json              # JSON处理
└── typing            # 类型注解

wps_form_generator.py
├── requests          # HTTP请求
├── hashlib           # MD5/SHA1签名
├── json              # JSON处理
├── datetime          # 日期格式化
├── uuid              # UUID生成
└── typing            # 类型注解

tencent_to_kdocs.py
├── requests          # HTTP请求
├── hashlib, hmac     # 签名计算
├── base64            # Base64编码
├── json              # JSON处理
├── datetime          # 日期处理
├── urllib.parse      # URL编码
└── typing            # 类型注解
```

### 5.2 函数调用链

#### 商品数据更新流程
```
main()
├── process_excel(excel_path)
│   ├── extract_images_from_excel(excel_path, images_dir)
│   │   └── openpyxl.load_workbook()
│   └── extract_price(raw_price) [每行调用]
└── push_to_github()
    ├── git add .
    ├── git commit -m '更新商品数据和图片'
    └── git push origin main
```

#### 腾讯智能表格生成流程
```
auto_generate_form(config)
├── TencentDocsAPI.__init__()
├── api.get_sheet_data(file_id, sheet_id, range)
├── analyze_headers(headers, FIELD_MAPPING)
├── api.create_smartsheet(title)
├── api.get_sheet_info(smartsheet_id)
├── create_smartsheet_fields(column_mapping, headers)
├── api.add_smartsheet_fields(file_id, sheet_id, fields)
├── convert_data_to_records(data, column_mapping)
├── api.add_smartsheet_records(file_id, sheet_id, records) [批量]
└── api.set_file_permission(file_id, "anyone_editable")
```

---

## 6. 数据流

### 6.1 商品数据流

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  Excel源文件 │ ──→ │ update_form │ ──→ │products_with│ ──→ │  index.html  │
│  .xlsx      │     │    .py      │     │_images.json  │     │  加载展示    │
└─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘
                           │
                           ↓
                    ┌─────────────┐
                    │   images/   │
                    │  商品图片   │
                    └─────────────┘
```

### 6.2 订单数据流

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│   用户下单   │ ──→ │ localStorage│ ──→ │ admin后台   │
│  index.html │     │ (浏览器存储) │     │ 查看/管理   │
└─────────────┘     └─────────────┘     └─────────────┘
                           │
                           ↓
                    ┌─────────────┐
                    │  CSV导出    │
                    │ (文件下载)  │
                    └─────────────┘
```

### 6.3 API数据流（腾讯文档→表单）

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  腾讯文档   │ ──→ │ Python脚本  │ ──→ │  收集表/表单 │
│  在线表格   │     │  API调用    │     │  (腾讯/金山) │
└─────────────┘     └─────────────┘     └─────────────┘
```

---

## 7. 依赖关系

### 7.1 外部依赖

| 依赖 | 版本要求 | 用途 | 安装命令 |
|------|----------|------|----------|
| Python | 3.7+ | 脚本运行环境 | - |
| openpyxl | latest | Excel文件处理 | `pip install openpyxl` |
| requests | latest | HTTP API调用 | `pip install requests` |
| Git | any | 代码推送 | - |

### 7.2 API平台依赖

| 平台 | 认证方式 | 所需凭证 |
|------|----------|----------|
| 腾讯文档 | OAuth2.0 / 直接Token | client_id, access_token, open_id |
| 金山文档 | WPS-2签名 / OAuth2.0 | app_id, app_key, access_token |
| WPS内部 | Cookie | wps_sid |

### 7.3 文件依赖关系

```
index.html
└── depends on: products_with_images.json
    └── generated by: update_form.py
        └── depends on: *.xlsx (Excel文件)
        └── outputs to: images/*.png

admin_x7k9m2p4q8n1.html
└── reads from: localStorage (key: 'product_orders_v1')
    └── written by: index.html (submit order)
```

---

## 8. 配置说明

### 8.1 腾讯文档配置 (TENCENT_CONFIG)

```python
TENCENT_CONFIG = {
    "client_id": "",        # 腾讯文档应用ID
    "file_id": "",          # 源表格文件ID
    "sheet_id": "",         # 工作表ID
    "range": "A1:H70",      # 数据范围
    "access_token": "",     # 访问令牌
    "open_id": "",          # 用户OpenID
    "output_title": "",     # 输出表格标题
}
```

### 8.2 金山文档配置 (KDOCS_CONFIG)

```python
KDOCS_CONFIG = {
    "app_id": "",           # 金山应用ID
    "app_key": "",          # 金山应用密钥
    "access_token": "",     # OAuth访问令牌
}
```

### 8.3 WPS Cookie配置

```python
WPS_SID = ""  # WPS登录后的Cookie值
```

---

## 9. 运行方式

### 9.1 前端页面运行

**本地预览**:
```bash
# 使用Python内置HTTP服务器
cd product_form
python -m http.server 8000

# 或使用Node.js的http-server
npx http-server -p 8000
```

**访问地址**:
- 订购页面: `http://localhost:8000/index.html`
- 管理后台: `http://localhost:8000/admin_x7k9m2p4q8n1.html`

### 9.2 数据处理脚本运行

**更新商品数据**:
```bash
# 自动查找Excel文件
python update_form.py

# 指定Excel文件
python update_form.py 20260510截止.xlsx
```

**生成腾讯智能表格**:
```bash
# 演示模式
python auto_form_generator.py --demo

# 正式运行（需先配置API密钥）
python auto_form_generator.py
```

**生成金山表单**:
```bash
# 演示模式
python wps_form_generator.py --demo

# 正式运行
python wps_form_generator.py
```

### 9.3 GitHub Pages部署

1. 将代码推送到GitHub仓库
2. 启用GitHub Pages（Settings → Pages → Source: main branch）
3. 访问 `https://your-username.github.io/repo-name/`

---

## 10. 部署说明

### 10.1 完整部署流程

```bash
# 1. 克隆仓库
git clone <repository-url>
cd product_form

# 2. 安装Python依赖
pip install openpyxl requests

# 3. 准备商品数据
# 将Excel文件放入项目目录

# 4. 运行数据处理
python update_form.py

# 5. 提交到GitHub
git add .
git commit -m "初始化商品数据"
git push origin main

# 6. 启用GitHub Pages（在GitHub网站设置中操作）
```

### 10.2 日常更新流程

```bash
# 1. 更新Excel文件
# 2. 运行脚本
python update_form.py

# 3. 脚本会自动推送，或手动推送
git push origin main

# 4. 等待GitHub Pages自动部署（约1-2分钟）
```

### 10.3 环境要求

| 环境 | 要求 |
|------|------|
| Python | 3.7 或更高 |
| 浏览器 | 支持ES6+ 和 localStorage |
| 网络 | 可访问GitHub（用于部署） |
| 可选 | 腾讯文档/金山文档开发者账号 |

---

## 附录

### A. 项目特点

1. **零成本部署**: 使用GitHub Pages免费托管静态页面
2. **数据本地化**: 订单数据存储在浏览器本地，无需数据库
3. **自动化处理**: Excel→JSON→GitHub全自动流水线
4. **多平台适配**: 支持腾讯文档、金山文档等多个平台API
5. **移动端友好**: 响应式设计，支持手机浏览器访问

### B. 注意事项

1. **数据安全**: `localStorage` 数据仅在当前浏览器有效，建议定期导出备份
2. **并发问题**: 多用户同时访问时，订单数据不共享
3. **API密钥**: 生产环境应使用环境变量存储敏感信息，避免硬编码
4. **图片路径**: 确保 `images/` 目录中的图片与 `products_with_images.json` 中的路径匹配
5. **价格处理**: `update_form.py` 会自动将价格+10，确认业务逻辑是否符合预期

### C. 扩展建议

1. **后端服务**: 可添加Node.js/Python后端，将订单存储到数据库
2. **用户认证**: 添加登录系统，支持多用户管理
3. **实时同步**: 使用WebSocket或轮询实现多设备订单同步
4. **邮件通知**: 订单提交后自动发送邮件通知
5. **微信支付**: 集成微信支付，实现完整电商闭环
