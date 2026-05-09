#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品订购系统后端服务

功能：
1. 接收前端提交的订单
2. 保存订单到本地JSON文件
3. 发送邮件通知到指定邮箱
4. 提供订单查询/删除/导出API

使用方法：
    本地运行:  python server.py
    生产部署:  gunicorn server:app

配置方式（优先使用环境变量，未设置则使用默认值）：
    SMTP_USER      - 发件QQ邮箱地址
    SMTP_PASSWORD  - QQ邮箱授权码
    NOTIFY_EMAIL   - 接收通知的邮箱（默认 820389501@qq.com）
    ADMIN_PASSWORD - 管理后台访问密码
    PORT           - 服务端口（默认 5000）
"""

import json
import os
import smtplib
import tempfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, redirect, url_for, session, make_response
from flask_cors import CORS

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.environ.get('DATA_DIR', SCRIPT_DIR)
os.makedirs(DATA_DIR, exist_ok=True)
ORDERS_FILE = os.path.join(DATA_DIR, 'orders.json')

SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', '')
NOTIFY_EMAIL = os.environ.get('NOTIFY_EMAIL', '820389501@qq.com')
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.qq.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', '')
SERVER_PORT = int(os.environ.get('PORT', '5000'))

app = Flask(__name__, static_folder=SCRIPT_DIR, static_url_path='')
app.secret_key = os.environ.get('SECRET_KEY', os.urandom(24).hex())
CORS(app)


def load_orders():
    if not os.path.exists(ORDERS_FILE):
        return []
    try:
        with open(ORDERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []


def save_orders(orders):
    with open(ORDERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(orders, f, ensure_ascii=False, indent=2)


def send_email_notification(order):
    if not SMTP_USER or not SMTP_PASSWORD:
        print("⚠ 邮件未配置，跳过发送。请设置环境变量 SMTP_USER 和 SMTP_PASSWORD")
        return False

    subject = f"新订单 - {order['user']['name']} - {datetime.now().strftime('%m/%d %H:%M')}"

    items_text = ""
    for i, item in enumerate(order['items'], 1):
        items_text += f"  {i}. {item['name']}\n"
        items_text += f"     单价：¥{item['price']} × 数量：{item['quantity']} = ¥{item['subtotal']}\n"

    body = f"""收到新订单！

【订购人信息】
姓名：{order['user']['name']}
手机号：{order['user']['phone']}
下单店铺：{order['user'].get('shop', '')}
备注：{order['user'].get('remark', '')}

【商品清单】
{items_text}
【订单总计】
商品种类：{len(order['items'])} 种
商品总数：{sum(item['quantity'] for item in order['items'])} 件
订单总价：¥{order['totalPrice']}

下单时间：{order['submitTime']}
"""

    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = NOTIFY_EMAIL
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        order_json = json.dumps(order, ensure_ascii=False, indent=2)
        json_attachment = MIMEText(order_json, 'plain', 'utf-8')
        json_attachment.add_header(
            'Content-Disposition',
            'attachment',
            filename=f"order_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        msg.attach(json_attachment)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, NOTIFY_EMAIL, msg.as_string())

        print(f"✓ 邮件通知已发送至 {NOTIFY_EMAIL}")
        return True
    except Exception as e:
        print(f"✗ 邮件发送失败: {e}")
        return False


def check_admin_auth():
    if not ADMIN_PASSWORD:
        return True
    return session.get('admin_authed') == True


@app.route('/')
def index():
    return send_from_directory(SCRIPT_DIR, 'index.html')


@app.route('/admin_x7k9m2p4q8n1.html')
def admin_page():
    if not check_admin_auth():
        return send_from_directory(SCRIPT_DIR, 'admin_login.html')
    return send_from_directory(SCRIPT_DIR, 'admin_x7k9m2p4q8n1.html')


@app.route('/api/admin/login', methods=['POST'])
def admin_login():
    pwd = request.get_json().get('password', '')
    if pwd == ADMIN_PASSWORD:
        session['admin_authed'] = True
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "密码错误"}), 401


@app.route('/api/admin/logout', methods=['POST'])
def admin_logout():
    session.pop('admin_authed', None)
    return jsonify({"success": True})


@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory(SCRIPT_DIR, filename)


@app.route('/api/orders', methods=['GET'])
def get_orders():
    if ADMIN_PASSWORD and not check_admin_auth():
        return jsonify({"error": "未授权"}), 401
    orders = load_orders()
    return jsonify(orders)


@app.route('/api/orders', methods=['POST'])
def submit_order():
    order = request.get_json()

    if not order or not order.get('user') or not order.get('items'):
        return jsonify({"error": "订单数据不完整"}), 400

    order['serverTime'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    orders = load_orders()
    orders.append(order)
    save_orders(orders)

    print(f"✓ 新订单: {order['user']['name']} | {len(order['items'])}种商品 | ¥{order['totalPrice']}")

    send_email_notification(order)

    return jsonify({"success": True, "message": "订单提交成功"})


@app.route('/api/orders/<int:order_index>', methods=['DELETE'])
def delete_order(order_index):
    if ADMIN_PASSWORD and not check_admin_auth():
        return jsonify({"error": "未授权"}), 401
    orders = load_orders()
    if 0 <= order_index < len(orders):
        orders.pop(order_index)
        save_orders(orders)
        return jsonify({"success": True})
    return jsonify({"error": "订单不存在"}), 404


@app.route('/api/orders/clear', methods=['POST'])
def clear_orders():
    if ADMIN_PASSWORD and not check_admin_auth():
        return jsonify({"error": "未授权"}), 401
    save_orders([])
    return jsonify({"success": True})


@app.route('/api/orders/export', methods=['GET'])
def export_orders():
    if ADMIN_PASSWORD and not check_admin_auth():
        return jsonify({"error": "未授权"}), 401
    orders = load_orders()
    if not orders:
        return jsonify({"error": "没有订单"}), 404

    csv_lines = ['\uFEFF']
    csv_lines.append('订单号,订购人,手机号,下单店铺,商品名称,单价,数量,小计,订单总价,下单时间,备注\n')

    for index, order in enumerate(orders):
        for item in order['items']:
            fields = [
                index + 1,
                order['user']['name'],
                order['user']['phone'],
                order['user'].get('shop', ''),
                item['name'],
                item['price'],
                item['quantity'],
                item['subtotal'],
                order['totalPrice'],
                order['submitTime'],
                order['user'].get('remark', '')
            ]
            csv_lines.append(','.join(f'"{str(f).replace(chr(34), chr(34)+chr(34))}"' for f in fields) + '\n')

    from flask import Response
    csv_content = ''.join(csv_lines)
    return Response(
        csv_content,
        mimetype='text/csv;charset=utf-8;',
        headers={'Content-Disposition': f'attachment; filename=orders_{datetime.now().strftime("%Y%m%d")}.csv'}
    )


if __name__ == '__main__':
    print("=" * 50)
    print("商品订购系统后端服务")
    print("=" * 50)

    if not SMTP_USER:
        print("\n⚠ 邮件通知未配置！")
        print("请设置环境变量：")
        print("  export SMTP_USER=你的QQ邮箱")
        print("  export SMTP_PASSWORD=QQ邮箱授权码")
        print("\n获取授权码：QQ邮箱 → 设置 → 账户 → POP3/SMTP服务 → 生成授权码")
        print()

    if ADMIN_PASSWORD:
        print(f"✓ 管理后台密码保护已启用")
    else:
        print("⚠ 管理后台无密码保护（建议设置 ADMIN_PASSWORD 环境变量）")

    print(f"订单保存位置: {ORDERS_FILE}")
    print(f"邮件通知目标: {NOTIFY_EMAIL}")
    print(f"\n访问地址: http://localhost:{SERVER_PORT}")
    print(f"管理后台: http://localhost:{SERVER_PORT}/admin_x7k9m2p4q8n1.html")
    print("=" * 50)

    app.run(host='0.0.0.0', port=SERVER_PORT, debug=True)
