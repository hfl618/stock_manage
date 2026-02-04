# -*- coding: utf-8 -*-
# -------------------------- 导入依赖（新增二维码和图片处理包） --------------------------
import os
import socket
import sys
import logging
import zipfile
import win32com.client
import pythoncom
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file, jsonify, \
    send_from_directory
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import uuid
import json
import qrcode

import os
os.environ['FLASK_ENV'] = 'development'
os.environ['FLASK_DEBUG'] = '1'

# -------------------------- 基础配置（稳定版，修复数据库路径核心问题） --------------------------
# 工具判断
IS_WINDOWS = sys.platform.startswith('win')
win32 = win32com.client if IS_WINDOWS else None

# 数量预警阈值（可自定义）
QUANTITY_WARN_LOW = 10  # ≤10 红色预警
QUANTITY_WARN_MID = 20  # 11-20 黄色预警

# 日志配置（简单实用）
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Flask初始化
app = Flask(__name__)
app.config['SECRET_KEY'] = 'component_stock_secure_2026'
# 核心修复：指定数据库绝对路径到项目根目录，避免Flask默认放到instance目录
BASE_DIR = os.path.dirname(__file__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(BASE_DIR, "instance/component.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10M文件上传限制
app.config['JSON_AS_ASCII'] = False
db = SQLAlchemy(app)

# 目录定义（自动创建，新增二维码目录）
STATIC_FOLDER = os.path.join(BASE_DIR, 'static')
IMG_FOLDER = os.path.join(STATIC_FOLDER, 'img')  # 图片目录
ATTACH_FOLDER = os.path.join(STATIC_FOLDER, 'attach')  # 附件目录
QRCODE_FOLDER = os.path.join(STATIC_FOLDER, 'qrcode')  # 二维码目录
BACKUP_FOLDER = os.path.join(BASE_DIR, 'backup')  # 备份目录

# 核心修改：数据库文件固定在Flask instance目录
DB_FILE = os.path.join(app.instance_path, 'component.db')

# 允许的文件格式
ALLOWED_IMG_EXT = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}
ALLOWED_ATTACH_EXT = {'pdf', 'doc', 'docx', 'xls', 'xlsx', 'zip', 'rar', 'txt', 'csv'}

# 自动创建必要目录（新增二维码目录）
for folder in [STATIC_FOLDER, IMG_FOLDER, ATTACH_FOLDER, QRCODE_FOLDER, BACKUP_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)
        logger.info(f"自动创建文件夹：{folder}")

# 默认帮助文件（简洁版，无多余内容）
HELP_FILE = os.path.join(BASE_DIR, 'help.txt')


# -------------------------- 数据库模型（新增二维码路径字段） --------------------------
class Component(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False, default='未知')  # 品类
    model = db.Column(db.String(200), nullable=False, default='未知型号')  # 型号规格
    package = db.Column(db.String(50), nullable=False, default='未知封装')  # 封装
    supplier = db.Column(db.String(100), default='未知供应商')  # 供应商
    quantity = db.Column(db.Integer, default=1)  # 数量（保留独立字段）
    unit = db.Column(db.String(20), default='个')  # 单位（保留独立字段）
    location = db.Column(db.String(100), default='未知位置')  # 存放位置
    price = db.Column(db.Float, default=0.00)  # 采购单价
    buy_time = db.Column(db.String(20), default=datetime.now().strftime('%Y-%m-%d'))  # 采购时间
    channel = db.Column(db.String(50), default='未知')  # 采购渠道
    remark = db.Column(db.String(200), default='无')  # 备注
    img_path = db.Column(db.String(255), default='')  # 图片路径
    attach_path = db.Column(db.String(255), default='')  # 附件路径
    qrcode_path = db.Column(db.String(255), default='')  # 二维码路径（新增）

    def __repr__(self):
        return f'<Component {self.id} - {self.category} {self.model}>'

    def get_file_prefix(self):
        """生成唯一文件前缀，避免重名"""
        invalid = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', ' ']
        pre = f"{self.category}_{self.model}_{self.package}"
        for c in invalid:
            pre = pre.replace(c, '_')
        return pre[:80]  # 限制长度，避免文件名过长

    def get_qr_data(self):
        """生成二维码数据字符串"""
        qr_data = {
            'id': self.id,
            'category': self.category,
            'model': self.model,
            'package': self.package,
            'supplier': self.supplier,
            'quantity': self.quantity,
            'unit': self.unit,
            'location': self.location,
            'price': self.price,
            'buy_time': self.buy_time,
            'channel': self.channel,
            'remark': self.remark
        }
        return json.dumps(qr_data, ensure_ascii=False)


# -------------------------- 核心工具函数（新增二维码功能） --------------------------
# 数量预警样式
def get_quantity_css(quantity):
    if quantity <= QUANTITY_WARN_LOW:
        return "text-danger fw-bold"
    elif quantity <= QUANTITY_WARN_MID:
        return "text-warning fw-bold"
    else:
        return "text-success"


# 检测重复元器件（料号/型号唯一匹配）- 修复：增强兼容性
def is_duplicate(item):
    """
    核心：判断单条BOM数据是否重复，基于【料号/型号】唯一匹配
    入参：字典格式
    出参：存在重复返回Component对象，不存在返回None
    """
    try:
        # 关键配置：基于料号/型号匹配
        unique_field = 'model'  # 改为你的实际字段：part_no/料号/code/sku等

        # 强校验：入参不是字典 → 直接返回None
        if not isinstance(item, dict):
            logger.debug(f"is_duplicate入参不是字典：{type(item)}")
            return None

        # 校验是否包含唯一关键字段，无值/无字段 → 视为新数据
        if unique_field not in item or not str(item[unique_field]).strip():
            return None

        # 数据库精准匹配：转字符串+去空格
        match_value = str(item[unique_field]).strip()
        duplicate_item = db.session.query(Component).filter(
            getattr(Component, unique_field) == match_value
        ).first()

        return duplicate_item
    except Exception as e:
        logger.error(f"单条BOM数据重复检测异常：{str(e)}", exc_info=True)
        return None  # 异常时视为新数据


# 表格内去重（BOM导入用）- 修复：适配字典格式入参
def remove_table_dup(data_list):
    """
    修复：入参为映射后的字典列表
    根据品类+型号+封装去重，保留第一条数据
    """
    if not isinstance(data_list, list) or not all(isinstance(d, dict) for d in data_list):
        logger.error("remove_table_dup入参不是「字典列表」")
        return data_list
    seen = set()
    unique = []
    for d in data_list:
        # 使用品类+型号+封装作为唯一键
        key = (
            d.get('category', '').strip(),
            d.get('model', '').strip(),
            d.get('package', '').strip()
        )
        if key not in seen and key[0] and key[1] and key[2]:  # 非空才去重
            seen.add(key)
            unique.append(d)
    return unique


# 文件验证
def allowed_file(filename, ext_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ext_set


# 保存上传文件
def save_file(file, save_dir, ext_set, comp_pre):
    if not file or file.filename == '':
        return ''
    if not allowed_file(file.filename, ext_set):
        flash(f"文件格式不支持！仅允许：{','.join(ext_set)}", "danger")
        return ''
    # 生成唯一文件名
    ext = file.filename.rsplit('.', 1)[1].lower()
    unique_name = f"{comp_pre}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:6]}.{ext}"
    file_path = os.path.join(save_dir, unique_name)
    file.save(file_path)
    # 返回相对路径
    rel_path = os.path.relpath(file_path, BASE_DIR).replace('\\', '/')
    logger.info(f"文件保存：{rel_path}")
    return rel_path


# 删除文件
def delete_file(file_path):
    if not file_path:
        return
    abs_path = os.path.join(BASE_DIR, file_path)
    if os.path.exists(abs_path):
        try:
            os.remove(abs_path)
            logger.info(f"文件删除：{abs_path}")
        except Exception as e:
            logger.warning(f"文件删除失败：{abs_path}，原因：{str(e)}")
            pass


# 生成二维码
def generate_qrcode(component):
    """为元器件生成二维码"""
    try:
        # 获取二维码数据
        qr_data = component.get_qr_data()

        # 创建二维码对象
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(qr_data)
        qr.make(fit=True)

        # 生成二维码图片
        img = qr.make_image(fill_color="black", back_color="white")

        # 生成文件名
        filename = f"QR_{component.id}_{component.model.replace('/', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"
        filepath = os.path.join(QRCODE_FOLDER, filename)

        # 保存图片
        img.save(filepath)

        # 返回相对路径
        rel_path = os.path.relpath(filepath, BASE_DIR).replace('\\', '/')

        # 更新组件二维码路径
        component.qrcode_path = rel_path
        db.session.commit()

        logger.info(f"二维码生成成功：{rel_path}")
        return rel_path
    except Exception as e:
        logger.error(f"二维码生成失败：{str(e)}")
        return ""


# 获取或生成二维码
def get_or_generate_qrcode(component_id):
    """获取二维码，如果没有则生成"""
    component = Component.query.get(component_id)
    if not component:
        return ""

    # 如果已有二维码文件且文件存在
    if component.qrcode_path and os.path.exists(os.path.join(BASE_DIR, component.qrcode_path)):
        return component.qrcode_path

    # 生成新二维码
    return generate_qrcode(component)


# 读取帮助文件
def get_help_content():
    try:
        with open(HELP_FILE, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "使用说明文件丢失，已自动重新创建！"


# 清理残留文件（修改：包含二维码）
def clean_residual_files():
    """清理无关联的图片、附件和二维码：数据库中不存在的文件直接删除"""
    try:
        # 获取数据库中所有有效文件路径
        valid_files = set()
        comps = Component.query.all()
        for c in comps:
            if c.img_path: valid_files.add(c.img_path)
            if c.attach_path: valid_files.add(c.attach_path)
            if c.qrcode_path: valid_files.add(c.qrcode_path)
        logger.info(f"数据库中有效文件数：{len(valid_files)}")

        # 扫描并清理图片目录
        img_del_count = 0
        for root, _, files in os.walk(IMG_FOLDER):
            for file in files:
                file_path = os.path.relpath(os.path.join(root, file), BASE_DIR).replace('\\', '/')
                if file_path not in valid_files:
                    delete_file(file_path)
                    img_del_count += 1

        # 扫描并清理附件目录
        attach_del_count = 0
        for root, _, files in os.walk(ATTACH_FOLDER):
            for file in files:
                file_path = os.path.relpath(os.path.join(root, file), BASE_DIR).replace('\\', '/')
                if file_path not in valid_files:
                    delete_file(file_path)
                    attach_del_count += 1

        # 扫描并清理二维码目录
        qrcode_del_count = 0
        for root, _, files in os.walk(QRCODE_FOLDER):
            for file in files:
                file_path = os.path.relpath(os.path.join(root, file), BASE_DIR).replace('\\', '/')
                if file_path not in valid_files:
                    delete_file(file_path)
                    qrcode_del_count += 1

        total_del = img_del_count + attach_del_count + qrcode_del_count
        logger.info(
            f"残留文件清理完成：图片{img_del_count}个，附件{attach_del_count}个，二维码{qrcode_del_count}个，总计{total_del}个")
        return True, f"清理成功！共删除残留文件{total_del}个（图片{img_del_count}个+附件{attach_del_count}个+二维码{qrcode_del_count}个）"
    except Exception as e:
        logger.error(f"残留文件清理失败：{str(e)}")
        return False, f"清理失败：{str(e)}"


# -------------------------- 开机自启函数（Windows专属，稳定版） --------------------------
def get_startup_path():
    return os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')


def is_auto_start():
    if not IS_WINDOWS or not win32:
        return False
    lnk_path = os.path.join(get_startup_path(), "元器件库存管理工具.lnk")
    return os.path.exists(lnk_path)


def create_auto_start():
    try:
        pythoncom.CoInitialize()
        startup = get_startup_path()
        if not startup:
            return False, "获取开机启动目录失败"
        lnk_path = os.path.join(startup, "元器件库存管理工具.lnk")
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        shortcut.TargetPath = sys.executable
        shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
        shortcut.WorkingDirectory = BASE_DIR
        shortcut.Save()
        pythoncom.CoUninitialize()
        return True, "开机自启设置成功（重启电脑生效，需管理员权限）"
    except Exception as e:
        return False, f"开机自启失败：{str(e)}（请以管理员身份运行程序）"


def delete_auto_start():
    if not IS_WINDOWS or not is_auto_start():
        return True, "未开启开机自启"
    try:
        os.remove(os.path.join(get_startup_path(), "元器件库存管理工具.lnk"))
        return True, "开机自启已关闭"
    except:
        return False, "关闭开机自启失败（请以管理员身份运行）"


# -------------------------- 备份恢复核心函数（修改：包含二维码） --------------------------
def backup_all_data():
    """修改：无下载弹窗，直接保存到backup默认目录，数据库指向instance，包含二维码"""
    try:
        # 1. 检测数据库文件是否存在（instance目录）
        if not os.path.exists(DB_FILE):
            flash("暂无数据可备份！数据库文件尚未创建", "warning")
            return None

        # 2. 检测是否有元器件数据
        comp_count = Component.query.count()
        if comp_count == 0:
            flash("暂无元器件数据，无需备份！添加数据后再尝试", "info")
            return None

        # 3. 生成备份文件名，保存到backup默认目录
        backup_name = f"元器件库存备份_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
        backup_path = os.path.join(BACKUP_FOLDER, backup_name)

        # 4. 打包备份（instance里的数据库+图片+附件+二维码）
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # 核心修改：打包instance目录下的数据库文件
            zf.write(DB_FILE, os.path.basename(DB_FILE))
            # 备份图片
            for root, _, files in os.walk(IMG_FOLDER):
                for file in files:
                    file_path = os.path.join(root, file)
                    zf.write(file_path, os.path.relpath(file_path, BASE_DIR))
            # 备份附件
            for root, _, files in os.walk(ATTACH_FOLDER):
                for file in files:
                    file_path = os.path.join(root, file)
                    zf.write(file_path, os.path.relpath(file_path, BASE_DIR))
            # 备份二维码（新增）
            for root, _, files in os.walk(QRCODE_FOLDER):
                for file in files:
                    file_path = os.path.join(root, file)
                    zf.write(file_path, os.path.relpath(file_path, BASE_DIR))

        logger.info(f"备份成功：{backup_path}，包含{comp_count}条元器件数据")
        flash(f"备份成功！共{comp_count}条数据，备份包已保存至【{BACKUP_FOLDER}】目录", "success")
        return backup_path
    except Exception as e:
        flash(f"备份失败：{str(e)}", "danger")
        logger.error(f"备份失败：{str(e)}")
        return None


def validate_backup_zip(zip_path):
    """简化验证：仅检测数据库文件"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            if 'component.db' in zf.namelist():
                return True, "备份文件验证通过"
            else:
                return False, "备份文件缺少数据库（component.db）"
    except:
        return False, "备份文件损坏或不是有效ZIP文件"


# -------------------------- BOM导入常量（稳定版+新增依旧导入选项） --------------------------
SYSTEM_FIELDS = [
    ('', '不映射（跳过）'),
    ('category', '品类【必填】'),
    ('model', '型号规格【必填】'),
    ('package', '封装【必填】'),
    ('supplier', '供应商'),
    ('quantity', '数量'),
    ('unit', '单位'),
    ('location', '存放位置'),
    ('price', '采购单价'),
    ('buy_time', '采购时间'),
    ('channel', '采购渠道'),
    ('remark', '备注'),
]
REQUIRED_FIELDS = ['category', 'model', 'package']


def parse_table_data(source, source_type):
    """解析粘贴/Excel数据"""
    columns, preview, raw_data = [], [], []
    try:
        if source_type == 'paste':
            lines = [l.strip() for l in source.split('\n') if l.strip()]
            if not lines:
                return columns, preview, raw_data, "无有效粘贴数据", 0
            # 按制表符分割
            col_num = len(lines[0].split('\t'))
            columns = [f'列{i + 1}' for i in range(col_num)]
            for line in lines:
                row = line.split('\t') + [''] * (col_num - len(line.split('\t')))
                raw_data.append(row[:col_num])
            preview = raw_data[:3]  # 只显示前三行
            row_count = len(raw_data)  # 计算行数
        elif source_type == 'excel':
            if not allowed_file(source.filename, {'xlsx'}):
                return columns, preview, raw_data, "仅支持xlsx格式Excel文件", 0
            df = pd.read_excel(source, engine='openpyxl')
            df.columns = [f'列{i + 1}' if pd.isna(c) or not str(c).strip() else str(c).strip() for i, c in
                          enumerate(df.columns)]
            columns = list(df.columns)
            raw_data = df.fillna('').values.tolist()
            preview = raw_data[:3]  # 只显示前三行
            row_count = len(raw_data)  # 计算行数
        return columns, preview, raw_data, "", row_count
    except Exception as e:
        logger.error(f"表格解析失败：{str(e)}")
        return columns, preview, raw_data, f"解析失败：{str(e)}", 0


def map_table_data(raw_data, columns, mapping, batch_vals):
    """映射表格数据为字典列表"""
    data_list, errors = [], []

    if not isinstance(mapping, dict) or not isinstance(batch_vals, dict) or not isinstance(raw_data, list):
        errors.append("数据格式错误")
        return data_list, errors

    # 检查必填字段是否映射
    mapped_fields = [v for v in mapping.values() if v]
    for f in REQUIRED_FIELDS:
        if f not in mapped_fields:
            errors.append(f"缺少必填字段映射：{f}")
    if errors:
        return data_list, errors

    # 处理批量值
    batch = {}
    for k, v in batch_vals.items():
        if not v:
            continue
        try:
            if k == 'quantity':
                batch[k] = int(v)
            elif k == 'price':
                batch[k] = float(v)
            else:
                batch[k] = v.strip()
        except ValueError:
            errors.append(f"批量{[i[1] for i in SYSTEM_FIELDS if i[0] == k][0]}必须为有效数字")
    if errors:
        return data_list, errors

    # 映射原始数据为字典
    for idx, row in enumerate(raw_data, 1):
        # 初始化默认值
        d = {
            'category': '未知', 'model': '未知型号', 'package': '未知封装',
            'supplier': '未知供应商', 'quantity': 1, 'unit': '个',
            'location': '未知位置', 'price': 0.00,
            'buy_time': datetime.now().strftime('%Y-%m-%d'),
            'channel': '未知', 'remark': '无'
        }
        # 按映射关系赋值
        for col, field in mapping.items():
            if not field or col not in columns:
                continue
            if columns.index(col) >= len(row):
                continue
            val = str(row[columns.index(col)]).strip()
            # 类型转换
            if field == 'quantity':
                d[field] = int(val) if val.isdigit() else 1
            elif field == 'price':
                d[field] = float(val) if val.replace('.', '').isdigit() else 0.00
            elif val:
                d[field] = val
        # 覆盖批量值
        d.update(batch)
        data_list.append(d)

    # 表格内去重
    data_list = remove_table_dup(data_list)
    return data_list, errors


# -------------------------- 前端模板（主页面修改：二维码列默认不显示，点击按钮才显示） --------------------------
# 主页面模板（修改：二维码列默认不显示，点击按钮才显示）
MAIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>元器件库存管理工具 - 稳定版</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f8f9fa; margin: 0; padding: 0; }
        .top-nav { background: #0d6efd; color: white; padding: 1rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; }
        .top-nav h4 { margin: 0; }
        .top-btn { color: white; text-decoration: none; margin-left: 1rem; padding: 0.3rem 0.8rem; border-radius: 4px; }
        .top-btn:hover { background: white; color: #0d6efd; }
        .container-main { padding: 1.5rem; max-width: 100%; }
        .oper-bar { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; margin-bottom: 1rem; }
        .table-box { background: white; border-radius: 8px; box-shadow: 0 0 5px rgba(0,0,0,0.1); overflow-x: auto; padding: 0.5rem; }
        .img-sm { width: 60px; height: 60px; object-fit: cover; border-radius: 4px; }
        .qrcode-sm { width: 60px; height: 60px; object-fit: cover; border-radius: 4px; border: 1px solid #ddd; }
        .qrcode-column { display: none; }  /* 默认不显示二维码列 */
        .alert { 
            position: fixed; top: 80px; right: 20px; z-index: 9999; 
            min-width: 320px; max-width: 400px; margin: 0; padding: 0.8rem 1.2rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        .file-link { color: #0d6efd; text-decoration: none; }
        .file-link:hover { text-decoration: underline; }
        .modal-backdrop { z-index: 1040 !important; }
        .modal { z-index: 1050 !important; }
        .show-qrcode .qrcode-column { display: table-cell; }  /* 显示二维码列的样式 */
        /* 修复扫码弹窗样式 */
        #videoElement {
            transform: scaleX(1);
        }
        #scanOverlay {
            animation: pulse 2s infinite;
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 70%;
            height: 70%;
            border: 2px solid red;
            box-sizing: border-box;
            pointer-events: none;
        }
        @keyframes pulse {
            0% { border-color: rgba(255, 0, 0, 0.7); }
            50% { border-color: rgba(255, 0, 0, 1); }
            100% { border-color: rgba(255, 0, 0, 0.7); }
        }
        #fileUploadArea.highlight {
            animation: highlight 2s ease;
        }
        @keyframes highlight {
            0% { background-color: rgba(13, 110, 253, 0.1); }
            100% { background-color: transparent; }
        }
    </style>
</head>
<body>
    <div class="top-nav">
        <h4>元器件库存管理工具 - v1</h4>
        <div>
            <a href="#" class="top-btn" data-bs-toggle="modal" data-bs-target="#settingModal">工具设置</a>
            <a href="#" class="top-btn" data-bs-toggle="modal" data-bs-target="#helpModal">使用说明</a>
            <a href="#" class="top-btn" data-bs-toggle="modal" data-bs-target="#qrcodeModal">扫码管理</a>
        </div>
    </div>

    <div class="container-main">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for c, m in messages %}
                    <div class="alert alert-{{c}} alert-dismissible fade show auto-close" role="alert" data-delay="3000">
                        {{m}}<button class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="oper-bar">
            <div>
                <input type="checkbox" id="checkAll" onclick="toggleAll()">
                <label for="checkAll" class="me-2">全选</label>
                <span id="selectCount" class="text-muted">已选：0 条</span>
                <button class="btn btn-primary btn-sm me-1" onclick="batchEdit()">批量编辑</button>
                <button class="btn btn-danger btn-sm me-1" onclick="batchDel()">批量删除</button>
                <button class="btn btn-success btn-sm me-1" data-bs-toggle="modal" data-bs-target="#addModal">添加元器件</button>
                <button class="btn btn-info btn-sm me-1" onclick="openBOM()">BOM批量导入</button>
                <button class="btn btn-warning btn-sm" id="exportBtn" disabled onclick="openExport()">导出选中</button>
                <button class="btn btn-outline-info btn-sm" onclick="toggleQRColumn()" id="toggleQRBtn">显示二维码列</button>
            </div>
            <form method="GET" class="d-flex gap-1">
                <input type="hidden" name="selected" id="selectedIds" value="{{selected|join(',')}}">
                <input type="text" name="kw" class="form-control form-control-sm" placeholder="搜索：品类/型号/封装/供应商" value="{{kw}}">
                <button type="submit" class="btn btn-primary btn-sm">搜索</button>
                <button type="button" class="btn btn-outline-primary btn-sm" data-bs-toggle="modal" data-bs-target="#advSearchModal">高级搜索</button>
                {% if adv_params %}
                <a href="{{url_for('index')}}" class="btn btn-light btn-sm border">清空搜索</a>
                {% endif %}
            </form>
        </div>

        <div class="table-box" id="componentTable">
            <table class="table table-striped table-hover table-sm">
                <thead class="table-dark">
                    <tr>
                        <th width="5%">选择</th>
                        <th width="8%">图片</th>
                        <th>品类</th>
                        <th>型号规格</th>
                        <th>封装</th>
                        <th>供应商</th>
                        <th>数量</th>
                        <th>存放位置</th>
                        <th>单价(¥)</th>
                        <th width="10%">附件</th>
                        <th width="8%" class="qrcode-column">二维码</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    {% for comp in components %}
                    <tr>
                        <td><input type="checkbox" class="compCheck" value="{{comp.id}}" {% if comp.id|string in selected %}checked{% endif %}></td>
                        <td>
                            {% if comp.img_path %}
                            <a href="/{{comp.img_path}}" target="_blank"><img src="/{{comp.img_path}}" class="img-sm"></a>
                            {% else %}
                            <span class="text-muted">无</span>
                            {% endif %}
                        </td>
                        <td>{{comp.category}}</td>
                        <td>{{comp.model}}</td>
                        <td>{{comp.package}}</td>
                        <td>{{comp.supplier}}</td>
                        <td class="{{get_quantity_css(comp.quantity)}}">{{comp.quantity}} {{comp.unit}}</td>
                        <td>{{comp.location}}</td>
                        <td>{{comp.price|round(2)}}</td>
                        <td>
                            {% if comp.attach_path %}
                            <a href="/{{comp.attach_path}}" target="_blank" class="file-link">{{comp.attach_path.split('/')[-1]|truncate(12)}}</a>
                            {% else %}
                            <span class="text-muted">无</span>
                            {% endif %}
                        </td>
                        <td class="qrcode-column">
                            {% if comp.qrcode_path %}
                            <a href="/{{comp.qrcode_path}}" target="_blank"><img src="/{{comp.qrcode_path}}" class="qrcode-sm" title="点击查看大图"></a>
                            {% else %}
                            <a href="{{url_for('generate_qrcode', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-outline-info btn-sm" title="生成二维码">生成</a>
                            {% endif %}
                        </td>
                        <td>
                            <a href="{{url_for('edit', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-warning btn-sm">编辑</a>
                            <a href="{{url_for('delete', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-danger btn-sm" onclick="return confirm('确定删除？将同步删除图片/附件/二维码！')">删除</a>
                        </td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="12" class="text-center text-muted py-3">暂无数据，点击「添加元器件」或「BOM批量导入」录入</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- 添加元器件弹窗 -->
    <div class="modal fade" id="addModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-success text-white">
                    <h5 class="modal-title">添加元器件</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{url_for('add', selected=selected|join(','), kw=kw)}}" enctype="multipart/form-data">
                    <div class="modal-body">
                        <div class="row g-3">
                            <div class="col-md-4"><label>品类<span class="text-danger">*</span></label><input type="text" name="category" class="form-control" required></div>
                            <div class="col-md-4"><label>型号规格<span class="text-danger">*</span></label><input type="text" name="model" class="form-control" required></div>
                            <div class="col-md-4"><label>封装<span class="text-danger">*</span></label><input type="text" name="package" class="form-control" required></div>
                            <div class="col-md-4"><label>供应商</label><input type="text" name="supplier" class="form-control" value="未知供应商"></div>
                            <div class="col-md-2"><label>数量</label><input type="number" name="quantity" class="form-control" min="0" value="1"></div>
                            <div class="col-md-2"><label>单位</label><input type="text" name="unit" class="form-control" value="个"></div>
                            <div class="col-md-4"><label>存放位置</label><input type="text" name="location" class="form-control" value="未知位置"></div>
                            <div class="col-md-2"><label>单价(¥)</label><input type="number" name="price" class="form-control" min="0" step="0.01" value="0.00"></div>
                            <div class="col-md-2"><label>采购时间</label><input type="date" name="buy_time" class="form-control" value="{{today}}"></div>
                            <div class="col-md-4"><label>采购渠道</label><input type="text" name="channel" class="form-control" value="未知"></div>
                            <div class="col-md-12"><label>备注</label><textarea name="remark" class="form-control" rows="2">无</textarea></div>
                            <div class="col-md-6">
                                <label>元器件图片</label>
                                <input type="file" name="img" class="form-control" accept=".png,.jpg,.jpeg,.gif,.bmp,.webp">
                                <p class="text-muted small">支持png/jpg等，单文件≤10M</p>
                            </div>
                            <div class="col-md-6">
                                <label>相关附件</label>
                                <input type="file" name="attach" class="form-control" accept=".pdf,.doc,.docx,.xls,.xlsx,.zip,.txt,.csv">
                                <p class="text-muted small">支持pdf/Excel/zip等，单文件≤10M</p>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                        <button type="submit" class="btn btn-success">保存</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 批量编辑弹窗 -->
    <div class="modal fade" id="batchEditModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-primary text-white">
                    <h5 class="modal-title">批量编辑</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{url_for('batch_edit', kw=kw)}}">
                    <div class="modal-body">
                        <p class="text-warning">仅修改填写的字段，未填写保留原值（暂不支持批量修改图片/附件）</p>
                        <div class="row g-3">
                            <div class="col-md-4"><label>供应商</label><input type="text" name="supplier" class="form-control"></div>
                            <div class="col-md-2"><label>数量</label><input type="number" name="quantity" class="form-control" min="0"></div>
                            <div class="col-md-2"><label>单位</label><input type="text" name="unit" class="form-control"></div>
                            <div class="col-md-4"><label>存放位置</label><input type="text" name="location" class="form-control"></div>
                            <div class="col-md-2"><label>单价(¥)</label><input type="number" name="price" class="form-control" min="0" step="0.01"></div>
                            <div class="col-md-2"><label>采购时间</label><input type="date" name="buy_time" class="form-control"></div>
                            <div class="col-md-4"><label>采购渠道</label><input type="text" name="channel" class="form-control"></div>
                            <div class="col-md-8"><label>备注</label><textarea name="remark" class="form-control" rows="2"></textarea></div>
                            <input type="hidden" name="ids" id="batchEditIds" value="{{selected|join(',')}}">
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                        <button type="submit" class="btn btn-primary">保存修改</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 高级搜索弹窗 -->
    <div class="modal fade" id="advSearchModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-primary text-white">
                    <h5 class="modal-title">高级多条件搜索</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="GET" action="{{url_for('index')}}">
                    <div class="modal-body">
                        <input type="hidden" name="selected" value="{{selected|join(',')}}">
                        <div class="row g-3">
                            <div class="col-md-3"><label>品类</label><input type="text" name="adv_cate" class="form-control-sm" value="{{adv_params.adv_cate or ''}}"></div>
                            <div class="col-md-3"><label>型号规格</label><input type="text" name="adv_model" class="form-control-sm" value="{{adv_params.adv_model or ''}}"></div>
                            <div class="col-md-3"><label>封装</th><label><input type="text" name="adv_pack" class="form-control-sm" value="{{adv_params.adv_pack or ''}}"></div>
                            <div class="col-md-3"><label>供应商</label><input type="text" name="adv_sup" class="form-control-sm" value="{{adv_params.adv_sup or ''}}"></div>
                            <div class="col-md-3"><label>存放位置</label><input type="text" name="adv_loc" class="form-control-sm" value="{{adv_params.adv_loc or ''}}"></div>
                            <div class="col-md-3"><label>采购渠道</label><input type="text" name="adv_chan" class="form-control-sm" value="{{adv_params.adv_chan or ''}}"></div>
                            <div class="col-md-3"><label>采购时间-开始</label><input type="date" name="adv_start" class="form-control-sm" value="{{adv_params.adv_start or ''}}"></div>
                            <div class="col-md-3"><label>采购时间-结束</label><input type="date" name="adv_end" class="form-control-sm" value="{{adv_params.adv_end or ''}}"></div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                        <button type="reset" class="btn btn-light border">重置</button>
                        <button type="submit" class="btn btn-primary">
                        搜索</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 导出弹窗（修改：增加二维码和图片导出选项） -->
    <div class="modal fade" id="exportModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-warning text-white">
                    <h5 class="modal-title">导出配置</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{url_for('export', kw=kw)}}">
                    <div class="modal-body">
                        <input type="hidden" name="ids" id="exportIds">
                        <div class="row g-2 mb-3">
                            <div class="col-md-2"><input type="checkbox" name="fields" value="category" checked> 品类</div>
                            <div class="col-md-3"><input type="checkbox" name="fields" value="model" checked> 型号规格</div>
                            <div class="col-md-2"><input type="checkbox" name="fields" value="package" checked> 封装</div>
                            <div class="col-md-3"><input type="checkbox" name="fields" value="supplier" checked> 供应商</div>
                            <div class="col-md-1"><input type="checkbox" name="fields" value="quantity" checked> 数量</div>
                            <div class="col-md-1"><input type="checkbox" name="fields" value="unit" checked> 单位</div>
                            <div class="col-md-3"><input type="checkbox" name="fields" value="location" checked> 存放位置</div>
                            <div class="col-md-2"><input type="checkbox" name="fields" value="price"> 单价</div>
                            <div class="col-md-3"><input type="checkbox" name="fields" value="buy_time"> 采购时间</div>
                            <div class="col-md-3"><input type="checkbox" name="fields" value="channel"> 采购渠道</div>
                            <div class="col-md-4"><input type="checkbox" name="fields" value="remark"> 备注</div>
                        </div>
                        <div class="mb-3">
                            <label>导出格式</label>
                            <div class="form-check form-check-inline">
                                <input type="radio" name="format" value="xlsx" checked class="form-check-input">
                                <label class="form-check-label">Excel(xlsx)</label>
                            </div>
                            <div class="form-check form-check-inline">
                                <input type="radio" name="format" value="csv" class="form-check-input">
                                <label class="form-check-label">CSV</label>
                            </div>
                            <div class="form-check form-check-inline">
                                <input type="radio" name="format" value="zip" class="form-check-input">
                                <label class="form-check-label">打包ZIP(包含二维码和图片)</label>
                            </div>
                        </div>
                        <div class="mb-3">
                            <label>额外导出内容</label>
                            <div class="form-check">
                                <input type="checkbox" name="export_qrcode" value="1" class="form-check-input" checked>
                                <label class="form-check-label">导出二维码</label>
                            </div>
                            <div class="form-check">
                                <input type="checkbox" name="export_img" value="1" class="form-check-input">
                                <label class="form-check-label">导出元器件图片</label>
                            </div>
                        </div>
                        <div>
                            <div class="form-check form-check-inline">
                                <input type="radio" name="action" value="export" checked class="form-check-input">
                                <label class="form-check-label">导出文件</label>
                            </div>
                            <div class="form-check form-check-inline">
                                <input type="radio" name="action" value="print" class="form-check-input">
                                <label class="form-check-label">打印数据</label>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                        <button type="submit" class="btn btn-warning">确认导出</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 工具设置弹窗 -->
    <div class="modal fade" id="settingModal" tabindex="-1">
        <div class="modal-dialog modal-md">
            <div class="modal-content">
                <div class="modal-header bg-secondary text-white">
                    <h5 class="modal-title">工具设置</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="d-grid gap-2">
                        <a href="{{url_for('backup', selected=selected|join(','), kw=kw)}}" class="btn btn-primary">📥 立即备份数据</a>
                        <a href="{{url_for('restore_page', kw=kw)}}" class="btn btn-warning">🔄 备份恢复（覆盖当前数据）</a>
                        <a href="{{url_for('clean_residual')}}" class="btn btn-danger">🗑️ 清理残留文件（无关联图片/附件/二维码）</a>
                        <a href="{{url_for('auto_start', op='open')}}" class="btn btn-info">📌 开启开机自启（Windows）</a>
                        <a href="{{url_for('auto_start', op='close')}}" class="btn btn-dark">❌ 关闭开机自启（Windows）</a>
                    </div>
                </div>
                <div class="modal-footer">
                    <p class="text-muted small w-100 text-center">备份文件保存在backup目录，恢复会覆盖当前所有数据</p>
                </div>
            </div>
        </div>
    </div>

    <!-- 使用说明弹窗 -->
    <div class="modal fade" id="helpModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-info text-white">
                    <h5 class="modal-title">使用说明</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <pre class="bg-light p-3 rounded">{{help_content}}</pre>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal">关闭</button>
                </div>
            </div>
        </div>
    </div>

    <!-- 扫码管理弹窗（修复版） -->
    <div class="modal fade" id="qrcodeModal" tabindex="-1" data-bs-backdrop="static">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-success text-white">
                    <h5 class="modal-title">扫码管理</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal" onclick="stopScanner()"></button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <div class="col-md-6">
                            <h6>📱 扫码读取元器件信息</h6>
                            <p class="text-muted small">使用手机扫描元器件二维码，快速查看详细信息</p>

                            <!-- 摄像头状态显示 -->
                            <div id="cameraStatus" class="alert alert-info">
                                <p><strong>📷 摄像头准备就绪</strong></p>
                                <p>点击"开始扫描"启动摄像头</p>
                            </div>

                            <!-- 摄像头容器 -->
                            <div id="cameraContainer" class="text-center" style="display: none;">
                                <div id="videoContainer" style="position: relative; display: inline-block;">
                                    <video id="videoElement" width="100%" style="max-width: 300px; border: 2px solid #0d6efd; border-radius: 4px; background: #000;"></video>
                                    <div id="scanOverlay" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 70%; height: 70%; border: 2px solid red; box-sizing: border-box; pointer-events: none;"></div>
                                </div>
                                <div id="scanControls" class="mt-3">
                                    <button id="startScanBtn" class="btn btn-primary btn-sm" onclick="startScanner()">
                                        <span class="spinner-border spinner-border-sm d-none" id="scanSpinner"></span>
                                        <span id="scanBtnText">开始扫描</span>
                                    </button>
                                    <button id="stopScanBtn" class="btn btn-secondary btn-sm" onclick="stopScanner()" style="display: none;">停止扫描</button>
                                    <button id="switchCameraBtn" class="btn btn-outline-secondary btn-sm" onclick="switchCamera()" style="display: none;">切换摄像头</button>
                                </div>
                            </div>

                            <!-- 文件上传区域 -->
                            <div id="fileUploadArea" class="mt-3">
                                <label class="fw-bold">📤 或上传二维码图片扫描：</label>
                                <div class="input-group mt-2">
                                    <input type="file" id="qrcodeFile" accept="image/*" class="form-control form-control-sm">
                                    <button class="btn btn-primary btn-sm" onclick="scanQRCodeFromFile()">上传并扫描</button>
                                </div>
                                <p class="text-muted small mt-1">支持PNG、JPG格式的二维码图片</p>
                            </div>

                            <!-- 扫描结果区域 -->
                            <div id="scanResult" class="mt-3 p-3 border rounded" style="min-height: 120px; background: #f8f9fa;">
                                <h6>📋 扫描结果：</h6>
                                <div id="resultContent" class="text-center text-muted py-3">
                                    等待扫描结果...
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <h6>📄 批量二维码操作</h6>
                            <div class="d-grid gap-2">
                                <button class="btn btn-outline-primary" onclick="batchGenerateQRCodes()">批量生成二维码</button>
                                <a href="{{url_for('batch_generate_qrcodes')}}" class="btn btn-outline-info">为所有元器件生成二维码</a>
                            </div>

                            <div class="mt-4">
                                <h6>💡 使用说明：</h6>
                                <div class="accordion" id="qrHelpAccordion">
                                    <div class="accordion-item">
                                        <h6 class="accordion-header">
                                            <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#qrMethod1">
                                                方法1：摄像头扫码
                                            </button>
                                        </h6>
                                        <div id="qrMethod1" class="accordion-collapse collapse">
                                            <div class="accordion-body small">
                                                <ol>
                                                    <li>点击<b>"开始扫描"</b>按钮启动摄像头</li>
                                                    <li>将二维码对准红色扫描框内</li>
                                                    <li>工具会自动识别并显示元器件信息</li>
                                                    <li>识别成功后可继续扫描下一个</li>
                                                </ol>
                                            </div>
                                        </div>
                                    </div>

                                    <div class="accordion-item">
                                        <h6 class="accordion-header">
                                            <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#qrMethod2">
                                                方法2：上传图片扫描
                                            </button>
                                        </h6>
                                        <div id="qrMethod2" class="accordion-collapse collapse">
                                            <div class="accordion-body small">
                                                <ol>
                                                    <li>点击<b>"选择文件"</b>按钮</li>
                                                    <li>选择保存的二维码图片</li>
                                                    <li>点击<b>"上传并扫描"</b></li>
                                                    <li>工具会自动解析二维码内容</li>
                                                </ol>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- 设备兼容性提示 -->
                            <div class="mt-4 alert alert-warning small">
                                <h6>⚠️ 注意事项：</h6>
                                <ul class="mb-0">
                                    <li>请确保已授予浏览器摄像头权限</li>
                                    <li>苹果设备需要使用Safari浏览器</li>
                                    <li>确保摄像头未被其他程序占用</li>
                                    <li>光线充足，二维码清晰无遮挡</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" data-bs-dismiss="modal" onclick="stopScanner()">关闭</button>
                </div>
            </div>
        </div>
    </div>

    <!-- 打印区域 -->
    <div id="printArea" class="d-none p-4">
        <h4 class="text-center mb-4">元器件库存数据</h4>
        <table class="table table-striped table-bordered" id="printTable"></table>
    </div>

    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@zxing/library@0.19.1"></script>
    <script>
        // 自动关闭提示框（修改：支持自动关闭）
        function initAutoClose() {
            document.querySelectorAll('.auto-close').forEach(alertEl => {
                const delay = alertEl.getAttribute('data-delay') || 3000;
                setTimeout(() => {
                    if (alertEl) {
                        const bsAlert = new bootstrap.Alert(alertEl);
                        bsAlert.close();
                    }
                }, parseInt(delay));
            });
        }

        // 页面加载后初始化
        window.onload = function() {
            initAutoClose();
            updateSelect();
            document.querySelectorAll('.compCheck').forEach(c => {
                c.addEventListener('change', updateSelect);
            });
        }

        // 获取选中ID
        function getSelected() {
            let ids = [];
            document.querySelectorAll('.compCheck:checked').forEach(c => ids.push(c.value));
            return ids;
        }

        // 更新选中状态
        function updateSelect() {
            let ids = getSelected();
            document.getElementById('selectCount').innerText = `已选：${ids.length} 条`;
            document.getElementById('exportBtn').disabled = ids.length === 0;
            document.getElementById('selectedIds').value = ids.join(',');
        }

        // 全选/取消
        function toggleAll() {
            let isCheck = document.getElementById('checkAll').checked;
            document.querySelectorAll('.compCheck').forEach(c => c.checked = isCheck);
            updateSelect();
        }

        // 批量编辑
        function batchEdit() {
            let ids = getSelected();
            if (ids.length === 0) {alert('请先选择元器件！'); return;}
            document.getElementById('batchEditIds').value = ids.join(',');
            new bootstrap.Modal(document.getElementById('batchEditModal')).show();
        }

        // 批量删除
        function batchDel() {
            let ids = getSelected();
            if (ids.length === 0) {alert('请先选择元器件！'); return;}
            if (confirm(`确定删除选中的${ids.length}条数据？将同时删除关联的图片、附件和二维码！`)) {
                window.location.href = "{{url_for('batch_delete', kw=kw)}}&ids=" + ids.join(',');
            }
        }

        // 打开BOM导入
        function openBOM() {
            window.open("{{url_for('bom_import')}}", "_blank", "width=1000,height=800,top=100,left=200");
        }

        // 打开导出弹窗
        function openExport() {
            let ids = getSelected();
            document.getElementById('exportIds').value = ids.join(',');
            new bootstrap.Modal(document.getElementById('exportModal')).show();
        }

        // 切换二维码列显示
        function toggleQRColumn() {
            const table = document.getElementById('componentTable');
            const button = document.getElementById('toggleQRBtn');

            if (table.classList.contains('show-qrcode')) {
                table.classList.remove('show-qrcode');
                button.innerText = '显示二维码列';
                button.classList.remove('btn-info');
                button.classList.add('btn-outline-info');
            } else {
                table.classList.add('show-qrcode');
                button.innerText = '隐藏二维码列';
                button.classList.remove('btn-outline-info');
                button.classList.add('btn-info');
            }
        }

        // 批量生成二维码
        function batchGenerateQRCodes() {
            let ids = getSelected();
            if (ids.length === 0) {alert('请先选择元器件！'); return;}
            if (confirm(`确定要为选中的${ids.length}个元器件生成二维码吗？`)) {
                window.location.href = "{{url_for('batch_generate_qrcodes')}}?ids=" + ids.join(',');
            }
        }

        // ===================== 扫码功能修复 =====================

        // 扫码相关变量
        let videoStream = null;
        let currentCamera = 'environment'; // 'user'为前置，'environment'为后置
        let isScanning = false;
        let codeReader = null;

        // 初始化扫码管理弹窗
        document.getElementById('qrcodeModal').addEventListener('shown.bs.modal', function() {
            console.log('扫码弹窗打开，初始化摄像头');
            initCamera();
        });

        document.getElementById('qrcodeModal').addEventListener('hidden.bs.modal', function() {
            console.log('扫码弹窗关闭，清理资源');
            stopScanner();
        });

        // 初始化摄像头
        function initCamera() {
            const cameraStatus = document.getElementById('cameraStatus');
            const cameraContainer = document.getElementById('cameraContainer');

            // 检查浏览器支持
            if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
                cameraStatus.innerHTML = `
                    <div class="alert alert-warning">
                        <p><strong>⚠️ 浏览器不支持摄像头功能</strong></p>
                        <p class="mb-0">请升级浏览器或使用文件上传方式</p>
                    </div>
                `;
                return;
            }

            // 检查HTTPS/localhost（iOS要求）
            const isLocalhost = window.location.hostname === 'localhost' || 
                               window.location.hostname === '127.0.0.1';
            const isHttps = window.location.protocol === 'https:';
            const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);

            if (isIOS && !isHttps && !isLocalhost) {
                cameraStatus.innerHTML = `
                    <div class="alert alert-warning">
                        <p><strong>⚠️ iOS设备限制</strong></p>
                        <p>iOS设备需要HTTPS或localhost环境</p>
                        <p class="mb-0">请使用文件上传方式扫描二维码</p>
                    </div>
                `;
                return;
            }

            // 显示摄像头可用
            cameraStatus.innerHTML = `
                <div class="alert alert-success">
                    <p><strong>✅ 摄像头可用</strong></p>
                    <p class="mb-0">点击下方"开始扫描"按钮启动摄像头</p>
                </div>
            `;
            cameraContainer.style.display = 'block';
        }

        // 启动扫码器
        async function startScanner() {
            console.log('开始启动扫码器...');

            const startBtn = document.getElementById('startScanBtn');
            const stopBtn = document.getElementById('stopScanBtn');
            const switchBtn = document.getElementById('switchCameraBtn');
            const spinner = document.getElementById('scanSpinner');
            const btnText = document.getElementById('scanBtnText');
            const videoElement = document.getElementById('videoElement');
            const resultContent = document.getElementById('resultContent');

            try {
                // 显示加载状态
                startBtn.disabled = true;
                spinner.classList.remove('d-none');
                btnText.textContent = '准备中...';
                resultContent.innerHTML = '<div class="text-primary"><span class="spinner-border spinner-border-sm"></span> 正在启动摄像头...</div>';

                // 停止之前的流
                if (videoStream) {
                    videoStream.getTracks().forEach(track => track.stop());
                    videoStream = null;
                }

                // 获取摄像头权限
                videoStream = await navigator.mediaDevices.getUserMedia({
                    video: {
                        facingMode: currentCamera,
                        width: { ideal: 640 },
                        height: { ideal: 480 },
                        frameRate: { ideal: 30 }
                    }
                });

                console.log('摄像头权限获取成功');

                // 设置视频元素
                videoElement.srcObject = videoStream;
                videoElement.setAttribute('playsinline', true); // iOS需要
                videoElement.setAttribute('autoplay', true);
                videoElement.setAttribute('muted', true);

                // 等待视频加载
                await new Promise((resolve) => {
                    videoElement.onloadedmetadata = () => {
                        videoElement.play();
                        resolve();
                    };
                });

                console.log('视频流播放成功');

                // 初始化ZXing
                codeReader = new ZXing.BrowserMultiFormatReader();

                // 更新UI状态
                startBtn.style.display = 'none';
                stopBtn.style.display = 'inline-block';
                switchBtn.style.display = 'inline-block';
                spinner.classList.add('d-none');
                btnText.textContent = '开始扫描';
                startBtn.disabled = false;

                // 开始解码
                isScanning = true;
                decodeContinuously();

            } catch (error) {
                console.error('启动摄像头失败:', error);

                let errorMsg = error.message || '未知错误';
                if (error.name === 'NotAllowedError') {
                    errorMsg = '摄像头权限被拒绝，请允许摄像头访问';
                } else if (error.name === 'NotFoundError') {
                    errorMsg = '未找到摄像头设备';
                } else if (error.name === 'NotReadableError') {
                    errorMsg = '摄像头被其他程序占用';
                }

                resultContent.innerHTML = `
                    <div class="alert alert-danger">
                        <p><strong>❌ 启动摄像头失败</strong></p>
                        <p class="mb-2">${errorMsg}</p>
                        <button class="btn btn-sm btn-outline-primary" onclick="useFileUpload()">改用文件上传</button>
                    </div>
                `;

                // 重置UI
                startBtn.disabled = false;
                spinner.classList.add('d-none');
                btnText.textContent = '开始扫描';
            }
        }

        // 连续解码
        function decodeContinuously() {
            if (!codeReader || !isScanning) return;

            const videoElement = document.getElementById('videoElement');
            const resultContent = document.getElementById('resultContent');

            codeReader.decodeFromVideoDevice(null, videoElement, (result, error) => {
                if (result) {
                    console.log('解码成功:', result.text);
                    // 成功解码
                    onScanSuccess(result.text);

                    // 暂停2秒后继续扫描
                    setTimeout(() => {
                        if (isScanning) {
                            decodeContinuously();
                        }
                    }, 2000);
                    return;
                }

                if (error) {
                    // 忽略"未找到二维码"的错误
                    if (!(error instanceof ZXing.NotFoundException)) {
                        console.warn('解码错误:', error);
                    }
                }

                // 继续扫描
                if (isScanning) {
                    requestAnimationFrame(decodeContinuously);
                }
            });
        }

        // 停止扫码器
        function stopScanner() {
            console.log('停止扫码器');
            isScanning = false;

            // 停止视频流
            if (videoStream) {
                videoStream.getTracks().forEach(track => {
                    track.stop();
                });
                videoStream = null;
            }

            // 停止ZXing解码
            if (codeReader) {
                codeReader.reset();
                codeReader = null;
            }

            // 重置UI
            const startBtn = document.getElementById('startScanBtn');
            const stopBtn = document.getElementById('stopScanBtn');
            const switchBtn = document.getElementById('switchCameraBtn');
            const videoElement = document.getElementById('videoElement');

            startBtn.style.display = 'inline-block';
            stopBtn.style.display = 'none';
            switchBtn.style.display = 'none';
            startBtn.disabled = false;
            document.getElementById('scanSpinner').classList.add('d-none');
            document.getElementById('scanBtnText').textContent = '开始扫描';

            // 清空视频
            videoElement.srcObject = null;
            videoElement.pause();
        }

        // 切换摄像头
        function switchCamera() {
            currentCamera = currentCamera === 'environment' ? 'user' : 'environment';
            stopScanner();
            setTimeout(startScanner, 500);
        }

        // 扫描成功处理
        function onScanSuccess(decodedText) {
            const resultContent = document.getElementById('resultContent');
            console.log('解析到的数据:', decodedText);

            try {
                const data = JSON.parse(decodedText);

                resultContent.innerHTML = `
                    <div class="alert alert-success">
                        <h6>✅ 扫码成功！</h6>
                        <div class="row mt-2">
                            <div class="col-6">
                                <p class="mb-1"><strong>ID：</strong>${data.id || 'N/A'}</p>
                                <p class="mb-1"><strong>品类：</strong>${data.category || 'N/A'}</p>
                                <p class="mb-1"><strong>型号：</strong>${data.model || 'N/A'}</p>
                                <p class="mb-1"><strong>封装：</strong>${data.package || 'N/A'}</p>
                            </div>
                            <div class="col-6">
                                <p class="mb-1"><strong>数量：</strong>${data.quantity || 0} ${data.unit || ''}</p>
                                <p class="mb-1"><strong>位置：</strong>${data.location || 'N/A'}</p>
                                <p class="mb-1"><strong>单价：</strong>¥${parseFloat(data.price || 0).toFixed(2)}</p>
                                <p class="mb-1"><strong>供应商：</strong>${data.supplier || 'N/A'}</p>
                            </div>
                        </div>
                        <div class="d-grid gap-2 mt-3">
                            <a href="/edit/${data.id}" class="btn btn-sm btn-primary" target="_blank">
                                📝 查看详情
                            </a>
                            <button class="btn btn-sm btn-outline-secondary" onclick="continueScanning()">继续扫描</button>
                        </div>
                    </div>
                `;

            } catch (e) {
                console.error('二维码解析失败:', e);
                resultContent.innerHTML = `
                    <div class="alert alert-danger">
                        <p><strong>❌ 二维码解析失败</strong></p>
                        <p class="small mb-2">可能原因：非本工具生成的二维码或数据格式错误</p>
                        <p class="small mb-2">原始数据：${decodedText.substring(0, 100)}${decodedText.length > 100 ? '...' : ''}</p>
                        <button class="btn btn-sm btn-secondary" onclick="continueScanning()">重新扫描</button>
                    </div>
                `;
            }
        }

        // 继续扫描
        function continueScanning() {
            const resultContent = document.getElementById('resultContent');
            resultContent.innerHTML = '<div class="text-primary"><span class="spinner-border spinner-border-sm"></span> 重新开始扫描...</div>';

            setTimeout(() => {
                if (isScanning) {
                    decodeContinuously();
                }
            }, 500);
        }

        // 从文件扫描二维码
        async function scanQRCodeFromFile() {
            const fileInput = document.getElementById('qrcodeFile');
            const file = fileInput.files[0];
            const resultContent = document.getElementById('resultContent');

            if (!file) {
                resultContent.innerHTML = `
                    <div class="alert alert-warning">
                        <p>❌ 请先选择二维码图片文件</p>
                    </div>
                `;
                return;
            }

            // 显示加载状态
            resultContent.innerHTML = `
                <div class="text-center py-3">
                    <div class="spinner-border spinner-border-sm text-primary"></div>
                    <span class="ms-2">正在识别二维码...</span>
                </div>
            `;

            try {
                const codeReader = new ZXing.BrowserMultiFormatReader();
                const img = new Image();

                const result = await new Promise((resolve, reject) => {
                    img.onload = function() {
                        codeReader.decodeFromImage(img)
                            .then(resolve)
                            .catch(reject);
                    };
                    img.onerror = () => reject(new Error('图片加载失败'));
                    img.src = URL.createObjectURL(file);
                });

                console.log('文件扫描成功:', result.text);
                onScanSuccess(result.text);

            } catch (err) {
                console.error('二维码识别失败:', err);
                resultContent.innerHTML = `
                    <div class="alert alert-danger">
                        <p>❌ 二维码识别失败</p>
                        <p class="small mb-1">${err.message || '无法识别二维码内容'}</p>
                        <button class="btn btn-sm btn-secondary" onclick="scanQRCodeFromFile()">重新尝试</button>
                    </div>
                `;
            }
        }

        // 使用文件上传（备用方案）
        function useFileUpload() {
            const fileUploadArea = document.getElementById('fileUploadArea');
            const qrcodeFile = document.getElementById('qrcodeFile');

            // 高亮显示文件上传区域
            fileUploadArea.style.border = '2px solid #0d6efd';
            fileUploadArea.style.borderRadius = '8px';
            fileUploadArea.style.padding = '15px';
            fileUploadArea.style.transition = 'all 0.3s';

            // 滚动到文件上传区域
            fileUploadArea.scrollIntoView({ behavior: 'smooth' });

            // 触发文件选择
            qrcodeFile.click();
        }

        // 打印数据
        document.querySelector('form[action="{{url_for('export', kw=kw)}}"]').addEventListener('submit', async function(e) {
            let action = document.querySelector('input[name="action"]:checked').value;
            if (action === 'print') {
                e.preventDefault();
                let ids = getSelected();
                let fields = [];
                document.querySelectorAll('input[name="fields"]:checked').forEach(f => fields.push(f.value));
                if (fields.length === 0) {alert('请选择打印字段！'); return;}
                // 获取打印数据
                let res = await fetch(`{{url_for('get_print_data')}}?ids=${ids.join(',')}&fields=${fields.join(',')}`);
                let data = await res.json();
                if (data.code !== 1) {alert(data.error); return;}
                // 渲染打印表格
                let fieldMap = {
                    category:'品类',model:'型号规格',package:'封装',supplier:'供应商',
                    quantity:'数量',unit:'单位',location:'存放位置',price:'单价(¥)',
                    buy_time:'采购时间',channel:'采购渠道',remark:'备注'
                };
                let table = document.getElementById('printTable');
                table.innerHTML = '<thead class="table-dark"><tr></tr></thead><tbody></tbody>';
                // 表头
                let theadTr = table.querySelector('thead tr');
                fields.forEach(f => {
                    let th = document.createElement('th');
                    th.innerText = fieldMap[f] || f;
                    theadTr.appendChild(th);
                });
                // 表体
                let tbody = table.querySelector('tbody');
                data.data.forEach(row => {
                    let tr = document.createElement('tr');
                    fields.forEach(f => {
                        let td = document.createElement('td');
                        td.innerText = f === 'price' ? '¥' + parseFloat(row[f]).toFixed(2) : row[f];
                        tr.appendChild(td);
                    });
                    tbody.appendChild(tr);
                });
                // 打印
                document.getElementById('printArea').classList.remove('d-none');
                window.print();
                document.getElementById('printArea').classList.add('d-none');
            }
        });
    </script>
</body>
</html>
'''

# 编辑页面模板（修改：包含二维码操作）
EDIT_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>编辑元器件 - 稳定版</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f8f9fa; padding: 2rem; }
        .container { background: white; padding: 2rem; border-radius: 8px; box-shadow: 0 0 5px rgba(0,0,0,0.1); }
        .img-preview { max-width: 200px; max-height: 200px; margin-top: 1rem; border-radius: 4px; }
        .qrcode-preview { max-width: 150px; max-height: 150px; margin-top: 1rem; border-radius: 4px; border: 1px solid #ddd; }
        .file-link { color: #0d6efd; text-decoration: none; }
        .file-link:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="container">
        <h4 class="text-primary">编辑元器件 [ID: {{comp.id}}]</h4>
        <form method="POST" enctype="multipart/form-data">
            <div class="row g-3">
                <div class="col-md-4"><label>品类<span class="text-danger">*</span></label><input type="text" name="category" class="form-control" required value="{{comp.category}}"></div>
                <div class="col-md-4"><label>型号规格<span class="text-danger">*</span></label><input type="text" name="model" class="form-control" required value="{{comp.model}}"></div>
                <div class="col-md-4"><label>封装<span class="text-danger">*</span></label><input type="text" name="package" class="form-control" required value="{{comp.package}}"></div>
                <div class="col-md-4"><label>供应商</label><input type="text" name="supplier" class="form-control" value="{{comp.supplier}}"></div>
                <div class="col-md-2"><label>数量</label><input type="number" name="quantity" class="form-control" min="0" value="{{comp.quantity}}"></div>
                <div class="col-md-2"><label>单位</label><input type="text" name="unit" class="form-control" value="{{comp.unit}}"></div>
                <div class="col-md-4"><label>存放位置</label><input type="text" name="location" class="form-control" value="{{comp.location}}"></div>
                <div class="col-md-2"><label>单价(¥)</label><input type="number" name="price" class="form-control" min="0" step="0.01" value="{{comp.price}}"></div>
                <div class="col-md-2"><label>采购时间</label><input type="date" name="buy_time" class="form-control" value="{{comp.buy_time}}"></div>
                <div class="col-md-4"><label>采购渠道</label><input type="text" name="channel" class="form-control" value="{{comp.channel}}"></div>
                <div class="col-md-12"><label>备注</label><textarea name="remark" class="form-control" rows="2">{{comp.remark}}</textarea></div>

                <div class="col-md-4">
                    <label>元器件图片（重新上传覆盖原有，勾选清空则删除）</label>
                    <input type="file" name="img" class="form-control" accept=".png,.jpg,.jpeg,.gif,.bmp,.webp">
                    {% if comp.img_path %}
                    <div class="mt-2">
                        <a href="/{{comp.img_path}}" target="_blank"><img src="/{{comp.img_path}}" class="img-preview"></a>
                        <div class="form-check mt-2">
                            <input type="checkbox" name="clear_img" class="form-check-input" id="clear_img">
                            <label class="form-check-label" for="clear_img">清空当前图片</label>
                        </div>
                    </div>
                    {% else %}
                    <p class="text-muted mt-2">暂无图片</p>
                    {% endif %}
                </div>

                <div class="col-md-4">
                    <label>相关附件（重新上传覆盖原有，勾选清空则删除）</label>
                    <input type="file" name="attach" class="form-control" accept=".pdf,.doc,.docx,.xls,.xlsx,.zip,.txt,.csv">
                    {% if comp.attach_path %}
                    <div class="mt-2">
                        <p>当前附件：<a href="/{{comp.attach_path}}" target="_blank" class="file-link">{{comp.attach_path.split('/')[-1]}}</a></p>
                        <div class="form-check mt-2">
                            <input type="checkbox" name="clear_attach" class="form-check-input" id="clear_attach">
                            <label class="form-check-label" for="clear_attach">清空当前附件</label>
                        </div>
                    </div>
                    {% else %}
                    <p class="text-muted mt-2">暂无附件</p>
                    {% endif %}
                </div>

                <div class="col-md-4">
                    <label>二维码管理</label>
                    {% if comp.qrcode_path %}
                    <div class="mt-2">
                        <a href="/{{comp.qrcode_path}}" target="_blank"><img src="/{{comp.qrcode_path}}" class="qrcode-preview"></a>
                        <div class="d-flex gap-2 mt-2">
                            <a href="{{url_for('regenerate_qrcode', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-outline-info btn-sm">重新生成</a>
                            <a href="{{url_for('delete_qrcode', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-outline-danger btn-sm" onclick="return confirm('确定删除二维码？')">删除</a>
                        </div>
                        <p class="text-muted small mt-1">扫描二维码查看元器件信息</p>
                    </div>
                    {% else %}
                    <p class="text-muted mt-2">暂无二维码</p>
                    <a href="{{url_for('generate_qrcode', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-outline-success btn-sm">生成二维码</a>
                    {% endif %}
                </div>
            </div>

            <div class="mt-4">
                <a href="{{url_for('index', selected=selected|join(','), kw=kw)}}" class="btn btn-secondary">返回</a>
                <button type="submit" class="btn btn-primary ms-2">保存修改</button>
            </div>
        </form>
    </div>
    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''

# BOM批量导入模板（修复：添加加载中转页面和步骤指示器）
BOM_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BOM批量导入 - 稳定版</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f8f9fa; padding: 1.5rem; }
        .container { background: white; padding: 2rem; border-radius: 8px; box-shadow: 0 0 5px rgba(0,0,0,0.1); max-width: 1200px; }
        .paste-area { width: 100%; min-height: 150px; resize: vertical; padding: 0.5rem; border-radius: 4px; }
        .mapping-table { font-size: 0.9rem; }
        .preview-table { font-size: 0.85rem; }
        .step { margin-bottom: 2rem; }
        .hidden { display: none; }
        .duplicate-item { padding: 1rem; border: 1px solid #ffc107; border-radius: 6px; background: #fff3cd; margin-bottom: 1rem; }
        .duplicate-title { font-weight: bold; color: #d97706; }
        .preview-container { max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 0.5rem; border-radius: 4px; }
        .import-detail { background: #f8f9fa; padding: 1rem; border-radius: 6px; border-left: 4px solid #0d6efd; }
        .import-detail p { margin: 0.3rem 0; }
        .alert { 
            position: fixed; top: 80px; right: 20px; z-index: 9999; 
            min-width: 320px; max-width: 400px; margin: 0; padding: 0.8rem 1.2rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        /* 加载动画样式 */
        .spinner-border {
            animation: spinner-border 0.75s linear infinite;
        }
        @keyframes spinner-border {
            to { transform: rotate(360deg); }
        }
        /* 步骤指示器 */
        .step-indicator {
            display: flex;
            justify-content: space-between;
            margin-bottom: 2rem;
            position: relative;
        }
        .step-indicator::before {
            content: '';
            position: absolute;
            top: 20px;
            left: 10%;
            right: 10%;
            height: 2px;
            background: #dee2e6;
            z-index: 1;
        }
        .step-item {
            text-align: center;
            position: relative;
            z-index: 2;
            flex: 1;
        }
        .step-number {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            background: #dee2e6;
            color: #6c757d;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 0.5rem;
            font-weight: bold;
            transition: all 0.3s ease;
        }
        .step-item.active .step-number {
            background: #0d6efd;
            color: white;
            transform: scale(1.1);
        }
        .step-label {
            font-size: 0.85rem;
            color: #6c757d;
        }
        .step-item.active .step-label {
            color: #0d6efd;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <h4 class="text-primary mb-4">BOM批量导入</h4>
        <!-- 步骤指示器 -->
        <div class="step-indicator mb-4">
            <div class="step-item active" data-step="1">
                <div class="step-number">1</div>
                <div class="step-label">选择导入方式</div>
            </div>
            <div class="step-item" data-step="2">
                <div class="step-number">2</div>
                <div class="step-label">字段映射</div>
            </div>
            <div class="step-item" data-step="3">
                <div class="step-number">3</div>
                <div class="step-label">处理重复</div>
            </div>
            <div class="step-item" data-step="4">
                <div class="step-number">4</div>
                <div class="step-label">完成导入</div>
            </div>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for c, m in messages %}
                    <div class="alert alert-{{c}} alert-dismissible fade show auto-close" role="alert" data-delay="3000">
                        {{m}}<button class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <!-- 步骤0：加载中转页面 -->
        <div class="step hidden" id="step0">
            <h5 class="text-secondary">数据加载中...</h5>
            <div class="text-center py-5">
                <div class="spinner-border text-primary" role="status" style="width: 3rem; height: 3rem;">
                    <span class="visually-hidden">加载中...</span>
                </div>
                <p class="mt-3">正在处理导入数据，请稍候...</p>
                <p class="text-muted small">这可能需要几秒钟时间，具体取决于数据量大小</p>
            </div>
        </div>

        <!-- 步骤1：选择导入方式 -->
        <div class="step" id="step1">
            <h5 class="text-secondary">步骤1：选择导入方式</h5>
            <ul class="nav nav-tabs mt-3" id="importTab" role="tablist">
                <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#pasteTab">表格粘贴导入</button></li>
                <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#excelTab">Excel文件导入</button></li>
            </ul>
            <div class="tab-content mt-3">
                <div class="tab-pane fade show active" id="pasteTab">
                    <textarea class="paste-area" id="pasteData" placeholder="直接复制Excel/表格数据粘贴到这里（制表符分隔）"></textarea>
                    <button class="btn btn-primary mt-2" onclick="parseData('paste')">解析数据</button>
                </div>
                <div class="tab-pane fade" id="excelTab">
                    <input type="file" id="excelFile" class="form-control" accept=".xlsx">
                    <p class="text-muted small mt-1">仅支持xlsx格式，请勿打开文件时上传</p>
                    <button class="btn btn-primary mt-2" onclick="parseData('excel')">解析数据</button>
                </div>
            </div>
        </div>

        <!-- 步骤2：字段映射 -->
        <div class="step hidden" id="step2">
            <h5 class="text-secondary">步骤2：字段映射（红色为必填）</h5>
            <!-- 数据统计信息 -->
            <div class="alert alert-light mb-3" id="dataStats">
                <div class="d-flex justify-content-between align-items-center">
                    <div>
                        📊 数据统计：共 <span class="badge bg-primary" id="rowCount">0</span> 条数据， 
                        <span class="badge bg-secondary" id="colCount">0</span> 列
                    </div>
                    <div class="text-muted small">显示前3行预览数据</div>
                </div>
            </div>

            <!-- 前三行预览 -->
            <div class="mb-3">
                <h6>前三行数据预览：</h6>
                <div class="preview-container" id="previewContainer">
                    <table class="table table-sm table-bordered">
                        <thead id="previewHeader"></thead>
                        <tbody id="previewBody"></tbody>
                    </table>
                </div>
            </div>

            <div class="table-responsive mt-3">
                <table class="table table-bordered mapping-table">
                    <thead class="table-dark">
                        <tr><th>表格列</th><th>映射为工具字段</th><th>预览数据</th></tr>
                    </thead>
                    <tbody id="mappingTbody"></tbody>
                </table>
            </div>
            <div class="mt-3 p-3 bg-light rounded">
                <p class="fw-bold mb-2">批量设置未映射字段（可选）</p>
                <div class="d-flex flex-wrap gap-3">
                    <div><label>供应商：</label><input type="text" id="batch_sup" class="form-control-sm" value="未知供应商"></div>
                    <div><label>单位：</label><input type="text" id="batch_unit" class="form-control-sm" value="个"></div>
                    <div><label>存放位置：</label><input type="text" id="batch_loc" class="form-control-sm" value="未知位置"></div>
                    <div><label>采购渠道：</label><input type="text" id="batch_chan" class="form-control-sm" value="未知"></div>
                </div>
            </div>
            <div class="mt-3">
                <button class="btn btn-secondary" onclick="backToStep1()">返回上一步</button>
                <button class="btn btn-primary ms-2" onclick="checkDuplicate()">确认映射并检测重复</button>
            </div>
        </div>

        <!-- 步骤3：重复数据处理 -->
        <div class="step hidden" id="step3">
            <h5 class="text-secondary">步骤3：重复数据处理</h5>
            <div class="alert alert-info mb-3">
                检测结果：共<span id="totalCount">0</span>条数据，其中<span id="dupCount" class="text-warning fw-bold">0</span>条重复，<span id="newCount" class="text-success fw-bold">0</span>条新增
            </div>
            <div id="noDup" class="alert alert-success hidden">🎉 无重复数据，所有数据均为新增！</div>
            <div id="dupList" class="mt-3"></div>
            <div class="mt-3">
                <button class="btn btn-secondary" onclick="backToStep2()">返回上一步</button>
                <button class="btn btn-success ms-2" onclick="doImport()">确认导入</button>
            </div>
        </div>

        <!-- 步骤4：导入完成（修改：显示完整导入详情） -->
        <div class="step hidden" id="step4">
            <h5 class="text-secondary">步骤4：导入完成</h5>
            {% if import_res %}
            <div class="import-detail mb-4">
                <h6>📋 导入结果详情：</h6>
                <p><strong>总计处理：</strong>{{import_res.total}} 条数据</p>
                <p><strong>新增数据：</strong><span class="text-success fw-bold">{{import_res.added}}</span> 条（全新元器件）</p>
                <p><strong>合并数量：</strong><span class="text-warning fw-bold">{{import_res.merged}}</span> 条（与库内数据数量相加）</p>
                <p><strong>覆盖更新：</strong><span class="text-info fw-bold">{{import_res.covered}}</span> 条（替换库内原有数据）</p>
                <p><strong>跳过数据：</strong><span class="text-secondary fw-bold">{{import_res.skipped}}</span> 条（保留库内原有数据）</p>

                <!-- 验证总数是否匹配 -->
                <div class="mt-3 p-2 border rounded {% if import_res.added + import_res.merged + import_res.covered + import_res.skipped == import_res.total %}bg-success-subtle{% else %}bg-danger-subtle{% endif %}">
                    <p class="mb-1"><strong>数据验证：</strong>
                        新增({{import_res.added}}) + 合并({{import_res.merged}}) + 覆盖({{import_res.covered}}) + 跳过({{import_res.skipped}}) = 
                        <span class="fw-bold">{{import_res.added + import_res.merged + import_res.covered + import_res.skipped}}</span>
                        {% if import_res.added + import_res.merged + import_res.covered + import_res.skipped == import_res.total %}
                        ✅ 与总计({{import_res.total}}) 匹配
                        {% else %}
                        ❌ 与总计({{import_res.total}}) 不匹配
                        {% endif %}
                    </p>
                </div>

                {% if import_res.new_items %}
                <div class="mt-3">
                    <h6>📦 新增元器件详情 ({{import_res.new_items|length}}个)：</h6>
                    <div class="table-responsive">
                        <table class="table table-sm table-bordered">
                            <thead class="table-light">
                                <tr>
                                    <th>序号</th>
                                    <th>品类</th>
                                    <th>型号规格</th>
                                    <th>封装</th>
                                    <th>数量</th>
                                    <th>位置</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for item in import_res.new_items %}
                                <tr>
                                    <td>{{loop.index}}</td>
                                    <td>{{item.category}}</td>
                                    <td>{{item.model}}</td>
                                    <td>{{item.package}}</td>
                                    <td>{{item.quantity}} {{item.unit}}</td>
                                    <td>{{item.location}}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
                {% endif %}

                {% if import_res.updated_items %}
                <div class="mt-3">
                    <h6>🔄 更新元器件详情 ({{import_res.updated_items|length}}个)：</h6>
                    <div class="table-responsive">
                        <table class="table table-sm table-bordered">
                            <thead class="table-light">
                                <tr>
                                    <th>序号</th>
                                    <th>品类</th>
                                    <th>型号规格</th>
                                    <th>封装</th>
                                    <th>更新类型</th>
                                    <th>数量变化</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for item in import_res.updated_items %}
                                <tr>
                                    <td>{{loop.index}}</td>
                                    <td>{{item.category}}</td>
                                    <td>{{item.model}}</td>
                                    <td>{{item.package}}</td>
                                    <td>
                                        {% if item.type %}
                                        {{item.type}}
                                        {% elif item.old_quantity is defined %}
                                        合并数量
                                        {% else %}
                                        覆盖更新
                                        {% endif %}
                                    </td>
                                    <td>
                                        {% if item.old_quantity is defined and item.new_quantity is defined %}
                                        {{item.old_quantity}} → {{item.new_quantity}}
                                        {% else %}
                                        -
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
                {% endif %}

                {% if import_res.skipped_items %}
                <div class="mt-3">
                    <h6>⏭️ 跳过的元器件 ({{import_res.skipped_items|length}}个)：</h6>
                    <div class="table-responsive">
                        <table class="table table-sm table-bordered">
                            <thead class="table-light">
                                <tr>
                                    <th>序号</th>
                                    <th>品类</th>
                                    <th>型号规格</th>
                                    <th>封装</th>
                                    <th>原因</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for item in import_res.skipped_items %}
                                <tr>
                                    <td>{{loop.index}}</td>
                                    <td>{{item.category}}</td>
                                    <td>{{item.model}}</td>
                                    <td>{{item.package}}</td>
                                    <td>{{item.reason}}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
                {% endif %}
            </div>
            {% endif %}
            <div class="alert alert-success">
                🎉 导入成功！总处理<span id="total">{{import_res.total if import_res else 0}}</span>条，
                新增<span id="add">{{import_res.added if import_res else 0}}</span>条，
                合并<span id="merge">{{import_res.merged if import_res else 0}}</span>条，
                覆盖<span id="cover">{{import_res.covered if import_res else 0}}</span>条，
                跳过<span id="skip">{{import_res.skipped if import_res else 0}}</span>条
            </div>
            <div class="mt-3">
                <button class="btn btn-secondary" onclick="backToStep1()">重新导入</button>
                <button class="btn btn-primary ms-2" onclick="closeWin()">关闭并返回主界面</button>
            </div>
        </div>

        <!-- 隐藏表单 -->
        <form id="importForm" class="hidden" method="POST" action="{{url_for('do_bom_import')}}">
            <input type="hidden" name="raw_data" id="rawData">
            <input type="hidden" name="mapping" id="mapping">
            <input type="hidden" name="batch_vals" id="batchVals">
            <input type="hidden" name="dup_oper" id="dupOper">
            <input type="hidden" name="unique_data" id="uniqueData">
        </form>
    </div>

    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
    <script>
        let parseRes = {columns:[], preview:[], raw_data:[], error:'', row_count:0};
        let mapping = {};
        let dupData = [];
        let newData = [];

        // 步骤切换
        function showStep(s, showLoading = false) {
            // 更新步骤指示器
            document.querySelectorAll('.step-item').forEach(item => {
                item.classList.remove('active');
            });

            if (s > 0) {
                const stepItem = document.querySelector(`.step-item[data-step="${s}"]`);
                if (stepItem) {
                    stepItem.classList.add('active');
                }
            }

            if (showLoading) {
                // 显示加载中转页面
                document.querySelectorAll('.step').forEach(el => el.classList.add('hidden'));
                document.getElementById('step0').classList.remove('hidden');

                // 延迟显示目标步骤
                setTimeout(() => {
                    document.getElementById('step0').classList.add('hidden');
                    if (s > 0) {
                        document.getElementById('step'+s).classList.remove('hidden');
                    }
                }, 800);
            } else {
                document.querySelectorAll('.step').forEach(el => el.classList.add('hidden'));
                if (s > 0) {
                    document.getElementById('step'+s).classList.remove('hidden');
                }
            }
        }

        function backToStep1() {showStep(1); parseRes = {columns:[], preview:[], raw_data:[], error:'', row_count:0};}
        function backToStep2() {showStep(2);}
        function closeWin() {window.opener.location.reload(); window.close();}

        // 自动关闭提示框
        function initAutoClose() {
            document.querySelectorAll('.auto-close').forEach(alertEl => {
                const delay = alertEl.getAttribute('data-delay') || 3000;
                setTimeout(() => {
                    if (alertEl) {
                        const bsAlert = new bootstrap.Alert(alertEl);
                        bsAlert.close();
                    }
                }, parseInt(delay));
            });
        }

        // 解析数据
        function parseData(type) {
            let formData = new FormData();
            formData.append('type', type);
            if (type === 'paste') {
                let data = document.getElementById('pasteData').value.trim();
                if (!data) {alert('请粘贴数据！'); return;}
                formData.append('paste_data', data);
            } else {
                let file = document.getElementById('excelFile').files[0];
                if (!file) {alert('请选择Excel文件！'); return;}
                formData.append('excel_file', file);
            }
            fetch("{{url_for('parse_bom_data')}}", {method:'POST', body:formData})
            .then(res => res.json())
            .then(data => {
                if (data.code !== 1) {alert(data.error); return;}
                parseRes = data.data;
                // 显示读取了多少行数据
                alert(`✅ 解析成功！共读取 ${parseRes.row_count} 条数据`);
                renderMapping();
                showStep(2, true); // 显示加载中转后跳转到步骤2
                initAutoClose();
            }).catch(err => alert('解析失败：'+err.message));
        }

        // 渲染映射表格和预览
        function renderMapping() {
            let tbody = document.getElementById('mappingTbody');
            tbody.innerHTML = '';
            mapping = {};
            let fields = {{SYSTEM_FIELDS|tojson}};

            // 更新数据统计
            document.getElementById('rowCount').innerText = parseRes.row_count;
            document.getElementById('colCount').innerText = parseRes.columns.length;

            // 渲染预览表格
            renderPreview();

            // 渲染映射表格
            parseRes.columns.forEach(col => {
                mapping[col] = '';
                let tr = document.createElement('tr');
                // 表格列
                let td1 = document.createElement('td'); 
                td1.innerText = col; 
                tr.appendChild(td1);
                // 下拉框
                let td2 = document.createElement('td');
                let select = document.createElement('select'); 
                select.className = 'form-select form-select-sm';
                fields.forEach(f => {
                    let opt = document.createElement('option');
                    opt.value = f[0]; 
                    opt.innerText = f[1];
                    if (['category','model','package'].includes(f[0])) {
                        opt.style.color = 'red'; 
                        opt.style.fontWeight = 'bold';
                    }
                    select.appendChild(opt);
                });
                select.onchange = function() {mapping[col] = this.value;};
                td2.appendChild(select); 
                tr.appendChild(td2);
                // 预览
                let td3 = document.createElement('td');
                let val = parseRes.preview.length > 0 ? parseRes.preview[0][parseRes.columns.indexOf(col)] : '';
                td3.innerText = val || '无'; 
                tr.appendChild(td3);
                tbody.appendChild(tr);
            });
        }

        // 渲染预览表格
        function renderPreview() {
            const previewHeader = document.getElementById('previewHeader');
            const previewBody = document.getElementById('previewBody');

            previewHeader.innerHTML = '';
            previewBody.innerHTML = '';

            // 表头
            let headerTr = document.createElement('tr');
            parseRes.columns.forEach((col, idx) => {
                let th = document.createElement('th');
                th.innerText = `列${idx+1}: ${col}`;
                headerTr.appendChild(th);
            });
            previewHeader.appendChild(headerTr);

            // 表体（前三行）
            const maxRows = Math.min(3, parseRes.preview.length);
            for (let i = 0; i < maxRows; i++) {
                let tr = document.createElement('tr');
                parseRes.preview[i].forEach(cell => {
                    let td = document.createElement('td');
                    td.innerText = cell || '';
                    tr.appendChild(td);
                });
                previewBody.appendChild(tr);
            }
        }

        // 检测重复数据
        function checkDuplicate() {
            // 获取批量值
            let batchVals = {
                supplier: document.getElementById('batch_sup').value.trim() || '未知供应商',
                unit: document.getElementById('batch_unit').value.trim() || '个',
                location: document.getElementById('batch_loc').value.trim() || '未知位置',
                channel: document.getElementById('batch_chan').value.trim() || '未知'
            };
            // 显示加载中转
            showStep(0, true);
            // 提交检测
            let formData = new FormData();
            formData.append('raw_data', JSON.stringify(parseRes.raw_data));
            formData.append('mapping', JSON.stringify(mapping));
            formData.append('batch_vals', JSON.stringify(batchVals));
            fetch("{{url_for('check_bom_dup')}}", {method:'POST', body:formData})
            .then(res => res.json())
            .then(data => {
                if (data.code !== 1) {alert(data.error); return;}
                dupData = data.data.duplicate || [];
                newData = data.data.unique || [];
                renderDup();
                showStep(3, true); // 显示加载中转后跳转到步骤3
            }).catch(err => {
                alert('检测失败：'+err.message);
                showStep(2); // 出错时返回步骤2
            });
        }

        // 渲染重复数据【四选项：跳过/合并/依旧导入/新增一条】
        function renderDup() {
            let total = dupData.length + newData.length;
            document.getElementById('totalCount').innerText = total;
            document.getElementById('dupCount').innerText = dupData.length;
            document.getElementById('newCount').innerText = newData.length;

            let dupList = document.getElementById('dupList');
            let noDup = document.getElementById('noDup');

            if (dupData.length === 0) {
                dupList.classList.add('hidden'); 
                noDup.classList.remove('hidden'); 
                return;
            }

            dupList.classList.remove('hidden'); 
            noDup.classList.add('hidden');
            dupList.innerHTML = '<h6>重复数据处理选项：</h6><p class="text-muted small">共检测到' + dupData.length + '条重复数据，请为每条数据选择处理方式：</p>';

            dupData.forEach((item, idx) => {
                let div = document.createElement('div');
                div.className = 'duplicate-item';
                div.innerHTML = `
                    <div class="duplicate-title">重复数据 #${idx+1}：${item.data.category} - ${item.data.model} - ${item.data.package}</div>
                    <table class="table table-sm table-bordered mt-2">
                        <tr class="table-secondary">
                            <th>字段</th><th>库内原有</th><th>待导入</th><th>处理方式</th>
                        </tr>
                        <tr><td>数量</td><td>${item.old_data.quantity} ${item.old_data.unit}</td><td>${item.data.quantity} ${item.data.unit}</td><td rowspan="5">
                            <select class="form-select form-select-sm dupOper" data-id="${item.old_data.id}">
                                <option value="skip">跳过（保留原有）</option>
                                <option value="merge" selected>合并（数量相加）</option>
                                <option value="cover">依旧导入（覆盖原有）</option>
                                <option value="new">新增一条（创建新记录）</option>
                            </select>
                        </td></tr>
                        <tr><td>供应商</td><td>${item.old_data.supplier}</td><td>${item.data.supplier}</td></tr>
                        <tr><td>单价</td><td>¥${parseFloat(item.old_data.price).toFixed(2)}</td><td>¥${parseFloat(item.data.price).toFixed(2)}</td></tr>
                        <tr><td>位置</td><td>${item.old_data.location}</td><td>${item.data.location}</td></tr>
                        <tr><td>备注</td><td>${item.old_data.remark}</td><td>${item.data.remark}</td></tr>
                    </table>
                `;
                dupList.appendChild(div);
            });
        }

        // 执行导入
        function doImport() {
            // 获取重复数据的处理方式
            let dupOper = {};
            let newItems = [];
            document.querySelectorAll('.dupOper').forEach(sel => {
                if (sel.value === 'new') {
                    // 对于"新增一条"选项，将数据添加到newData中
                    const oldId = sel.dataset.id;
                    const dupItem = dupData.find(item => item.old_data.id == oldId);
                    if (dupItem) {
                        newData.push(dupItem.data);
                    }
                    dupOper[oldId] = 'skip'; // 原来的跳过
                } else {
                    dupOper[sel.dataset.id] = sel.value;
                }
            });

            // 获取批量值
            let batchVals = {
                supplier: document.getElementById('batch_sup').value.trim() || '未知供应商',
                unit: document.getElementById('batch_unit').value.trim() || '个',
                location: document.getElementById('batch_loc').value.trim() || '未知位置',
                channel: document.getElementById('batch_chan').value.trim() || '未知'
            };

            // 赋值隐藏表单
            document.getElementById('rawData').value = JSON.stringify(parseRes.raw_data);
            document.getElementById('mapping').value = JSON.stringify(mapping);
            document.getElementById('batchVals').value = JSON.stringify(batchVals);
            document.getElementById('dupOper').value = JSON.stringify(dupOper);
            document.getElementById('uniqueData').value = JSON.stringify(newData);

            // 显示加载中转
            showStep(0, true);

            // 延迟提交表单，确保加载界面显示
            setTimeout(() => {
                document.getElementById('importForm').submit();
            }, 1000);
        }

        // 页面加载时初始化
        window.onload = function() {
            initAutoClose();
            {% if import_res %}
                // 导入完成时直接显示步骤4，不经过中转
                document.querySelectorAll('.step').forEach(el => el.classList.add('hidden'));
                // 更新步骤指示器
                document.querySelectorAll('.step-item').forEach(item => {
                    item.classList.remove('active');
                });
                const stepItem = document.querySelector('.step-item[data-step="4"]');
                if (stepItem) {
                    stepItem.classList.add('active');
                }
                document.getElementById('step4').classList.remove('hidden');
                initAutoClose();
            {% endif %}
        }
    </script>
</body>
</html>
'''

# 备份恢复页面模板（保持不变）
RESTORE_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>备份恢复 - 稳定版</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f8f9fa; padding: 2rem; }
        .container { background: white; padding: 2rem; border-radius: 8px; box-shadow: 0 0 5px rgba(0,0,0,0.1); max-width: 800px; }
        .warn { color: #dc3545; font-weight: bold; margin: 2rem 0; padding: 1rem; background: #f8d7da; border-radius: 6px; }
    </style>
</head>
<body>
    <div class="container">
        <h4 class="text-primary mb-4">备份恢复</h4>
        <div class="warn">
            ⚠️ 警告：恢复操作会<strong>覆盖当前所有数据</strong>（数据库+图片+附件+二维码），请确保已备份当前重要数据后再执行！
        </div>
        <form method="POST" enctype="multipart/form-data">
            <div class="mb-3">
                <label class="form-label">选择备份ZIP文件（来自backup目录）</label>
                <input type="file" name="backup_zip" class="form-control" accept=".zip" required>
                <div class="form-text mt-2">仅支持本工具生成的备份文件（命名以「元器件库存备份_」开头）</div>
            </div>
            <div class="d-flex justify-content-between">
                <a href="{{url_for('index', kw=kw)}}" class="btn btn-secondary">返回主界面</a>
                <button type="submit" class="btn btn-danger">确认恢复（覆盖当前数据）</button>
            </div>
        </form>
    </div>
    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''

# -------------------------- 视图函数（修复BOM导入统计逻辑） --------------------------
# 全局模板函数
app.add_template_global(get_quantity_css, 'get_quantity_css')
app.add_template_global(SYSTEM_FIELDS, 'SYSTEM_FIELDS')


# 解析请求参数
def parse_args(req):
    selected = [s for s in req.args.get('selected', '').split(',') if s.strip().isdigit()]
    kw = req.args.get('kw', '').strip()
    adv_params = {
        'adv_cate': req.args.get('adv_cate', '').strip(),
        'adv_model': req.args.get('adv_model', '').strip(),
        'adv_pack': req.args.get('adv_pack', '').strip(),
        'adv_sup': req.args.get('adv_sup', '').strip(),
        'adv_loc': req.args.get('adv_loc', '').strip(),
        'adv_chan': req.args.get('adv_chan', '').strip(),
        'adv_start': req.args.get('adv_start', '').strip(),
        'adv_end': req.args.get('adv_end', '').strip()
    }
    adv_params = {k: v for k, v in adv_params.items() if v}
    return selected, kw, adv_params


# 首页
@app.route('/')
def index():
    selected, kw, adv_params = parse_args(request)
    query = Component.query
    if kw:
        query = query.filter(db.or_(
            Component.category.like(f'%{kw}%'),
            Component.model.like(f'%{kw}%'),
            Component.package.like(f'%{kw}%'),
            Component.supplier.like(f'%{kw}%'),
            Component.location.like(f'%{kw}%')
        ))
    if adv_params:
        if 'adv_cate' in adv_params: query = query.filter(Component.category.like(f'%{adv_params["adv_cate"]}%'))
        if 'adv_model' in adv_params: query = query.filter(Component.model.like(f'%{adv_params["adv_model"]}%'))
        if 'adv_pack' in adv_params: query = query.filter(Component.package.like(f'%{adv_params["adv_pack"]}%'))
        if 'adv_sup' in adv_params: query = query.filter(Component.supplier.like(f'%{adv_params["adv_sup"]}%'))
        if 'adv_loc' in adv_params: query = query.filter(Component.location.like(f'%{adv_params["adv_loc"]}%'))
        if 'adv_chan' in adv_params: query = query.filter(Component.channel.like(f'%{adv_params["adv_chan"]}%'))
        if 'adv_start' in adv_params and 'adv_end' in adv_params:
            query = query.filter(Component.buy_time.between(adv_params["adv_start"], adv_params["adv_end"]))
        elif 'adv_start' in adv_params:
            query = query.filter(Component.buy_time >= adv_params["adv_start"])
        elif 'adv_end' in adv_params:
            query = query.filter(Component.buy_time <= adv_params["adv_end"])
    components = query.order_by(Component.id.desc()).all()
    return render_template_string(MAIN_TEMPLATE,
                                  components=components, selected=selected, kw=kw, adv_params=adv_params,
                                  today=datetime.now().strftime('%Y-%m-%d'), help_content=get_help_content()
                                  )


# 添加元器件（修改：自动生成二维码）
@app.route('/add', methods=['POST'])
def add():
    selected, kw, _ = parse_args(request)
    form = {
        'category': request.form.get('category', '').strip(),
        'model': request.form.get('model', '').strip(),
        'package': request.form.get('package', '').strip(),
        'supplier': request.form.get('supplier', '未知供应商').strip(),
        'quantity': int(request.form.get('quantity', 1) or 1),
        'unit': request.form.get('unit', '个').strip(),
        'location': request.form.get('location', '未知位置').strip(),
        'price': float(request.form.get('price', 0.00) or 0.00),
        'buy_time': request.form.get('buy_time', datetime.now().strftime('%Y-%m-%d')).strip(),
        'channel': request.form.get('channel', '未知').strip(),
        'remark': request.form.get('remark', '无').strip()
    }
    if not form['category'] or not form['package']:
        flash("品类和封装为必填项！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    old_comp = is_duplicate(form)
    if old_comp:
        flash("添加失败：该品类+封装的元器件已存在！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    comp = Component(**form)
    db.session.add(comp)
    db.session.flush()
    pre = comp.get_file_prefix()
    comp.img_path = save_file(request.files.get('img'), IMG_FOLDER, ALLOWED_IMG_EXT, pre)
    comp.attach_path = save_file(request.files.get('attach'), ATTACH_FOLDER, ALLOWED_ATTACH_EXT, pre)
    # 自动生成二维码
    generate_qrcode(comp)
    db.session.commit()
    flash(f"元器件「{form['category']}-{form['model']}」添加成功！已自动生成二维码。", "success")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 编辑元器件
@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    if request.method == 'GET':
        return render_template_string(EDIT_TEMPLATE, comp=comp, selected=selected, kw=kw)
    form = {
        'category': request.form.get('category', '').strip(),
        'model': request.form.get('model', '').strip(),
        'package': request.form.get('package', '').strip(),
        'supplier': request.form.get('supplier', '未知供应商').strip(),
        'quantity': int(request.form.get('quantity', 1) or 1),
        'unit': request.form.get('unit', '个').strip(),
        'location': request.form.get('location', '未知位置').strip(),
        'price': float(request.form.get('price', 0.00) or 0.00),
        'buy_time': request.form.get('buy_time', datetime.now().strftime('%Y-%m-%d')).strip(),
        'channel': request.form.get('channel', '未知').strip(),
        'remark': request.form.get('remark', '无').strip()
    }
    if not form['category'] or not form['package']:
        flash("品类和封装为必填项！", "danger")
        return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))
    old_comp = is_duplicate(form)
    if old_comp and old_comp.id != id:
        flash("修改失败：该品类+封装的元器件已存在！", "danger")
        return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))
    if request.form.get('clear_img'):
        delete_file(comp.img_path)
        comp.img_path = ''
    if request.form.get('clear_attach'):
        delete_file(comp.attach_path)
        comp.attach_path = ''
    pre = comp.get_file_prefix()
    new_img = save_file(request.files.get('img'), IMG_FOLDER, ALLOWED_IMG_EXT, pre)
    new_attach = save_file(request.files.get('attach'), ATTACH_FOLDER, ALLOWED_ATTACH_EXT, pre)
    if new_img:
        delete_file(comp.img_path)
        comp.img_path = new_img
    if new_attach:
        delete_file(comp.attach_path)
        comp.attach_path = new_attach
    for k, v in form.items():
        setattr(comp, k, v)
    # 如果数据有修改，重新生成二维码
    if any(getattr(comp, k) != v for k, v in form.items() if k != 'id'):
        generate_qrcode(comp)
    db.session.commit()
    flash(f"元器件「{comp.category}-{comp.model}」修改成功！二维码已更新。", "success")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 删除单个元器件（修改：删除二维码）
@app.route('/delete/<int:id>')
def delete(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    delete_file(comp.img_path)
    delete_file(comp.attach_path)
    delete_file(comp.qrcode_path)  # 删除二维码
    db.session.delete(comp)
    db.session.commit()
    flash(f"元器件「{comp.category}-{comp.model}」删除成功！", "success")
    selected = [s for s in selected if s != str(id)]
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 批量删除（修改：删除二维码）
@app.route('/batch_delete')
def batch_delete():
    selected, kw, _ = parse_args(request)
    ids = [int(i) for i in request.args.get('ids', '').split(',') if i.strip().isdigit()]
    if not ids:
        flash("请选择要删除的元器件！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    comps = Component.query.filter(Component.id.in_(ids)).all()
    for c in comps:
        delete_file(c.img_path)
        delete_file(c.attach_path)
        delete_file(c.qrcode_path)  # 删除二维码
        db.session.delete(c)
    db.session.commit()
    flash(f"批量删除成功！共删除 {len(comps)} 条数据", "success")
    return redirect(url_for('index', selected='', kw=kw))


# 批量编辑
@app.route('/batch_edit', methods=['POST'])
def batch_edit():
    kw = request.args.get('kw', '').strip()
    ids = [int(i) for i in request.form.get('ids', '').split(',') if i.strip().isdigit()]
    if not ids:
        flash("请选择要编辑的元器件！", "danger")
        return redirect(url_for('index', kw=kw))
    form = {}
    if request.form.get('supplier', '').strip(): form['supplier'] = request.form.get('supplier').strip()
    if request.form.get('quantity', '').strip():
        try:
            form['quantity'] = int(request.form.get('quantity'))
        except:
            flash("批量数量必须为数字！", "danger");
            return redirect(url_for('index', kw=kw))
    if request.form.get('unit', '').strip(): form['unit'] = request.form.get('unit').strip()
    if request.form.get('location', '').strip(): form['location'] = request.form.get('location').strip()
    if request.form.get('price', '').strip():
        try:
            form['price'] = float(request.form.get('price'))
        except:
            flash("批量单价必须为数字！", "danger");
            return redirect(url_for('index', kw=kw))
    if request.form.get('buy_time', '').strip(): form['buy_time'] = request.form.get('buy_time').strip()
    if request.form.get('channel', '').strip(): form['channel'] = request.form.get('channel').strip()
    if request.form.get('remark', '').strip(): form['remark'] = request.form.get('remark').strip()
    if not form:
        flash("请填写要修改的字段！", "warning")
        return redirect(url_for('index', kw=kw))
    Component.query.filter(Component.id.in_(ids)).update(form, synchronize_session=False)

    # 批量更新后重新生成二维码
    comps = Component.query.filter(Component.id.in_(ids)).all()
    for comp in comps:
        generate_qrcode(comp)

    db.session.commit()
    flash(f"批量编辑成功！共修改 {len(ids)} 条数据，二维码已更新", "success")
    return redirect(url_for('index', selected=','.join([str(i) for i in ids]), kw=kw))


# 导出/打印（修改：支持导出二维码和图片）
@app.route('/export', methods=['POST'])
def export():
    kw = request.args.get('kw', '').strip()
    ids = [int(i) for i in request.form.get('ids', '').split(',') if i.strip().isdigit()]
    fields = request.form.getlist('fields')
    fmt = request.form.get('format', 'xlsx')
    export_qrcode = request.form.get('export_qrcode') == '1'
    export_img = request.form.get('export_img') == '1'

    if not ids or not fields:
        flash("请选择元器件和导出字段！", "danger")
        return redirect(url_for('index', kw=kw))

    # 如果是ZIP格式，打包导出
    if fmt == 'zip':
        try:
            # 创建临时ZIP文件
            zip_filename = f"元器件库存导出_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            zip_path = os.path.join(BACKUP_FOLDER, zip_filename)

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                # 1. 导出Excel数据
                field_cn = {
                    'category': '品类', 'model': '型号规格', 'package': '封装', 'supplier': '供应商',
                    'quantity': '数量', 'unit': '单位', 'location': '存放位置', 'price': '采购单价(¥)',
                    'buy_time': '采购时间', 'channel': '采购渠道', 'remark': '备注'
                }
                comps = Component.query.filter(Component.id.in_(ids)).all()
                data = []
                for c in comps:
                    row = {}
                    for f in fields:
                        val = getattr(c, f, '')
                        if f == 'price': val = round(float(val), 2)
                        row[field_cn[f]] = val
                    data.append(row)

                # 创建Excel文件
                df = pd.DataFrame(data)
                excel_filename = f"元器件数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                excel_buffer = BytesIO()
                df.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                zf.writestr(excel_filename, excel_buffer.getvalue())

                # 2. 导出二维码
                if export_qrcode:
                    qrcode_count = 0
                    for c in comps:
                        qrcode_path = get_or_generate_qrcode(c.id)
                        if qrcode_path and os.path.exists(os.path.join(BASE_DIR, qrcode_path)):
                            zf.write(os.path.join(BASE_DIR, qrcode_path),
                                     f"qrcodes/QR_{c.id}_{c.model.replace('/', '_')}.png")
                            qrcode_count += 1

                # 3. 导出图片
                if export_img:
                    img_count = 0
                    for c in comps:
                        if c.img_path and os.path.exists(os.path.join(BASE_DIR, c.img_path)):
                            zf.write(os.path.join(BASE_DIR, c.img_path),
                                     f"images/IMG_{c.id}_{c.model.replace('/', '_')}{os.path.splitext(c.img_path)[1]}")
                            img_count += 1

                # 4. 添加导出说明
                readme = f"""元器件库存导出包
导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
导出数量: {len(comps)} 个元器件
包含内容:
  - Excel数据文件: {excel_filename}
  - 二维码文件: {qrcode_count if export_qrcode else 0} 个
  - 元器件图片: {img_count if export_img else 0} 个
"""
                zf.writestr("README.txt", readme)

            # 发送ZIP文件
            return send_file(zip_path, mimetype='application/zip',
                             as_attachment=True, download_name=zip_filename)
        except Exception as e:
            flash(f"打包导出失败：{str(e)}", "danger")
            return redirect(url_for('index', kw=kw))

    # 普通Excel/CSV导出
    field_cn = {
        'category': '品类', 'model': '型号规格', 'package': '封装', 'supplier': '供应商',
        'quantity': '数量', 'unit': '单位', 'location': '存放位置', 'price': '采购单价(¥)',
        'buy_time': '采购时间', 'channel': '采购渠道', 'remark': '备注'
    }
    comps = Component.query.filter(Component.id.in_(ids)).all()
    data = []
    for c in comps:
        row = {}
        for f in fields:
            val = getattr(c, f, '')
            if f == 'price': val = round(float(val), 2)
            row[field_cn[f]] = val
        data.append(row)
    df = pd.DataFrame(data)
    output = BytesIO()
    filename = f"元器件库存_{datetime.now().strftime('%Y%m%d%H%M%S')}"
    if fmt == 'xlsx':
        df.to_excel(output, index=False, engine='openpyxl')
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        filename += '.xlsx'
    else:
        df.to_csv(output, index=False, encoding='utf-8-sig')
        mimetype = 'text/csv'
        filename += '.csv'
    output.seek(0)
    return send_file(output, mimetype=mimetype, as_attachment=True, download_name=filename)


# 获取打印数据
@app.route('/get_print_data')
def get_print_data():
    ids = [int(i) for i in request.args.get('ids', '').split(',') if i.strip().isdigit()]
    fields = request.args.get('fields', '').split(',')
    if not ids or not fields:
        return jsonify({'code': 0, 'error': '无有效数据/字段'})
    comps = Component.query.filter(Component.id.in_(ids)).all()
    data = []
    for c in comps:
        row = {}
        for f in fields:
            row[f] = getattr(c, f, '')
        data.append(row)
    return jsonify({'code': 1, 'data': data})


# 清理残留文件（修改：包含二维码）
@app.route('/clean_residual')
def clean_residual():
    selected, kw, _ = parse_args(request)
    is_success, msg = clean_residual_files()
    flash(msg, "success" if is_success else "danger")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 数据备份
@app.route('/backup')
def backup():
    selected, kw, _ = parse_args(request)
    backup_all_data()
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 备份恢复页面
@app.route('/restore')
def restore_page():
    kw = request.args.get('kw', '').strip()
    return render_template_string(RESTORE_TEMPLATE, kw=kw)


# 执行备份恢复
@app.route('/restore', methods=['POST'])
def do_restore():
    import tempfile
    import shutil
    kw = request.args.get('kw', '').strip()
    backup_zip = request.files.get('backup_zip')
    if not backup_zip or backup_zip.filename == '':
        flash("请选择备份ZIP文件！", "danger")
        return redirect(url_for('restore_page', kw=kw))

    temp_zip = os.path.join(tempfile.gettempdir(), f"temp_backup_{uuid.uuid4().hex[:8]}.zip")
    backup_zip.save(temp_zip)

    is_valid, msg = validate_backup_zip(temp_zip)
    if not is_valid:
        os.remove(temp_zip)
        flash(f"备份文件验证失败：{msg}", "danger")
        return redirect(url_for('restore_page', kw=kw))

    temp_unzip = os.path.join(tempfile.gettempdir(), f"temp_unzip_{uuid.uuid4().hex[:8]}")
    os.makedirs(temp_unzip, exist_ok=True)

    try:
        with zipfile.ZipFile(temp_zip, 'r') as zf:
            zf.extractall(temp_unzip)

        db_temp_path = os.path.join(temp_unzip, 'component.db')
        if not os.path.exists(db_temp_path):
            for root, _, files in os.walk(temp_unzip):
                if 'component.db' in files:
                    db_temp_path = os.path.join(root, 'component.db')
                    break
        if not os.path.exists(db_temp_path):
            flash("备份包中未找到数据库文件（component.db）！", "danger")
            return redirect(url_for('restore_page', kw=kw))

        shutil.copy2(db_temp_path, DB_FILE)
        logger.info(f"数据库文件已从【{db_temp_path}】还原到【{DB_FILE}】")

        shutil.copytree(os.path.join(temp_unzip, 'static'), STATIC_FOLDER, dirs_exist_ok=True)
        logger.info(f"图片/附件/二维码已还原到【{STATIC_FOLDER}】")

        flash("数据恢复成功！已将数据库还原到instance目录，图片/附件/二维码还原到原路径，请刷新页面查看", "success")
    except Exception as e:
        flash(f"恢复失败：{str(e)}", "danger")
        logger.error(f"恢复失败：{str(e)}")
    finally:
        os.remove(temp_zip)
        shutil.rmtree(temp_unzip, ignore_errors=True)
    return redirect(url_for('index', kw=kw))


# 开机自启
@app.route('/auto_start/<op>')
def auto_start(op):
    selected, kw, _ = parse_args(request)
    if op == 'open':
        is_success, msg = create_auto_start()
    elif op == 'close':
        is_success, msg = delete_auto_start()
    else:
        is_success, msg = False, "无效操作"
    flash(msg, "success" if is_success else "danger")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# BOM导入页面
@app.route('/bom_import')
def bom_import():
    return render_template_string(BOM_TEMPLATE)


# 解析BOM数据
@app.route('/parse_bom_data', methods=['POST'])
def parse_bom_data():
    try:
        source_type = request.form.get('type', '')
        if not source_type in ['paste', 'excel']:
            return jsonify({'code': 0, 'error': '无效的导入类型'})
        source = request.form.get('paste_data', '') if source_type == 'paste' else request.files.get('excel_file')
        columns, preview, raw_data, error, row_count = parse_table_data(source, source_type)
        if error:
            return jsonify({'code': 0, 'error': error})
        return jsonify({
            'code': 1,
            'data': {
                'columns': columns,
                'preview': preview,
                'raw_data': raw_data,
                'row_count': row_count
            }
        })
    except Exception as e:
        logger.error(f"BOM解析异常：{str(e)}")
        return jsonify({'code': 0, 'error': f"解析异常：{str(e)}"})


@app.route('/check_bom_dup', methods=['POST'])
def check_bom_dup():
    """
    BOM重复检测接口：正确区分重复数据和新增数据
    """
    try:
        # 接收并解析前端传参
        raw_data_str = request.form.get('raw_data', '[]')
        mapping_str = request.form.get('mapping', '{}')
        batch_vals_str = request.form.get('batch_vals', '{}')

        try:
            raw_data = json.loads(raw_data_str)
            mapping = json.loads(mapping_str)
            batch_vals = json.loads(batch_vals_str)
        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败：{str(e)}")
            return jsonify({'code': 0, 'error': f"数据格式错误：{str(e)}"})

        # 验证数据格式
        if not isinstance(raw_data, list) or not isinstance(mapping, dict):
            return jsonify({'code': 0, 'error': "数据格式错误：原始数据必须是列表，映射必须是字典"})

        # 获取列名（从mapping的键中提取）
        columns = list(mapping.keys())

        # 映射数据为字典列表
        data_list, errors = map_table_data(raw_data, columns, mapping, batch_vals)
        if errors:
            return jsonify({'code': 0, 'error': '; '.join(errors)})

        # 检测重复数据 - 正确区分重复和新增
        duplicate = []  # 重复数据
        unique = []  # 新增数据

        for item in data_list:
            old_component = is_duplicate(item)
            if old_component:
                # 获取完整的老数据信息
                old_data = {
                    'id': old_component.id,
                    'category': old_component.category,
                    'model': old_component.model,
                    'package': old_component.package,
                    'supplier': old_component.supplier,
                    'quantity': old_component.quantity,
                    'unit': old_component.unit,
                    'location': old_component.location,
                    'price': float(old_component.price),
                    'buy_time': old_component.buy_time,
                    'channel': old_component.channel,
                    'remark': old_component.remark
                }
                duplicate.append({
                    'old_data': old_data,
                    'data': item
                })
            else:
                # 这是新增数据
                unique.append(item)

        # 记录日志
        logger.info(f"BOM重复检测完成：共{len(data_list)}条，重复{len(duplicate)}条，新增{len(unique)}条")

        # 返回结果
        return jsonify({
            'code': 1,
            'data': {
                'duplicate': duplicate,
                'unique': unique
            },
            'error': ''
        })

    except Exception as e:
        logger.error(f"BOM重复检测失败：{str(e)}", exc_info=True)
        return jsonify({
            'code': 0,
            'error': f"检测失败：{str(e)[:200]}"
        })


# 执行BOM导入（修复：正确统计各种操作类型）
@app.route('/do_bom_import', methods=['POST'])
def do_bom_import():
    try:
        import json
        # 解析JSON数据
        try:
            raw_data = json.loads(request.form.get('raw_data', '[]'))
            mapping = json.loads(request.form.get('mapping', '{}'))
            batch_vals = json.loads(request.form.get('batch_vals', '{}'))
            dup_oper = json.loads(request.form.get('dup_oper', '{}'))
            unique_data = json.loads(request.form.get('unique_data', '[]'))
        except json.JSONDecodeError as e:
            logger.error(f"BOM导入：JSON解析失败 {str(e)}")
            flash("数据格式错误，导入失败", "danger")
            return render_template_string(BOM_TEMPLATE,
                                          import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0,
                                                      'skipped': 0, 'new_items': [], 'updated_items': [],
                                                      'skipped_items': []})

        # 基础类型校验
        if not isinstance(raw_data, list) or not isinstance(dup_oper, dict) or not isinstance(unique_data, list):
            flash("数据格式错误，导入失败", "danger")
            return render_template_string(BOM_TEMPLATE,
                                          import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0,
                                                      'skipped': 0, 'new_items': [], 'updated_items': [],
                                                      'skipped_items': []})

        # 获取列名（从mapping的键中提取）
        columns = list(mapping.keys())

        # 映射数据（内部已完成表格去重）
        data_list, errors = map_table_data(raw_data, columns, mapping, batch_vals)
        if errors:
            logger.error(f"BOM导入映射失败：{';'.join(errors)}")
            flash(f"导入失败：{';'.join(errors)}", "danger")
            return render_template_string(BOM_TEMPLATE,
                                          import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0,
                                                      'skipped': 0, 'new_items': [], 'updated_items': [],
                                                      'skipped_items': []})

        # 初始化统计和详情
        total = len(data_list)
        added = merged = covered = skipped = 0
        new_items = []  # 新增的元器件
        updated_items = []  # 更新的元器件
        skipped_items = []  # 跳过的元器件

        # 1. 先处理重复数据（根据用户选择的处理方式）
        duplicate_data = []
        for d in data_list:
            old_comp = is_duplicate(d)
            if old_comp:
                duplicate_data.append((old_comp, d))

        for old_comp, new_data in duplicate_data:
            op = dup_oper.get(str(old_comp.id), 'merge')
            if op == 'skip':
                skipped += 1
                skipped_items.append({
                    'category': old_comp.category,
                    'model': old_comp.model,
                    'package': old_comp.package,
                    'reason': '用户选择跳过'
                })
                continue
            elif op == 'merge':
                # 记录原始数量
                old_quantity = old_comp.quantity
                # 数量相加，其余字段保留原有
                old_comp.quantity += new_data.get('quantity', 0)
                merged += 1
                updated_items.append({
                    'category': old_comp.category,
                    'model': old_comp.model,
                    'package': old_comp.package,
                    'old_quantity': old_quantity,
                    'new_quantity': old_comp.quantity,
                    'type': '合并数量'
                })
            elif op == 'cover':
                # 全量覆盖原有数据
                for k, v in new_data.items():
                    if hasattr(old_comp, k) and k != 'id':
                        setattr(old_comp, k, v)
                covered += 1
                updated_items.append({
                    'category': old_comp.category,
                    'model': old_comp.model,
                    'package': old_comp.package,
                    'type': '覆盖更新'
                })

        # 2. 再处理新增数据（从unique_data中获取，这是前端传过来的新增数据）
        for new_item in unique_data:
            if isinstance(new_item, dict):
                # 创建新的元器件
                new_comp = Component(**new_item)
                db.session.add(new_comp)
                added += 1
                new_items.append({
                    'category': new_comp.category,
                    'model': new_comp.model,
                    'package': new_comp.package,
                    'quantity': new_comp.quantity,
                    'unit': new_comp.unit,
                    'location': new_comp.location,
                    'price': new_comp.price
                })

        # 3. 处理"新增一条"选项（从重复数据中创建新记录）
        for old_comp, new_data in duplicate_data:
            op = dup_oper.get(str(old_comp.id), 'merge')
            if op == 'new':
                # 创建新记录（稍微修改型号以避免唯一性冲突）
                new_item_copy = new_data.copy()
                # 在型号后添加时间戳以确保唯一性
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                new_item_copy['model'] = f"{new_item_copy['model']}_dup_{timestamp}"

                new_comp = Component(**new_item_copy)
                db.session.add(new_comp)
                added += 1
                new_items.append({
                    'category': new_comp.category,
                    'model': new_comp.model,
                    'package': new_comp.package,
                    'quantity': new_comp.quantity,
                    'unit': new_comp.unit,
                    'location': new_comp.location,
                    'price': new_comp.price,
                    'note': '从重复数据创建'
                })

        # 提交所有更改
        db.session.commit()

        # 为新添加的元器件生成二维码
        new_comps = Component.query.filter(Component.qrcode_path == '').all()
        qrcode_generated = 0
        for comp in new_comps:
            if generate_qrcode(comp):
                qrcode_generated += 1

        # 导入结果统计
        import_res = {
            'total': total,
            'added': added,
            'merged': merged,
            'covered': covered,
            'skipped': skipped,
            'new_items': new_items,  # 显示所有新增
            'updated_items': updated_items,  # 显示所有更新
            'skipped_items': skipped_items,  # 显示所有跳过
            'qrcode_generated': qrcode_generated
        }

        # 验证统计数据是否正确
        calculated_total = added + merged + covered + skipped
        if calculated_total != total:
            logger.warning(f"统计数据不匹配：计算总数({calculated_total}) != 实际总数({total})")

        logger.info(f"BOM导入完成：新增{added}，合并{merged}，覆盖{covered}，跳过{skipped}，总计{total}")

        flash(
            f"导入成功！共处理{total}条数据，新增{added}条，合并{merged}条，覆盖{covered}条，跳过{skipped}条，生成{qrcode_generated}个二维码",
            "success")

        return render_template_string(BOM_TEMPLATE, import_res=import_res)

    except Exception as e:
        # 发生错误时回滚
        db.session.rollback()
        logger.error(f"BOM导入执行失败：{str(e)}", exc_info=True)
        flash(f"导入失败：{str(e)}", "danger")
        return render_template_string(BOM_TEMPLATE,
                                      import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0,
                                                  'skipped': 0, 'new_items': [], 'updated_items': [],
                                                  'skipped_items': []})


# 二维码相关路由
@app.route('/generate_qrcode/<int:id>')
def generate_qrcode_route(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    qrcode_path = generate_qrcode(comp)
    if qrcode_path:
        flash(f"二维码生成成功！", "success")
    else:
        flash("二维码生成失败", "warning")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


@app.route('/regenerate_qrcode/<int:id>')
def regenerate_qrcode(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    # 删除旧二维码
    if comp.qrcode_path:
        delete_file(comp.qrcode_path)
    # 生成新二维码
    qrcode_path = generate_qrcode(comp)
    if qrcode_path:
        flash(f"二维码重新生成成功！", "success")
    else:
        flash("二维码重新生成失败", "warning")
    return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))


@app.route('/delete_qrcode/<int:id>')
def delete_qrcode(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    if comp.qrcode_path:
        delete_file(comp.qrcode_path)
        comp.qrcode_path = ''
        db.session.commit()
        flash("二维码删除成功", "success")
    else:
        flash("没有二维码可删除", "warning")
    return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))


@app.route('/batch_generate_qrcodes')
def batch_generate_qrcodes():
    selected, kw, _ = parse_args(request)
    ids = [int(i) for i in request.args.get('ids', '').split(',') if i.strip().isdigit()]

    if not ids:
        # 为所有元器件生成二维码
        comps = Component.query.all()
    else:
        comps = Component.query.filter(Component.id.in_(ids)).all()

    success_count = 0
    for comp in comps:
        if not comp.qrcode_path or not os.path.exists(os.path.join(BASE_DIR, comp.qrcode_path)):
            if generate_qrcode(comp):
                success_count += 1

    flash(f"批量生成二维码完成！共为 {success_count}/{len(comps)} 个元器件生成二维码", "success")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 扫码读取接口
@app.route('/scan_qrcode', methods=['POST'])
def scan_qrcode():
    try:
        data = request.get_json()
        qr_data = data.get('data', '')

        # 解析二维码数据
        try:
            comp_data = json.loads(qr_data)
            comp_id = comp_data.get('id')

            if comp_id:
                comp = Component.query.get(comp_id)
                if comp:
                    return jsonify({
                        'code': 1,
                        'data': {
                            'id': comp.id,
                            'category': comp.category,
                            'model': comp.model,
                            'package': comp.package,
                            'supplier': comp.supplier,
                            'quantity': comp.quantity,
                            'unit': comp.unit,
                            'location': comp.location,
                            'price': comp.price,
                            'buy_time': comp.buy_time,
                            'channel': comp.channel,
                            'remark': comp.remark,
                            'img_path': comp.img_path,
                            'qrcode_path': comp.qrcode_path
                        }
                    })

            return jsonify({'code': 0, 'error': '未找到对应的元器件'})
        except json.JSONDecodeError:
            # 如果不是JSON，尝试按ID直接查找
            if qr_data.isdigit():
                comp = Component.query.get(int(qr_data))
                if comp:
                    return jsonify({
                        'code': 1,
                        'data': {
                            'id': comp.id,
                            'category': comp.category,
                            'model': comp.model,
                            'package': comp.package,
                            'supplier': comp.supplier,
                            'quantity': comp.quantity,
                            'unit': comp.unit,
                            'location': comp.location,
                            'price': comp.price,
                            'buy_time': comp.buy_time,
                            'channel': comp.channel,
                            'remark': comp.remark
                        }
                    })

            return jsonify({'code': 0, 'error': '无效的二维码数据'})
    except Exception as e:
        logger.error(f"扫码读取失败：{str(e)}")
        return jsonify({'code': 0, 'error': f'扫码读取失败：{str(e)}'})


# 静态文件访问
@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory(STATIC_FOLDER, path)


# -------------------------- 数据库迁移函数 --------------------------
def migrate_database():
    """检查并更新数据库表结构"""
    with app.app_context():
        try:
            from sqlalchemy import inspect, text

            # 创建inspector对象
            inspector = inspect(db.engine)

            # 检查表是否存在
            if 'component' in inspector.get_table_names():
                # 获取现有列
                columns = [col['name'] for col in inspector.get_columns('component')]
                logger.info(f"当前数据库列: {columns}")

                # 检查是否缺少必要的列
                missing_columns = []
                expected_columns = ['id', 'category', 'model', 'package', 'supplier', 'quantity', 'unit',
                                    'location', 'price', 'buy_time', 'channel', 'remark', 'img_path',
                                    'attach_path', 'qrcode_path']

                for col in expected_columns:
                    if col not in columns:
                        missing_columns.append(col)

                if missing_columns:
                    logger.info(f"检测到缺失的列: {missing_columns}")

                    # 尝试添加缺失的列
                    for col in missing_columns:
                        if col == 'qrcode_path':
                            try:
                                db.session.execute(
                                    text('ALTER TABLE component ADD COLUMN qrcode_path VARCHAR(255) DEFAULT ""'))
                                logger.info(f"已添加列: {col}")
                            except Exception as e:
                                logger.error(f"添加列 {col} 失败: {e}")
                    db.session.commit()
                else:
                    logger.info("数据库表结构完整")
            else:
                # 表不存在，创建新表
                logger.info("component表不存在，将创建新表")
                db.create_all()

        except Exception as e:
            logger.error(f"数据库迁移检查失败: {e}")
            # 如果出错，重新创建所有表
            try:
                db.drop_all()
                db.create_all()
                logger.info("数据库表已重建")
            except Exception as e2:
                logger.error(f"重建数据库表失败: {e2}")


# -------------------------- 程序入口（智能浏览器检测与打开） --------------------------
if __name__ == '__main__':
    # ... 之前的初始化代码保持不变 ...

    # 导入必要的模块
    import threading
    import time
    import subprocess
    import sys


    def detect_available_browsers():
        """检测工具中可用的浏览器"""
        browsers = []

        if IS_WINDOWS:
            # Windows工具浏览器检测
            browser_paths = [
                ("Chrome", [
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                    os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
                ]),
                ("Edge", [
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                ]),
                ("Firefox", [
                    r"C:\Program Files\Mozilla Firefox\firefox.exe",
                    r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                ]),
                ("Opera", [
                    r"C:\Program Files\Opera\launcher.exe",
                    r"C:\Program Files (x86)\Opera\launcher.exe"
                ]),
                ("Brave", [
                    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
                    r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
                ])
            ]
        else:
            # Linux/Mac工具浏览器检测
            browser_paths = [
                ("Chrome", ["google-chrome", "chrome", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]),
                ("Firefox", ["firefox", "/Applications/Firefox.app/Contents/MacOS/firefox"]),
                ("Safari", ["safari", "open -a Safari"]),
                ("Edge", ["microsoft-edge", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"]),
                ("Opera", ["opera", "/Applications/Opera.app/Contents/MacOS/Opera"])
            ]

        # 检测可用的浏览器
        for browser_name, paths in browser_paths:
            for path in paths:
                if isinstance(path, str) and (path.startswith("/") or "\\" in path or ":" in path):
                    # 完整路径检测
                    if os.path.exists(path):
                        browsers.append((browser_name, path))
                        break
                else:
                    # 命令检测（Linux/Mac）
                    try:
                        if IS_WINDOWS:
                            # Windows命令检测
                            if " " in path:
                                # 有空格的需要特殊处理
                                if os.path.exists(path.split()[0]):
                                    browsers.append((browser_name, path))
                                    break
                            else:
                                result = subprocess.run(['where', path], capture_output=True, text=True, shell=True)
                                if result.returncode == 0 and result.stdout.strip():
                                    browsers.append((browser_name, result.stdout.strip().split('\n')[0]))
                                    break
                        else:
                            # Linux/Mac命令检测
                            result = subprocess.run(['which', path], capture_output=True, text=True)
                            if result.returncode == 0:
                                browsers.append((browser_name, result.stdout.strip()))
                                break
                    except:
                        continue

        return browsers


    def open_browser_smart():
        """智能打开浏览器"""
        time.sleep(2.5)  # 等待服务器启动

        url = "http://localhost:5000"

        print(f"\n{'=' * 80}")
        print("元器件库存管理工具 - 服务已启动")
        print("=" * 80)
        print(f"🎯 服务地址: {url}")
        print(f"🌐 网络地址: http://{get_local_ip()}:5000")
        print("\n🔄 正在检测可用浏览器...")

        # 检测可用浏览器
        available_browsers = detect_available_browsers()

        if available_browsers:
            print(f"✓ 检测到 {len(available_browsers)} 个可用浏览器:")
            for i, (name, path) in enumerate(available_browsers, 1):
                print(f"  {i}. {name} ({os.path.basename(path)})")

            # 按优先级尝试打开浏览器
            browser_priority = ["Chrome", "Edge", "Firefox", "Brave", "Opera", "Safari"]
            opened = False

            for priority_name in browser_priority:
                for browser_name, browser_path in available_browsers:
                    if browser_name == priority_name:
                        try:
                            print(f"\n正在尝试打开 {browser_name}...")

                            if IS_WINDOWS:
                                # Windows: 使用完整路径
                                subprocess.Popen([browser_path, '--new-window', url])
                            else:
                                # Linux/Mac: 根据路径类型处理
                                if browser_path.startswith('/Applications'):
                                    # Mac应用程序
                                    subprocess.Popen(
                                        ['open', '-a', browser_path.replace('/Contents/MacOS/', '').rsplit('/', 1)[0],
                                         url])
                                else:
                                    # Linux可执行文件
                                    subprocess.Popen([browser_path, '--new-window', url])

                            print(f"✓ 已启动 {browser_name} 浏览器")
                            opened = True
                            break
                        except Exception as e:
                            print(f"✗ {browser_name} 启动失败: {e}")

                if opened:
                    break
        else:
            print("⚠ 未检测到常见浏览器")

        # 如果特定浏览器打开失败，尝试工具默认方式
        if not opened:
            print("\n尝试使用工具默认方式打开...")
            try:
                # 方法1: 使用os.startfile (Windows)
                if IS_WINDOWS:
                    os.startfile(url)
                    print("✓ 使用工具默认方式打开")
                    opened = True
                else:
                    # 方法2: 使用webbrowser模块
                    import webbrowser
                    webbrowser.open_new(url)
                    print("✓ 调用默认浏览器")
                    opened = True
            except Exception as e:
                print(f"✗ 工具默认方式失败: {e}")

        # 如果还是失败，提供详细指引
        if not opened:
            print("\n" + "!" * 80)
            print("❌ 自动打开浏览器失败")
            print("!" * 80)
            print("\n请手动执行以下操作:")
            print(f"1. 打开任意浏览器（Chrome/Edge/Firefox/360/QQ浏览器等）")
            print(f"2. 在地址栏输入: {url}")
            print(f"3. 或扫描下方二维码访问（如果支持）")

            # 尝试生成访问二维码（可选）
            try:
                import qrcode
                qr = qrcode.QRCode(version=1, box_size=2, border=2)
                qr.add_data(url)
                qr.make(fit=True)

                print("\n访问二维码:")
                qr.print_ascii(invert=True)
            except:
                pass

            print("\n💡 提示: 按 Ctrl+C 停止服务器")
            print("=" * 80)
        else:
            print(f"\n✅ 浏览器已成功打开!")
            print(f"💡 如果页面没有显示，请手动访问: {url}")
            print("=" * 80)


    def get_local_ip():
        """获取本地IP地址"""
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('8.8.8.8', 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return '127.0.0.1'


    # 显示启动信息
    print("\n" + "=" * 80)
    print("元器件库存管理工具 v1.0")
    print("=" * 80)
    print("🚀 启动流程:")
    print("  ✓ 数据库初始化完成")
    print("  ✓ 目录结构就绪")
    print("  ▶ 启动Web服务器...")
    print("  ⏳ 正在检测浏览器...")
    print("=" * 80)

    # 启动浏览器线程
    browser_thread = threading.Thread(target=open_browser_smart, daemon=True)
    browser_thread.start()

    # 启动Flask服务
    try:
        app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"⚠ 端口5000被占用，尝试5001端口...")


            # 更新浏览器打开的URL
            def open_on_5001():
                time.sleep(2.5)
                url_5001 = "http://localhost:5001"
                try:
                    if IS_WINDOWS:
                        os.startfile(url_5001)
                    else:
                        import webbrowser
                        webbrowser.open_new(url_5001)
                    print(f"✓ 请在浏览器中访问: {url_5001}")
                except:
                    print(f"请手动访问: {url_5001}")


            threading.Thread(target=open_on_5001, daemon=True).start()

            app.run(debug=True, host='0.0.0.0', port=5000)