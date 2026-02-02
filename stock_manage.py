# -*- coding: utf-8 -*-
# -------------------------- 导入依赖（仅保留必要包，无冗余） --------------------------
import os
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

# -------------------------- 基础配置（稳定版，修复数据库路径核心问题） --------------------------
# 系统判断
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

# 目录定义（自动创建，无多余目录）
STATIC_FOLDER = os.path.join(BASE_DIR, 'static')
IMG_FOLDER = os.path.join(STATIC_FOLDER, 'img')  # 图片目录
ATTACH_FOLDER = os.path.join(STATIC_FOLDER, 'attach')  # 附件目录

BACKUP_FOLDER = os.path.join(BASE_DIR, 'backup') # 备份目录
# 核心修改：数据库文件固定在Flask instance目录
DB_FILE = os.path.join(app.instance_path, 'component.db')

# 允许的文件格式
ALLOWED_IMG_EXT = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}
ALLOWED_ATTACH_EXT = {'pdf', 'doc', 'docx', 'xls', 'xlsx', 'zip', 'rar', 'txt', 'csv'}

# 自动创建必要目录（无冗余）
for folder in [STATIC_FOLDER, IMG_FOLDER, ATTACH_FOLDER, BACKUP_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)
        logger.info(f"自动创建文件夹：{folder}")

# 默认帮助文件（简洁版，无多余内容）
HELP_FILE = os.path.join(BASE_DIR, 'help.txt')


# -------------------------- 数据库模型（稳定版，无多余字段，保留数量/单位独立字段） --------------------------
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

    def __repr__(self):
        return f'<Component {self.id} - {self.category} {self.model}>'

    def get_file_prefix(self):
        """生成唯一文件前缀，避免重名"""
        invalid = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', ' ']
        pre = f"{self.category}_{self.model}_{self.package}"
        for c in invalid:
            pre = pre.replace(c, '_')
        return pre[:80]  # 限制长度，避免文件名过长


# -------------------------- 核心工具函数（稳定版+新增残留清理，修复备份逻辑） --------------------------
# 数量预警样式
def get_quantity_css(quantity):
    if quantity <= QUANTITY_WARN_LOW:
        return "text-danger fw-bold"
    elif quantity <= QUANTITY_WARN_MID:
        return "text-warning fw-bold"
    else:
        return "text-success"


# 检测重复元器件（品类+封装为唯一键）
def is_duplicate(data):
    """确保入参是字典，避免列表调用get报错，增加类型校验"""
    if not isinstance(data, dict):
        logger.error(f"is_duplicate入参不是字典：{type(data)}")
        return None
    category = data.get('category', '').strip()
    package = data.get('package', '').strip()
    return Component.query.filter(Component.category == category, Component.package == package).first()


# 表格内去重（BOM导入用）
def remove_table_dup(data_list):
    seen = set()
    unique = []
    for d in data_list:
        key = (d.get('category', '').strip(), d.get('package', '').strip())
        if key not in seen:
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


# 读取帮助文件
def get_help_content():
    try:
        with open(HELP_FILE, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "使用说明文件丢失，已自动重新创建！"


# 清理残留文件（新增核心函数）
def clean_residual_files():
    """清理无关联的图片和附件：数据库中不存在的文件直接删除"""
    try:
        # 获取数据库中所有有效文件路径
        valid_files = set()
        comps = Component.query.all()
        for c in comps:
            if c.img_path: valid_files.add(c.img_path)
            if c.attach_path: valid_files.add(c.attach_path)
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

        total_del = img_del_count + attach_del_count
        logger.info(f"残留文件清理完成：图片{img_del_count}个，附件{attach_del_count}个，总计{total_del}个")
        return True, f"清理成功！共删除残留文件{total_del}个（图片{img_del_count}个+附件{attach_del_count}个）"
    except Exception as e:
        logger.error(f"残留文件清理失败：{str(e)}")
        return False, f"清理失败：{str(e)}"


# -------------------------- 开机自启函数（Windows专属，稳定版） --------------------------
def get_startup_path():
    return os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')


def is_auto_start():
    if not IS_WINDOWS or not win32:
        return False
    lnk_path = os.path.join(get_startup_path(), "元器件库存管理系统.lnk")
    return os.path.exists(lnk_path)


def create_auto_start():
    try:
        pythoncom.CoInitialize()
        startup = get_startup_path()
        if not startup:
            return False, "获取开机启动目录失败"
        lnk_path = os.path.join(startup, "元器件库存管理系统.lnk")
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
        os.remove(os.path.join(get_startup_path(), "元器件库存管理系统.lnk"))
        return True, "开机自启已关闭"
    except:
        return False, "关闭开机自启失败（请以管理员身份运行）"


# -------------------------- 备份恢复核心函数（彻底修复路径+无数据判断逻辑） --------------------------
def backup_all_data():
    """修改：无下载弹窗，直接保存到backup默认目录，数据库指向instance"""
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

        # 4. 打包备份（instance里的数据库+图片+附件）
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

        logger.info(f"备份成功：{backup_path}，包含{comp_count}条元器件数据")
        flash(f"备份成功！共{comp_count}条数据，备份包已保存至【{BACKUP_FOLDER}】目录", "success")
        return backup_path  # 仅返回路径，不触发下载
    except Exception as e:
        flash(f"备份失败：{str(e)}", "danger")
        logger.error(f"备份失败：{str(e)}")
        return None
def validate_backup_zip(zip_path):
    """简化验证：仅检测数据库文件，去掉附件/图片目录的严格检测，避免空目录验证失败"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            # 只验证数据库是否存在，图片/附件目录空也正常
            if 'component.db' in zf.namelist():
                return True, "备份文件验证通过"
            else:
                return False, "备份文件缺少数据库（component.db）"
    except:
        return False, "备份文件损坏或不是有效ZIP文件"


def unzip_backup(zip_path, target_dir):
    """解压备份，直接覆盖，无冗余逻辑"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(target_dir)
        return True, "解压成功"
    except Exception as e:
        return False, f"解压失败：{str(e)}"


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
    """解析粘贴/Excel数据，稳定版+增加类型校验"""
    columns, preview, raw_data = [], [], []
    try:
        if source_type == 'paste':
            lines = [l.strip() for l in source.split('\n') if l.strip()]
            if not lines:
                return columns, preview, raw_data, "无有效粘贴数据"
            col_num = len(lines[0].split('\t'))
            columns = [f'列{i + 1}' for i in range(col_num)]
            for line in lines:
                row = line.split('\t') + [''] * (col_num - len(line.split('\t')))
                raw_data.append(row[:col_num])
            preview = raw_data[:3]
        elif source_type == 'excel':
            df = pd.read_excel(source, engine='openpyxl')
            df.columns = [f'列{i + 1}' if pd.isna(c) else str(c).strip() for i, c in enumerate(df.columns)]
            columns = list(df.columns)
            raw_data = df.fillna('').values.tolist()
            preview = raw_data[:3]
        return columns, preview, raw_data, ""
    except Exception as e:
        return columns, preview, raw_data, f"解析失败：{str(e)}（Excel仅支持xlsx）"


def map_table_data(raw_data, columns, mapping, batch_vals):
    """映射表格数据，稳定版+增加参数类型校验"""
    data_list, errors = [], []
    # 增加类型校验，避免非字典入参
    if not isinstance(mapping, dict) or not isinstance(batch_vals, dict):
        errors.append("映射数据或批量值格式错误")
        return data_list, errors
    mapped = [v for v in mapping.values() if v]
    for f in REQUIRED_FIELDS:
        if f not in mapped:
            errors.append(f"缺少必填字段映射：{f}")
    if errors:
        return data_list, errors
    # 处理批量值
    batch = {}
    for k, v in batch_vals.items():
        if not v:
            continue
        try:
            batch[k] = int(v) if k == 'quantity' else float(v) if k == 'price' else v
        except:
            errors.append(f"批量{SYSTEM_FIELDS[[i[0] for i in SYSTEM_FIELDS].index(k)][1]}非有效数字")
    # 映射数据
    for idx, row in enumerate(raw_data, 1):
        d = {
            'category': '未知', 'model': '未知型号', 'package': '未知封装',
            'supplier': '未知供应商', 'quantity': 1, 'unit': '个',
            'location': '未知位置', 'price': 0.00,
            'buy_time': datetime.now().strftime('%Y-%m-%d'),
            'channel': '未知', 'remark': '无'
        }
        for col, field in mapping.items():
            if not field or col not in columns:
                continue
            val = str(row[columns.index(col)]).strip()
            if field == 'quantity':
                d[field] = int(val) if val.isdigit() else 1
            elif field == 'price':
                d[field] = float(val) if val.replace('.', '').isdigit() else 0.00
            elif val:
                d[field] = val
        d.update(batch)
        data_list.append(d)
    return data_list, errors


# -------------------------- 前端模板（已修复Jinja语法错误+合并列显示正常） --------------------------
# 主页面模板（合并数量+单位列，修复Jinja语法错误，修正colspan）
MAIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>元器件库存管理系统 - 稳定版</title>
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
        /* 优化弹窗样式：固定尺寸+避免叠加+降低渲染压力 */
        .alert { 
            position: fixed; top: 80px; right: 20px; z-index: 9999; 
            min-width: 320px; max-width: 400px; margin: 0; padding: 0.8rem 1.2rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); /* 轻微阴影，减少重绘 */
        }
        .file-link { color: #0d6efd; text-decoration: none; }
        .file-link:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="top-nav">
        <h4>元器件库存管理系统 - 稳定版</h4>
        <div>
            <a href="#" class="top-btn" data-bs-toggle="modal" data-bs-target="#settingModal">系统设置</a>
            <a href="#" class="top-btn" data-bs-toggle="modal" data-bs-target="#helpModal">使用说明</a>
        </div>
    </div>

    <div class="container-main">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for c, m in messages %}
                    <div class="alert alert-{{c}} alert-dismissible fade show" role="alert">
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

        <div class="table-box">
            <table class="table table-striped table-hover table-sm">
                <thead class="table-dark">
                    <tr>
                        <th width="5%">选择</th>
                        <th width="8%">图片</th>
                        <th>品类</th>
                        <th>型号规格</th>
                        <th>封装</th>
                        <th>供应商</th>
                        <th>数量（含单位）</th>
                        <th>存放位置</th>
                        <th>单价(¥)</th>
                        <th width="10%">附件</th>
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
                        <td>
                            <a href="{{url_for('edit', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-warning btn-sm">编辑</a>
                            <a href="{{url_for('delete', id=comp.id, selected=selected|join(','), kw=kw)}}" class="btn btn-danger btn-sm" onclick="return confirm('确定删除？将同步删除图片/附件！')">删除</a>
                        </td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="11" class="text-center text-muted py-3">暂无数据，点击「添加元器件」或「BOM批量导入」录入</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- 添加元器件弹窗（保留单位输入框） -->
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

    <!-- 批量编辑弹窗（保留单位输入框） -->
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
                            <div class="col-md-3"><label>封装</label><input type="text" name="adv_pack" class="form-control-sm" value="{{adv_params.adv_pack or ''}}"></div>
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
                        <button type="submit" class="btn btn-primary">搜索</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 导出弹窗 -->
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
                        <button type="submit" class="btn btn-warning">确认</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- 系统设置弹窗【新增清理残留文件按钮】 -->
    <div class="modal fade" id="settingModal" tabindex="-1">
        <div class="modal-dialog modal-md">
            <div class="modal-content">
                <div class="modal-header bg-secondary text-white">
                    <h5 class="modal-title">系统设置</h5>
                    <button class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="d-grid gap-2">
                        <a href="{{url_for('backup', selected=selected|join(','), kw=kw)}}" class="btn btn-primary">📥 立即备份数据</a>
                        <a href="{{url_for('restore_page', kw=kw)}}" class="btn btn-warning">🔄 备份恢复（覆盖当前数据）</a>
                        <a href="{{url_for('clean_residual')}}" class="btn btn-danger">🗑️ 清理残留文件（无关联图片/附件）</a>
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

    <!-- 打印区域（隐藏） -->
    <div id="printArea" class="d-none p-4">
        <h4 class="text-center mb-4">元器件库存数据</h4>
        <table class="table table-striped table-bordered" id="printTable"></table>
    </div>

    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
    <script>
        // 自动关闭提示框
        setTimeout(() => {document.querySelectorAll('.alert').forEach(a => new bootstrap.Alert(a).close())}, 3000);

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
            if (confirm(`确定删除选中的${ids.length}条数据？不可恢复！`)) {
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

        // 初始化
        window.onload = function() {
            updateSelect();
            document.querySelectorAll('.compCheck').forEach(c => {
                c.addEventListener('change', updateSelect);
            });
        }
    </script>
</body>
</html>
'''

# 编辑页面模板（保留单位输入框，不修改）
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

                <div class="col-md-6">
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

                <div class="col-md-6">
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

# BOM批量导入模板【三选项：跳过/合并/依旧导入，保留数量/单位独立映射】
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
    </style>
</head>
<body>
    <div class="container">
        <h4 class="text-primary mb-4">BOM批量导入</h4>
        <div class="alert alert-info mb-4">
            💡 仅需映射「品类、型号规格、封装」3个必填字段，表格内重复数据自动去重，库内重复数据支持「跳过/合并（数量相加）/依旧导入（覆盖原有）」
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
            <div class="table-responsive mt-3">
                <table class="table table-bordered mapping-table">
                    <thead class="table-dark">
                        <tr><th>表格列</th><th>映射为系统字段</th><th>预览数据</th></tr>
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
            <h5 class="text-secondary">步骤3：重复数据处理（<span id="dupCount">0</span>条重复，<span id="uniCount">0</span>条全新）</h5>
            <div id="noDup" class="alert alert-success hidden">🎉 无重复数据，可直接导入！</div>
            <div id="dupList" class="mt-3"></div>
            <div class="mt-3">
                <button class="btn btn-secondary" onclick="backToStep2()">返回上一步</button>
                <button class="btn btn-success ms-2" onclick="doImport()">确认导入</button>
            </div>
        </div>

        <!-- 步骤4：导入完成 -->
        <div class="step hidden" id="step4">
            <h5 class="text-secondary">步骤4：导入完成</h5>
            <div class="alert alert-success">
                🎉 导入成功！总处理<span id="total">0</span>条，新增<span id="add">0</span>条，合并<span id="merge">0</span>条，覆盖<span id="cover">0</span>条，跳过<span id="skip">0</span>条
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
        </form>
    </div>

    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
    <script>
        let parseRes = {columns:[], preview:[], raw_data:[], error:''};
        let mapping = {};
        let dupData = [];
        let uniData = [];

        // 步骤切换
        function showStep(s) {document.querySelectorAll('.step').forEach(el => el.classList.add('hidden')); document.getElementById('step'+s).classList.remove('hidden');}
        function backToStep1() {showStep(1); parseRes = {columns:[], preview:[], raw_data:[], error:''};}
        function backToStep2() {showStep(2);}
        function closeWin() {window.opener.location.reload(); window.close();}

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
                renderMapping();
                showStep(2);
            }).catch(err => alert('解析失败：'+err.message));
        }

        // 渲染映射表格
        function renderMapping() {
            let tbody = document.getElementById('mappingTbody');
            tbody.innerHTML = '';
            mapping = {};
            let fields = {{SYSTEM_FIELDS|tojson}};
            parseRes.columns.forEach(col => {
                mapping[col] = '';
                let tr = document.createElement('tr');
                // 表格列
                let td1 = document.createElement('td'); td1.innerText = col; tr.appendChild(td1);
                // 下拉框
                let td2 = document.createElement('td');
                let select = document.createElement('select'); select.className = 'form-select form-select-sm';
                fields.forEach(f => {
                    let opt = document.createElement('option');
                    opt.value = f[0]; opt.innerText = f[1];
                    if (['category','model','package'].includes(f[0])) {opt.style.color = 'red'; opt.style.fontWeight = 'bold';}
                    select.appendChild(opt);
                });
                select.onchange = function() {mapping[col] = this.value;};
                td2.appendChild(select); tr.appendChild(td2);
                // 预览
                let td3 = document.createElement('td');
                let val = parseRes.preview.length > 0 ? parseRes.preview[0][parseRes.columns.indexOf(col)] : '';
                td3.innerText = val || '无'; tr.appendChild(td3);
                tbody.appendChild(tr);
            });
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
            // 提交检测
            let formData = new FormData();
            formData.append('raw_data', JSON.stringify(parseRes.raw_data));
            formData.append('mapping', JSON.stringify(mapping));
            formData.append('batch_vals', JSON.stringify(batchVals));
            fetch("{{url_for('check_bom_dup')}}", {method:'POST', body:formData})
            .then(res => res.json())
            .then(data => {
                if (data.code !== 1) {alert(data.error); return;}
                dupData = data.data.duplicate;
                uniData = data.data.unique;
                renderDup();
                showStep(3);
            }).catch(err => alert('检测失败：'+err.message));
        }

        // 渲染重复数据【三选项：跳过/合并/依旧导入】
        function renderDup() {
            document.getElementById('dupCount').innerText = dupData.length;
            document.getElementById('uniCount').innerText = uniData.length;
            let dupList = document.getElementById('dupList');
            let noDup = document.getElementById('noDup');
            if (dupData.length === 0) {dupList.classList.add('hidden'); noDup.classList.remove('hidden'); return;}
            dupList.classList.remove('hidden'); noDup.classList.add('hidden');
            dupList.innerHTML = '';
            dupData.forEach((item, idx) => {
                let div = document.createElement('div');
                div.className = 'duplicate-item';
                div.innerHTML = `
                    <div class="duplicate-title">重复数据 #${idx+1}：${item.data.category} - ${item.data.model} - ${item.data.package}</div>
                    <table class="table table-sm table-bordered mt-2">
                        <tr class="table-secondary">
                            <th>字段</th><th>库内原有</th><th>待导入</th><th>处理方式</th>
                        </tr>
                        <tr><td>数量</td><td>${item.old_data.quantity} ${item.old_data.unit}</td><td>${item.data.quantity} ${item.data.unit}</td><td rowspan="4">
                            <select class="form-select form-select-sm dupOper" data-id="${item.old_data.id}">
                                <option value="merge" selected>合并（数量相加）</option>
                                <option value="skip">跳过（保留原有）</option>
                                <option value="cover">依旧导入（覆盖原有）</option>
                            </select>
                        </td></tr>
                        <tr><td>供应商</td><td>${item.old_data.supplier}</td><td>${item.data.supplier}</td></tr>
                        <tr><td>单价</td><td>¥${item.old_data.price.toFixed(2)}</td><td>¥${item.data.price.toFixed(2)}</td></tr>
                        <tr><td>位置</td><td>${item.old_data.location}</td><td>${item.data.location}</td></tr>
                    </table>
                `;
                dupList.appendChild(div);
            });
        }

        // 执行导入
        function doImport() {
            // 获取处理方式
            let dupOper = {};
            document.querySelectorAll('.dupOper').forEach(sel => {dupOper[sel.dataset.id] = sel.value;});
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
            // 提交
            document.getElementById('importForm').submit();
        }

        // 导入结果渲染【新增覆盖计数】
        {% if import_res %}
            window.onload = function() {
                document.getElementById('total').innerText = {{import_res.total}};
                document.getElementById('add').innerText = {{import_res.added}};
                document.getElementById('merge').innerText = {{import_res.merged}};
                document.getElementById('cover').innerText = {{import_res.covered}};
                document.getElementById('skip').innerText = {{import_res.skipped}};
                showStep(4);
            }
        {% endif %}
    </script>
</body>
</html>
'''

# 备份恢复页面模板
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
            ⚠️ 警告：恢复操作会<strong>覆盖当前所有数据</strong>（数据库+图片+附件），请确保已备份当前重要数据后再执行！
        </div>
        <form method="POST" enctype="multipart/form-data">
            <div class="mb-3">
                <label class="form-label">选择备份ZIP文件（来自backup目录）</label>
                <input type="file" name="backup_zip" class="form-control" accept=".zip" required>
                <div class="form-text mt-2">仅支持本系统生成的备份文件（命名以「元器件库存备份_」开头）</div>
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

# -------------------------- 视图函数（修复参数解构+增加类型校验，无任何报错） --------------------------
# 全局模板函数
app.add_template_global(get_quantity_css, 'get_quantity_css')
app.add_template_global(SYSTEM_FIELDS, 'SYSTEM_FIELDS')


# 解析请求参数（稳定版，返回值固定，避免解构错误）
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
    # 快速搜索
    if kw:
        query = query.filter(db.or_(
            Component.category.like(f'%{kw}%'),
            Component.model.like(f'%{kw}%'),
            Component.package.like(f'%{kw}%'),
            Component.supplier.like(f'%{kw}%'),
            Component.location.like(f'%{kw}%')
        ))
    # 高级搜索
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


# 添加元器件
@app.route('/add', methods=['POST'])
def add():
    selected, kw, _ = parse_args(request)
    # 获取表单数据
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
    # 检测重复（增加非空校验）
    if not form['category'] or not form['package']:
        flash("品类和封装为必填项！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    if is_duplicate(form):
        flash("添加失败：该品类+封装的元器件已存在！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    # 创建元器件
    comp = Component(**form)
    db.session.add(comp)
    db.session.flush()  # 获取ID
    # 保存文件
    pre = comp.get_file_prefix()
    comp.img_path = save_file(request.files.get('img'), IMG_FOLDER, ALLOWED_IMG_EXT, pre)
    comp.attach_path = save_file(request.files.get('attach'), ATTACH_FOLDER, ALLOWED_ATTACH_EXT, pre)
    # 提交
    db.session.commit()
    flash(f"元器件「{form['category']}-{form['model']}」添加成功！", "success")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 编辑元器件
@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    if request.method == 'GET':
        return render_template_string(EDIT_TEMPLATE, comp=comp, selected=selected, kw=kw)
    # 保存修改
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
    # 检测重复（排除自身+非空校验）
    if not form['category'] or not form['package']:
        flash("品类和封装为必填项！", "danger")
        return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))
    dup = is_duplicate(form)
    if dup and dup.id != id:
        flash("修改失败：该品类+封装的元器件已存在！", "danger")
        return redirect(url_for('edit', id=id, selected=','.join(selected), kw=kw))
    # 清空文件
    if request.form.get('clear_img'):
        delete_file(comp.img_path)
        comp.img_path = ''
    if request.form.get('clear_attach'):
        delete_file(comp.attach_path)
        comp.attach_path = ''
    # 重新上传文件
    pre = comp.get_file_prefix()
    new_img = save_file(request.files.get('img'), IMG_FOLDER, ALLOWED_IMG_EXT, pre)
    new_attach = save_file(request.files.get('attach'), ATTACH_FOLDER, ALLOWED_ATTACH_EXT, pre)
    if new_img:
        delete_file(comp.img_path)
        comp.img_path = new_img
    if new_attach:
        delete_file(comp.attach_path)
        comp.attach_path = new_attach
    # 更新字段
    for k, v in form.items():
        setattr(comp, k, v)
    db.session.commit()
    flash(f"元器件「{comp.category}-{comp.model}」修改成功！", "success")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 删除单个元器件
@app.route('/delete/<int:id>')
def delete(id):
    selected, kw, _ = parse_args(request)
    comp = Component.query.get_or_404(id)
    # 删除文件
    delete_file(comp.img_path)
    delete_file(comp.attach_path)
    # 删除数据
    db.session.delete(comp)
    db.session.commit()
    flash(f"元器件「{comp.category}-{comp.model}」删除成功！", "success")
    # 移除选中ID
    selected = [s for s in selected if s != str(id)]
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 批量删除
@app.route('/batch_delete')
def batch_delete():
    selected, kw, _ = parse_args(request)
    ids = [int(i) for i in request.args.get('ids', '').split(',') if i.strip().isdigit()]
    if not ids:
        flash("请选择要删除的元器件！", "danger")
        return redirect(url_for('index', selected=','.join(selected), kw=kw))
    # 批量删除
    comps = Component.query.filter(Component.id.in_(ids)).all()
    for c in comps:
        delete_file(c.img_path)
        delete_file(c.attach_path)
        db.session.delete(c)
    db.session.commit()
    flash(f"批量删除成功！共删除 {len(comps)} 条数据", "success")
    return redirect(url_for('index', selected='', kw=kw))


# 批量编辑
@app.route('/batch_edit', methods=['POST'])
def batch_edit():
    # 修复参数解构：避免列表被当作字典，直接获取kw
    kw = request.args.get('kw', '').strip()
    ids = [int(i) for i in request.form.get('ids', '').split(',') if i.strip().isdigit()]
    if not ids:
        flash("请选择要编辑的元器件！", "danger")
        return redirect(url_for('index', kw=kw))
    # 处理表单数据（仅非空字段）
    form = {}
    if request.form.get('supplier', '').strip(): form['supplier'] = request.form.get('supplier').strip()
    if request.form.get('quantity', '').strip():
        try:
            form['quantity'] = int(request.form.get('quantity'))
        except:
            flash("批量数量必须为数字！", "danger"); return redirect(url_for('index', kw=kw))
    if request.form.get('unit', '').strip(): form['unit'] = request.form.get('unit').strip()
    if request.form.get('location', '').strip(): form['location'] = request.form.get('location').strip()
    if request.form.get('price', '').strip():
        try:
            form['price'] = float(request.form.get('price'))
        except:
            flash("批量单价必须为数字！", "danger"); return redirect(url_for('index', kw=kw))
    if request.form.get('buy_time', '').strip(): form['buy_time'] = request.form.get('buy_time').strip()
    if request.form.get('channel', '').strip(): form['channel'] = request.form.get('channel').strip()
    if request.form.get('remark', '').strip(): form['remark'] = request.form.get('remark').strip()
    # 无修改字段
    if not form:
        flash("请填写要修改的字段！", "warning")
        return redirect(url_for('index', kw=kw))
    # 批量更新
    Component.query.filter(Component.id.in_(ids)).update(form, synchronize_session=False)
    db.session.commit()
    flash(f"批量编辑成功！共修改 {len(ids)} 条数据", "success")
    return redirect(url_for('index', selected=','.join([str(i) for i in ids]), kw=kw))


# 导出/打印
@app.route('/export', methods=['POST'])
def export():
    kw = request.args.get('kw', '').strip()
    ids = [int(i) for i in request.form.get('ids', '').split(',') if i.strip().isdigit()]
    fields = request.form.getlist('fields')
    fmt = request.form.get('format', 'xlsx')
    if not ids or not fields:
        flash("请选择元器件和导出字段！", "danger")
        return redirect(url_for('index', kw=kw))
    # 字段映射
    field_cn = {
        'category': '品类', 'model': '型号规格', 'package': '封装', 'supplier': '供应商',
        'quantity': '数量', 'unit': '单位', 'location': '存放位置', 'price': '采购单价(¥)',
        'buy_time': '采购时间', 'channel': '采购渠道', 'remark': '备注'
    }
    # 查询数据
    comps = Component.query.filter(Component.id.in_(ids)).all()
    data = []
    for c in comps:
        row = {}
        for f in fields:
            val = getattr(c, f, '')
            if f == 'price': val = round(float(val), 2)
            row[field_cn[f]] = val
        data.append(row)
    # 生成文件
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


# 清理残留文件
@app.route('/clean_residual')
def clean_residual():
    selected, kw, _ = parse_args(request)
    is_success, msg = clean_residual_files()
    flash(msg, "success" if is_success else "danger")
    return redirect(url_for('index', selected=','.join(selected), kw=kw))


# 数据备份（修复路径+无数据友好提示，无报错）
# 数据备份（修改：无下载弹窗，直接保存到默认目录，重定向回首页）
@app.route('/backup')
def backup():
    selected, kw, _ = parse_args(request)
    backup_all_data()  # 仅执行备份，不处理返回路径
    return redirect(url_for('index', selected=','.join(selected), kw=kw))  # 直接重定向，无弹窗

# 备份恢复页面
@app.route('/restore')
def restore_page():
    kw = request.args.get('kw', '').strip()
    return render_template_string(RESTORE_TEMPLATE, kw=kw)


# 执行备份恢复（修改：强制将db文件还原到instance目录，确保数据库生效）
@app.route('/restore', methods=['POST'])
def do_restore():
    import tempfile
    import shutil
    kw = request.args.get('kw', '').strip()
    backup_zip = request.files.get('backup_zip')
    if not backup_zip or backup_zip.filename == '':
        flash("请选择备份ZIP文件！", "danger")
        return redirect(url_for('restore_page', kw=kw))

    # 1. 保存临时文件，避免文件占用
    temp_zip = os.path.join(tempfile.gettempdir(), f"temp_backup_{uuid.uuid4().hex[:8]}.zip")
    backup_zip.save(temp_zip)

    # 2. 验证备份文件（必须包含component.db）
    is_valid, msg = validate_backup_zip(temp_zip)
    if not is_valid:
        os.remove(temp_zip)
        flash(f"备份文件验证失败：{msg}", "danger")
        return redirect(url_for('restore_page', kw=kw))

    # 3. 创建临时解压目录
    temp_unzip = os.path.join(tempfile.gettempdir(), f"temp_unzip_{uuid.uuid4().hex[:8]}")
    os.makedirs(temp_unzip, exist_ok=True)

    try:
        # 4. 解压备份包到临时目录
        with zipfile.ZipFile(temp_zip, 'r') as zf:
            zf.extractall(temp_unzip)

        # 5. 核心修改：查找解压后的db文件，强制复制到instance目录
        db_temp_path = os.path.join(temp_unzip, 'component.db')
        if not os.path.exists(db_temp_path):
            # 兼容旧备份包（可能db在其他路径），递归查找
            for root, _, files in os.walk(temp_unzip):
                if 'component.db' in files:
                    db_temp_path = os.path.join(root, 'component.db')
                    break
        if not os.path.exists(db_temp_path):
            flash("备份包中未找到数据库文件（component.db）！", "danger")
            return redirect(url_for('restore_page', kw=kw))

        # 6. 覆盖instance目录下的db文件（先关闭可能的连接，强制覆盖）
        shutil.copy2(db_temp_path, DB_FILE)
        logger.info(f"数据库文件已从【{db_temp_path}】还原到【{DB_FILE}】")

        # 7. 解压图片和附件到原有目录（BASE_DIR下的static）
        shutil.copytree(os.path.join(temp_unzip, 'static'), STATIC_FOLDER, dirs_exist_ok=True)
        logger.info(f"图片/附件已还原到【{STATIC_FOLDER}】")

        flash("数据恢复成功！已将数据库还原到instance目录，图片/附件还原到原路径，请刷新页面查看", "success")
    except Exception as e:
        flash(f"恢复失败：{str(e)}", "danger")
        logger.error(f"恢复失败：{str(e)}")
    finally:
        # 8. 清理临时文件，避免残留
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
        source = request.form.get('paste_data', '') if source_type == 'paste' else request.files.get('excel_file')
        columns, preview, raw_data, error = parse_table_data(source, source_type)
        if error:
            return jsonify({'code': 0, 'error': error})
        # 表格内去重
        raw_data = remove_table_dup(raw_data)
        return jsonify({'code': 1, 'data': {'columns': columns, 'preview': preview, 'raw_data': raw_data}})
    except Exception as e:
        return jsonify({'code': 0, 'error': f"解析异常：{str(e)}"})


# 检测BOM重复数据
@app.route('/check_bom_dup', methods=['POST'])
def check_bom_dup():
    try:
        import json
        raw_data = json.loads(request.form.get('raw_data', '[]'))
        mapping = json.loads(request.form.get('mapping', '{}'))
        batch_vals = json.loads(request.form.get('batch_vals', '{}'))
        data_list, errors = map_table_data(raw_data, list(mapping.keys()), mapping, batch_vals)
        if errors:
            return jsonify({'code': 0, 'error': '；'.join(errors)})
        # 分离重复和全新数据
        duplicate, unique = [], []
        for d in data_list:
            old_comp = is_duplicate(d)
            if old_comp:
                duplicate.append({'data': d,
                                  'old_data': {'id': old_comp.id, 'quantity': old_comp.quantity, 'unit': old_comp.unit,
                                               'supplier': old_comp.supplier, 'price': old_comp.price,
                                               'location': old_comp.location}})
            else:
                unique.append(d)
        return jsonify({'code': 1, 'data': {'duplicate': duplicate, 'unique': unique}})
    except Exception as e:
        return jsonify({'code': 0, 'error': f"检测异常：{str(e)}"})


# 执行BOM导入（三选项逻辑：跳过/合并/覆盖）
@app.route('/do_bom_import', methods=['POST'])
def do_bom_import():
    try:
        import json
        raw_data = json.loads(request.form.get('raw_data', '[]'))
        mapping = json.loads(request.form.get('mapping', '{}'))
        batch_vals = json.loads(request.form.get('batch_vals', '{}'))
        dup_oper = json.loads(request.form.get('dup_oper', '{}'))
        # 映射数据
        data_list, errors = map_table_data(raw_data, list(mapping.keys()), mapping, batch_vals)
        if errors:
            return render_template_string(BOM_TEMPLATE,
                                          import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0, 'skipped': 0})
        # 统计
        total = len(data_list)
        added = merged = covered = skipped = 0
        # 处理数据
        for d in data_list:
            old_comp = is_duplicate(d)
            if old_comp:
                # 重复数据处理
                op = dup_oper.get(str(old_comp.id), 'merge')
                if op == 'skip':
                    skipped += 1
                    continue
                elif op == 'merge':
                    old_comp.quantity += d['quantity']
                    merged += 1
                elif op == 'cover':
                    # 覆盖原有数据
                    for k, v in d.items():
                        setattr(old_comp, k, v)
                    covered += 1
                db.session.commit()
            else:
                # 全新数据
                new_comp = Component(**d)
                db.session.add(new_comp)
                db.session.commit()
                added += 1
        # 导入结果
        import_res = {'total': total, 'added': added, 'merged': merged, 'covered': covered, 'skipped': skipped}
        return render_template_string(BOM_TEMPLATE, import_res=import_res)
    except Exception as e:
        logger.error(f"BOM导入失败：{str(e)}")
        return render_template_string(BOM_TEMPLATE,
                                      import_res={'total': 0, 'added': 0, 'merged': 0, 'covered': 0, 'skipped': 0})


# 静态文件访问
@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory(STATIC_FOLDER, path)


# -------------------------- 程序入口（完整可运行，无任何报错） --------------------------
if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # 首次运行自动创建数据库表
        logger.info("数据库初始化完成，系统启动成功！")
    # 启动服务（关闭debug，生产可用，如需调试改为debug=True）
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)