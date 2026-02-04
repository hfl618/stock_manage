import sys
print("Python版本:", sys.version[:6])

modules = [
    'flask', 'flask_sqlalchemy', 'sqlalchemy',
    'pandas', 'numpy', 'qrcode', 'PIL',
    'openpyxl', 'win32com', 'pythoncom'
]

for module in modules:
    try:
        __import__(module)
        print(f"✓ {module} 导入成功")
    except ImportError as e:
        print(f"✗ {module} 导入失败: {e}")