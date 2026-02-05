"""
元器件AI识别导入工具 - 完整单文件版
功能：多步骤图片上传 + OCR识别 + AI解析 + 表格管理
特点：自动读取.env + 手动配置API + 自动弹出浏览器
"""

import os
import json
import re
import time
import base64
import warnings
import threading
import platform
import subprocess
import webbrowser
from datetime import datetime

import requests
import pandas as pd
import numpy as np
from PIL import Image, ImageEnhance, ImageOps
try:
    import easyocr
except Exception:
    easyocr = None

# 可选使用 OpenCV 做更强的预处理
try:
    import cv2
except Exception:
    cv2 = None
from flask import Flask, render_template_string, request, jsonify, session, send_file
from werkzeug.utils import secure_filename
from dotenv import load_dotenv, set_key, dotenv_values
from difflib import get_close_matches

warnings.filterwarnings("ignore")


# ==================== 配置文件 ====================
class Config:
    """应用配置"""
    # 环境文件路径
    ENV_FILE = '.env'

    # Flask配置
    SECRET_KEY = os.getenv('SECRET_KEY', os.urandom(24).hex())
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 20 * 1024 * 1024  # 20MB
    JSON_AS_ASCII = False

    # AI配置
    API_KEY = os.getenv('API_KEY', '')
    API_URL = os.getenv('API_URL', 'https://api.openai.com/v1')
    AI_MODEL = os.getenv('AI_MODEL', 'gpt-3.5-turbo')
    AI_TEMPERATURE = 0.1
    AI_MAX_TOKENS = 4000
    AI_TIMEOUT = 60

    # 图片配置
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

    # 浏览器配置
    BROWSER_DELAY = 2.0
    TARGET_URL = "http://127.0.0.1:5000"

    # Flask运行配置
    HOST = '0.0.0.0'
    PORT = 5000
    DEBUG = True

    @classmethod
    def load_env(cls):
        """加载环境变量"""
        try:
            # 优先从 .env 文件读取
            env = dotenv_values(cls.ENV_FILE) if os.path.exists(cls.ENV_FILE) else {}
            # 兼容 .env 中只有裸 key 的情况（例如只有 sk-xxxx 行）
            if os.path.exists(cls.ENV_FILE):
                try:
                    with open(cls.ENV_FILE, 'r', encoding='utf-8') as f:
                        raw = f.read()
                except:
                    raw = ''

                # 如果没有明确的 API_KEY 键，尝试在裸文本中查找 sk- 前缀的 key
                if not env.get('API_KEY'):
                    m = re.search(r"\b(sk-[A-Za-z0-9\-_.]{20,})\b", raw)
                    if m:
                        env = dict(env)
                        env['API_KEY'] = m.group(1)

            # 作为兜底，加载环境变量
            load_dotenv(cls.ENV_FILE)

            cls.API_KEY = env.get('API_KEY') or os.getenv('API_KEY', cls.API_KEY)
            cls.API_URL = env.get('API_URL') or os.getenv('API_URL', cls.API_URL)
            cls.AI_MODEL = env.get('AI_MODEL') or os.getenv('AI_MODEL', cls.AI_MODEL)
            # SECRET_KEY 也可通过环境注入
            cls.SECRET_KEY = os.getenv('SECRET_KEY', cls.SECRET_KEY)
            return True
        except Exception as e:
            print(f"⚠️ 加载{cls.ENV_FILE}失败: {e}")
            return False

    @classmethod
    def save_env(cls, api_key, api_url, ai_model):
        """保存配置到环境文件"""
        try:
            # 读取现有键，更新后写回，兼容任意现有格式
            existing = dotenv_values(cls.ENV_FILE) if os.path.exists(cls.ENV_FILE) else {}
            new_env = dict(existing)
            new_env['API_KEY'] = api_key or ''
            new_env['API_URL'] = api_url or ''
            new_env['AI_MODEL'] = ai_model or ''

            # 保留已有 SECRET_KEY（如果存在）或使用当前值
            if 'SECRET_KEY' not in new_env or not new_env.get('SECRET_KEY'):
                new_env['SECRET_KEY'] = cls.SECRET_KEY or os.getenv('SECRET_KEY', '')

            # 写入文件（覆盖写法，保证格式正确）
            lines = []
            for k, v in new_env.items():
                # 如果值包含特殊字符，保持原样（不加引号）
                lines.append(f"{k}={v}")

            with open(cls.ENV_FILE, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines) + '\n')

            # 同步到运行时配置
            cls.API_KEY = api_key
            cls.API_URL = api_url
            cls.AI_MODEL = ai_model
            return True
        except Exception as e:
            print(f"❌ 保存.env失败: {e}")
            return False


# 加载配置
Config.load_env()

# ==================== OCR初始化 ====================
try:
    OCR_READER = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    print("[OK] EasyOCR initialized successfully")
except Exception as e:
    print(f"[ERROR] EasyOCR init failed: {e}")
    OCR_READER = None

# ==================== Flask应用初始化 ====================
app = Flask(__name__)
app.config['SECRET_KEY'] = Config.SECRET_KEY
app.config['UPLOAD_FOLDER'] = Config.UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = Config.MAX_CONTENT_LENGTH
app.config['JSON_AS_ASCII'] = Config.JSON_AS_ASCII
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


# ==================== OCR功能函数 ====================
def allowed_file(filename):
    """检查文件类型是否允许"""
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS


def save_base64_image(base64_str, filename):
    """保存Base64图片到文件"""
    try:
        # 移除Base64前缀
        if ',' in base64_str:
            base64_str = base64_str.split(',')[1]

        # 解码并保存
        image_data = base64.b64decode(base64_str)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        with open(filepath, 'wb') as f:
            f.write(image_data)

        print(f"✅ Base64图片保存成功: {filepath}")
        return filepath
    except Exception as e:
        print(f"❌ 保存Base64图片失败: {e}")
        return None


def preprocess_image(image_path):
    """预处理图片以提高OCR准确率"""
    try:
        print(f"📷 预处理图片: {image_path}")

        # 打开图片
        img = Image.open(image_path)
        original_size = img.size
        print(f"  原始尺寸: {original_size[0]}x{original_size[1]}")

        # 确保图片为RGB模式
        if img.mode != 'RGB':
            img = img.convert('RGB')
            print("  转换为RGB模式")

        # 计算缩放因子，确保最小边至少600像素
        min_target_size = 600
        max_target_size = 2000

        # 计算缩放因子
        width, height = original_size
        if min(width, height) >= min_target_size:
            scale = 1.0
        else:
            scale = min_target_size / min(width, height)

        # 如果缩放后太大，限制最大尺寸
        new_width = int(width * scale)
        new_height = int(height * scale)

        if max(new_width, new_height) > max_target_size:
            scale = max_target_size / max(width, height)
            new_width = int(width * scale)
            new_height = int(height * scale)

        # 缩放图片
        if scale != 1.0:
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            print(f"  缩放图片: {new_width}x{new_height} (缩放因子: {scale:.2f})")

        # 增强对比度
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.8)
        print("  对比度增强: 1.8x")

        # 增强锐度
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(1.5)
        print("  锐度增强: 1.5x")

        # 转换为灰度（PIL）
        img = img.convert('L')
        print("  转换为灰度图")

        # 使用 OpenCV 进行更强的去噪/对比度和二值化（如果可用）
        if cv2 is not None:
            arr = np.array(img)

            # 去噪：双边滤波保留边缘
            try:
                arr = cv2.bilateralFilter(arr, d=9, sigmaColor=75, sigmaSpace=75)
            except Exception:
                pass

            # 自适应直方图均衡（CLAHE）增强局部对比度
            try:
                clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
                arr = clahe.apply(arr)
            except Exception:
                pass

            # 自适应阈值二值化，减少背景噪声
            try:
                arr = cv2.adaptiveThreshold(arr, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                            cv2.THRESH_BINARY, 15, 8)
            except Exception:
                # 退回 Otsu
                try:
                    _, arr = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                except Exception:
                    pass

            # 尝试去除小噪点（开操作）
            try:
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
                arr = cv2.morphologyEx(arr, cv2.MORPH_OPEN, kernel)
            except Exception:
                pass

            # 尝试检测并校正倾斜角度（deskew）
            try:
                coords = np.column_stack(np.where(arr < 255))
                if coords.size > 0:
                    angle = cv2.minAreaRect(coords)[-1]
                    if angle < -45:
                        angle = -(90 + angle)
                    else:
                        angle = -angle
                    if abs(angle) > 0.1:
                        (h, w) = arr.shape
                        center = (w // 2, h // 2)
                        M = cv2.getRotationMatrix2D(center, angle, 1.0)
                        arr = cv2.warpAffine(arr, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
                        print(f"  校正倾斜角度: {angle:.2f}°")
            except Exception:
                pass

            # 返回 PIL 图像
            try:
                img = Image.fromarray(arr)
            except Exception:
                pass

        return img
    except Exception as e:
        print(f"⚠️ 图片预处理失败: {e}")
        try:
            return Image.open(image_path).convert('L')
        except:
            return None


def image_to_text(image_path):
    """从图片提取文本"""
    global OCR_READER
    # 懒初始化 OCR_READER（如果 easyocr 可用且尚未初始化）
    if OCR_READER is None:
        if easyocr is not None:
            try:
                OCR_READER = easyocr.Reader(['ch_sim', 'en'], gpu=False)
                print("✅ EasyOCR 懒初始化成功")
            except Exception as e:
                print(f"⚠️ EasyOCR 懒初始化失败: {e}")
                OCR_READER = None
        else:
            return "", "OCR未初始化"

    try:
        print(f"🔍 开始OCR识别: {image_path}")

        # 预处理图片
        img = preprocess_image(image_path)
        if img is None:
            return "", "图片处理失败"

        # 转换为numpy数组
        img_array = np.array(img)
        print(f"  NumPy数组形状: {img_array.shape}")

        # OCR识别：优先使用 detail=1 获取每段置信度以过滤低置信结果
        print("🤖 调用EasyOCR识别 (detail=1)...")
        try:
            raw = OCR_READER.readtext(img_array, detail=1, paragraph=False)
        except Exception as e:
            print(f"⚠️ EasyOCR 细节识别失败: {e}, 回退 paragraph=true")
            raw = OCR_READER.readtext(img_array, detail=0, paragraph=True)

        # raw (detail=1) 格式：[(bbox, text, confidence), ...]
        text_segments = []
        if isinstance(raw, list) and raw and isinstance(raw[0], (list, tuple)) and len(raw[0]) >= 2:
            # 过滤低置信的单元
            for item in raw:
                try:
                    if len(item) == 3:
                        bbox, txt, conf = item
                    else:
                        # 当返回格式不同，尝试展开
                        txt = str(item[1]) if len(item) > 1 else str(item[0])
                        conf = 1.0
                    # 过滤非常短或置信度极低的识别
                    if txt and str(txt).strip():
                        if isinstance(conf, (int, float)):
                            if conf >= 0.30:  # 阈值可调
                                text_segments.append(str(txt).strip())
                        else:
                            text_segments.append(str(txt).strip())
                except Exception:
                    continue

        # 如果没有获得分段结果，回退为 paragraph=true 的输出
        if not text_segments:
            try:
                print("🤖 回退到 paragraph=True 识别")
                para = OCR_READER.readtext(img_array, detail=0, paragraph=True)
                if isinstance(para, list):
                    text_segments = [str(t).strip() for t in para if str(t).strip()]
                else:
                    text_segments = [str(para).strip()]
            except Exception as e:
                print(f"❌ 回退识别也失败: {e}")
                text_segments = []

        # 合并文本
        text = ' '.join(text_segments)

        # 清理文本
        text = clean_ocr_text(text)

        print(f"✅ OCR识别完成，文字长度: {len(text)}")
        if text:
            print(f"📝 OCR结果预览: {text[:100]}...")

        return text, ""

    except Exception as e:
        print(f"❌ OCR识别失败: {e}")
        return "", str(e)


def clean_ocr_text(text):
    """清理OCR文本，过滤标点符号和乱码"""
    if not text:
        return ""

    # 替换常见错别字
    corrections = {
        '肖待基': '肖特基',
        'BATS4C': 'BAT54C',
        '半': '',  # 移除"半"字符
        'SOT- 23': 'SOT-23',
        '500-323': 'SOD-323',
        '85819WS': 'BAV99WS',
        'RTO8OSBRDO7IOKL': 'RT0805BRD0710KL',
        'RTO8O5BRD071OORL': 'RT0805BRD07100RL'
    }

    for wrong, correct in corrections.items():
        text = text.replace(wrong, correct)

    # 移除控制字符和非常规字符
    import re
    # 只保留字母、数字、中文、空格、常见符号
    allowed_pattern = re.compile(r'[^\w\s\u4e00-\u9fff.,/\-:%Ω℃°μAVFHKMWHz#@$*()\[\]{}_=+<>]')
    cleaned_text = allowed_pattern.sub('', text)

    # 修复常见格式问题
    cleaned_text = re.sub(r'(\d+)\s*([KMG]?Ω)', r'\1\2', cleaned_text)
    cleaned_text = re.sub(r'(\d+)\s*([munp]?F)', r'\1\2', cleaned_text)
    cleaned_text = re.sub(r'(\d+)\s*([mun]?H)', r'\1\2', cleaned_text)

    # 移除多余空格
    cleaned_text = ' '.join(cleaned_text.split())

    # 进一步删除看起来像乱码的长串（例如过多的非字母数字的连续字符）
    cleaned_text = re.sub(r'[\W_]{5,}', ' ', cleaned_text)

    return cleaned_text


# ==================== AI解析功能函数 ====================
def enhanced_ai_parse_component(prompt):
    """
    增强的AI解析函数，处理JSON格式问题
    """
    # 获取配置
    api_key = Config.API_KEY
    api_url = Config.API_URL
    ai_model = Config.AI_MODEL

    if not api_key or not api_url:
        return [], "❌ 请先配置API密钥和接口地址"

    try:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        # 系统提示词
        system_prompt = """你是一名经验丰富的电子元器件识别专家，擅长从OCR文本中精确提取电子元器件信息。

### 识别规则：
1. **品类名/型号**：元器件的具体型号（如：4.7K、1N4148、STM32H750）
2. **类别**：元器件的分类（如：电阻、电容、二极管、三极管、IC芯片等）
3. **封装**：元器件的物理封装（如：0603、SOT-23、SOP-8、QFP-100）
4. **规格**：元器件的技术参数（如：4.7KΩ ±1%、100V 150mA、32-bit ARM Cortex-M7）
5. **数量**：物料数量，必须是整数（如：100、50、1）
6. **单位**：计数单位（如：个、只、片等，通常是"个"）
7. **单价**：单个元器件价格，必须是数字（如：0.0413、5.5、120.0），如无法识别则留空
8. **供应商**：经销商名称（如：VISHAY、DIODES、立创商城等），**如无法从OCR文本识别出来，则留空**
9. **购买渠道**：购买来源（如：立创商城、淘宝、官方代理商等），**如无法从OCR文本识别出来，则留空**
10. **备注**：其他信息备注

### 输出格式：
返回一个JSON数组，每个元素是包含10个字段的数组：
[
  ["型号1", "类别1", "封装1", "规格1", 数量1, "单位1", 单价1, "供应商1", "渠道1", "备注1"],
  ["型号2", "类别2", "封装2", "规格2", 数量2, "单位2", 单价2, "供应商2", "渠道2", "备注2"]
]

### 重要规则：
- **数量必须是整数**，不能有小数
- **单价必须是浮点数**，可以有小数位
- **未识别的字段必须设为空字符串""，不要使用"未知"或其他默认值**
- 只提取JSON数组，不要有其他文字
- 如果无法识别某个信息，设置为空字符串而不是添加默认值"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ]

        payload = {
            "model": ai_model,
            "messages": messages,
            "temperature": Config.AI_TEMPERATURE,
            "max_tokens": Config.AI_MAX_TOKENS,
        }

        print(f"🔄 发送AI请求到: {api_url}/chat/completions")

        # 调用API
        response = requests.post(
            url=f"{api_url.rstrip('/')}/chat/completions",
            headers=headers,
            json=payload,
            timeout=Config.AI_TIMEOUT
        )

        print(f"✅ AI响应状态码: {response.status_code}")

        # 调试：打印响应前500字符
        print(f"📝 响应内容预览: {response.text[:500]}")

        if response.status_code not in [200, 201]:
            error_msg = f"API请求失败: {response.status_code} - {response.text[:200]}"
            print(f"❌ {error_msg}")
            return [], error_msg

        # 尝试解析JSON
        try:
            result = response.json()
        except json.JSONDecodeError as e:
            print(f"❌ JSON解析错误: {e}")
            print(f"📝 原始响应内容: {response.text[:500]}")

            # 尝试清理响应，提取可能的JSON部分
            cleaned_text = response.text.strip()

            # 移除可能的HTML标签
            import re
            cleaned_text = re.sub(r'<[^>]+>', '', cleaned_text)

            # 查找JSON开始和结束
            json_start = cleaned_text.find('{')
            json_end = cleaned_text.rfind('}')

            if json_start != -1 and json_end != -1:
                try:
                    json_str = cleaned_text[json_start:json_end + 1]
                    result = json.loads(json_str)
                    print("✅ 从清理后的文本中提取JSON成功")
                except:
                    return [], "❌ API返回非JSON格式，请检查API服务"
            else:
                return [], "❌ API返回非JSON格式，请检查API服务"

        # 处理可能的错误响应格式
        if 'error' in result:
            error_msg = f"API错误: {result['error']}"
            if isinstance(result['error'], dict) and 'message' in result['error']:
                error_msg = f"API错误: {result['error']['message']}"
            print(f"❌ {error_msg}")
            return [], error_msg

        # 检查是否包含choices字段
        if 'choices' not in result or len(result['choices']) == 0:
            return [], "❌ API返回格式不正确，缺少choices字段"

        ai_response = result['choices'][0]['message']['content'].strip()
        print(f"📄 AI返回内容前200字符: {ai_response[:200]}...")

        # 清理和提取JSON部分
        cleaned_response = clean_ai_response(ai_response)
        print(f"🧹 清理后响应前200字符: {cleaned_response[:200]}...")

        try:
            # 尝试解析JSON
            data = json.loads(cleaned_response)

            # 验证和标准化组件数据
            validated_components = []
            if isinstance(data, list):
                for comp in data:
                    if isinstance(comp, list) and len(comp) >= 4:
                        # 确保有10个字段
                        while len(comp) < 10:
                            comp.append("")

                        # 标准化处理
                        standardized_comp = standardize_component(comp)
                        validated_components.append(standardized_comp)

            if not validated_components:
                return [], "AI没有提取到有效的元器件信息"

            print(f"✅ 成功提取 {len(validated_components)} 个元器件")
            return validated_components, f"成功提取 {len(validated_components)} 个元器件"

        except json.JSONDecodeError as e:
            print(f"❌ JSON解析错误: {str(e)}")
            print(f"📝 尝试提取的文本: {cleaned_response[:200]}")

            # 尝试从AI响应中提取表格数据
            try:
                import re
                # 查找类似数组的结构
                array_pattern = r'\[\s*\[[^\]]+\](?:,\s*\[[^\]]+\])*\s*\]'
                match = re.search(array_pattern, ai_response, re.DOTALL)
                if match:
                    json_str = match.group(0)
                    # 清理JSON字符串
                    json_str = json_str.replace("'", '"')
                    data = json.loads(json_str)

                    validated_components = []
                    for comp in data:
                        if isinstance(comp, list) and len(comp) >= 4:
                            while len(comp) < 10:
                                comp.append("")
                            validated_components.append(standardize_component(comp))

                    if validated_components:
                        print(f"✅ 通过正则提取 {len(validated_components)} 个元器件")
                        return validated_components, f"成功提取 {len(validated_components)} 个元器件"
            except Exception as regex_error:
                print(f"❌ 正则提取失败: {regex_error}")

            return [], f"AI返回的不是有效JSON: {str(e)}"
        except Exception as e:
            print(f"❌ 处理AI响应时出错: {str(e)}")
            return [], f"处理AI响应时出错: {str(e)}"

    except requests.exceptions.Timeout:
        return [], "API请求超时，请检查网络或稍后重试"
    except requests.exceptions.ConnectionError:
        return [], "网络连接错误，请检查API地址和网络"
    except Exception as e:
        print(f"❌ AI解析失败: {str(e)}")
        return [], f"AI解析失败: {str(e)}"


def clean_ai_response(response):
    """
    清理AI响应，提取JSON部分
    """
    response = str(response).strip()

    # 移除可能的markdown代码块
    if response.startswith('```json'):
        response = response[7:]
    elif response.startswith('```'):
        response = response[3:]

    if response.endswith('```'):
        response = response[:-3]

    # 移除首尾空白和换行
    response = response.strip()

    # 查找JSON开始和结束位置
    json_start = response.find('[')
    json_end = response.rfind(']')

    if json_start != -1 and json_end != -1 and json_end > json_start:
        response = response[json_start:json_end + 1]

    # 修复常见的JSON格式问题
    response = response.replace("'", '"')  # 单引号转双引号
    response = response.replace("，", ",")  # 中文逗号转英文逗号
    response = response.replace("：", ":")  # 中文冒号转英文冒号

    # 移除控制字符
    response = ''.join(char for char in response if ord(char) >= 32 or char in '\n\r\t')

    return response


def standardize_component(comp):
    """
    标准化元器件数据，纠正错别字，并确保返回字段类型：
    [model(str), category(str), package(str), spec(str), quantity(int), unit(str), price(float|''), supplier(str), channel(str), remark(str)]
    """
    #（保留原有前置处理逻辑）
    # 以下为替换实现（类型严格化）
    if len(comp) < 10:
        while len(comp) < 10:
            comp.append("")

    # 型号和类别、封装、规格 字符串化、修正
    model = str(comp[0]).strip()
    model = model.replace('BATS4C', 'BAT54C').replace('肖待基', '肖特基').replace('？', '').replace('?', '')
    special_chars = '_*/\\|"\'`~!@#$%^&()[]{}<>'
    for char in special_chars:
        model = model.replace(char, '')

    category = str(comp[1]).strip()
    category = category.replace('肖待基二极管', '肖特基二极管').replace('肖待基', '肖特基')
    package = str(comp[2]).strip().upper()
    spec = str(comp[3]).strip()

    # 数量：尽可能转为 int，失败时默认 1
    try:
        quantity_str = str(comp[4]).replace('个', '').replace('PCS', '').strip()
        import re
        quantity_num = re.sub(r'[^\d\-]', '', quantity_str)
        if quantity_num == '' or quantity_num == '-':
            quantity = 1
        else:
            # 允许负号但最终以正数为准
            quantity = int(abs(int(quantity_num)))
            if quantity <= 0:
                quantity = 1
    except Exception:
        quantity = 1

    # 单位：强制字符串，空则 '个'
    unit = str(comp[5]).strip() or '个'

    # 单价：提取数字并返回 float（保留两位），否则返回空字符串
    price_raw = str(comp[6]).strip()
    price_num = ""
    try:
        price_candidate = price_raw.replace('¥', '').replace('￥', '').replace(',', '').strip()
        # 移除非数字但保留小数点和负号
        import re
        m = re.search(r'(-?\d+\.?\d*)', price_candidate)
        if m:
            price_num = round(float(m.group(1)), 2)
        else:
            price_num = ""
    except Exception:
        price_num = ""

    # 供应商/渠道/备注 强制字符串
    supplier = str(comp[7]).strip() if comp[7] is not None else ""
    channel = str(comp[8]).strip() if comp[8] is not None else ""
    remark = str(comp[9]).strip() if comp[9] is not None else ""

    # 如果 category 为空，基于型号猜测（保留之前逻辑）
    if not category and model:
        mu = model.upper()
        if any(prefix in mu for prefix in ['1N', 'BAT', 'BAV', 'LL', 'MMSZ']):
            category = '二极管'
        elif any(prefix in mu for prefix in ['2N', 'BC', 'S8', 'MMBT']):
            category = '三极管'
        elif any(prefix in mu for prefix in ['IRF', 'SI', 'AO', 'FDN']):
            category = 'MOS管'
        elif any(prefix in mu for prefix in ['STM', 'AT', 'PIC', 'L', 'LM', 'MAX']):
            category = 'IC芯片'
        elif 'K' in mu or 'Ω' in mu or '电阻' in category:
            category = '电阻'
        elif 'U' in mu or 'F' in mu or '电容' in category:
            category = '电容'
        elif 'H' in mu or '电感' in category:
            category = '电感'
        elif 'MHZ' in mu or '晶振' in category:
            category = '晶振'
        else:
            category = '其他'

    # 返回类型规范化（price 保持 float 或 空字符串）
    return [model, category, package, spec, int(quantity), str(unit), (price_num if price_num != "" else ""), str(supplier), str(channel), str(remark)]


def infer_supplier_from_model(model):
    """根据型号推测供应商"""
    model_upper = model.upper()

    supplier_mapping = {
        # ST微电子
        'STM': 'ST',
        'STM32': 'ST',

        # TI德州仪器
        'LM': 'TI',
        'TL': 'TI',
        'SN': 'TI',
        'TMS': 'TI',

        # ON安森美
        'NCP': 'ON',
        'NCV': 'ON',
        'FAN': 'ON',

        # Infineon英飞凌
        'IRF': 'Infineon',
        'BTS': 'Infineon',

        # NXP恩智浦
        'LPC': 'NXP',
        'S': 'NXP',

        # Microchip
        'PIC': 'Microchip',
        'AT': 'Microchip',

        # Diodes公司
        'BAT': 'DIODES',
        'BAV': 'DIODES',
        'LL': 'DIODES',

        # Vishay威世
        '1N': 'VISHAY',
        'MMSZ': 'VISHAY',

        # 国巨
        'RC': '国巨',
        'RK': '国巨',
        'RT': '国巨',

        # 三星
        'CL': '三星',

        # 村田
        'GRM': '村田',
        'LQG': '村田',

        # TDK
        'C': 'TDK',
        'B': 'TDK',
    }

    for prefix, supplier in supplier_mapping.items():
        if model_upper.startswith(prefix):
            return supplier

    return "未知"


def standardize_channel(channel):
    """标准化购买渠道"""
    if not channel:
        return "立创商城"  # 默认渠道

    channel_lower = channel.lower()

    if any(x in channel_lower for x in ['立创', 'lcsc', 'jlc']):
        return "立创商城"
    elif any(x in channel_lower for x in ['得捷', 'digi', 'digikey']):
        return "得捷"
    elif any(x in channel_lower for x in ['贸泽', 'mouser']):
        return "贸泽"
    elif any(x in channel_lower for x in ['淘宝', 'taobao', 'tb']):
        return "淘宝"
    elif any(x in channel_lower for x in ['京东', 'jd']):
        return "京东"
    elif any(x in channel_lower for x in ['华强', 'hq']):
        return "华强北"
    else:
        return channel


# ==================== 表格功能函数 ====================
def generate_empty_table():
    """生成空的库存表"""
    try:
        columns = [
            '品类名', '类别', '封装', '规格',
            '数量', '单位', '单价(元)', '供应商',
            '购买渠道', '备注', '入库时间'
        ]

        df = pd.DataFrame(columns=columns)
        df.to_excel('元器件库存表.xlsx', index=False)
        print("✅ 已创建新的元器件库存表")
        return '元器件库存表.xlsx', "✅ 已创建新的元器件库存表"
    except Exception as e:
        print(f"❌ 创建库存表失败: {e}")
        return None, f"❌ 创建库存表失败: {str(e)}"


def get_table_info():
    """获取表格信息"""
    try:
        if os.path.exists('元器件库存表.xlsx'):
            df = pd.read_excel('元器件库存表.xlsx')
            return {
                '文件存在': True,
                '文件路径': os.path.abspath('元器件库存表.xlsx'),
                '数据条数': len(df),
                '列名': df.columns.tolist()
            }
        else:
            return {
                '文件存在': False,
                '文件路径': '元器件库存表.xlsx',
                '数据条数': 0,
                '列名': []
            }
    except Exception as e:
        print(f"❌ 获取表格信息失败: {e}")
        return {
            '文件存在': False,
            '文件路径': '元器件库存表.xlsx',
            '数据条数': 0,
            '列名': []
        }


def append_data_to_table(components):
    """将数据追加到表格，保存前强制列类型，单价转为两位小数的数值类型，文件保存至docs目录"""
    try:
        if not components:
            return None, "没有数据可保存"

        # 列名定义（保持原有顺序）
        columns = [
            '品类名', '类别', '封装', '规格',
            '数量', '单位', '单价(元)', '供应商',
            '购买渠道', '备注'
        ]

        # 创建DataFrame，保证每一列类型尽量规范
        df_new = pd.DataFrame(components, columns=columns)

        # 强制数量为整数
        if '数量' in df_new.columns:
            df_new['数量'] = pd.to_numeric(df_new['数量'], errors='coerce').fillna(1).astype(int)

        # 强制单位为字符串
        if '单位' in df_new.columns:
            df_new['单位'] = df_new['单位'].astype(str).fillna('个')

        # 强制单价为浮点数，两位小数；无法解析则设为 NaN
        if '单价(元)' in df_new.columns:
            df_new['单价(元)'] = pd.to_numeric(df_new['单价(元)'], errors='coerce')
            # 保留两位小数（数值类型）
            df_new['单价(元)'] = df_new['单价(元)'].round(2)

        # 供应商与渠道强制为字符串（避免excel数字识别）
        for col in ['供应商', '购买渠道', '备注', '品类名', '类别', '封装', '规格']:
            if col in df_new.columns:
                df_new[col] = df_new[col].astype(str).fillna('')

        # 添加入库时间
        df_new['入库时间'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')

        # 创建docs目录并读取或创建表格
        os.makedirs('docs', exist_ok=True)
        file_path = os.path.join('docs', '元器件库存表.xlsx')
        
        if os.path.exists(file_path):
            df_existing = pd.read_excel(file_path)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        # 保存到docs目录下的Excel
        df_combined.to_excel(file_path, index=False)

        print(f"✅ 成功保存 {len(df_new)} 条数据到表格")
        return file_path, f"成功保存 {len(df_new)} 条数据"

    except Exception as e:
        print(f"❌ 保存数据失败: {e}")
        return None, f"保存数据失败: {str(e)}"


# ==================== Flask路由 ====================
@app.route('/')
def index():
    """主页面"""
    return render_template_string(HTML_TEMPLATE)


@app.route('/save_api_config', methods=['POST'])
def save_api_config():
    """保存API配置到session和.env文件"""
    try:
        data = request.get_json()
        api_key = data.get('api_key', '').strip()
        api_url = data.get('api_url', '').strip()
        ai_model = data.get('ai_model', 'gpt-3.5-turbo').strip()

        if not api_key or not api_url:
            return jsonify({'success': False, 'message': '❌ API密钥和接口地址不能为空'})

        # 保存到.env文件
        if Config.save_env(api_key, api_url, ai_model):
            save_message = '✅ API配置已保存到本地配置文件'
        else:
            save_message = '✅ API配置已保存（本地保存失败）'

        # 测试连接
        try:
            response = requests.post(
                url=f"{api_url.rstrip('/')}/chat/completions",
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {api_key}"
                },
                json={
                    "model": ai_model,
                    "messages": [{"role": "user", "content": "hello"}],
                    "max_tokens": 5
                },
                timeout=15
            )
            if response.status_code in [200, 201]:
                return jsonify({'success': True, 'message': f'{save_message}，连接测试成功！'})
            else:
                return jsonify({'success': False, 'message': f'{save_message}，但连接测试失败：{response.status_code}'})
        except Exception as e:
            return jsonify({'success': False, 'message': f'{save_message}，但连接测试失败：{str(e)[:100]}'})

    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ 保存失败：{str(e)[:100]}'})


@app.route('/get_env_config')
def get_env_config():
    """获取环境变量配置"""
    try:
        return jsonify({
            'success': True,
            'config': {
                'api_key': Config.API_KEY[:8] + "..." if Config.API_KEY else "",
                'api_url': Config.API_URL,
                'ai_model': Config.AI_MODEL
            }
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ 读取配置失败：{str(e)}'})


@app.route('/upload_images', methods=['POST'])
def upload_images():
    """上传图片并OCR识别"""
    try:
        # 检查是否有base64数据（安全解析 JSON）
        try:
            data = request.get_json(silent=True)
        except Exception:
            data = None
        base64_images = data.get('base64_images', []) if data else []
        file_uploads = request.files.getlist('images')

        ocr_results = []
        temp_files = []

        # 处理base64图片
        for idx, base64_img in enumerate(base64_images):
            filename = f"paste_{idx}_{int(time.time())}.jpg"
            filepath = save_base64_image(base64_img, filename)
            if filepath:
                temp_files.append(filepath)

                ocr_text, error = image_to_text(filepath)
                if not error and ocr_text:
                    # 清理OCR文本
                    cleaned_text = clean_ocr_text(ocr_text)
                    ocr_results.append({
                        'filename': f"粘贴图片_{idx + 1}.jpg",
                        'text': cleaned_text[:500] + '...' if len(cleaned_text) > 500 else cleaned_text,
                        'full_text': cleaned_text
                    })

        # 处理文件上传
        for idx, file in enumerate(file_uploads):
            if file and allowed_file(file.filename):
                filename = secure_filename(f"{idx}_{int(time.time())}_{file.filename}")
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                temp_files.append(filepath)

                ocr_text, error = image_to_text(filepath)
                if not error and ocr_text:
                    # 清理OCR文本
                    cleaned_text = clean_ocr_text(ocr_text)
                    ocr_results.append({
                        'filename': file.filename,
                        'text': cleaned_text[:500] + '...' if len(cleaned_text) > 500 else cleaned_text,
                        'full_text': cleaned_text
                    })

        # 清理临时文件
        for filepath in temp_files:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass

        if not ocr_results:
            return jsonify({'success': False, 'message': '❌ 图片识别失败，请检查图片质量'})

        # 合并所有OCR文本
        full_text = '\n'.join([r['full_text'] for r in ocr_results])
        session['ocr_full_text'] = full_text

        return jsonify({
            'success': True,
            'message': f'✅ 成功识别 {len(ocr_results)} 张图片',
            'ocr_results': ocr_results
        })

    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ 处理失败：{str(e)}'})


@app.route('/ai_parse', methods=['POST'])
def api_ai_parse():
    """AI解析"""
    try:
        ocr_text = session.get('ocr_full_text', '')
        user_prompt = request.json.get('user_prompt', '') if request.json else ''

        if not ocr_text:
            return jsonify({'success': False, 'message': '❌ 没有OCR结果'})

        # 构建更严格的提示词，要求准确识别或留空
        full_prompt = f"""请从以下OCR文字中精确提取电子元器件信息，返回JSON数组格式：

OCR文字：
{ocr_text[:2000]}

重要规则：
1. 返回JSON数组，每个元素是包含10个字段的数组
2. 字段顺序：[品类名, 类别, 封装, 规格, 数量, 单位, 单价, 供应商, 购买渠道, 备注]
3. 数量必须是整数（如：100）
4. 单价必须是浮点数或整数（如：0.05, 10.5）
5. 未识别到的字段必须设为空字符串""，不要填写"未知"、"立创商城"等默认值
6. 供应商只有在OCR文字中明确出现时才填写，否则留空
7. 购买渠道只有在OCR文字中明确出现时才填写，否则留空
8. 只返回JSON数组，不要有其他文字

示例输出格式：
[
  ["4.7K", "电阻", "0603", "4.7KΩ ±1%", 100, "只", 0.05, "", "", ""],
  ["STM32F103C8T6", "MCU", "LQFP-48", "ARM Cortex-M3", 10, "个", 15.80, "ST", "", ""]
]
请提取所有元器件信息："""

        # 合并用户额外提示（如果存在）
        if user_prompt:
            full_prompt += "\n\n用户额外说明:\n" + user_prompt

        print("🤖 开始调用AI解析函数...")
        print(f"📝 输入提示长度: {len(full_prompt)}")

        components, msg = enhanced_ai_parse_component(full_prompt)

        if not components:
            return jsonify({'success': False, 'message': msg})

        # 保存解析结果到 session，供后续保存使用
        session['parsed_components'] = components

        return jsonify({'success': True, 'message': f'✅ 成功解析出 {len(components)} 个元器件', 'components': components})

    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ AI解析失败：{str(e)}'})


def append_data_to_table(components, export_dir=None, export_name=None, mode='append'):
    """将数据追加到表格。

    参数：
      - components: 元器件列表
      - export_dir: 导出目录（相对 project 根目录 或 绝对路径），默认 'docs' 子目录
      - export_name: 导出文件名（可带或不带 .xlsx）
      - mode: 'append'（追加到默认/指定文件），'new'（每次创建新文件），'overwrite'（覆盖目标文件）
    返回: (file_path, message)
    """
    try:
        if not components:
            return None, "没有数据可保存"

        # 基本路径
        base_dir = os.path.abspath(os.path.dirname(__file__))
        default_exports = os.path.join(base_dir, 'docs')

        # 计算导出目录
        if export_dir:
            # 如果传入相对路径，基于 base_dir
            if os.path.isabs(export_dir):
                export_dir_full = export_dir
            else:
                export_dir_full = os.path.join(base_dir, export_dir)
        else:
            export_dir_full = default_exports

        os.makedirs(export_dir_full, exist_ok=True)

        # 创建DataFrame
        columns = [
            '品类名', '类别', '封装', '规格',
            '数量', '单位', '单价(元)', '供应商',
            '购买渠道', '备注'
        ]
        df_new = pd.DataFrame(components, columns=columns)

        # 强制类型：数量 整数，单位 字符串，单价 浮点或 NaN，供应商/渠道/备注 字符串
        if '数量' in df_new.columns:
            df_new['数量'] = pd.to_numeric(df_new['数量'], errors='coerce').fillna(1).astype(int)
        if '单位' in df_new.columns:
            df_new['单位'] = df_new['单位'].astype(str).fillna('个')
        if '单价(元)' in df_new.columns:
            df_new['单价(元)'] = pd.to_numeric(df_new['单价(元)'], errors='coerce')
            df_new['单价(元)'] = df_new['单价(元)'].round(2)
        for col in ['供应商', '购买渠道', '备注', '品类名', '类别', '封装', '规格']:
            if col in df_new.columns:
                df_new[col] = df_new[col].astype(str).fillna('')

        # 添加入库时间
        df_new['入库时间'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')

        # 决定目标文件路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if export_name:
            if not export_name.lower().endswith('.xlsx'):
                export_name = export_name + '.xlsx'
            target_file = os.path.join(export_dir_full, export_name)
        else:
            if mode == 'new':
                target_file = os.path.join(export_dir_full, f"元器件库存表_{timestamp}.xlsx")
            else:
                target_file = os.path.join(export_dir_full, '元器件库存表.xlsx')

        # 读取现有表格并合并（除非 mode == 'new' 或 mode == 'overwrite'）
        if mode == 'new':
            df_combined = df_new
        else:
            if os.path.exists(target_file) and mode == 'append':
                try:
                    df_existing = pd.read_excel(target_file)
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                except Exception:
                    df_combined = df_new
            elif os.path.exists(target_file) and mode == 'overwrite':
                df_combined = df_new
            else:
                # 文件不存在，直接使用新数据
                df_combined = df_new

        # 保存到 Excel
        df_combined.to_excel(target_file, index=False)

        print(f"✅ 成功保存 {len(df_new)} 条数据到表格: {target_file}")
        return target_file, f"成功保存 {len(df_new)} 条数据"

    except Exception as e:
        print(f"❌ 保存数据失败: {e}")
        return None, f"保存数据失败: {str(e)}"


@app.route('/save_to_table', methods=['POST'])
def save_to_table():
    """保存到表格，支持指定导出目录/文件名与模式（append/new/overwrite）"""
    try:
        components = session.get('parsed_components', [])

        if not components:
            return jsonify({'success': False, 'message': '❌ 没有数据可保存'})

        data = request.get_json(silent=True) or {}
        export_dir = data.get('export_dir')
        export_name = data.get('export_name')
        mode = data.get('mode', 'append')

        file_path, msg = append_data_to_table(components, export_dir=export_dir, export_name=export_name, mode=mode)

        if not file_path:
            return jsonify({'success': False, 'message': msg})

        # 清空session
        session.pop('ocr_full_text', None)
        session.pop('parsed_components', None)

        return jsonify({
            'success': True,
            'message': f'✅ 成功导入 {len(components)} 个元器件',
            'file_path': file_path
        })

    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ 保存失败：{str(e)}'})


@app.route('/get_table_info')
def api_get_table_info():
    """获取表格信息"""
    info = get_table_info()
    return jsonify({'success': True, 'data': info})


@app.route('/download_table')
def download_table():
    """下载表格"""
    if os.path.exists('元器件库存表.xlsx'):
        return send_file(
            '元器件库存表.xlsx',
            as_attachment=True,
            download_name=f"元器件库存表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
    else:
        return jsonify({'success': False, 'message': '❌ 表格文件不存在'}), 404


@app.route('/test_api', methods=['POST'])
def test_api():
    """测试API连接"""
    try:
        data = request.get_json()
        api_key = data.get('api_key', '').strip()
        api_url = data.get('api_url', '').strip()
        ai_model = data.get('ai_model', 'gpt-3.5-turbo').strip()

        if not api_key or not api_url:
            return jsonify({'success': False, 'message': '❌ API密钥和接口地址不能为空'})

        # 测试连接
        response = requests.post(
            url=f"{api_url.rstrip('/')}/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            json={
                "model": ai_model,
                "messages": [{"role": "user", "content": "hello"}],
                "max_tokens": 5
            },
            timeout=10
        )

        if response.status_code in [200, 201]:
            return jsonify({'success': True, 'message': '✅ 连接测试成功！'})
        else:
            return jsonify({'success': False, 'message': f'❌ 连接测试失败：{response.status_code}'})

    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ 连接测试失败：{str(e)[:100]}'})

# 重点：methods 是复数，值是 ['POST']（大写），不能漏、不能错
@app.route('/parse_paste_data', methods=['POST'])
def parse_paste_data():
    # 函数内容（用之前给你的容错版代码）
    try:
        raw_data = request.get_data(as_text=True)
        print(f"【粘贴解析】收到原始请求体：{raw_data[:200]}...")
        
        try:
            data = request.get_json(force=True)
        except Exception as e:
            print(f"【粘贴解析】JSON解析失败：{e}")
            return jsonify({'success': False, 'message': '❌ 请求格式错误，请确保是JSON格式'})
        
        paste_text = data.get('paste_text', '').strip()
        if not paste_text:
            return jsonify({'success': False, 'message': '❌ 请输入数据'})

        # 后续解析逻辑（和之前的容错版一致）
        clean_text = paste_text.replace('\n', '').replace(' ', '').rstrip(',').strip()
        if clean_text.startswith('[') and clean_text.endswith(']') and not clean_text.startswith('[['):
            clean_text = f"[{clean_text}]"

        try:
            parsed = json.loads(clean_text)
        except json.JSONDecodeError as e:
            return jsonify({'success': False, 'message': f'❌ JSON格式错误：{str(e)[:50]}'})

        if not isinstance(parsed, list):
            return jsonify({'success': False, 'message': '❌ 数据必须是数组格式'})

        components = [item for item in parsed if isinstance(item, list) and len(item) >= 4]
        if not components:
            return jsonify({'success': False, 'message': '❌ 未提取到有效元器件数据'})

        session['parsed_components'] = components
        return jsonify({'success': True, 'message': f'✅ 成功解析 {len(components)} 个元器件', 'components': components})

    except Exception as e:
        print(f"【粘贴解析】未知错误：{e}", exc_info=True)
        return jsonify({'success': False, 'message': f'❌ 解析失败：{str(e)[:80]}'})

# ==================== 浏览器自动打开函数 ====================
def open_browser():
    """自动打开浏览器"""
    time.sleep(2.0)

    url = Config.TARGET_URL
    print(f"🌐 尝试自动打开浏览器: {url}")

    # 检查操作系统
    system_name = platform.system()

    if system_name == "Windows":
        print("🪟 检测到Windows系统...")

        # Windows系统常见浏览器路径
        browser_paths = [
            # Chrome
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",

            # Microsoft Edge
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",

            # Firefox
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        ]

        # 尝试每个浏览器路径
        for browser_path in browser_paths:
            if os.path.exists(browser_path):
                try:
                    print(f"✅ 找到浏览器: {browser_path}")
                    subprocess.Popen([browser_path, url])
                    print(f"✅ 已使用C盘浏览器打开: {url}")
                    return True
                except Exception as e:
                    print(f"⚠️ 使用路径打开失败 {browser_path}: {e}")

    # 通用方法：使用webbrowser模块
    try:
        result = webbrowser.open(url, new=0, autoraise=True)
        if result:
            print(f"✅ 已使用系统默认浏览器打开: {url}")
            return True
    except Exception as e:
        print(f"⚠️ webbrowser.open失败: {e}")

    print(f"❌ 无法自动打开浏览器，请手动访问: {url}")
    return False


# ==================== HTML模板 ====================
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>元器件AI识别导入工具</title>
    <link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.bootcdn.net/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        :root {
            --primary: #4e73df;
            --success: #1cc88a;
            --info: #36b9cc;
            --warning: #f6c23e;
            --danger: #e74a3b;
        }

        body {
            background: #f8f9fc;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        /* API配置面板 */
        .api-config-panel {
            background: white;
            border-radius: 10px;
            box-shadow: 0 0.15rem 1.75rem 0 rgba(58, 59, 69, 0.15);
            padding: 20px;
            margin-bottom: 25px;
            border-left: 5px solid var(--primary);
        }

        .api-status {
            display: inline-flex;
            align-items: center;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: 600;
        }
        .api-status.success {
            background: #d4edda;
            color: #155724;
        }
        .api-status.warning {
            background: #fff3cd;
            color: #856404;
        }
        .api-status.danger {
            background: #f8d7da;
            color: #721c24;
        }

        .step-indicator {
            display: flex;
            justify-content: space-between;
            margin-bottom: 30px;
            position: relative;
        }

        .step-indicator::before {
            content: '';
            position: absolute;
            top: 20px;
            left: 10%;
            right: 10%;
            height: 2px;
            background: #e3e6f0;
            z-index: 1;
        }

        .step-item {
            text-align: center;
            position: relative;
            z-index: 2;
            flex: 1;
        }

        .step-item.active .step-number {
            background: var(--primary);
            color: white;
            border-color: var(--primary);
        }

        .step-item.completed .step-number {
            background: var(--success);
            color: white;
            border-color: var(--success);
        }

        .step-number {
            width: 40px;
            height: 40px;
            line-height: 36px;
            border-radius: 50%;
            background: white;
            border: 2px solid #e3e6f0;
            color: #b7b9cc;
            font-weight: bold;
            margin: 0 auto 8px;
            transition: all 0.3s;
        }

        .step-label {
            color: #6e707e;
            font-size: 14px;
            font-weight: 600;
        }

        .step-item.active .step-label {
            color: var(--primary);
        }

        .step-content {
            background: white;
            border-radius: 10px;
            box-shadow: 0 0.15rem 1.75rem 0 rgba(58, 59, 69, 0.15);
            padding: 25px;
            margin-bottom: 25px;
            display: none;
        }

        .step-content.active {
            display: block;
            animation: fadeIn 0.5s;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .card-header {
            background: #f8f9fc;
            border-bottom: 1px solid #e3e6f0;
            padding: 18px 25px;
            font-weight: 700;
            color: #4e73df;
        }

        .btn-primary {
            background: var(--primary);
            border-color: var(--primary);
            padding: 8px 25px;
            font-weight: 600;
        }

        .btn-success {
            background: var(--success);
            border-color: var(--success);
            padding: 8px 25px;
            font-weight: 600;
        }

        .form-control, .form-select {
            padding: 10px 12px;
            border: 1px solid #d1d3e2;
            border-radius: 6px;
            font-size: 14px;
        }

        .alert-flash {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            min-width: 300px;
            max-width: 400px;
            box-shadow: 0 0.15rem 1.75rem 0 rgba(58, 59, 69, 0.15);
            animation: slideInRight 0.3s, fadeOut 0.5s 2.7s forwards;
        }

        @keyframes slideInRight {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }

        @keyframes fadeOut {
            to { opacity: 0; }
        }

        .upload-area {
            border: 2px dashed var(--primary);
            border-radius: 10px;
            padding: 30px 20px;
            text-align: center;
            background: #f8f9fe;
            cursor: pointer;
            transition: all 0.3s;
            margin-bottom: 15px;
            position: relative;
        }

        .upload-area:hover {
            background: #eef1ff;
            border-color: #2e59d9;
        }

        .upload-area.paste-active {
            border-color: var(--success);
            background: #e8f7f1;
        }

        .image-preview {
            max-height: 80px;
            object-fit: contain;
            border-radius: 4px;
            margin: 3px;
            border: 1px solid #e3e6f0;
            background: white;
        }

        .ocr-preview {
            max-height: 200px;
            overflow-y: auto;
            background: #f8f9fc;
            border: 1px solid #e3e6f0;
            border-radius: 5px;
            padding: 12px;
            font-family: 'Consolas', monospace;
            font-size: 12px;
            line-height: 1.4;
            white-space: pre-wrap;
            word-wrap: break-word;
        }

        .components-table {
            width: 100%;
            font-size: 13px;
            border-collapse: collapse;
        }

        .components-table th {
            background: #f8f9fc;
            padding: 10px;
            font-weight: 600;
            border-bottom: 2px solid #e3e6f0;
            text-align: center;
            white-space: nowrap;
        }

        .components-table td {
            padding: 8px 10px;
            border-bottom: 1px solid #e3e6f0;
            text-align: center;
            vertical-align: middle;
        }

        .components-table tr:hover {
            background: #f8f9fc;
        }

        .components-table .text-muted {
            color: #858796 !important;
        }

        .action-buttons {
            display: flex;
            justify-content: space-between;
            margin-top: 25px;
            padding-top: 15px;
            border-top: 1px solid #e3e6f0;
        }

        .loading-spinner {
            display: none;
            text-align: center;
            padding: 30px;
        }

        .table-responsive {
            max-height: 400px;
            overflow-y: auto;
            border: 1px solid #e3e6f0;
            border-radius: 5px;
        }

        .badge-count {
            font-size: 0.8em;
            padding: 3px 8px;
            border-radius: 10px;
        }

        .upload-tips {
            font-size: 12px;
            color: #6c757d;
            margin-top: 10px;
        }

        .upload-tips ul {
            text-align: left;
            display: inline-block;
            margin: 10px auto;
        }

        .paste-indicator {
            position: absolute;
            top: 10px;
            right: 10px;
            background: var(--success);
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 12px;
            display: none;
        }

        .upload-area.paste-active .paste-indicator {
            display: block;
        }

        .password-toggle {
            position: relative;
        }

        .password-toggle .toggle-icon {
            position: absolute;
            right: 10px;
            top: 50%;
            transform: translateY(-50%);
            cursor: pointer;
            color: #6c757d;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- API配置面板 -->
        <div class="api-config-panel">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h5 class="mb-0"><i class="fas fa-cog me-2"></i>API配置</h5>
                <span id="apiStatus" class="api-status warning">未配置</span>
            </div>

            <div class="row g-3">
                <div class="col-md-4 password-toggle">
                    <label class="form-label">API密钥 <span class="text-danger">*</span></label>
                    <input type="password" class="form-control" id="api_key" 
                           placeholder="sk-xxxxxxxxxxxxxxxx" value="">
                    <span class="toggle-icon" id="toggle-password">
                        <i class="fas fa-eye"></i>
                    </span>
                </div>
                <div class="col-md-4">
                    <label class="form-label">API接口地址 <span class="text-danger">*</span></label>
                    <input type="text" class="form-control" id="api_url" 
                           placeholder="https://api.example.com/v1" 
                           value="https://api.openai.com/v1">
                </div>
                <div class="col-md-3">
                    <label class="form-label">AI模型</label>
                    <select class="form-select" id="ai_model">
                        <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
                        <option value="gpt-4">GPT-4</option>
                        <option value="gpt-4-turbo">GPT-4 Turbo</option>
                        <option value="deepseek-chat" selected>deepseek Chat</option>
                        <option value="glm-4">GLM-4</option>
                        <option value="qwen-turbo">Qwen Turbo</option>
                    </select>
                </div>
                <div class="col-md-1 d-flex align-items-end">
                    <button class="btn btn-primary w-100" id="test-api-btn">
                        <i class="fas fa-plug me-2"></i>测试
                    </button>
                </div>
            </div>

            <div class="mt-3">
                <small class="text-muted">
                    <i class="fas fa-info-circle me-1"></i>
                    支持OpenAI、DeepSeek、智谱AI、讯飞星火等兼容OpenAI API的模型
                </small>
            </div>
        </div>

        <!-- 主标题 -->
        <h4 class="text-primary mb-4">
            <i class="fas fa-microchip me-2"></i>元器件AI识别导入工具
        </h4>

        <!-- 模式选择 -->
        <div class="card mb-4 border-info">
            <div class="card-body">
                <h6 class="card-title mb-3">选择导入模式</h6>
                <div class="btn-group w-100" role="group">
                    <input type="radio" class="btn-check" name="mode-selector" id="mode-ocr" value="ocr" checked>
                    <label class="btn btn-outline-primary" for="mode-ocr">
                        <i class="fas fa-image me-2"></i>识图模式（OCR识别）
                    </label>
                    
                    <input type="radio" class="btn-check" name="mode-selector" id="mode-paste" value="paste">
                    <label class="btn btn-outline-primary" for="mode-paste">
                        <i class="fas fa-paste me-2"></i>粘贴模式（手动输入）
                    </label>
                </div>
            </div>
        </div>

        <!-- 步骤指示器 -->
        <div class="step-indicator">
            <div class="step-item active" id="step1-indicator">
                <div class="step-number">1</div>
                <div class="step-label">上传图片</div>
            </div>
            <div class="step-item" id="step2-indicator">
                <div class="step-number">2</div>
                <div class="step-label">OCR识别</div>
            </div>
            <div class="step-item" id="step3-indicator">
                <div class="step-number">3</div>
                <div class="step-label">AI解析</div>
            </div>
            <div class="step-item" id="step4-indicator">
                <div class="step-number">4</div>
                <div class="step-label">导入完成</div>
            </div>
        </div>

        <!-- Flash消息容器 -->
        <div id="flash-container"></div>

        <!-- 步骤1: 上传图片 -->
        <div class="step-content active" id="step1-content">
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-images me-2"></i>步骤1: 上传元器件图片
                </div>
                <div class="card-body">
                    <p class="text-muted mb-3">支持多种方式上传图片</p>

                    <div class="upload-area" id="upload-area">
                        <div class="paste-indicator" id="paste-indicator">
                            <i class="fas fa-paste me-1"></i>按Ctrl+V粘贴
                        </div>
                        <i class="fas fa-cloud-upload-alt fa-2x text-primary mb-3"></i>
                        <h5>拖放图片、复制粘贴或点击选择</h5>
                        <p class="text-muted mb-3">支持 JPG, PNG, GIF, BMP 格式</p>
                        <button class="btn btn-outline-primary" id="select-images-btn">
                            <i class="fas fa-folder-open me-2"></i>选择图片文件
                        </button>
                        <input type="file" id="image-upload" class="d-none" multiple accept="image/*">

                        <div class="upload-tips">
                            <p>支持以下方式：</p>
                            <ul>
                                <li>点击"选择图片文件"按钮</li>
                                <li>拖拽图片到此处（支持多张）</li>
                                <li>按Ctrl+V粘贴剪贴板图片</li>
                                <li>从文件管理器拖拽文件</li>
                            </ul>
                        </div>
                    </div>

                    <div id="image-preview-container" class="mb-3" style="display:none;">
                        <h6>已选择 <span id="image-count" class="badge bg-primary badge-count">0</span> 张图片</h6>
                        <div id="image-previews" class="d-flex flex-wrap"></div>
                    </div>

                    <div class="mb-3">
                        <label class="form-label">额外需求说明（可选）</label>
                        <textarea class="form-control" id="user_prompt" rows="2" 
                                  placeholder="例如：主要提取二极管和电阻，忽略螺丝等机械件..."></textarea>
                    </div>

                    <div class="loading-spinner" id="upload-loading">
                        <div class="spinner-border text-primary" role="status"></div>
                        <p class="mt-2">正在处理图片...</p>
                    </div>

                    <div class="action-buttons">
                        <button class="btn btn-secondary" id="step1-prev-btn" disabled>
                            <i class="fas fa-arrow-left me-2"></i>上一步
                        </button>
                        <button class="btn btn-primary" id="step1-next-btn" disabled>
                            开始识别 <i class="fas fa-robot ms-2"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- 步骤1-粘贴模式 -->
        <div class="step-content" id="step1-paste-content" style="display:none;">
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-paste me-2"></i>步骤1: 输入元器件数据
                </div>
                <div class="card-body">
                    <p class="text-muted mb-3">粘贴或输入格式化的元器件数据</p>

                    <div class="mb-3">
                        <label class="form-label">数据格式说明</label>
                        <div class="alert alert-info small">
                            <p class="mb-2"><strong>支持两种格式：</strong></p>
                            <p class="mb-2">1. JSON数组格式（一条数据）：<br>
                            <code>["品类名", "类别", "封装", "规格", 数量, "单位", 单价, "供应商", "渠道", "备注"]</code></p>
                            <p class="mb-0">2. 多条数据（每行一个JSON数组）</p>
                        </div>
                    </div>

                    <div class="mb-3">
                        <label class="form-label">粘贴数据 <span class="text-danger">*</span></label>
                        <textarea class="form-control" id="paste-textarea" rows="10" 
                                  placeholder='["AO3401", "场效应管(MOSFET)", "SOT-23", "P沟道 30V 4.2A", 20, "个", "", "Hottech", "", ""]
["SS8050", "三极管(BJT)", "SOT-23", "NPN 电流1.5A", 50, "个", 0.0777, "UMW", "嘉立创", ""]
["1N4148", "开关二极管", "SOD-323", "", 100, "只", 0.05, "", "", ""]'></textarea>
                    </div>

                    <div class="alert alert-warning small">
                        <i class="fas fa-lightbulb me-2"></i>
                        <strong>提示：</strong>数量必须是整数，单价可以是数字或留空。供应商、渠道、备注如果没有可留空字符串""。
                    </div>

                    <div class="loading-spinner" id="parse-loading">
                        <div class="spinner-border text-primary" role="status"></div>
                        <p class="mt-2">正在解析数据...</p>
                    </div>

                    <div class="action-buttons">
                        <button class="btn btn-secondary" id="step1-paste-prev-btn" disabled>
                            <i class="fas fa-arrow-left me-2"></i>上一步
                        </button>
                        <button class="btn btn-primary" id="step1-paste-next-btn" disabled>
                            开始解析 <i class="fas fa-check ms-2"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- 步骤2: OCR识别 -->
        <div class="step-content" id="step2-content">
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-eye me-2"></i>步骤2: OCR识别结果
                </div>
                <div class="card-body">
                    <div class="mb-4">
                        <h6 class="mb-2">
                            <i class="fas fa-file-alt me-2"></i>OCR识别结果
                        </h6>
                        <div class="ocr-preview" id="ocr-text-display">等待识别...</div>
                    </div>

                    <div class="action-buttons">
                        <button class="btn btn-secondary" id="step2-prev-btn">
                            <i class="fas fa-arrow-left me-2"></i>上一步
                        </button>
                        <button class="btn btn-primary" id="step2-next-btn">
                            开始AI解析 <i class="fas fa-robot ms-2"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- 步骤3: AI解析 -->
        <div class="step-content" id="step3-content">
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-robot me-2"></i>步骤3: AI解析结果
                </div>
                <div class="card-body">
                    <div id="ai-results-section" style="display:none;">
                        <h6 class="mb-3">
                            <i class="fas fa-microchip me-2"></i>解析结果
                            <span id="component-count" class="badge bg-success badge-count ms-2">0</span>
                        </h6>

                        <div class="table-responsive">
                            <table class="components-table">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>品类名</th>
                                        <th>类别</th>
                                        <th>封装</th>
                                        <th>规格</th>
                                        <th>数量</th>
                                        <th>单价(元)</th>
                                        <th>供应商</th>
                                        <th>购买渠道</th>
                                        <th>备注</th>
                                    </tr>
                                </thead>
                                <tbody id="components-body">
                                    <!-- 数据动态插入 -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <div class="loading-spinner" id="ai-loading">
                        <div class="spinner-border text-primary" role="status"></div>
                        <p class="mt-2">AI正在解析...</p>
                    </div>

                    <div class="action-buttons">
                        <button class="btn btn-secondary" id="step3-prev-btn">
                            <i class="fas fa-arrow-left me-2"></i>上一步
                        </button>
                        <button class="btn btn-success" id="step3-next-btn" disabled>
                            确认导入 <i class="fas fa-save ms-2"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- 步骤4: 完成导入 -->
        <div class="step-content" id="step4-content">
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-check-circle me-2"></i>步骤4: 导入完成
                </div>
                <div class="card-body text-center">
                    <div class="mb-4">
                        <i class="fas fa-check-circle text-success" style="font-size: 70px;"></i>
                    </div>

                    <h4 class="text-success mb-3">导入成功！</h4>

                    <div class="alert alert-success mb-4" id="import-summary">
                        <!-- 摘要动态显示 -->
                    </div>

                    <div class="row mb-4">
                        <div class="col-md-6 mb-3">
                            <div class="card border-primary h-100">
                                <div class="card-body">
                                    <h6 class="card-title">库存表信息</h6>
                                    <div id="table-info" class="text-start">
                                        <!-- 表格信息动态显示 -->
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="card border-info h-100">
                                <div class="card-body">
                                    <h6 class="card-title">下一步操作</h6>
                                    <p class="card-text">您可以继续导入或下载库存表</p>
                                    <div class="mt-3">
                                        <button class="btn btn-outline-primary w-100 mb-2" id="new-import-btn">
                                            <i class="fas fa-redo me-2"></i>开始新的导入
                                        </button>
                                        <button class="btn btn-primary w-100" id="download-table-btn">
                                            <i class="fas fa-download me-2"></i>下载库存表
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.bootcdn.net/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
    <script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>

    <script>
        $(document).ready(function() {
            let currentStep = 1;
            let selectedImages = [];
            let selectedBase64Images = [];
            let ocrResults = [];
            let aiComponents = [];
            let apiConfigured = false;
            let currentMode = 'ocr';  // 'ocr' 或 'paste'

            // 页面加载时检查API配置
            checkApiConfig();

            // 模式切换处理
            $('input[name="mode-selector"]').on('change', function() {
                currentMode = $(this).val();
                if (currentMode === 'ocr') {
                    $('#step1-content').show();
                    $('#step1-paste-content').hide();
                    resetOCRMode();
                } else {
                    $('#step1-content').hide();
                    $('#step1-paste-content').show();
                    resetPasteMode();
                }
            });

            // 重置OCR模式
            function resetOCRMode() {
                selectedImages = [];
                selectedBase64Images = [];
                $('#image-preview-container').hide();
                $('#image-previews').empty();
                $('#image-upload').val('');
                $('#user_prompt').val('');
                $('#step1-next-btn').prop('disabled', true);
            }

            // 重置粘贴模式
            function resetPasteMode() {
                $('#paste-textarea').val('');
                $('#step1-paste-next-btn').prop('disabled', true);
            }

            // 显示提示消息
            function showFlash(message, type = 'info') {
                const icons = {
                    'success': 'fa-check-circle',
                    'error': 'fa-exclamation-circle',
                    'warning': 'fa-exclamation-triangle',
                    'info': 'fa-info-circle'
                };

                const alertClass = {
                    'success': 'alert-success',
                    'error': 'alert-danger',
                    'warning': 'alert-warning',
                    'info': 'alert-info'
                }[type];

                const flash = `
                    <div class="alert alert-flash ${alertClass} alert-dismissible fade show">
                        <i class="fas ${icons[type] || 'fa-info-circle'} me-2"></i>
                        ${message}
                        <button class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                `;

                $('#flash-container').html(flash);

                setTimeout(() => {
                    $('.alert-flash').alert('close');
                }, 3000);
            }

            // 检查API配置
            function checkApiConfig() {
                $.get('/get_env_config', function(response) {
                    if (response.success) {
                        const config = response.config;

                        if (config.api_key && config.api_key.length > 10) {
                            // 有配置，更新状态
                            updateApiStatus(true, '已配置');
                            apiConfigured = true;
                        } else {
                            // 无配置或配置不完整
                            updateApiStatus(false, '未配置');
                            apiConfigured = false;

                            // 如果有部分配置，填充到表单
                            if (config.api_url) {
                                $('#api_url').val(config.api_url);
                            }
                            if (config.ai_model) {
                                $('#ai_model').val(config.ai_model);
                            }
                        }
                    }
                }).fail(function() {
                    updateApiStatus(false, '检查失败');
                    apiConfigured = false;
                });
            }

            // 更新API状态显示
            function updateApiStatus(isConfigured, message) {
                const apiStatus = $('#apiStatus');
                if (isConfigured) {
                    apiStatus.removeClass('warning danger').addClass('success');
                    apiStatus.html('<i class="fas fa-check-circle me-1"></i>' + message);
                } else {
                    apiStatus.removeClass('success').addClass('warning');
                    apiStatus.html('<i class="fas fa-exclamation-triangle me-1"></i>' + message);
                }
            }

            // 切换步骤
            function goToStep(step) {
                $('.step-content').removeClass('active');
                $('.step-item').removeClass('active completed');

                $(`#step${step}-content`).addClass('active');

                for (let i = 1; i <= 4; i++) {
                    const indicator = $(`#step${i}-indicator`);
                    if (i < step) {
                        indicator.addClass('completed');
                    } else if (i === step) {
                        indicator.addClass('active');
                    }
                }

                currentStep = step;
            }

            // 密码显示/隐藏切换
            $('#toggle-password').click(function() {
                const passwordInput = $('#api_key');
                const toggleIcon = $(this).find('i');

                if (passwordInput.attr('type') === 'password') {
                    passwordInput.attr('type', 'text');
                    toggleIcon.removeClass('fa-eye').addClass('fa-eye-slash');
                } else {
                    passwordInput.attr('type', 'password');
                    toggleIcon.removeClass('fa-eye-slash').addClass('fa-eye');
                }
            });

            // 测试API连接
            $('#test-api-btn').click(function() {
                const apiKey = $('#api_key').val().trim();
                const apiUrl = $('#api_url').val().trim();
                const aiModel = $('#ai_model').val();

                if (!apiKey || !apiUrl) {
                    showFlash('❌ API密钥和接口地址不能为空', 'error');
                    return;
                }

                // 显示测试中状态
                updateApiStatus(false, '测试中...');

                $.ajax({
                    url: '/test_api',
                    method: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({
                        api_key: apiKey,
                        api_url: apiUrl,
                        ai_model: aiModel
                    }),
                    success: function(response) {
                        if (response.success) {
                            // 测试成功，保存配置
                            $.ajax({
                                url: '/save_api_config',
                                method: 'POST',
                                contentType: 'application/json',
                                data: JSON.stringify({
                                    api_key: apiKey,
                                    api_url: apiUrl,
                                    ai_model: aiModel
                                }),
                                success: function(saveResponse) {
                                    if (saveResponse.success) {
                                        updateApiStatus(true, '配置成功');
                                        apiConfigured = true;
                                        showFlash('✅ API配置保存成功！', 'success');
                                    } else {
                                        updateApiStatus(false, '保存失败');
                                        showFlash(saveResponse.message, 'error');
                                    }
                                },
                                error: function() {
                                    updateApiStatus(false, '保存失败');
                                    showFlash('❌ 保存配置失败', 'error');
                                }
                            });
                        } else {
                            updateApiStatus(false, '连接失败');
                            showFlash(response.message, 'error');
                        }
                    },
                    error: function() {
                        updateApiStatus(false, '请求失败');
                        showFlash('❌ 网络请求失败', 'error');
                    }
                });
            });

            // 步骤1: 上传图片
            $('#select-images-btn, #upload-area').click(() => $('#image-upload').click());

            $('#image-upload').change(function(e) {
                handleFileSelection(Array.from(e.target.files));
            });

            // 拖放上传
            $('#upload-area')
                .on('dragover', e => { 
                    e.preventDefault(); 
                    e.stopPropagation();
                    $(e.currentTarget).addClass('border-primary bg-light'); 
                })
                .on('dragleave', e => { 
                    e.preventDefault(); 
                    e.stopPropagation();
                    $(e.currentTarget).removeClass('border-primary bg-light'); 
                })
                .on('drop', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    $(this).removeClass('border-primary bg-light');

                    const files = Array.from(e.originalEvent.dataTransfer.files);
                    handleFileSelection(files);
                });

            // 粘贴图片功能
            let pasteTimeout = null;
            $(document).on('paste', function(e) {
                const items = e.originalEvent.clipboardData.items;
                const pasteArea = $('#upload-area');

                if (!items) return;

                let hasImages = false;
                const imageFiles = [];

                for (let item of items) {
                    if (item.type.indexOf('image') !== -1) {
                        hasImages = true;
                        const blob = item.getAsFile();
                        imageFiles.push(blob);
                    }
                }

                if (hasImages && currentStep === 1) {
                    // 高亮显示粘贴区域
                    pasteArea.addClass('paste-active');
                    clearTimeout(pasteTimeout);
                    pasteTimeout = setTimeout(() => {
                        pasteArea.removeClass('paste-active');
                    }, 2000);

                    // 处理粘贴的图片
                    processPastedImages(imageFiles);
                }
            });

            function processPastedImages(imageBlobs) {
                if (imageBlobs.length === 0) return;

                // 将Blob转换为Base64
                const base64Promises = imageBlobs.map(blob => {
                    return new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onload = function(e) {
                            resolve(e.target.result);
                        };
                        reader.readAsDataURL(blob);
                    });
                });

                Promise.all(base64Promises).then(base64Images => {
                    // 添加到预览
                    base64Images.forEach((base64, idx) => {
                        const timestamp = new Date().getTime();
                        const filename = `粘贴图片_${timestamp}_${idx}.png`;

                        selectedBase64Images.push(base64);

                        // 添加到预览
                        $('#image-previews').append(`
                            <div style="width: 80px; margin: 2px;" class="position-relative">
                                <img src="${base64}" class="image-preview w-100">
                                <span class="badge bg-success position-absolute top-0 end-0" style="font-size: 0.6em">粘贴</span>
                            </div>
                        `);
                    });

                    // 更新计数
                    const totalCount = selectedImages.length + selectedBase64Images.length;
                    $('#image-count').text(totalCount);
                    $('#image-preview-container').show();
                    $('#step1-next-btn').prop('disabled', false);

                    showFlash(`✅ 已粘贴 ${imageBlobs.length} 张图片`, 'success');
                });
            }

            function handleFileSelection(files) {
                if (files.length === 0) return;

                // 过滤图片文件
                const imageFiles = files.filter(file => 
                    file.type.startsWith('image/') && 
                    ['image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/webp'].includes(file.type.toLowerCase())
                );

                if (imageFiles.length === 0) {
                    showFlash('❌ 请选择图片文件', 'error');
                    return;
                }

                // 添加到已选文件
                selectedImages = selectedImages.concat(imageFiles);

                // 显示预览
                handleImageSelection();
            }

            function handleImageSelection() {
                const totalCount = selectedImages.length + selectedBase64Images.length;

                if (totalCount > 0) {
                    $('#image-preview-container').show();
                    $('#image-count').text(totalCount);

                    // 只显示前12张预览
                    $('#image-previews').empty();

                    // 显示文件图片预览
                    selectedImages.slice(0, 8).forEach((file, idx) => {
                        const reader = new FileReader();
                        reader.onload = e => {
                            $('#image-previews').append(`
                                <div style="width: 80px; margin: 2px;" class="position-relative">
                                    <img src="${e.target.result}" class="image-preview w-100">
                                    <span class="badge bg-primary position-absolute top-0 end-0" style="font-size: 0.6em">${idx+1}</span>
                                </div>
                            `);
                        };
                        reader.readAsDataURL(file);
                    });

                    // 显示粘贴图片预览（如果有）
                    selectedBase64Images.slice(0, 4).forEach((base64, idx) => {
                        $('#image-previews').append(`
                            <div style="width: 80px; margin: 2px;" class="position-relative">
                                <img src="${base64}" class="image-preview w-100">
                                <span class="badge bg-success position-absolute top-0 end-0" style="font-size: 0.6em">粘贴</span>
                            </div>
                        `);
                    });

                    if (totalCount > 12) {
                        $('#image-previews').append(`
                            <div style="width: 80px; margin: 2px;" class="position-relative d-flex align-items-center justify-content-center">
                                <span class="badge bg-info">+${totalCount - 12}张</span>
                            </div>
                        `);
                    }

                    $('#step1-next-btn').prop('disabled', false);
                    showFlash(`✅ 已选择 ${totalCount} 张图片`, 'success');
                }
            }

            $('#step1-next-btn').click(function() {
                if (!apiConfigured) {
                    showFlash('❌ 请先配置API', 'error');
                    return;
                }

                const totalCount = selectedImages.length + selectedBase64Images.length;
                if (totalCount === 0) {
                    showFlash('❌ 请先选择图片', 'error');
                    return;
                }

                $('#upload-loading').show();
                $('#step1-next-btn').prop('disabled', true);

                const formData = new FormData();

                // 添加文件
                selectedImages.forEach(file => formData.append('images', file));

                // 准备发送base64图片
                const uploadData = {
                    base64_images: selectedBase64Images
                };

                // 如果有base64图片，使用JSON方式发送
                if (selectedBase64Images.length > 0) {
                    $.ajax({
                        url: '/upload_images',
                        method: 'POST',
                        contentType: 'application/json',
                        data: JSON.stringify(uploadData),
                        success: handleUploadResponse,
                        error: handleUploadError
                    });
                } else {
                    // 只有文件时使用FormData
                    $.ajax({
                        url: '/upload_images',
                        method: 'POST',
                        data: formData,
                        contentType: false,
                        processData: false,
                        success: handleUploadResponse,
                        error: handleUploadError
                    });
                }
            });

            // 粘贴模式处理
            $('#paste-textarea').on('input', function() {
                const hasContent = $(this).val().trim().length > 0;
                $('#step1-paste-next-btn').prop('disabled', !hasContent);
            });

            $('#step1-paste-next-btn').click(function() {
                const pasteText = $('#paste-textarea').val().trim();
                if (!pasteText) {
                    showFlash('❌ 请输入数据', 'error');
                    return;
                }

                $('#parse-loading').show();
                $('#step1-paste-next-btn').prop('disabled', true);

                $.ajax({
                    url: '/parse_paste_data',
                    method: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({ paste_text: pasteText }),
                    success: function(response) {
                        $('#parse-loading').hide();

                        if (response.success) {
                            aiComponents = response.components || [];
                            showFlash(response.message, 'success');
                            
                            // 直接跳到步骤3（AI解析结果），因为粘贴模式已经是结构化数据
                            goToStep(3);
                            setTimeout(() => {
                                displayComponents(aiComponents);
                                $('#component-count').text(aiComponents.length);
                                $('#ai-results-section').show();
                                $('#step3-next-btn').prop('disabled', false);
                            }, 300);
                        } else {
                            showFlash(response.message, 'error');
                            $('#step1-paste-next-btn').prop('disabled', false);
                        }
                    },
                    error: function() {
                        $('#parse-loading').hide();
                        showFlash('❌ 解析失败，请检查数据格式', 'error');
                        $('#step1-paste-next-btn').prop('disabled', false);
                    }
                });
            });

            $('#step1-paste-prev-btn').click(() => {
                currentMode = 'ocr';
                $('input[name="mode-selector"]').val(['ocr']);
                $('#mode-ocr').prop('checked', true);
                $('#step1-content').show();
                $('#step1-paste-content').hide();
            });

            function handleUploadResponse(response) {
                $('#upload-loading').hide();

                if (response.success) {
                    ocrResults = response.ocr_results || [];
                    showFlash(response.message, 'success');

                    if (ocrResults.length > 0) {
                        let ocrText = '';
                        ocrResults.forEach((r, i) => {
                            ocrText += `【图片${i+1}】${r.text}\\n\\n`;
                        });
                        $('#ocr-text-display').text(ocrText);
                    }

                    goToStep(2);
                } else {
                    showFlash(response.message, 'error');
                    $('#step1-next-btn').prop('disabled', false);
                }
            }

            function handleUploadError() {
                $('#upload-loading').hide();
                showFlash('❌ 上传失败，请检查网络连接', 'error');
                $('#step1-next-btn').prop('disabled', false);
            }

            // 步骤2: OCR结果
            $('#step2-prev-btn').click(() => goToStep(1));
            $('#step2-next-btn').click(function() {
                goToStep(3);
                setTimeout(startAIparse, 500);
            });

            // 步骤3: AI解析
            function startAIparse() {
                $('#ai-loading').show();
                const userPrompt = $('#user_prompt').val();

                $.ajax({
                    url: '/ai_parse',
                    method: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({ user_prompt: userPrompt }),
                    success: function(response) {
                        $('#ai-loading').hide();

                        if (response.success) {
                            aiComponents = response.components || [];
                            showFlash(response.message, 'success');

                            if (aiComponents.length > 0) {
                                displayComponents(aiComponents);
                                $('#component-count').text(aiComponents.length);
                                $('#ai-results-section').show();
                                $('#step3-next-btn').prop('disabled', false);
                            } else {
                                showFlash('⚠️ AI解析完成但没有提取到元器件信息', 'warning');
                            }
                        } else {
                            showFlash(response.message, 'error');
                        }
                    },
                    error: function(xhr, status, error) {
                        $('#ai-loading').hide();
                        showFlash('❌ AI解析请求失败: ' + error, 'error');
                        console.error('AI解析错误:', error);
                    }
                });
            }

            function displayComponents(components) {
                const tbody = $('#components-body');
                tbody.empty();

                components.forEach((row, idx) => {
                    const tr = $('<tr></tr>');
                    tr.append(`<td class="text-muted">${idx + 1}</td>`);

                    row.forEach((cell, cellIdx) => {
                        let display = cell || '';
                        if (cellIdx === 4) display = display || 0;  // 数量
                        if (cellIdx === 6 && cell) {  // 单价
                            display = parseFloat(cell).toFixed(2);
                        }
                        if (!display && cellIdx !== 4 && cellIdx !== 6) {
                            display = '<span class="text-muted">-</span>';
                        }
                        tr.append(`<td>${display}</td>`);
                    });

                    tbody.append(tr);
                });
            }

            $('#step3-prev-btn').click(() => goToStep(2));

            $('#step3-next-btn').click(function() {
                if (aiComponents.length === 0) {
                    showFlash('❌ 没有数据可导入', 'error');
                    return;
                }

                $.ajax({
                    url: '/save_to_table',
                    method: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({}),
                    success: function(response) {
                        if (response.success) {
                            showFlash(response.message, 'success');
                            goToStep(4);

                            $('#import-summary').html(`
                                <h6 class="alert-heading">导入摘要</h6>
                                <p>成功导入 <strong>${aiComponents.length}</strong> 个元器件</p>
                                <p class="mb-0">文件: <code>${response.file_path || '元器件库存表.xlsx'}</code></p>
                            `);

                            updateTableInfo();
                        } else {
                            showFlash(response.message, 'error');
                        }
                    },
                    error: function() {
                        showFlash('❌ 保存失败', 'error');
                    }
                });
            });

            function updateTableInfo() {
                $.get('/get_table_info', function(response) {
                    if (response.success) {
                        const info = response.data;
                        $('#table-info').html(`
                            <p><strong>文件路径:</strong><br><small>${info.文件路径}</small></p>
                            <p><strong>状态:</strong> ${info.文件存在 ? '✅ 存在' : '❌ 不存在'}</p>
                            <p><strong>总数据条数:</strong> ${info.数据条数}</p>
                        `);
                    }
                });
            }

            // 步骤4: 完成
            $('#new-import-btn').click(function() {
                // 重置所有状态
                selectedImages = [];
                selectedBase64Images = [];
                ocrResults = [];
                aiComponents = [];

                $('#image-preview-container').hide();
                $('#image-previews').empty();
                $('#image-upload').val('');
                $('#user_prompt').val('');
                $('#ocr-text-display').text('等待识别...');
                $('#ai-results-section').hide();
                $('#components-body').empty();
                $('#step1-next-btn').prop('disabled', true);
                $('#step3-next-btn').prop('disabled', true);

                goToStep(1);
                showFlash('🔄 开始新的导入', 'info');
            });

            $('#download-table-btn').click(() => window.open('/download_table', '_blank'));

            // 初始化
            updateTableInfo();

            // 页面加载后检查API配置并填充表单
            setTimeout(() => {
                $.get('/get_env_config', function(response) {
                    if (response.success) {
                        const config = response.config;
                        if (config.api_key && config.api_key.length > 10) {
                            $('#api_key').val(config.api_key);
                        }
                        if (config.api_url) {
                            $('#api_url').val(config.api_url);
                        }
                        if (config.ai_model) {
                            $('#ai_model').val(config.ai_model);
                        }
                    }
                });
            }, 1000);
        });
    </script>
</body>
</html>
'''

# ==================== 主程序入口 ====================
if __name__ == '__main__':
    print("=" * 60)
    print("元器件AI识别导入工具 - 完整单文件版")
    print("=" * 60)
    print("访问地址: http://127.0.0.1:5000")
    print("网络地址: http://192.168.3.51:5000")
    print("按 Ctrl+C 停止程序")
    print("-" * 60)

    # 显示当前配置状态
    if Config.SECRET_KEY:
        masked_key = Config.SECRET_KEY[:8] + "..." + Config.SECRET_KEY[-4:] if len(Config.SECRET_KEY) > 12 else "***"
        print(f"✅ API密钥: {masked_key}")
    else:
        print("⚠️ API密钥: 未配置")

    print(f"🌐 API地址: {Config.API_URL}")
    print(f"🤖 AI模型: {Config.AI_MODEL}")

    # 检查是否有库存表，如果没有则创建
    if not os.path.exists('元器件库存表.xlsx'):
        generate_empty_table()
        print("📊 已创建新的元器件库存表")

    print("-" * 60)
    print("功能说明：")
    print("1. 支持拖拽多张图片上传")
    print("2. 支持Ctrl+V粘贴图片")
    print("3. 支持选择文件上传")
    print("4. 自动保存API配置到.env文件")
    print("5. 支持多种AI模型")
    print("-" * 60)
    print("🔧 启动中...")

    # 在新线程中打开浏览器
    try:
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()
        print("🌐 已启动浏览器自动打开线程...")
    except Exception as e:
        print(f"⚠️ 启动浏览器线程失败: {e}")
        print(f"💡 请手动访问: http://127.0.0.1:5000")

    # 运行Flask应用
    try:
        app.run(host=Config.HOST, port=Config.PORT, debug=Config.DEBUG, use_reloader=False)
    except KeyboardInterrupt:
        print("\n👋 程序已停止")
    except Exception as e:
        print(f"❌ Flask启动失败: {e}")