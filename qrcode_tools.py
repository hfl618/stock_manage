# -*- coding: utf-8 -*-
"""
二维码扫码管理（独立模块）
提供：独立扫码页面片段 HTML_TEMPLATE 与更健壮的扫码解析处理函数 scan_qrcode_handler()
避免在模块顶层导入 stock_manage，路由内按需导入 Component 以防循环依赖。
"""
from flask import jsonify, request, render_template_string
import json
import re
import urllib.parse

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>扫码管理</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { padding: 12px; font-family: Arial, sans-serif; }
        #videoElement { width: 100%; max-height: 360px; background: #000; border-radius:6px; }
        #resultBox { min-height: 80px; }
    </style>
</head>
<body>
    <div class="d-flex justify-content-between align-items-center mb-2">
        <h5 class="mb-0">扫码管理</h5>
        <div>
            <button class="btn btn-sm btn-outline-secondary" onclick="closePanel()">关闭</button>
        </div>
    </div>

    <div class="mb-2">
        <div class="d-flex justify-content-between mb-1">
            <div id="cameraStatus" class="alert alert-secondary p-2" style="flex:1;margin-right:8px;">摄像头状态：未初始化</div>
            <div>
                <button class="btn btn-sm btn-outline-primary me-1" onclick="window.open(location.href,'_blank')">在新窗口打开</button>
                <button class="btn btn-sm btn-outline-secondary" onclick="closePanel()">关闭</button>
            </div>
        </div>

        <div style="position:relative;">
            <video id="videoElement" playsinline muted style="width:100%; border-radius:6px; background:#000;"></video>
            <div id="scanOverlay" style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:60%;height:50%;border:3px dashed rgba(255,0,0,0.9);border-radius:8px;pointer-events:none;box-shadow:0 0 0 4px rgba(255,0,0,0.05);"></div>
        </div>

        <div class="mt-2 d-flex gap-2">
            <button id="startBtn" class="btn btn-primary btn-sm" onclick="startScanner()">开始扫描</button>
            <button id="stopBtn" class="btn btn-secondary btn-sm" onclick="stopScanner()" style="display:none;">停止扫描</button>
            <button id="refreshBtn" class="btn btn-outline-secondary btn-sm" onclick="refreshScanner()">刷新摄像头</button>
        </div>
    </div>

    <div class="mb-2">
        <label class="form-label">或上传二维码图片：</label>
        <div class="input-group">
            <input type="file" id="fileInput" accept="image/*" class="form-control form-control-sm">
            <button class="btn btn-primary btn-sm" onclick="scanFromFile()">上传并识别</button>
        </div>
    </div>

    <div id="resultBox" class="border rounded p-2">
        <div id="resultInner">等待扫码...</div>
    </div>

    <script src="https://unpkg.com/@zxing/library@latest"></script>
    <script>
        let codeReader = null;
        let videoStream = null;

        function setStatus(msg, cls='secondary'){
            const el = document.getElementById('cameraStatus');
            el.className = 'alert alert-' + cls + ' p-2';
            el.textContent = '摄像头状态：' + msg;
        }

        async function startScanner(){
            const startBtn = document.getElementById('startBtn');
            const stopBtn = document.getElementById('stopBtn');
            const video = document.getElementById('videoElement');
            try{
                setStatus('准备中...', 'info');
                if (videoStream){
                    videoStream.getTracks().forEach(t=>t.stop());
                    videoStream = null;
                }
                const constraints = { video: { facingMode: 'environment', width: { ideal: 640 }, height: { ideal: 480 } } };
                videoStream = await navigator.mediaDevices.getUserMedia(constraints);
                video.srcObject = videoStream;
                await video.play();

                codeReader = new ZXing.BrowserMultiFormatReader();
                setStatus('已启动', 'success');
                startBtn.style.display='none'; stopBtn.style.display='inline-block';

                codeReader.decodeFromVideoDevice(null, video, (result, err) => {
                    if (result) {
                        onDecode(result.text);
                    }
                    // ignore errors
                });

            }catch(e){
                console.error(e);
                setStatus(e.message || '启动失败', 'danger');
            }
        }

        function refreshScanner(){ stopScanner(); setTimeout(startScanner, 500); }

        function stopScanner(){
            const startBtn = document.getElementById('startBtn');
            const stopBtn = document.getElementById('stopBtn');
            if (codeReader){ try{ codeReader.reset(); }catch{} codeReader=null; }
            if (videoStream){ videoStream.getTracks().forEach(t=>t.stop()); videoStream=null; }
            const video = document.getElementById('videoElement'); try{ video.pause(); video.srcObject = null; }catch{}
            setStatus('已停止', 'secondary');
            startBtn.style.display='inline-block'; stopBtn.style.display='none';
        }

        async function scanFromFile(){
            const fi = document.getElementById('fileInput');
            if (!fi.files || fi.files.length===0){ alert('请选择图片文件'); return; }
            const file = fi.files[0];
            const img = new Image();
            const url = URL.createObjectURL(file);
            img.src = url;
            img.onload = async ()=>{
                try{
                    const reader = new ZXing.BrowserMultiFormatReader();
                    const res = await reader.decodeFromImage(img);
                    onDecode(res.text);
                }catch(err){
                    console.error(err);
                    showResult('识别失败：' + (err.message || err), false);
                }finally{ URL.revokeObjectURL(url); }
            };
            img.onerror = ()=>{ showResult('图片加载失败', false); URL.revokeObjectURL(url); };
        }

        async function onDecode(text){
            showResult('已识别，正在查询...', true);
            try{
                const resp = await fetch('/scan_qrcode', {
                    method:'POST', headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({data: text})
                });
                const j = await resp.json();
                if (j && j.code==1 && j.data){
                    const d = j.data;
                    // 显示详细信息
                    let html = `<div><strong>识别成功</strong></div>`;
                    html += `<div class="mt-2">ID: ${d.id || 'N/A'}</div>`;
                    html += `<div>品类: ${d.category || ''} &nbsp; 类别: ${d.type || ''}</div>`;
                    html += `<div>型号: ${d.model || ''} &nbsp; 封装: ${d.package || ''}</div>`;
                    html += `<div class="mt-2">数量: ${d.quantity || 0} ${d.unit || ''}</div>`;
                    html += `<div class="mt-2"><button class="btn btn-sm btn-primary" onclick="openEdit(${d.id})">打开详情</button> <button class="btn btn-sm btn-outline-secondary" onclick="continueScanning()">继续扫描</button></div>`;
                    showResult(html, true, true);
                    // 自动跳转到编辑页（延迟0.6s），若不希望自动跳转可注释下一行
                    setTimeout(()=>{ try{ window.parent.location.href = '/edit/' + d.id; }catch(e){ window.location.href = '/edit/' + d.id; } }, 600);
                    return;
                }else{
                    showResult('未找到对应元器件：' + (j && j.error? j.error : JSON.stringify(j)), false);
                }
            }catch(e){
                console.error(e);
                showResult('请求失败：' + e.message, false);
            }
        }

        function showResult(msg, ok, rawHtml){
            const box = document.getElementById('resultInner');
            if(rawHtml){ box.innerHTML = '<div class="alert alert-success">' + msg + '</div>'; return; }
            box.innerHTML = ok ? ('<div class="alert alert-success">' + msg + '</div>') : ('<div class="alert alert-danger">' + msg + '</div>');
        }

        function openEdit(id){
            try{ window.parent.location.href = '/edit/' + id; }catch(e){ window.location.href = '/edit/' + id; }
        }

        function closePanel(){
            try{ window.parent.hideQRCodeTools(); }catch(e){ window.close(); }
        }
    </script>
</body>
</html>
"""

def parse_qrcode_data(raw_data):
    """纯粹的二维码数据解析函数（无数据库操作），返回 (comp_id, comp_obj)"""
    comp_id = None
    comp_obj = None
    
    # 1. 处理已经是字典的情况
    if isinstance(raw_data, dict):
        comp_id = raw_data.get('id') or raw_data.get('ID') or raw_data.get('Id')
        return comp_id, raw_data

    # 字节转字符串
    if isinstance(raw_data, bytes):
        raw_data = raw_data.decode('utf-8', errors='ignore')
    raw_data = str(raw_data).strip()
    
    # 2. 尝试解析JSON
    try:
        # 尝试标准 JSON 解析
        parsed = json.loads(raw_data)
        if isinstance(parsed, dict):
            comp_id = parsed.get('id') or parsed.get('ID') or parsed.get('Id')
            comp_obj = parsed
    except Exception:
        # 3. 尝试解析 Python 字典字符串 (解决单引号问题)
        try:
            import ast
            parsed = ast.literal_eval(raw_data)
            if isinstance(parsed, dict):
                comp_id = parsed.get('id') or parsed.get('ID') or parsed.get('Id')
                comp_obj = parsed
        except:
            pass

    # 4. 如果不是JSON/Dict，尝试从URL或文本中提取ID
    if not comp_id:
        # 匹配纯数字
        m_digits = re.fullmatch(r'^\d+$', raw_data)
        if m_digits:
            try:
                comp_id = int(raw_data)
            except:
                pass
        else:
            # 尝试匹配 URL 中的 ID (例如 .../edit/123)
            m_url = re.search(r'/edit/(\d+)', raw_data)
            if m_url:
                try:
                    comp_id = int(m_url.group(1))
                except:
                    pass
            
            if not comp_id:
                # 尝试匹配 id=123
                m_param = re.search(r'[?&]id=(\d+)', raw_data, re.IGNORECASE)
                if m_param:
                    try:
                        comp_id = int(m_param.group(1))
                    except:
                        pass

            if not comp_id:
                # 最后尝试提取任意连续数字（如果只有一组）
                all_digits = re.findall(r'\b(\d{1,10})\b', raw_data)
                if len(all_digits) == 1:
                    try:
                        comp_id = int(all_digits[0])
                    except:
                        pass

    return comp_id, comp_obj
    """
    纯粹的二维码数据解析函数（无数据库操作）
    返回 (comp_id, comp_obj) 元组
    """
    comp_id = None
    comp_obj = None

    # bytes -> str
    if isinstance(raw_data, bytes):
        raw_data = raw_data.decode('utf-8', errors='ignore')
    raw_data = str(raw_data).strip()

    # 尝试解析为 JSON
    try:
        parsed = json.loads(raw_data)
        if isinstance(parsed, dict):
            comp_id = parsed.get('id') or parsed.get('ID') or parsed.get('Id')
            comp_obj = parsed
    except Exception:
        pass

    # 不是 JSON 时尝试提取纯数字 ID
    if not comp_id:
        m_digits = re.match(r'^\d+$', raw_data)
        if m_digits:
            try:
                comp_id = int(raw_data)
            except:
                comp_id = None
        else:
            m2 = re.search(r'\b(\d{1,10})\b', raw_data)
            if m2:
                try:
                    comp_id = int(m2.group(1))
                except:
                    comp_id = None

    return comp_id, comp_obj


def extract_qrcode_data(request):
    """从请求中提取二维码数据（支持JSON/Form/原始Body），返回 (raw_data, error_resp)"""
    try:
        raw_data = None
        # 1. 优先解析JSON请求
        if request.is_json:
            payload = request.get_json(silent=True)
            if isinstance(payload, dict) and payload.get('data') is not None:
                raw_data = payload.get('data')
            else:
                raw_data = json.dumps(payload) if payload else None
        # 2. 解析Form表单
        if raw_data is None and request.form:
            raw_data = request.form.get('data') or request.form.get('qr')
        # 3. 解析原始Body（处理urlencoded）
        if raw_data is None:
            raw_body = request.get_data(as_text=True).strip()
            if raw_body:
                raw_data = urllib.parse.unquote_plus(raw_body)
        # 校验数据是否为空
        if not raw_data:
            return None, (jsonify({'code':0, 'error':'无效的二维码数据'}), 400)
        return raw_data, None
    except Exception as e:
        return None, (jsonify({'code':0, 'error':f'提取数据失败：{str(e)}'}), 500)
    """
    从请求中提取二维码数据（支持多种格式）
    返回 (raw_data_str, error_response)
    """
    try:
        raw_data = None

        # 1) JSON 请求优先
        if request.is_json:
            payload = request.get_json(silent=True)
            if isinstance(payload, dict) and payload.get('data') is not None:
                raw_data = payload.get('data')
            else:
                raw_data = json.dumps(payload) if payload is not None else None

        # 2) form 数据
        if raw_data is None and request.form:
            raw_data = request.form.get('data') or request.form.get('qr') or None

        # 3) 原始 body（文本或 urlencoded）
        if raw_data is None:
            raw_body = request.get_data(as_text=True) or ''
            raw_body = raw_body.strip()
            if raw_body:
                decoded = urllib.parse.unquote_plus(raw_body)
                raw_body = decoded
                m = re.search(r'(\{[\s\S]*\}|\[[\s\S]*\])', raw_body)
                if m:
                    raw_data = m.group(1)
                else:
                    raw_data = raw_body

        if not raw_data:
            return None, (jsonify({'code': 0, 'error': '无效的二维码数据'}), 400)

        return raw_data, None

    except Exception as e:
        return None, (jsonify({'code': 0, 'error': f'提取二维码数据失败：{str(e)}'}), 500)
