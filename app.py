#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Author: lhchlsy2000@163.com
Date: 2024
"""

from flask import Flask, request, jsonify, send_file, render_template, Response, stream_with_context, send_from_directory
import argparse
import os
import tempfile
from md2ppt import MarkdownToPPT
import requests
import json
import pandas as pd
from werkzeug.utils import secure_filename
import config

app = Flask(__name__)

# 自定义LLM API配置
LLM_API_URL = config.LLM_API_URL

# 配置上传文件存储路径
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
ALLOWED_EXCEL_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def call_llm_api(messages, stream=False):
    """调用自定义LLM API"""
    try:
        headers = {
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": config.LLM_MODEL,
            "messages": messages,
            "stream": stream,
            "temperature": config.LLM_REQUEST_CONFIG["temperature"],
            "max_tokens": config.LLM_REQUEST_CONFIG["max_tokens"]
        }
        
        response = requests.post(LLM_API_URL, headers=headers, json=payload, stream=stream)
        response.raise_for_status()
        
        if stream:
            return response.iter_lines()
        else:
            result = response.json()
            return result["choices"][0]["message"]["content"]
    except Exception as e:
        raise Exception(f"LLM API调用失败: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/generate_topics', methods=['POST'])
def generate_topics():
    data = request.json
    role = data.get('role')
    title = data.get('title')
    topic_num = data.get('topicNum')
    
    prompt = f"""
    我希望你作为{role}，以{title}为主题生成{topic_num}个PPT标题，要求能吸引人的注意。
    请用中文生成标题，标题要简洁有力，富有吸引力。
    以下是返回的一些要求：
    1.【The response should be a list of {topic_num} items separated by "\\n"】
    2. 每个标题必须是中文
    3. 标题要简洁，不超过15个字
    """
    
    try:
        content = call_llm_api([{"role": "user", "content": prompt}])
        topics = content.strip().split('\n')
        return jsonify({'topics': topics})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate_outline', methods=['POST'])
def generate_outline():
    data = request.json
    title = data.get('title')
    title_requirement = data.get('requirement', '内容完整，层次分明')
    
    prompt = f"""我希望你使用markdown的格式根据{title}生成一个只有标题的大纲，并且请遵循以下要求:
    1. 如果要创建标题，请在单词或短语前面添加井号 (#) 。# 的数量代表了标题的级别。
    2. 不能使用无序或者有序列表,必须全部使用添加井号 (#)的方式表示大纲结构。
    3. 第一级(#)表示大纲的标题,第二级(##)表示章节的标题,第三级(###)表示章节的重点。
    4. 对大纲来说，要求{title_requirement}。
    5. 大纲的第一章是{title}的简介，最后一章是总结。
    6. 所有标题必须使用中文。
    7. 标题要简洁明了，每个标题不超过15个字。
    """
    
    def generate():
        try:
            for line in call_llm_api([{"role": "user", "content": prompt}], stream=True):
                if line:
                    line = line.decode('utf-8')
                    if line.startswith('data: '):
                        try:
                            data = json.loads(line[6:])
                            if 'choices' in data and len(data['choices']) > 0:
                                content = data['choices'][0].get('delta', {}).get('content', '')
                                if content:
                                    yield f"data: {json.dumps({'content': content})}\n\n"
                        except json.JSONDecodeError:
                            continue
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
        finally:
            yield "data: [DONE]\n\n"
    
    return Response(stream_with_context(generate()), mimetype='text/event-stream')

@app.route('/api/generate_content', methods=['POST'])
def generate_content():
    data = request.json
    outline = data.get('outline')
    requirement = data.get('requirement', '内容专业、通俗易懂')
    
    prompt = f"""我对大纲进行了如下修改,这是修改后的大纲:
    {outline}
    请根据大纲生成的PPT文本的正文内容,我希望你同样以markdown的格式返回,并且请遵循以下要求:
    1. 不要丢失原有的大纲markdown信息和格式。
    2. 每个标题必须独占一行，标题前后不能有其他内容。
    3. 每个段落必须使用<p></p>标签包围，并且独占一行。
    4. 对正文来说，要求{requirement}。
    5. 你需要把生成的段落放在正确的位置。
    6. 所有内容必须使用中文。
    7. 每个段落要简洁明了，控制在50-100字之间。
    8. 使用专业但通俗易懂的语言。
    9. 确保每个标题和段落之间有适当的空行。
    10. 不要在大纲标题和内容之间添加额外的空行。
    """
    
    def generate():
        try:
            for line in call_llm_api([{"role": "user", "content": prompt}], stream=True):
                if line:
                    line = line.decode('utf-8')
                    if line.startswith('data: '):
                        try:
                            data = json.loads(line[6:])
                            if 'choices' in data and len(data['choices']) > 0:
                                content = data['choices'][0].get('delta', {}).get('content', '')
                                if content:
                                    yield f"data: {json.dumps({'content': content})}\n\n"
                        except json.JSONDecodeError:
                            continue
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
        finally:
            yield "data: [DONE]\n\n"
    
    return Response(stream_with_context(generate()), mimetype='text/event-stream')

@app.route('/api/generate_ppt', methods=['POST'])
def generate_ppt():
    try:
        # 获取上传的内容和模板
        content = request.form.get('content')
        template = request.files.get('template')
        
        if not content:
            return jsonify({'error': '请先生成内容'}), 400
        
        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as md_file:
            md_file.write(content)
            md_path = md_file.name
        
        # 创建输出文件路径
        output_path = os.path.join(tempfile.gettempdir(), 'output.pptx')
        
        # 创建转换器实例
        converter = MarkdownToPPT()
        
        # 如果有模板，使用模板
        if template:
            template_path = os.path.join(tempfile.gettempdir(), 'template.pptx')
            template.save(template_path)
            converter = MarkdownToPPT(template_path)
        
        # 设置图片目录
        converter.set_image_dir(UPLOAD_FOLDER)
        
        # 处理Markdown内容
        converter.process_markdown(content)
        
        # 保存PPT
        converter.save(output_path)
        
        # 发送文件
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='output.pptx'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # 清理临时文件
        try:
            if 'md_path' in locals():
                os.unlink(md_path)
            if 'template_path' in locals():
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)
        except Exception as e:
            print(f"清理临时文件失败: {str(e)}")

@app.route('/api/upload_image', methods=['POST'])
def upload_image():
    try:
        if 'image' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400
            
        file = request.files['image']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
            
        if not allowed_file(file.filename, ALLOWED_IMAGE_EXTENSIONS):
            return jsonify({'error': '不支持的文件类型'}), 400
            
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # 返回相对URL和markdown格式
        url = f'/uploads/{filename}'
        markdown = f'![{filename}]({url})'
        return jsonify({'url': url, 'markdown': markdown})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    try:
        if 'excel' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400
            
        file = request.files['excel']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
            
        if not allowed_file(file.filename, ALLOWED_EXCEL_EXTENSIONS):
            return jsonify({'error': '不支持的文件类型'}), 400
            
        # 读取Excel文件
        df = pd.read_excel(file)
        
        # 生成Markdown表格
        markdown_table = "| " + " | ".join(df.columns) + " |\n"
        markdown_table += "| " + " | ".join(["---" for _ in df.columns]) + " |\n"
        
        for _, row in df.iterrows():
            markdown_table += "| " + " | ".join(str(cell) for cell in row) + " |\n"
            
        return jsonify({'markdown': markdown_table})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 添加静态文件服务
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

if __name__ == '__main__':
    # app.run(debug=True) 
    parser = argparse.ArgumentParser(description="Run the Flask app with custom host and port.")
    
    parser.add_argument('--host', type=str, default='127.0.0.1', help='The host to bind the server to (default: 127.0.0.1)')
    parser.add_argument('--port', type=int, default=5000, help='The port to bind the server to (default: 5000)')
    args = parser.parse_args()
    app.run(host=args.host, port=args.port, debug=True)