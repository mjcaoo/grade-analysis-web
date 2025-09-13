# -*- coding: utf-8 -*-
# @Author  : mjcaoo
# @File    : app.py
# @Time    : 2025/9/13

from flask import Flask, render_template, request, jsonify, send_file, session, Blueprint
import os
import uuid
from datetime import datetime
import json
import pandas as pd
from grade_analyzer import GradeAnalyzer
from werkzeug.utils import secure_filename
from pytz import timezone

# 创建Flask应用，支持子路径部署
app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # 在生产环境中应该使用环境变量

# 配置子路径前缀
URL_PREFIX = '/zongce'

# 创建Blueprint
main_bp = Blueprint('main', __name__, url_prefix=URL_PREFIX)

# 处理反向代理配置
class ProxyFix:
    def __init__(self, app):
        self.app = app
    
    def __call__(self, environ, start_response):
        # 处理反向代理的路径
        script_name = environ.get('HTTP_X_SCRIPT_NAME', '')
        if script_name:
            environ['SCRIPT_NAME'] = script_name
            path_info = environ['PATH_INFO']
            if path_info.startswith(script_name):
                environ['PATH_INFO'] = path_info[len(script_name):]
        return self.app(environ, start_response)

# 应用代理修复
app.wsgi_app = ProxyFix(app.wsgi_app)

# 配置
UPLOAD_FOLDER = 'uploads'
RESULTS_FOLDER = 'results'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传和结果目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@main_bp.route('/')
def index():
    """主页 - 显示文件上传界面"""
    return render_template('index.html')

@main_bp.route('/upload', methods=['POST'])
def upload_files():
    """处理文件上传"""
    try:
        # 检查是否有成绩文件和主要课程文件
        if 'grade_files' not in request.files:
            return jsonify({'error': '请选择成绩文件'}), 400
        
        if 'main_course_file' not in request.files:
            return jsonify({'error': '请选择主要课程列表文件'}), 400
        
        # 检查主要课程文件是否为空
        main_course_file = request.files.get('main_course_file')
        if not main_course_file or main_course_file.filename == '':
            return jsonify({'error': '主要课程列表文件是必需的'}), 400
        
        # 生成会话ID
        session_id = str(uuid.uuid4())
        session['session_id'] = session_id
        
        # 创建会话专用的上传目录
        session_upload_dir = os.path.join(UPLOAD_FOLDER, session_id)
        os.makedirs(session_upload_dir, exist_ok=True)
        
        uploaded_files = []
        
        # 处理成绩文件（可能有多个）
        grade_files = request.files.getlist('grade_files')
        for file in grade_files:
            if file and file.filename and file.filename != '' and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(session_upload_dir, filename)
                file.save(file_path)
                uploaded_files.append({
                    'type': 'grade',
                    'filename': filename,
                    'path': file_path
                })
        
        # 处理主要课程列表文件（必需）
        if not main_course_file.filename or not allowed_file(main_course_file.filename):
            return jsonify({'error': '主要课程列表文件格式不正确，请选择Excel文件'}), 400
            
        filename = secure_filename(main_course_file.filename)
        file_path = os.path.join(session_upload_dir, filename)
        main_course_file.save(file_path)
        uploaded_files.append({
            'type': 'main_course',
            'filename': filename,
            'path': file_path
        })
        
        if not uploaded_files:
            return jsonify({'error': '没有有效的文件上传'}), 400
        
        # 保存上传文件信息到会话
        session['uploaded_files'] = uploaded_files
        
        return jsonify({
            'message': '文件上传成功',
            'session_id': session_id,
            'files': [f['filename'] for f in uploaded_files]
        })
        
    except Exception as e:
        return jsonify({'error': f'文件上传失败: {str(e)}'}), 500

@main_bp.route('/analyze', methods=['POST'])
def analyze():
    """执行成绩分析"""
    try:
        # 检查会话
        if 'session_id' not in session or 'uploaded_files' not in session:
            return jsonify({'error': '请先上传文件'}), 400
        
        session_id = session['session_id']
        uploaded_files = session['uploaded_files']
        
        # 分离成绩文件和主要课程文件
        grade_file_paths = []
        main_course_file_path = None
        
        for file_info in uploaded_files:
            if file_info['type'] == 'grade':
                grade_file_paths.append(file_info['path'])
            elif file_info['type'] == 'main_course':
                main_course_file_path = file_info['path']
        
        if not grade_file_paths:
            return jsonify({'error': '没有找到成绩文件'}), 400
        
        if not main_course_file_path:
            return jsonify({'error': '没有找到主要课程列表文件'}), 400
        
        # 创建分析器实例
        analyzer = GradeAnalyzer()
        
        # 加载主要课程列表（必需）
        analyzer.load_main_courses(main_course_file_path)
        if not analyzer.main_courses:
            return jsonify({'error': '主要课程列表为空或加载失败'}), 400
        
        # 执行分析
        results_df = analyzer.process_combined_data(grade_file_paths)
        
        if results_df.empty:
            return jsonify({'error': '分析失败，没有有效的数据'}), 400
        
        # 生成结果文件
        shanghai_tz = timezone('Asia/Shanghai')
        timestamp = datetime.now(shanghai_tz).strftime('%Y%m%d_%H%M%S')
        result_filename = f'成绩分析结果_{timestamp}.xlsx'
        result_path = os.path.join(RESULTS_FOLDER, result_filename)
        
        # 保存Excel结果
        with pd.ExcelWriter(result_path, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='综合成绩', index=False)
        
        # 将结果转换为JSON格式返回
        results_json = results_df.to_dict('records')
        
        # 保存分析结果到会话（只保存必要信息，减少session大小）
        session['analysis_results'] = {
            'result_file': result_filename,
            'result_path': result_path,
            'timestamp': timestamp,
            'student_count': len(results_json)
        }
        
        return jsonify({
            'message': '分析完成',
            'student_count': len(results_json),
            'result_file': result_filename,
            'preview': results_json[:10] if len(results_json) > 10 else results_json  # 返回前10条预览
        })
        
    except Exception as e:
        return jsonify({'error': f'分析失败: {str(e)}'}), 500

@main_bp.route('/results')
def get_results():
    """获取分析结果"""
    try:
        if 'analysis_results' not in session:
            return jsonify({'error': '没有找到分析结果'}), 404
        
        analysis_results = session['analysis_results']
        
        # 从Excel文件重新读取数据（因为session中不再存储完整数据）
        result_path = analysis_results['result_path']
        if os.path.exists(result_path):
            df = pd.read_excel(result_path)
            results_data = df.to_dict('records')
        else:
            return jsonify({'error': '结果文件不存在'}), 404
        
        return jsonify({
            'data': results_data,
            'student_count': analysis_results['student_count'],
            'result_file': analysis_results['result_file'],
            'timestamp': analysis_results['timestamp']
        })
        
    except Exception as e:
        return jsonify({'error': f'获取结果失败: {str(e)}'}), 500

@main_bp.route('/download/<path:filename>')
def download_file(filename):
    """下载分析结果文件"""
    try:
        if 'analysis_results' not in session:
            return jsonify({'error': '没有找到分析结果'}), 404
        
        result_path = session['analysis_results']['result_path']
        
        if not os.path.exists(result_path):
            return jsonify({'error': '结果文件不存在'}), 404
        
        # 从URL解码文件名
        import urllib.parse
        decoded_filename = urllib.parse.unquote(filename)
        
        return send_file(result_path, as_attachment=True, download_name=decoded_filename)
        
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500

@main_bp.route('/download_result')
def download_result():
    """直接下载分析结果文件（不通过文件名参数）"""
    try:
        if 'analysis_results' not in session:
            return jsonify({'error': '没有找到分析结果'}), 404
        
        analysis_results = session['analysis_results']
        result_path = analysis_results['result_path']
        result_filename = analysis_results['result_file']
        
        if not os.path.exists(result_path):
            return jsonify({'error': '结果文件不存在'}), 404
        
        return send_file(result_path, as_attachment=True, download_name=result_filename)
        
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500

@main_bp.route('/api/status')
def status():
    """API状态检查"""
    return jsonify({
        'status': 'running',
        'message': '成绩分析Web服务正在运行',
        'version': '1.0.0'
    })

@main_bp.route('/sample/<path:filename>')
def download_sample(filename):
    """下载示例文件"""
    try:
        sample_dir = os.path.join(app.root_path, 'static', 'samples')
        file_path = os.path.join(sample_dir, filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': '示例文件不存在'}), 404
        
        return send_file(file_path, as_attachment=True, download_name=filename)
        
    except Exception as e:
        return jsonify({'error': f'下载示例文件失败: {str(e)}'}), 500

# 注册Blueprint
app.register_blueprint(main_bp)

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': '页面不存在'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': '内部服务器错误'}), 500

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)