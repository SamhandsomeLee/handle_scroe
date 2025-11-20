"""
成绩导入转换工具 - Web UI 版本
基于 Flask 的现代化 Web 应用
"""
import os
import sys
import io
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from excel_handler import ExcelHandler

# 配置
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# 创建文件夹
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# 创建 Flask 应用
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 限制

def allowed_file(filename):
    """检查文件是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/api/info', methods=['GET'])
def get_info():
    """获取应用信息"""
    return jsonify({
        'name': '成绩导入转换工具',
        'version': '2.0.0',
        'description': '将包含两次考试数据的 Excel 文件转换为标准的8工作表格式'
    })

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传文件"""
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未选择文件'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式，请上传 .xls 或 .xlsx 文件'}), 400
        
        # 保存文件
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 获取文件信息
        info = ExcelHandler.get_file_info(filepath)
        
        if not info['success']:
            return jsonify({'success': False, 'error': info['error']}), 400
        
        return jsonify({
            'success': True,
            'filename': filename,
            'filepath': filepath,
            'sheets': info['sheets'],
            'total_sheets': info['total_sheets']
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/convert', methods=['POST'])
def convert_file():
    """转换文件"""
    try:
        data = request.json
        input_filepath = data.get('filepath')
        
        if not input_filepath or not os.path.exists(input_filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 400
        
        # 读取文件
        sheets = ExcelHandler.read_excel_file(input_filepath)
        
        # 生成输出文件名
        input_filename = os.path.basename(input_filepath)
        output_filename = f"转换_{input_filename}"
        output_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        
        # 确保输出目录存在
        os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
        
        # 转换
        ExcelHandler.write_excel_file(sheets, output_filepath)
        
        # 验证文件是否生成
        if not os.path.exists(output_filepath):
            return jsonify({'success': False, 'error': '文件生成失败'}), 500
        
        print(f"文件已生成: {output_filepath}")
        print(f"文件大小: {os.path.getsize(output_filepath)} bytes")
        
        return jsonify({
            'success': True,
            'message': '转换成功！',
            'output_filename': output_filename,
            'output_filepath': output_filepath
        })
    
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)}), 400
    except Exception as e:
        print(f"转换错误: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'转换失败: {str(e)}'}), 500

@app.route('/api/download/<path:filename>', methods=['GET'])
def download_file(filename):
    """下载文件"""
    try:
        # 直接使用 filename，不用 secure_filename
        filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
        
        print(f"下载请求: {filename}")
        print(f"完整路径: {filepath}")
        print(f"文件存在: {os.path.exists(filepath)}")
        
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': f'文件不存在: {filepath}'}), 404
        
        # 确保文件可读
        if not os.access(filepath, os.R_OK):
            return jsonify({'success': False, 'error': '无权访问文件'}), 403
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.ms-excel'
        )
    
    except Exception as e:
        print(f"Download error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/cleanup', methods=['POST'])
def cleanup():
    """清理临时文件"""
    try:
        # 清理上传文件夹
        for file in os.listdir(app.config['UPLOAD_FOLDER']):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file)
            if os.path.isfile(filepath):
                os.remove(filepath)
        
        # 清理下载文件夹（可选）
        # for file in os.listdir(app.config['DOWNLOAD_FOLDER']):
        #     filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], file)
        #     if os.path.isfile(filepath):
        #         os.remove(filepath)
        
        return jsonify({'success': True, 'message': '清理完成'})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.errorhandler(404)
def not_found(error):
    """404 错误处理"""
    return jsonify({'success': False, 'error': '页面不存在'}), 404

@app.errorhandler(500)
def internal_error(error):
    """500 错误处理"""
    return jsonify({'success': False, 'error': '服务器内部错误'}), 500

if __name__ == '__main__':
    print("=" * 80)
    print("成绩导入转换工具 - Web UI")
    print("=" * 80)
    print("\n启动服务器...")
    print("访问地址: http://localhost:5000")
    print("按 Ctrl+C 停止服务器\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
