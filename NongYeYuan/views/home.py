from flask import Blueprint, render_template, request, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename

hm = Blueprint('home', __name__)

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@hm.route('/home', methods=["GET", "POST"])
def home():
    print("起始位置！！！！！！！！！！")
    if request.method == "GET":
        return render_template("home.html")
    
    # 处理 POST 请求（文件上传）
    if 'file-input' not in request.files:
        return jsonify({'error': '没有文件部分'}), 400
    
    file = request.files['file-input']
    
    # 检查用户是否选择了文件
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    # 检查文件类型是否允许
    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件类型，只允许 .xlsx 和 .xls 文件'}), 400
    
    try:
        # 生成唯一文件名
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        
        # 保存文件
        file.save(file_path)
        
        # 读取 Excel 文件
        df = pd.read_excel(file_path)
        
        # 在终端输出文件信息
        print(f"\n=== 文件信息 ===")
        print(f"文件名: {filename}")
        print(f"文件路径: {file_path}")
        print(f"总行数: {len(df)}")
        print(f"列名: {', '.join(df.columns)}")
        
        # 输出数值列的统计信息
        print("\n=== 数值列统计信息 ===")
        for col in df.select_dtypes(include=['number']).columns:
            print(f"{col}:")
            print(f"  平均值: {df[col].mean()}")
            print(f"  最小值: {df[col].min()}")
            print(f"  最大值: {df[col].max()}")
            print(f"  标准差: {df[col].std()}")
        
        # 返回成功响应（不含下载链接）
        return jsonify({
            'success': True,
            'message': '文件上传成功，服务器已处理文件信息'
        }), 200
    
    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

print("终点位置！！！！！！！！！！")