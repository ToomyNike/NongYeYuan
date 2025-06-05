from flask import Blueprint, request, render_template
import os

hm = Blueprint('home', __name__)

# 配置上传文件夹
UPLOAD_FOLDER = 'input'
# 确保上传文件夹存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@hm.route('/home', methods=["GET", "POST"])
def home():
    if request.method == "GET":
        # 渲染 home.html 页面
        print("已渲染")
        return render_template('home.html')
    
    # 处理 POST 请求（文件上传）
    if 'file-input' not in request.files:
        print("未获取文件")
        return "No file part", 400
    
    file = request.files['file-input']
    
    # 检查用户是否选择了文件
    if file.filename == '':
        print("文件名为空")
        return "No selected file", 400
    
    try:
        # 获取文件名
        
        filename = file.filename
        print("上传文件：",filename)
        # 保存文件到上传文件夹
        file.save(os.path.join(UPLOAD_FOLDER, filename))
        
        return "File saved successfully", 200
    
    except Exception as e:
        print(5)
        return f"Error processing file: {str(e)}", 500