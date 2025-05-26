from flask import Blueprint, request, render_template

hm = Blueprint('home', __name__)

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'

@hm.route('/home', methods=["GET", "POST"])
def home():
    if request.method == "GET":
        # 渲染 home.html 页面
        return render_template('home.html')
    
    # 处理 POST 请求（文件上传）
    if 'file-input' not in request.files:
        return "No file part", 400
    
    file = request.files['file-input']
    
    # 检查用户是否选择了文件
    if file.filename == '':
        return "No selected file", 400
    
    try:
        # 获取文件名
        filename = file.filename
        # 输出文件名到终端
        print(f"Uploaded file name: {filename}")
        
        return "File name received and printed to terminal", 200
    
    except Exception as e:
        return f"Error processing file: {str(e)}", 500