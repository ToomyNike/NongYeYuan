from flask import Blueprint, request, render_template
import os
from Quanmian import runQuanmian
from HeHu import runHeHu
from XiaoQu import runXiaoQu

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
        print("上传文件：", filename)
        # 保存文件到上传文件夹
        file.save(os.path.join(UPLOAD_FOLDER, filename))

        # 获取选中的区域值
        selected_area = request.form.get('selected-area')
        if selected_area:
            
            if selected_area == "县区域":
                print("选中的区域:", selected_area)
                runQuanmian()
            elif selected_area == "河湖流域":
                print("选中的区域:", selected_area)
                runHeHu()
            elif selected_area == "重点小流域":
                print("选中的区域:", selected_area)
                runXiaoQu()

        return "File saved successfully", 200
    
    except Exception as e:
        print("Somethings went wrong:")
        return f"Error processing file: {str(e)}", 500