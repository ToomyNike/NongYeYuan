from flask import Blueprint, request, render_template, send_file
import os
from Quanmian import runQuanmian
from HeHu import runHeHu
from XiaoQu import runXiaoQu

hm = Blueprint('home', __name__)

# 配置上传
UPLOAD_FOLDER = 'input'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@hm.route('/home', methods=["GET", "POST"])
def home():
    if request.method == "GET":
        print("已渲染")
        return render_template('home.html')
    if 'file-input' not in request.files:
        print("未获取文件")
        return "No file part", 400
    
    file = request.files['file-input']
    
    if file.filename == '':
        print("文件名为空")
        return "No selected file", 400
    
    try:
        
        filename = file.filename
        print("上传文件：", filename)
        file.save(os.path.join(UPLOAD_FOLDER, filename))

  
        selected_area = request.form.get('selected-area')      # 获取选中的区域值
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
    
# 添加下载文件的路由
@hm.route('/download')
def download_file():
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        nongyeyuan_dir = os.path.dirname(current_dir)
        output_dir = os.path.join(os.path.dirname(nongyeyuan_dir), 'output')
        file_path = os.path.join(output_dir, 'result_test.xlsx')

        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return "文件不存在", 404
        
        print("弹出下载")
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        print(f"下载文件时出错: {e}")
        return "下载文件失败", 500