from flask import Flask,request

# 用当前脚本名称实例化Flask对象，方便flask从该脚本文件中获取需要的内容
app = Flask(__name__)

#处理url和视图函数之间的关系的程序就是"路由"，在Flask中，路由是通过@app.route装饰器(以@开头)来表示的
@app.route("/index")
#url映射的函数，要传参则在上述route（路由）中添加参数申明
def index():
    age=request.args.get("age") #get请求（就是地址栏直接get）
    # 通过request对象获取请求参数
    pwd = request.args.get("pwd")
    # 通过request对象获取请求参数
    print("age:", age)
    print("pwd:", pwd)
    return "Hello World!"

# 直属的第一个作为视图函数被绑定，第二个就是普通函数
# 路由与视图函数需要一一对应
# def not():
#     return "Not Hello World!"

# 启动一个本地开发服务器，激活该网页
app.run(host="127.0.0.1", port=5000, debug=True)

