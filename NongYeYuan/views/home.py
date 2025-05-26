from flask import Blueprint,render_template

hm = Blueprint('home', __name__)

@hm.route('/home') #/目录最好和方法同名
def home():
    return render_template("home.html")