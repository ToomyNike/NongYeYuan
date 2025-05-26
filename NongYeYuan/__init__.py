from flask import Flask

#创建Flask app 实例
def create_app():
    app = Flask(__name__)
    
    from .views import home 
    #imort的是方法不是文件
    app.register_blueprint(home.hm)

    return app
    # Load configuration
    # app.config.from_object('NongYeYuan.config.Config')

    # # Initialize extensions
    # with app.app_context():
    #     from NongYeYuan import db, migrate
    #     db.init_app(app)
    #     migrate.init_app(app, db)

    # # Register blueprints
    # from NongYeYuan.routes import main_bp
    # app.register_blueprint(main_bp)

    