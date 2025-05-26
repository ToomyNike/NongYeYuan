from NongYeYuan import create_app

import sys
print(sys.path)

app = create_app()

if __name__ == '__main__':
    app.run()