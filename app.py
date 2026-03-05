from app_modules.bootstrap import *


if __name__ == '__main__':
    ensure_db()
    app.run(debug=True, use_reloader=False)
