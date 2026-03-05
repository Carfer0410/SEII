from .accounting_reports import *  # noqa: F401,F403


if __name__ == "__main__":
    import os
    import sys

    project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    from app import app, ensure_db

    ensure_db()
    app.run(debug=True)
