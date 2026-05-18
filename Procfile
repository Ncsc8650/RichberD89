web: python -c "import os, app; app.init_db(); app.ThreadingHTTPServer(('0.0.0.0', int(os.environ.get('PORT','8000'))), app.App).serve_forever()"
