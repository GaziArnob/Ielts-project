PythonAnywhere free deployment steps

1. Create a free PythonAnywhere account.
2. Open Bash and clone this repo:
   `git clone https://github.com/GaziArnob/Ielts-project.git`
3. Create a web app using Flask with Python 3.11.
4. In the web app WSGI file, replace the contents with:

```python
import os
import sys
from pathlib import Path

project_home = Path("/home/YOUR_USERNAME/Ielts-project").resolve()
if str(project_home) not in sys.path:
    sys.path.insert(0, str(project_home))

os.environ["APP_DATA_DIR"] = str(project_home / "data")
os.environ["FLASK_SECRET_KEY"] = "change-this-secret"
os.environ["IELTS_ADMIN_PASSWORD"] = "change-this-password"

from wsgi import application
```

5. In Static Files, map:
   URL: `/static/`
   Directory: `/home/YOUR_USERNAME/Ielts-project/static`
6. Reload the web app.

Notes

- Uploaded speaking audio, exports, and the SQLite database will be stored inside `data/`.
- The free PythonAnywhere URL will look like:
  `https://YOUR_USERNAME.pythonanywhere.com`
