import subprocess
import sys
import os

here = os.path.dirname(os.path.abspath(__file__))
app_path = os.path.join(here, "impact_index_app.py")

subprocess.run([sys.executable, "-m", "streamlit", "run", app_path])
