# run.py - Convenient Root Runner for Cyberpunk Task Extractor Web Server
import sys
from pathlib import Path

# Add backend to python path
BACKEND_DIR = Path(__file__).resolve().parent / "backend"
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from app import start_server

if __name__ == "__main__":
    start_server(port=5000)
