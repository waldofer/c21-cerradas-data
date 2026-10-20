import sys
from pathlib import Path

# Los tests importan modulos que viven en la raiz del repo.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
