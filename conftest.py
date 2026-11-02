import sys
from pathlib import Path

# Scripts live in the repository root and import each other by module name
sys.path.insert(0, str(Path(__file__).parent))
