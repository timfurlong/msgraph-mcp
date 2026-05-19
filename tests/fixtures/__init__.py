import json
from pathlib import Path

_FIXTURES = Path(__file__).parent


def load(name: str) -> dict:
    return json.loads((_FIXTURES / name).read_text())
