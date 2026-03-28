"""세션 관리: 상태 저장/복원, 작업 이력, 프리셋"""
import json
from datetime import datetime
from pathlib import Path

# 설정 디렉토리
_CONFIG_DIR = Path.home() / '.mss_converter'
_SESSION_FILE = _CONFIG_DIR / 'session.json'
_HISTORY_FILE = _CONFIG_DIR / 'history.json'
_PRESETS_FILE = _CONFIG_DIR / 'presets.json'

MAX_HISTORY = 100


def _ensure_dir():
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)


def _read_json(path: Path) -> dict | list:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding='utf-8'))
    except (json.JSONDecodeError, OSError):
        pass
    return {} if path == _SESSION_FILE or path == _PRESETS_FILE else []


def _write_json(path: Path, data):
    _ensure_dir()
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


# ================================================================
#  세션 상태 저장/복원 (E1)
# ================================================================
def save_session(state: dict):
    """현재 세션 상태를 저장.

    state 예시:
    {
        'tab_a': {
            'file1': '/path/to/file1.xlsx',
            'file2': '/path/to/file2.xlsx',
            'file3': '/path/to/file3.xlsx',
            'save_path': '/path/to/output',
            'brand_mode': 'all' | 'select',
            'selected_brands': ['브랜드A', '브랜드B'],
        },
        'tab_b': {
            'files': ['/path1.xlsx', '/path2.xlsx'],
            'matching_file': '/path/to/match.xlsx',
            'save_path': '/path/to/output',
        },
        'last_tab': 0,
    }
    """
    _write_json(_SESSION_FILE, state)


def load_session() -> dict:
    """저장된 세션 상태를 복원."""
    data = _read_json(_SESSION_FILE)
    return data if isinstance(data, dict) else {}


# ================================================================
#  작업 이력 (P1)
# ================================================================
def add_history_entry(entry: dict):
    """작업 이력 추가.

    entry 예시:
    {
        'timestamp': '2026-03-29 14:30:00',
        'type': 'A' | 'B',
        'brands': ['브랜드A'],
        'success': 3,
        'fail': 0,
        'total': 3,
        'files': {'file1': 'name1.xlsx', 'file2': 'name2.xlsx', 'file3': 'name3.xlsx'},
        'output_path': '/path/to/output',
        'warnings': 2,
    }
    """
    history = get_history()
    entry.setdefault('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    history.insert(0, entry)
    # 최대 이력 수 제한
    if len(history) > MAX_HISTORY:
        history = history[:MAX_HISTORY]
    _write_json(_HISTORY_FILE, history)


def get_history() -> list[dict]:
    """작업 이력 조회."""
    data = _read_json(_HISTORY_FILE)
    return data if isinstance(data, list) else []


def clear_history():
    """작업 이력 전체 삭제."""
    _write_json(_HISTORY_FILE, [])


# ================================================================
#  프리셋 (P3)
# ================================================================
def save_preset(name: str, config: dict):
    """프리셋 저장.

    config 예시:
    {
        'file1': '/path/to/file1.xlsx',
        'file2': '/path/to/file2.xlsx',
        'file3': '/path/to/file3.xlsx',
        'save_path': '/path/to/output',
        'brand_mode': 'select',
        'selected_brands': ['브랜드A', '브랜드B'],
    }
    """
    presets = get_presets()
    config['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    presets[name] = config
    _write_json(_PRESETS_FILE, presets)


def get_presets() -> dict:
    """모든 프리셋 조회. {이름: config}"""
    data = _read_json(_PRESETS_FILE)
    return data if isinstance(data, dict) else {}


def delete_preset(name: str):
    """프리셋 삭제."""
    presets = get_presets()
    presets.pop(name, None)
    _write_json(_PRESETS_FILE, presets)


def load_preset(name: str) -> dict | None:
    """특정 프리셋 로드."""
    presets = get_presets()
    return presets.get(name)
