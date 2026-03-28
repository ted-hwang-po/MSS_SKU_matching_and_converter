"""UI 헬퍼: Tooltip, 사용자 친화적 에러 메시지, 파일 검증"""
import re
import tkinter as tk
from tkinter import ttk
from pathlib import Path


# ================================================================
#  Tooltip
# ================================================================
class ToolTip:
    """위젯에 마우스를 올리면 표시되는 툴팁"""

    def __init__(self, widget: tk.Widget, text: str, delay: int = 400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._tip_window = None
        self._after_id = None
        widget.bind('<Enter>', self._schedule)
        widget.bind('<Leave>', self._hide)
        widget.bind('<ButtonPress>', self._hide)

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip_window:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self._tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f'+{x}+{y}')
        label = tk.Label(
            tw, text=self.text, justify=tk.LEFT,
            background='#fffde7', foreground='#333',
            relief=tk.SOLID, borderwidth=1,
            font=('', 9), wraplength=350, padx=6, pady=4,
        )
        label.pack()

    def _hide(self, event=None):
        self._cancel()
        if self._tip_window:
            self._tip_window.destroy()
            self._tip_window = None

    def update_text(self, text: str):
        self.text = text


# ================================================================
#  사용자 친화적 에러 메시지 변환
# ================================================================
_ERROR_PATTERNS = [
    # (정규식 패턴, 사용자 메시지 템플릿)
    (r"'(.+?)' 컬럼을 찾을 수 없습니다\. 후보: \[(.+?)\], 실제 컬럼: \[(.+?)\]",
     "'{0}' 컬럼을 찾을 수 없습니다.\n\n"
     "확인사항:\n"
     "  - 올바른 파일을 선택했는지 확인해주세요\n"
     "  - 파일에 다음 컬럼 중 하나가 있어야 합니다: {1}"),

    (r"총 발주 수량 데이터셋에 '브랜드명' 컬럼이 없습니다",
     "파일 1에 '브랜드명' 컬럼이 없습니다.\n\n"
     "확인사항:\n"
     "  - 총 발주 수량 데이터셋이 맞는지 확인해주세요\n"
     "  - 파일 첫 행이 헤더(컬럼명)인지 확인해주세요"),

    (r"헤더를 찾을 수 없습니다: (.+)",
     "파일에서 헤더 행을 찾을 수 없습니다.\n"
     "파일: {0}\n\n"
     "확인사항:\n"
     "  - 발주 파일 형식이 맞는지 확인해주세요\n"
     "  - '브랜드명' 컬럼이 포함되어 있어야 합니다"),

    (r"Permission denied|PermissionError",
     "파일 접근이 거부되었습니다.\n\n"
     "확인사항:\n"
     "  - 해당 파일이 다른 프로그램(Excel 등)에서 열려 있지 않은지 확인해주세요\n"
     "  - 파일/폴더 접근 권한을 확인해주세요"),

    (r"No such file or directory|FileNotFoundError",
     "파일을 찾을 수 없습니다.\n\n"
     "확인사항:\n"
     "  - 파일이 삭제되거나 이동되지 않았는지 확인해주세요\n"
     "  - 파일 경로가 올바른지 확인해주세요"),

    (r"openpyxl.*not.*support|InvalidFileException|BadZipFile",
     "파일 형식을 읽을 수 없습니다.\n\n"
     "확인사항:\n"
     "  - .xlsx 또는 .csv 형식의 파일인지 확인해주세요\n"
     "  - 파일이 손상되지 않았는지 확인해주세요"),

    (r"out of range|KeyError|IndexError",
     "데이터 처리 중 예상치 못한 값이 발견되었습니다.\n\n"
     "확인사항:\n"
     "  - 입력 파일의 데이터가 비어있지 않은지 확인해주세요\n"
     "  - 파일 형식이 올바른지 확인해주세요"),
]


def friendly_error(exc: Exception) -> str:
    """예외를 사용자 친화적 메시지로 변환. 매칭 실패 시 원본 메시지 반환."""
    error_str = str(exc)
    for pattern, template in _ERROR_PATTERNS:
        match = re.search(pattern, error_str, re.IGNORECASE)
        if match:
            groups = match.groups()
            try:
                return template.format(*groups)
            except (IndexError, KeyError):
                return template
    # 매칭되지 않는 에러는 타입 + 메시지만 표시 (traceback 제외)
    type_name = type(exc).__name__
    return f"오류가 발생했습니다 ({type_name}):\n{error_str}"


# ================================================================
#  파일 스키마 검증
# ================================================================
_FILE_SCHEMAS = {
    1: {
        'name': '총 발주 수량 데이터셋',
        'required_any': [['브랜드명']],
        'expected': ['브랜드명', '88코드', '상품명', '수량(오프)'],
    },
    2: {
        'name': '상품코드-바코드 매칭 데이터셋',
        'required_any': [['바코드'], ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], ['스타일번호']],
        'expected': ['바코드', '상품코드', '스타일번호'],
    },
    3: {
        'name': '상품별 옵션 정보',
        'required_any': [['사이즈유형'], ['사이즈유형명']],
        'expected': ['사이즈유형', '사이즈유형명', 'Size01'],
    },
}


def validate_file_schema(filepath: str, file_num: int) -> dict:
    """파일의 스키마를 검증하고 결과를 반환.

    Returns:
        {
            'valid': bool,
            'message': str,  # 상태 메시지
            'columns': list[str],  # 발견된 컬럼 목록
            'row_count': int,  # 데이터 행 수
            'missing': list[str],  # 누락된 필수 컬럼 그룹
        }
    """
    from .loader import load_order_data, load_matching_data, load_option_data, _find_header_row

    schema = _FILE_SCHEMAS.get(file_num)
    if not schema:
        return {'valid': True, 'message': '검증 스킵', 'columns': [], 'row_count': 0, 'missing': []}

    try:
        filepath = Path(filepath)
        if not filepath.exists():
            return {'valid': False, 'message': '파일이 존재하지 않습니다', 'columns': [], 'row_count': 0, 'missing': []}

        # 파일 로드 (각 파일별 적절한 로더 사용)
        if file_num == 1:
            df = load_order_data(filepath)
        elif file_num == 2:
            df = load_matching_data(filepath)
        else:
            df = load_option_data(filepath)

        columns = list(df.columns)
        row_count = len(df)

        # 필수 컬럼 그룹 검증
        missing_groups = []
        for group in schema['required_any']:
            found = any(c in columns for c in group)
            if not found:
                missing_groups.append(group)

        if missing_groups:
            missing_str = ', '.join(['/'.join(g) for g in missing_groups])
            return {
                'valid': False,
                'message': f"필수 컬럼 누락: {missing_str}",
                'columns': columns,
                'row_count': row_count,
                'missing': missing_groups,
            }

        return {
            'valid': True,
            'message': f"{row_count}행 로드됨",
            'columns': columns,
            'row_count': row_count,
            'missing': [],
        }

    except Exception as e:
        return {
            'valid': False,
            'message': friendly_error(e),
            'columns': [],
            'row_count': 0,
            'missing': [],
        }


# ================================================================
#  파일 상태 표시 위젯
# ================================================================
class FileStatusLabel(ttk.Frame):
    """파일 상태를 아이콘 + 텍스트로 표시하는 위젯"""

    STATUS_COLORS = {
        'none': '#999999',
        'loading': '#f59e0b',
        'valid': '#22c55e',
        'invalid': '#ef4444',
    }

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._status_label = ttk.Label(self, text='', font=('', 8))
        self._status_label.pack(side=tk.LEFT, padx=(5, 0))
        self.set_status('none', '')

    def set_status(self, status: str, message: str):
        icons = {'none': '  ', 'loading': '..', 'valid': 'OK', 'invalid': '!!'}
        icon = icons.get(status, '  ')
        color = self.STATUS_COLORS.get(status, '#999')
        self._status_label.configure(
            text=f"[{icon}] {message}" if message else '',
            foreground=color,
        )
