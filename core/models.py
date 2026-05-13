from dataclasses import dataclass, field
from typing import List, Dict

@dataclass
class Worker:
    id: int
    name: str
    is_only_day: bool = False
    dedicated_shift: str = ""  # "" / "day" / "night"
    shift_preference: str = "-"  # "-" / "주간선호" / "야간선호"
    preferred_day_days: List[int] = field(default_factory=list)
    preferred_night_days: List[int] = field(default_factory=list)
    fixed_leave_days: List[int] = field(default_factory=list)  # 연가(연)
    fixed_off_days: List[int] = field(default_factory=list)    # 휴무(휴) 등 고정
    fixed_day_days: List[int] = field(default_factory=list)    # 주간 강제 지정
    fixed_night_days: List[int] = field(default_factory=list)  # 야간 강제 지정
    fixed_off_night_days: List[int] = field(default_factory=list) # 비번(비) 강제 지정
    # [신규] 커스텀 기타 근무 저장 { "약어": [날짜리스트] }
    custom_shifts: Dict[str, List[int]] = field(default_factory=dict)
    prev_month_last_day_night: bool = False  # 지난달 마지막 날 야간 근무 여부
    target_hours: int = 160 # 목표 근무시간 (개인별)
    start_day: int = None   # 근무 시작일 (인덱스 0~, None이면 1일부터)
    end_day: int = None     # 근무 종료일 (인덱스 0~, None이면 말일까지)
