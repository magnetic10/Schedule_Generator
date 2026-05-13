# config.py

NUM_WORKERS = 8
DAYS_IN_MONTH = 30  # 예시, 실행 시 변경 가능

# 근무 형태 별칭
SHIFT_DAY = 0    # 주
SHIFT_NIGHT = 1  # 야
SHIFT_OFF_NIGHT = 2 # 비
SHIFT_OFF = 3    # 휴
SHIFT_LEAVE = 4  # 연

STR_TO_SHIFT = {
    '주': SHIFT_DAY,
    '야': SHIFT_NIGHT,
    '비': SHIFT_OFF_NIGHT,
    '휴': SHIFT_OFF,
    '연': SHIFT_LEAVE
}
SHIFT_TO_STR = {v: k for k, v in STR_TO_SHIFT.items()}

# 근무 시간
HOURS = {
    SHIFT_DAY: 8,
    SHIFT_NIGHT: 8,
    SHIFT_OFF_NIGHT: 8,
    SHIFT_OFF: 0,
    SHIFT_LEAVE: 8
}

# TO (필요 인원) 설정
MIN_DAY_WORKERS = 1
MAX_DAY_WORKERS = 3
TARGET_DAY_WORKERS = 2

MIN_NIGHT_WORKERS = 1
MAX_NIGHT_WORKERS = 2
TARGET_NIGHT_WORKERS = 2

TARGET_HOURS_PER_MONTH = 160 # 월 목표 근무 시간
