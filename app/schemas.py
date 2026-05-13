from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


ShiftCode = Literal["day", "night", "off_night", "off", "leave"]
PreferenceCode = Literal["day", "night"] | None
DedicatedShiftCode = Literal["day", "night"] | None


@dataclass(slots=True)
class WorkerInput:
    name: str
    start_day: int = 1
    end_day: int | None = None
    is_day_only: bool = False
    dedicated_shift: DedicatedShiftCode = None
    target_hours: int | None = None
    preference: PreferenceCode = None
    prev_month_last_day_night: bool = False
    fixed_shifts: dict[int, ShiftCode | str] = field(default_factory=dict)


@dataclass(slots=True)
class ScheduleSettings:
    target_day: int = 1
    target_night: int = 2
    min_day: int = 1
    max_day: int = 2
    min_night: int = 1
    max_night: int = 2
    use_emergency_range: bool = False
    emergency_min_day: int | None = None
    emergency_max_day: int | None = None
    emergency_min_night: int | None = None
    emergency_max_night: int | None = None
    use_preference: bool = False
    allow_leave_after_off_night: bool = False
    allow_double_night_cycle: bool = False
    use_advanced_settings: bool = False
    max_consecutive_day: int = 5
    max_consecutive_rest: int = 4
    allow_user_forced_rule_violations: bool = False
    penalty_order: list[str] = field(
        default_factory=lambda: [
            "five_rest_streak",
            "day_streak_over_default",
            "emergency_range",
            "double_night_cycle",
        ]
    )
    special_shifts: dict[str, int] = field(default_factory=dict)


@dataclass(slots=True)
class MonthInfo:
    year: int
    month: int
    days_in_month: int
    workday_count: int
    default_target_hours: int
    label: str


@dataclass(slots=True)
class SidebarDefaults:
    month_info: MonthInfo
    common_target_hours: int
    settings: ScheduleSettings
    use_individual_targets: bool = False
    use_day_only: bool = False
    use_work_period: bool = False
    show_shift_colors: bool = True


@dataclass(slots=True)
class ScheduleRequest:
    year: int
    month: int
    workers: list[WorkerInput]
    settings: ScheduleSettings = field(default_factory=ScheduleSettings)
    random_seed: int = 0


@dataclass(slots=True)
class WorkerScheduleRow:
    name: str
    days: list[str]
    raw_days: list[int | str]


@dataclass(slots=True)
class ScheduleResult:
    year: int
    month: int
    days_in_month: int
    status: str
    rows: list[WorkerScheduleRow]
    day_counts: list[int]
    night_counts: list[int]
    emergency_days: list[int] = field(default_factory=list)
    double_night_cycle_days: list[dict[str, object]] = field(default_factory=list)
    double_night_cycle_used: bool = False
    long_off_streak_fallback_used: bool = False
    long_off_streaks: list[dict[str, object]] = field(default_factory=list)
    repair_changed_count: int = 0
    repair_changed_cells: list[dict[str, object]] = field(default_factory=list)


@dataclass(slots=True)
class RepairEdit:
    worker_index: int
    day: int
    shift: ShiftCode | str


@dataclass(slots=True)
class ScheduleRepairRequest:
    request: ScheduleRequest
    result: ScheduleResult
    edits: list[RepairEdit]


@dataclass(slots=True)
class LeaveNeedGuide:
    year: int
    month: int
    days_in_month: int
    total_target_hours: int
    leave_credit_hours: int
    special_credit_hours: int
    non_to_credit_hours: int
    remaining_regular_target_hours: int
    max_regular_capacity_hours: int
    shortage_hours: int
    suggested_leave_days: int
    uses_individual_targets: bool
    uses_emergency_range: bool
    capacity_basis_label: str
    message: str


@dataclass(slots=True)
class TemplateWorkerLoadResult:
    workers: list[WorkerInput]
    recognized_count: int
    warning: str | None = None
    template_info: dict[str, object] = field(default_factory=dict)
    template_status: str = "valid_with_names"
    preserve_existing_workers: bool = False


@dataclass(slots=True)
class ExcelExportResult:
    output_path: str
    filename: str
    worker_count: int
    days_in_month: int
