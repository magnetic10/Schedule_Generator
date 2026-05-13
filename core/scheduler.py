from __future__ import annotations

from typing import Any, Dict, List, Tuple

from ortools.sat.python import cp_model

from .config import (
	MAX_DAY_WORKERS,
	MAX_NIGHT_WORKERS,
	MIN_DAY_WORKERS,
	MIN_NIGHT_WORKERS,
	SHIFT_DAY,
	SHIFT_LEAVE,
	SHIFT_NIGHT,
	SHIFT_OFF,
	SHIFT_OFF_NIGHT,
	TARGET_DAY_WORKERS,
	TARGET_NIGHT_WORKERS,
)
from .models import Worker


class SchedulerError(Exception):
	pass


class _SolveAttemptInfeasible(Exception):
	def __init__(self, status: int):
		super().__init__(status)
		self.status = status


EMERGENCY_DAY_PENALTY = 1_000_000_000
FIVE_REST_STREAK_PENALTY = 100_000_000
DOUBLE_NIGHT_CYCLE_PENALTY = 10_000_000
REPAIR_CHANGE_PENALTY = 1_000_000
TO_TARGET_DEVIATION_PENALTY = 5_000
FOUR_REST_STREAK_PENALTY = 2_500
THREE_REST_STREAK_PENALTY = 250
DAY_WORKLOAD_GAP_PENALTY = 30
PREFERENCE_MISMATCH_PENALTY = 2
SOLVER_MAX_TIME_SECONDS = 40.0


def _cp_status_name(status: int) -> str:
	return {
		cp_model.OPTIMAL: "OPTIMAL",
		cp_model.FEASIBLE: "FEASIBLE",
		cp_model.INFEASIBLE: "INFEASIBLE",
		cp_model.MODEL_INVALID: "MODEL_INVALID",
		cp_model.UNKNOWN: "UNKNOWN",
	}.get(status, str(status))


def _is_worker_active_on_day(worker: Worker, day: int) -> bool:
	if worker.start_day is not None and day < worker.start_day:
		return False
	if worker.end_day is not None and day > worker.end_day:
		return False
	return True


def _worker_dedicated_shift(worker: Worker) -> str:
	value = str(getattr(worker, "dedicated_shift", "") or "").strip().lower()
	if not value and getattr(worker, "is_only_day", False):
		value = "day"
	return value if value in ("day", "night") else ""


REST_STREAK_SHIFTS = {SHIFT_OFF, SHIFT_LEAVE}
SHIFT_LABELS = {
	SHIFT_DAY: "주간",
	SHIFT_NIGHT: "야간",
	SHIFT_OFF_NIGHT: "비번",
	SHIFT_OFF: "휴무",
	SHIFT_LEAVE: "연가",
}


def _shift_label(value: int | str | None) -> str:
	if value is None:
		return "미지정"
	return SHIFT_LABELS.get(value, str(value))


def _is_rest_streak_shift(value: int | str | None) -> bool:
	return value in REST_STREAK_SHIFTS


def _fixed_rest_runs(forced: Dict[int, int | str], num_days: int) -> List[Tuple[int, int]]:
	runs: List[Tuple[int, int]] = []
	start: int | None = None

	for day in range(num_days):
		if _is_rest_streak_shift(forced.get(day)):
			if start is None:
				start = day
			continue

		if start is not None:
			runs.append((start, day - 1))
			start = None

	if start is not None:
		runs.append((start, num_days - 1))

	return runs


def _rest_day_indicator(
	model: cp_model.CpModel,
	x: dict,
	w_idx: int,
	day: int,
) -> cp_model.IntVar:
	is_rest = model.NewBoolVar(f"rest_w{w_idx}_d{day}")
	rest_sum = x[w_idx, day, SHIFT_OFF] + x[w_idx, day, SHIFT_LEAVE]
	model.Add(rest_sum == 1).OnlyEnforceIf(is_rest)
	model.Add(rest_sum == 0).OnlyEnforceIf(is_rest.Not())
	return is_rest


def _rest_window_indicator(
	model: cp_model.CpModel,
	x: dict,
	w_idx: int,
	start: int,
	length: int,
) -> cp_model.IntVar:
	window = range(start, start + length)
	all_rest = model.NewBoolVar(f"rest_streak_w{w_idx}_d{start}_len{length}")
	rest_days = [
		_rest_day_indicator(model, x, w_idx, day)
		for day in window
	]
	for rest_day in rest_days:
		model.AddImplication(all_rest, rest_day)
	model.AddBoolOr([rest_day.Not() for rest_day in rest_days] + [all_rest])
	return all_rest


def _add_rest_streak_policy(
	model: cp_model.CpModel,
	x: dict,
	workers: List[Worker],
	forced_by_worker: List[Dict[int, int | str]],
	num_days: int,
	objective_terms: list,
	allow_auto_five_off_streaks: bool,
) -> None:
	for w_idx, worker in enumerate(workers):
		forced = forced_by_worker[w_idx]

		for run_start, run_end in _fixed_rest_runs(forced, num_days):
			if run_end - run_start + 1 < 2:
				continue
			for adjacent_day in (run_start - 1, run_end + 1):
				if adjacent_day < 0 or adjacent_day >= num_days:
					continue
				if not _is_worker_active_on_day(worker, adjacent_day):
					continue
				if _is_rest_streak_shift(forced.get(adjacent_day)):
					continue
				model.Add(x[w_idx, adjacent_day, SHIFT_OFF] == 0)

		for length, penalty in (
			(3, THREE_REST_STREAK_PENALTY),
			(4, FOUR_REST_STREAK_PENALTY),
			(5, FIVE_REST_STREAK_PENALTY),
		):
			if num_days < length:
				continue
			for start in range(0, num_days - length + 1):
				window = range(start, start + length)
				if any(not _is_worker_active_on_day(worker, day) for day in window):
					continue
				if all(_is_rest_streak_shift(forced.get(day)) for day in window):
					continue

				all_rest = _rest_window_indicator(model, x, w_idx, start, length)
				if length == 5 and not allow_auto_five_off_streaks:
					model.Add(all_rest == 0)
				else:
					objective_terms.append(all_rest * penalty)


def _build_to_config(to_cfg: dict | None) -> dict:
	cfg = to_cfg or {}
	allow_min_d = int(cfg.get("allow_min_d", MIN_DAY_WORKERS))
	allow_max_d = int(cfg.get("allow_max_d", MAX_DAY_WORKERS))
	allow_min_n = int(cfg.get("allow_min_n", MIN_NIGHT_WORKERS))
	allow_max_n = int(cfg.get("allow_max_n", MAX_NIGHT_WORKERS))
	return {
		"tgt_d": int(cfg.get("tgt_d", TARGET_DAY_WORKERS)),
		"tgt_n": int(cfg.get("tgt_n", TARGET_NIGHT_WORKERS)),
		"allow_min_d": allow_min_d,
		"allow_max_d": allow_max_d,
		"allow_min_n": allow_min_n,
		"allow_max_n": allow_max_n,
		"use_emergency_range": bool(cfg.get("use_emergency_range", False)),
		"em_min_d": int(cfg.get("em_min_d", allow_min_d)),
		"em_max_d": int(cfg.get("em_max_d", allow_max_d)),
		"em_min_n": int(cfg.get("em_min_n", allow_min_n)),
		"em_max_n": int(cfg.get("em_max_n", allow_max_n)),
		"special_shifts": dict(cfg.get("special_shifts", {})),
		"use_preference": bool(cfg.get("use_preference", False)),
		"allow_leave_after_off_night": bool(cfg.get("allow_leave_after_off_night", False)),
		"allow_double_night_cycle": bool(cfg.get("allow_double_night_cycle", False)),
	}


def _validate_day(day: int, num_days: int, worker_name: str, field_name: str) -> None:
	if day < 0 or day >= num_days:
		raise SchedulerError(
			f"{worker_name}: {field_name}에 범위를 벗어난 날짜 인덱스가 있습니다 ({day + 1}일)."
		)


def _collect_forced_assignments(worker: Worker, num_days: int) -> Dict[int, int | str]:
	forced: Dict[int, int | str] = {}

	def set_forced(day: int, value: int | str, field_name: str) -> None:
		_validate_day(day, num_days, worker.name, field_name)
		if day in forced and forced[day] != value:
			raise SchedulerError(
				f"{worker.name}: {day + 1}일에 서로 다른 고정 근무가 중복 지정되었습니다."
			)
		forced[day] = value

	for d in worker.fixed_day_days:
		set_forced(int(d), SHIFT_DAY, "fixed_day_days")
	for d in worker.fixed_night_days:
		set_forced(int(d), SHIFT_NIGHT, "fixed_night_days")
	for d in worker.fixed_off_night_days:
		set_forced(int(d), SHIFT_OFF_NIGHT, "fixed_off_night_days")
	for d in worker.fixed_off_days:
		set_forced(int(d), SHIFT_OFF, "fixed_off_days")
	for d in worker.fixed_leave_days:
		set_forced(int(d), SHIFT_LEAVE, "fixed_leave_days")

	for abbr, days in worker.custom_shifts.items():
		for d in days:
			set_forced(int(d), str(abbr), f"custom_shifts[{abbr}]")

	return forced


def diagnose_infeasibility(workers: List[Worker], num_days: int, cfg: dict) -> str | None:
    """스케줄 생성이 불가능한 원인을 분석하여 설명 메시지를 반환한다."""
    min_d, min_n = _diagnostic_min_staffing(cfg)
    basis_label = "예외 범위" if bool(cfg.get("use_emergency_range", False)) else "기본 범위"
    reasons = []

    # 1. 일자별 최소 인원 충족 여부 (가장 흔한 원인)
    for d in range(num_days):
        stats = _daily_capacity_stats(workers, num_days, d)
        fixed_day = stats["fixed_day"]
        fixed_night = stats["fixed_night"]
        fixed_off_night = stats["fixed_off_night"]
        required_off_night = _required_off_night_for_day(workers, num_days, d, min_n)
        day_need = max(0, min_d - fixed_day)
        night_need = max(0, min_n - fixed_night)
        off_night_need = max(0, required_off_night - fixed_off_night)
        day_only = stats["day_only"]
        night_only = stats["night_only"]
        both = stats["both"]
        day_possible = fixed_day + day_only + both
        night_possible = fixed_night + night_only + both
        off_night_possible = fixed_off_night + night_only + both
        combined_possible = fixed_day + fixed_night + fixed_off_night + day_only + night_only + both
        combined_needed = min_d + min_n + required_off_night
        day_reasons = []

        if day_possible < min_d:
            day_reasons.append(f"주간 최소 {min_d}명 중 최대 {day_possible}명만 가능")
        if night_possible < min_n:
            day_reasons.append(f"야간 최소 {min_n}명 중 최대 {night_possible}명만 가능")
        if off_night_possible < required_off_night:
            day_reasons.append(f"전날 야간에 따른 비번 최소 {required_off_night}명 중 최대 {off_night_possible}명만 가능")
        if night_need + off_night_need > night_only + both:
            day_reasons.append(
                f"야간/비번 필요 {night_need + off_night_need}명 중 최대 {night_only + both}명만 가능"
            )
        if day_need + night_need + off_night_need > day_only + night_only + both:
            day_reasons.append(
                f"남은 주/야/비 필요 {day_need + night_need + off_night_need}명 중 최대 {day_only + night_only + both}명만 가능"
            )
        if combined_possible < combined_needed:
            day_reasons.append(f"주간+야간+비번 최소 {combined_needed}명 중 최대 {combined_possible}명만 가능")

        if day_reasons:
            reasons.append(
                f"● {d+1}일: {basis_label} 기준으로 {', '.join(day_reasons)}합니다. "
                f"{_unavailable_summary(stats['unavailable'])} "
                "휴무/연가/비번/기타 근무 입력이나 기본 TO 설정을 조정해 주세요."
            )

    # 2. 고정 근무 간의 규칙 위반 (사용자 입력 오류)
    for w in workers:
        forced = _collect_forced_assignments(w, num_days)
        allow_leave_after_off_night = bool(cfg.get("allow_leave_after_off_night", False))
        allow_double_night_cycle = bool(cfg.get("allow_double_night_cycle", False))
        worker_label = w.name or "이름 없는 근무자"

        # 야-비-휴 규칙과 고정 근무 충돌 체크
        for d in range(num_days - 1):
            current = forced.get(d)
            next_shift = forced.get(d + 1)
            if current == SHIFT_NIGHT and next_shift is not None and next_shift != SHIFT_OFF_NIGHT:
                reasons.append(
                    f"● {worker_label}: {d+1}일 야간 다음 {d+2}일은 비번이어야 하는데 "
                    f"{_shift_label(next_shift)}이 지정되었습니다."
                )

            if current == SHIFT_OFF_NIGHT and next_shift is not None:
                allowed_next = {SHIFT_OFF}
                if allow_leave_after_off_night:
                    allowed_next.add(SHIFT_LEAVE)
                can_double_cycle = allow_double_night_cycle and d >= 1 and d + 4 < num_days
                if can_double_cycle:
                    allowed_next.add(SHIFT_NIGHT)
                if next_shift not in allowed_next:
                    allowed_text = "휴무"
                    if allow_leave_after_off_night:
                        allowed_text += " 또는 연가"
                    if can_double_cycle:
                        allowed_text += " 또는 야간"
                    reasons.append(
                        f"● {worker_label}: {d+1}일 비번 다음 {d+2}일은 {allowed_text}이어야 하는데 "
                        f"{_shift_label(next_shift)}이 지정되었습니다."
                    )

        for d in range(1, num_days):
            current = forced.get(d)
            prev_shift = forced.get(d - 1)
            if current == SHIFT_OFF_NIGHT and prev_shift is not None and prev_shift != SHIFT_NIGHT:
                reasons.append(
                    f"● {worker_label}: {d+1}일 비번은 전날 {d}일이 야간이어야 하는데 "
                    f"{_shift_label(prev_shift)}이 지정되었습니다."
                )
        
        # 전담 근무자 고정 근무 충돌 체크
        dedicated_shift = _worker_dedicated_shift(w)
        if dedicated_shift == "day":
            for d, val in forced.items():
                if val in (SHIFT_NIGHT, SHIFT_OFF_NIGHT):
                    reasons.append(f"● {worker_label}: 주간 전담인데 {d+1}일에 야간/비번 근무가 지정되었습니다.")
        elif dedicated_shift == "night":
            for d, val in forced.items():
                if val == SHIFT_DAY:
                    reasons.append(f"● {worker_label}: 야간 전담인데 {d+1}일에 주간 근무가 지정되었습니다.")

        # 근무 기간 외 근무 지정 체크
        if w.start_day is not None:
            for d in range(w.start_day):
                if forced.get(d) in [SHIFT_DAY, SHIFT_NIGHT]:
                    reasons.append(f"● {worker_label}: 근무 시작일({w.start_day+1}일) 이전인 {d+1}일에 근무가 지정되었습니다.")
        if w.end_day is not None:
            for d in range(w.end_day + 1, num_days):
                if forced.get(d) in [SHIFT_DAY, SHIFT_NIGHT]:
                    reasons.append(f"● {worker_label}: 근무 종료일({w.end_day+1}일) 이후인 {d+1}일에 근무가 지정되었습니다.")

    return "\n".join(reasons) if reasons else None


def _diagnostic_min_staffing(cfg: dict) -> tuple[int, int]:
    if bool(cfg.get("use_emergency_range", False)):
        return int(cfg.get("em_min_d", cfg["allow_min_d"])), int(cfg.get("em_min_n", cfg["allow_min_n"]))
    return int(cfg["allow_min_d"]), int(cfg["allow_min_n"])


def _daily_capacity_stats(workers: List[Worker], num_days: int, day: int) -> dict[str, object]:
    stats = {
        "fixed_day": 0,
        "fixed_night": 0,
        "fixed_off_night": 0,
        "day_only": 0,
        "night_only": 0,
        "both": 0,
        "unavailable": [],
    }

    for index, worker in enumerate(workers):
        label = worker.name or f"근무자{index + 1}"
        forced = _collect_forced_assignments(worker, num_days)
        current = forced.get(day)
        previous = forced.get(day - 1) if day > 0 else None

        if not _is_worker_active_on_day(worker, day):
            stats["unavailable"].append(f"{label}(기간 외)")
            continue

        forced_off_night = (day == 0 and worker.prev_month_last_day_night) or previous == SHIFT_NIGHT
        forced_off_after_off_night = (
            day == 1 and worker.prev_month_last_day_night
        ) or previous == SHIFT_OFF_NIGHT

        if current == SHIFT_DAY:
            stats["fixed_day"] += 1
            continue
        if current == SHIFT_NIGHT:
            stats["fixed_night"] += 1
            continue
        if current == SHIFT_OFF_NIGHT or forced_off_night:
            stats["fixed_off_night"] += 1
            stats["unavailable"].append(f"{label}(비번)")
            continue
        if current == SHIFT_OFF or forced_off_after_off_night:
            stats["unavailable"].append(f"{label}(휴무)")
            continue
        if current == SHIFT_LEAVE:
            stats["unavailable"].append(f"{label}(연가)")
            continue
        if isinstance(current, str):
            stats["unavailable"].append(f"{label}(기타 근무)")
            continue

        dedicated_shift = _worker_dedicated_shift(worker)
        if dedicated_shift == "day":
            stats["day_only"] += 1
        elif dedicated_shift == "night":
            stats["night_only"] += 1
        else:
            stats["both"] += 1

    return stats


def _required_off_night_for_day(workers: List[Worker], num_days: int, day: int, min_night: int) -> int:
    if day <= 0:
        return sum(1 for worker in workers if worker.prev_month_last_day_night)

    previous_day = day - 1
    fixed_previous_night = 0
    for worker in workers:
        if not _is_worker_active_on_day(worker, previous_day):
            continue
        forced = _collect_forced_assignments(worker, num_days)
        if forced.get(previous_day) == SHIFT_NIGHT:
            fixed_previous_night += 1
    return max(0, min_night, fixed_previous_night)


def _unavailable_summary(unavailable: list[str]) -> str:
    if not unavailable:
        return "배치 제외 인원은 없습니다."
    shown = ", ".join(unavailable[:5])
    suffix = f" 외 {len(unavailable) - 5}명" if len(unavailable) > 5 else ""
    return f"배치 제외 인원: {shown}{suffix}."


def _raise_schedule_failure(
	workers: List[Worker],
	num_days: int,
	cfg: dict,
	status: int,
) -> None:
	diag_msg = diagnose_infeasibility(workers, num_days, cfg)
	if diag_msg:
		raise SchedulerError(f"스케줄 생성이 불가능합니다.\n\n[분석 결과]\n{diag_msg}")

	if status == cp_model.INFEASIBLE:
		raise SchedulerError("현재 조건으로는 스케줄 생성이 불가능합니다. 기본 TO 설정이나 고정 근무를 완화해 주세요.")
	raise SchedulerError(f"제한 시간 안에 해를 찾지 못했습니다. solver 상태: {_cp_status_name(status)}")


def _validate_solve_inputs(workers: List[Worker], num_days: int, cfg: dict) -> None:
	if not workers:
		raise SchedulerError("근무자 목록이 비어 있습니다.")
	if num_days <= 0:
		raise SchedulerError("일 수가 올바르지 않습니다.")

	allow_min_d, allow_max_d = cfg["allow_min_d"], cfg["allow_max_d"]
	allow_min_n, allow_max_n = cfg["allow_min_n"], cfg["allow_max_n"]
	use_emergency_range = cfg["use_emergency_range"]
	em_min_d, em_max_d = cfg["em_min_d"], cfg["em_max_d"]
	em_min_n, em_max_n = cfg["em_min_n"], cfg["em_max_n"]
	tgt_d, tgt_n = cfg["tgt_d"], cfg["tgt_n"]

	if allow_min_d > allow_max_d or allow_min_n > allow_max_n:
		raise SchedulerError("허용 인원 범위(min/max)가 올바르지 않습니다.")

	if em_min_d > em_max_d or em_min_n > em_max_n:
		raise SchedulerError("예외 인원 범위(min/max)가 올바르지 않습니다.")

	if use_emergency_range and (
		em_min_d > allow_min_d
		or em_max_d < allow_max_d
		or em_min_n > allow_min_n
		or em_max_n < allow_max_n
	):
		raise SchedulerError("예외 범위는 기본 허용 범위를 포함해야 합니다.")

	if tgt_d < allow_min_d or tgt_d > allow_max_d or tgt_n < allow_min_n or tgt_n > allow_max_n:
		raise SchedulerError("목표 인원이 허용 범위를 벗어났습니다. 기본 TO 설정을 확인해 주세요.")


def _attempt_plans(cfg: dict) -> List[dict[str, bool]]:
	allow_double = bool(cfg.get("allow_double_night_cycle", False))
	use_emergency = bool(cfg.get("use_emergency_range", False))
	plans: List[dict[str, bool]] = []

	for allow_emergency_range in ([False, True] if use_emergency else [False]):
		for allow_auto_five_off_streaks in (False, True):
			for allow_double_night_cycle in ([False, True] if allow_double else [False]):
				plans.append(
					{
						"allow_emergency_range": allow_emergency_range,
						"allow_auto_five_off_streaks": allow_auto_five_off_streaks,
						"allow_double_night_cycle": allow_double_night_cycle,
					}
				)

	return plans


def _solve_with_optional_fallback(
	workers: List[Worker],
	num_days: int,
	cfg: dict,
	random_seed: int = 0,
	repair_context: dict[str, Any] | None = None,
) -> Tuple[List[List[int | str]], str, dict[str, Any]]:
	first_status: int | None = None
	last_status: int | None = None
	for plan in _attempt_plans(cfg):
		try:
			schedule, status_str, info = _solve_schedule_attempt(
				workers,
				num_days,
				cfg,
				random_seed,
				allow_auto_five_off_streaks=plan["allow_auto_five_off_streaks"],
				allow_double_night_cycle=plan["allow_double_night_cycle"],
				allow_emergency_range=plan["allow_emergency_range"],
				repair_context=repair_context,
			)
		except _SolveAttemptInfeasible as error:
			last_status = error.status
			if first_status is None:
				first_status = error.status
			if error.status != cp_model.INFEASIBLE:
				_raise_schedule_failure(workers, num_days, cfg, error.status)
			continue

		info["long_off_streak_fallback_used"] = bool(plan["allow_auto_five_off_streaks"])
		if plan["allow_auto_five_off_streaks"] and first_status is not None:
			info["first_without_long_off_status"] = _cp_status_name(first_status)
		info["double_night_cycle_fallback_used"] = bool(plan["allow_double_night_cycle"])
		info["emergency_range_enabled_in_attempt"] = bool(plan["allow_emergency_range"])
		return schedule, status_str, info

	_raise_schedule_failure(workers, num_days, cfg, last_status or cp_model.UNKNOWN)


def solve_schedule(
	workers: List[Worker],
	num_days: int,
	to_cfg: dict | None = None,
	random_seed: int = 0,
) -> Tuple[List[List[int | str]], str, dict[str, Any]]:
	cfg = _build_to_config(to_cfg)
	_validate_solve_inputs(workers, num_days, cfg)
	return _solve_with_optional_fallback(workers, num_days, cfg, random_seed)


def repair_schedule(
	workers: List[Worker],
	num_days: int,
	to_cfg: dict | None,
	baseline_schedule: List[List[int | str]],
	locked_assignments: Dict[Tuple[int, int], int | str],
	random_seed: int = 0,
) -> Tuple[List[List[int | str]], str, dict[str, Any]]:
	cfg = _build_to_config(to_cfg)
	_validate_solve_inputs(workers, num_days, cfg)
	if len(baseline_schedule) != len(workers):
		raise SchedulerError("부분 재생성 기준 결과의 근무자 수가 현재 입력과 다릅니다.")
	for row in baseline_schedule:
		if len(row) != num_days:
			raise SchedulerError("부분 재생성 기준 결과의 날짜 수가 현재 월과 다릅니다.")
	for (worker_index, day), _value in locked_assignments.items():
		if worker_index < 0 or worker_index >= len(workers) or day < 0 or day >= num_days:
			raise SchedulerError("부분 편집한 칸의 위치가 현재 근무표 범위를 벗어났습니다.")

	repair_context = {
		"baseline_schedule": baseline_schedule,
		"locked_assignments": locked_assignments,
	}
	schedule, status_str, info = _solve_with_optional_fallback(
		workers,
		num_days,
		cfg,
		random_seed,
		repair_context=repair_context,
	)
	info["repair_mode"] = True
	return schedule, status_str, info


def _solve_schedule_attempt(
	workers: List[Worker],
	num_days: int,
	cfg: dict,
	random_seed: int,
	allow_auto_five_off_streaks: bool,
	allow_double_night_cycle: bool,
	allow_emergency_range: bool,
	repair_context: dict[str, Any] | None = None,
) -> Tuple[List[List[int | str]], str, dict[str, Any]]:
	allow_min_d, allow_max_d = cfg["allow_min_d"], cfg["allow_max_d"]
	allow_min_n, allow_max_n = cfg["allow_min_n"], cfg["allow_max_n"]
	use_emergency_range = bool(allow_emergency_range and cfg["use_emergency_range"])
	em_min_d, em_max_d = cfg["em_min_d"], cfg["em_max_d"]
	em_min_n, em_max_n = cfg["em_min_n"], cfg["em_max_n"]
	tgt_d, tgt_n = cfg["tgt_d"], cfg["tgt_n"]
	special_shifts: Dict[str, int] = {
		str(k): int(v) for k, v in cfg["special_shifts"].items()
	}
	use_preference = cfg["use_preference"]
	allow_leave_after_off_night = cfg["allow_leave_after_off_night"]
	use_double_night_cycle = bool(allow_double_night_cycle and cfg["allow_double_night_cycle"])
	repair_context = repair_context or {}
	baseline_schedule = repair_context.get("baseline_schedule")
	locked_assignments: Dict[Tuple[int, int], int | str] = dict(
		repair_context.get("locked_assignments") or {}
	)

	model = cp_model.CpModel()
	assignable_shifts = [SHIFT_DAY, SHIFT_NIGHT, SHIFT_OFF_NIGHT, SHIFT_OFF]
	all_shifts = assignable_shifts + [SHIFT_LEAVE]

	forced_by_worker: List[Dict[int, int | str]] = []
	custom_credit_by_worker: List[int] = []

	for worker in workers:
		forced = _collect_forced_assignments(worker, num_days)
		w_idx = len(forced_by_worker)
		for (locked_worker, locked_day), locked_value in locked_assignments.items():
			if locked_worker != w_idx:
				continue
			if locked_value == "":
				raise SchedulerError("부분 편집 값이 비어 있습니다.")
			if locked_day in forced and forced[locked_day] != locked_value:
				raise SchedulerError(
					f"{worker.name or f'근무자{w_idx + 1}'}: {locked_day + 1}일은 입력표에서 이미 고정한 근무와 부분 편집 값이 충돌합니다."
				)
			forced[locked_day] = locked_value
		forced_by_worker.append(forced)

		credit = 0
		for d, val in forced.items():
			if isinstance(val, str):
				if val not in special_shifts:
					raise SchedulerError(
						f"{worker.name}: 기타 근무 '{val}'의 인정 시간이 설정에 없습니다."
					)
				credit += special_shifts[val]
		custom_credit_by_worker.append(credit)

	x = {}
	for w_idx in range(len(workers)):
		for d in range(num_days):
			for s in all_shifts:
				x[w_idx, d, s] = model.NewBoolVar(f"x_w{w_idx}_d{d}_s{s}")

	double_night_cycle_vars: list[tuple[int, int, cp_model.IntVar]] = []

	for w_idx, worker in enumerate(workers):
		forced = forced_by_worker[w_idx]
		for d in range(num_days):
			forced_value = forced.get(d)
			is_outside = False
			if worker.start_day is not None and d < worker.start_day:
				is_outside = True
			if worker.end_day is not None and d > worker.end_day:
				is_outside = True

			if is_outside:
				if forced_value is not None:
					raise SchedulerError(
						f"{worker.name or f'근무자{w_idx + 1}'}: 근무 기간 밖인 {d + 1}일에 고정 근무가 지정되었습니다."
					)
				for s in all_shifts:
					model.Add(x[w_idx, d, s] == 0)
				continue

			if isinstance(forced_value, str):
				for s in all_shifts:
					model.Add(x[w_idx, d, s] == 0)
			else:
				model.Add(sum(x[w_idx, d, s] for s in all_shifts) == 1)

				if isinstance(forced_value, int):
					model.Add(x[w_idx, d, forced_value] == 1)
				else:
					model.Add(x[w_idx, d, SHIFT_LEAVE] == 0)

			dedicated_shift = _worker_dedicated_shift(worker)
			if dedicated_shift == "day":
				model.Add(x[w_idx, d, SHIFT_NIGHT] == 0)
				model.Add(x[w_idx, d, SHIFT_OFF_NIGHT] == 0)
			elif dedicated_shift == "night":
				model.Add(x[w_idx, d, SHIFT_DAY] == 0)

		if worker.prev_month_last_day_night:
			if isinstance(forced.get(0), str):
				raise SchedulerError(
					f"{worker.name}: 지난달 말 야간 처리와 1일 기타 근무가 충돌합니다."
				)
			model.Add(x[w_idx, 0, SHIFT_OFF_NIGHT] == 1)

			if num_days >= 2:
				if isinstance(forced.get(1), str):
					raise SchedulerError(
						f"{worker.name}: 지난달 말 야간 처리와 2일 기타 근무가 충돌합니다."
					)
				model.Add(x[w_idx, 1, SHIFT_OFF] == 1)
		else:
			if forced.get(0) not in (SHIFT_OFF_NIGHT,) and not isinstance(forced.get(0), str):
				model.Add(x[w_idx, 0, SHIFT_OFF_NIGHT] == 0)

		for d in range(num_days - 1):
			model.AddBoolOr([
				x[w_idx, d, SHIFT_NIGHT].Not(),
				x[w_idx, d + 1, SHIFT_OFF_NIGHT],
			])

		for d in range(num_days - 1):
			allowed_after_off_night = [x[w_idx, d + 1, SHIFT_OFF]]
			if allow_leave_after_off_night:
				allowed_after_off_night.append(x[w_idx, d + 1, SHIFT_LEAVE])
			can_double_cycle = use_double_night_cycle and d >= 1 and d + 4 < num_days
			if can_double_cycle:
				allowed_after_off_night.append(x[w_idx, d + 1, SHIFT_NIGHT])
				double_cycle = model.NewBoolVar(f"double_night_cycle_w{w_idx}_d{d}")
				model.AddBoolAnd([
					x[w_idx, d, SHIFT_OFF_NIGHT],
					x[w_idx, d + 1, SHIFT_NIGHT],
				]).OnlyEnforceIf(double_cycle)
				model.AddBoolOr([
					x[w_idx, d, SHIFT_OFF_NIGHT].Not(),
					x[w_idx, d + 1, SHIFT_NIGHT].Not(),
					double_cycle,
				])
				model.Add(x[w_idx, d + 2, SHIFT_OFF_NIGHT] == 1).OnlyEnforceIf(double_cycle)
				for rest_day in (d + 3, d + 4):
					rest_options = [x[w_idx, rest_day, SHIFT_OFF]]
					if allow_leave_after_off_night:
						rest_options.append(x[w_idx, rest_day, SHIFT_LEAVE])
					model.Add(sum(rest_options) == 1).OnlyEnforceIf(double_cycle)
				double_night_cycle_vars.append((w_idx, d, double_cycle))
			model.AddBoolOr([
				x[w_idx, d, SHIFT_OFF_NIGHT].Not(),
				*allowed_after_off_night,
			])

		for d in range(1, num_days):
			model.AddBoolOr([
				x[w_idx, d, SHIFT_OFF_NIGHT].Not(),
				x[w_idx, d - 1, SHIFT_NIGHT],
			])

		for start in range(0, max(0, num_days - 5)):
			model.Add(
				sum(x[w_idx, day, SHIFT_DAY] for day in range(start, start + 6)) <= 5
			)

	day_counts = []
	night_counts = []
	emergency_used_by_day = []
	for d in range(num_days):
		day_cnt = sum(x[w_idx, d, SHIFT_DAY] for w_idx in range(len(workers)))
		night_cnt = sum(x[w_idx, d, SHIFT_NIGHT] for w_idx in range(len(workers)))

		if use_emergency_range:
			emergency_used = model.NewBoolVar(f"emergency_used_d{d}")
			model.Add(day_cnt >= allow_min_d).OnlyEnforceIf(emergency_used.Not())
			model.Add(day_cnt <= allow_max_d).OnlyEnforceIf(emergency_used.Not())
			model.Add(night_cnt >= allow_min_n).OnlyEnforceIf(emergency_used.Not())
			model.Add(night_cnt <= allow_max_n).OnlyEnforceIf(emergency_used.Not())
			model.Add(day_cnt >= em_min_d).OnlyEnforceIf(emergency_used)
			model.Add(day_cnt <= em_max_d).OnlyEnforceIf(emergency_used)
			model.Add(night_cnt >= em_min_n).OnlyEnforceIf(emergency_used)
			model.Add(night_cnt <= em_max_n).OnlyEnforceIf(emergency_used)
			emergency_used_by_day.append(emergency_used)
		else:
			model.Add(day_cnt >= allow_min_d)
			model.Add(day_cnt <= allow_max_d)
			model.Add(night_cnt >= allow_min_n)
			model.Add(night_cnt <= allow_max_n)

		day_counts.append(day_cnt)
		night_counts.append(night_cnt)

	objective_terms = []

	_add_rest_streak_policy(
		model,
		x,
		workers,
		forced_by_worker,
		num_days,
		objective_terms,
		allow_auto_five_off_streaks,
	)

	if baseline_schedule is not None:
		for w_idx, worker in enumerate(workers):
			for d in range(num_days):
				if not _is_worker_active_on_day(worker, d):
					continue
				baseline_value = baseline_schedule[w_idx][d]
				if baseline_value not in all_shifts:
					continue
				changed = model.NewBoolVar(f"repair_changed_w{w_idx}_d{d}")
				model.Add(x[w_idx, d, baseline_value] == 0).OnlyEnforceIf(changed)
				model.Add(x[w_idx, d, baseline_value] == 1).OnlyEnforceIf(changed.Not())
				objective_terms.append(changed * REPAIR_CHANGE_PENALTY)

	if use_emergency_range:
		for emergency_used in emergency_used_by_day:
			objective_terms.append(emergency_used * EMERGENCY_DAY_PENALTY)

	if use_double_night_cycle:
		for _w_idx, _day, double_cycle in double_night_cycle_vars:
			objective_terms.append(double_cycle * DOUBLE_NIGHT_CYCLE_PENALTY)

	for d in range(num_days):
		day_dev = model.NewIntVar(0, len(workers), f"day_dev_{d}")
		night_dev = model.NewIntVar(0, len(workers), f"night_dev_{d}")
		model.AddAbsEquality(day_dev, day_counts[d] - tgt_d)
		model.AddAbsEquality(night_dev, night_counts[d] - tgt_n)
		objective_terms.append(day_dev * TO_TARGET_DEVIATION_PENALTY)
		objective_terms.append(night_dev * TO_TARGET_DEVIATION_PENALTY)

	day_totals = []
	for w_idx, worker in enumerate(workers):
		day_total = model.NewIntVar(0, num_days, f"day_total_w{w_idx}")
		model.Add(day_total == sum(x[w_idx, d, SHIFT_DAY] for d in range(num_days)))
		day_totals.append(day_total)

		effective_hours = sum(
			8 * x[w_idx, d, SHIFT_DAY]
			+ 8 * x[w_idx, d, SHIFT_NIGHT]
			+ 8 * x[w_idx, d, SHIFT_OFF_NIGHT]
			+ 8 * x[w_idx, d, SHIFT_LEAVE]
			for d in range(num_days)
		) + custom_credit_by_worker[w_idx]
		raw_target_hours = getattr(worker, "target_hours", 160)
		target_hours = 160 if raw_target_hours is None else int(raw_target_hours)
		model.Add(effective_hours == target_hours)

		if use_preference:
			pref = str(getattr(worker, "shift_preference", "") or "").lower()
			if pref in ("day", "주간선호"):
				objective_terms.append(
					sum(x[w_idx, d, SHIFT_NIGHT] for d in range(num_days)) * PREFERENCE_MISMATCH_PENALTY
				)
			elif pref in ("night", "야간선호"):
				objective_terms.append(
					sum(x[w_idx, d, SHIFT_DAY] for d in range(num_days)) * PREFERENCE_MISMATCH_PENALTY
				)

	max_day = model.NewIntVar(0, num_days, "max_day_total")
	min_day = model.NewIntVar(0, num_days, "min_day_total")
	day_gap = model.NewIntVar(0, num_days, "day_gap")
	model.AddMaxEquality(max_day, day_totals)
	model.AddMinEquality(min_day, day_totals)
	model.Add(day_gap == max_day - min_day)
	objective_terms.append(day_gap * DAY_WORKLOAD_GAP_PENALTY)

	model.Minimize(sum(objective_terms))

	solver = cp_model.CpSolver()
	solver.parameters.max_time_in_seconds = SOLVER_MAX_TIME_SECONDS
	solver.parameters.random_seed = int(random_seed)
	solver.parameters.randomize_search = True
	solver.parameters.num_search_workers = 8

	status = solver.Solve(model)

	if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
		raise _SolveAttemptInfeasible(status)

	schedule: List[List[int | str]] = []
	for w_idx in range(len(workers)):
		forced = forced_by_worker[w_idx]
		row: List[int | str] = []
		for d in range(num_days):
			forced_value = forced.get(d)
			if (workers[w_idx].start_day is not None and d < workers[w_idx].start_day) or (
				workers[w_idx].end_day is not None and d > workers[w_idx].end_day
			):
				row.append("")
				continue
			if isinstance(forced_value, str):
				row.append(forced_value)
				continue

			assigned = SHIFT_OFF
			for s in all_shifts:
				if solver.Value(x[w_idx, d, s]) == 1:
					assigned = s
					break
			row.append(assigned)
		schedule.append(row)

	status_str = "OPTIMAL" if status == cp_model.OPTIMAL else "FEASIBLE"
	info: dict[str, Any] = {
		"emergency_days": [],
		"double_night_cycle_days": [],
		"double_night_cycle_used": False,
		"double_night_cycle_fallback_used": use_double_night_cycle,
		"double_night_cycle_policy": "enabled" if use_double_night_cycle else "disabled",
		"emergency_range_enabled_in_attempt": use_emergency_range,
		"long_off_streak_fallback_used": allow_auto_five_off_streaks,
		"long_off_streak_policy": "fallback" if allow_auto_five_off_streaks else "strict",
		"attempt_status": status_str,
	}
	if baseline_schedule is not None:
		changed_cells = []
		for w_idx, row in enumerate(schedule):
			for d, value in enumerate(row):
				if baseline_schedule[w_idx][d] != value:
					changed_cells.append(
						{
							"worker_index": w_idx,
							"day": d + 1,
							"from": baseline_schedule[w_idx][d],
							"to": value,
							"user_locked": (w_idx, d) in locked_assignments,
						}
					)
		info["repair_changed_cells"] = changed_cells
		info["repair_changed_count"] = len(changed_cells)
		info["repair_locked_count"] = len(locked_assignments)
	if use_emergency_range:
		info["emergency_days"] = [
			day + 1
			for day, emergency_used in enumerate(emergency_used_by_day)
			if solver.Value(emergency_used) == 1
		]
	if use_double_night_cycle:
		info["double_night_cycle_days"] = [
			{
				"worker_index": w_idx,
				"start_day": off_night_day,
				"end_day": off_night_day + 3,
			}
			for w_idx, off_night_day, double_cycle in double_night_cycle_vars
			if solver.Value(double_cycle) == 1
		]
		info["double_night_cycle_used"] = bool(info["double_night_cycle_days"])
	return schedule, status_str, info
