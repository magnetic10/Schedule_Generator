from __future__ import annotations

import os
import re
import sys
import uuid
import json
from dataclasses import asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.encoders import jsonable_encoder
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from core.scheduler import SchedulerError

from .schemas import (
    RepairEdit,
    ScheduleRequest,
    ScheduleRepairRequest,
    ScheduleResult,
    ScheduleSettings,
    WorkerInput,
    WorkerScheduleRow,
)
from .service import (
    build_default_worker_inputs,
    build_month_info,
    build_sidebar_defaults,
    default_excel_filename,
    export_schedule_result_to_excel,
    load_template_workers,
    reset_worker_schedules,
    solve_request,
    repair_request,
    build_leave_need_guide,
)


def _resource_root() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


BASE_DIR = _resource_root()
RUNTIME_DIR = _runtime_root()
WEB_DIR = BASE_DIR / "web"
STORAGE_DIR = RUNTIME_DIR / "storage"
UPLOAD_DIR = STORAGE_DIR / "uploads"
OUTPUT_DIR = STORAGE_DIR / "outputs"
DEBUG_DIR = STORAGE_DIR / "debug"
DEFAULT_TEMPLATE_CANDIDATES = (
    BASE_DIR / "templates" / "default_template.xlsx",
    RUNTIME_DIR / "templates" / "default_template.xlsx",
)

for directory in (UPLOAD_DIR, OUTPUT_DIR, DEBUG_DIR):
    directory.mkdir(parents=True, exist_ok=True)


app = FastAPI(title="근무표 생성기 V3")

if WEB_DIR.exists():
    app.mount("/web", StaticFiles(directory=str(WEB_DIR)), name="web")


@app.get("/")
def index() -> FileResponse:
    index_path = WEB_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="웹 UI 파일을 찾을 수 없습니다.")
    return FileResponse(index_path)


@app.get("/api/month-info")
def month_info(year: int, month: int) -> dict[str, Any]:
    return _to_dict(build_month_info(year, month))


@app.get("/api/default-state")
def default_state(year: int | None = None, month: int | None = None) -> dict[str, Any]:
    defaults = build_sidebar_defaults(year, month)
    workers = build_default_worker_inputs()
    return {
        "defaults": _to_dict(defaults),
        "workers": _to_dict(workers),
        "shift_options": _shift_options({}),
    }


@app.post("/api/reset-input")
def reset_input(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    workers = [_worker_input_from_dict(item) for item in payload.get("workers", [])]
    return {"workers": _to_dict(reset_worker_schedules(workers))}


@app.post("/api/reset-settings")
def reset_settings(payload: dict[str, Any] = Body(default={})) -> dict[str, Any]:
    year = payload.get("year")
    month = payload.get("month")
    defaults = build_sidebar_defaults(year, month) if year and month else build_sidebar_defaults()
    return {"defaults": _to_dict(defaults)}


@app.post("/api/leave-guide")
def leave_guide(payload: dict[str, Any] = Body(...)):
    try:
        req = _schedule_request_from_dict(payload)
        return {"success": True, "guide": _to_dict(build_leave_need_guide(req))}
    except (SchedulerError, ValueError) as exc:
        return _error_response(str(exc))


@app.post("/api/solve")
def solve(payload: dict[str, Any] = Body(...)):
    try:
        req = _schedule_request_from_dict(payload)
        result = solve_request(req)
        response = {"success": True, "result": _to_dict(result)}
        _write_solve_debug_log(payload, response)
        return response
    except (SchedulerError, ValueError) as exc:
        response = {"success": False, "error": str(exc)}
        _write_solve_debug_log(payload, response)
        return _error_response(str(exc))


@app.post("/api/repair")
def repair(payload: dict[str, Any] = Body(...)):
    try:
        req = _schedule_repair_request_from_dict(payload)
        result = repair_request(req)
        response = {"success": True, "result": _to_dict(result)}
        _write_solve_debug_log(payload, response)
        return response
    except (SchedulerError, ValueError) as exc:
        response = {"success": False, "error": str(exc)}
        _write_solve_debug_log(payload, response)
        return _error_response(str(exc))


@app.post("/api/upload-template")
async def upload_template(
    file: UploadFile = File(...),
    previous_file_id: str | None = Form(default=None),
) -> dict[str, Any]:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=".xlsx 파일만 업로드할 수 있습니다.")
    if not isinstance(previous_file_id, str):
        previous_file_id = None

    file_id = uuid.uuid4().hex
    target_path = UPLOAD_DIR / f"{file_id}.xlsx"
    contents = await file.read()
    _prune_uploads(keep_ids=[previous_file_id] if previous_file_id else [])
    _prune_outputs(max_files=12)
    try:
        target_path.write_bytes(contents)
    except OSError as exc:
        if previous_file_id:
            _delete_upload(previous_file_id)
        _prune_uploads()
        _prune_outputs(max_files=4)
        try:
            target_path.write_bytes(contents)
        except OSError:
            target_path.unlink(missing_ok=True)
            return _error_response(
                f"엑셀 파일을 임시 저장하지 못했습니다. 저장 공간 또는 파일 접근 권한을 확인해 주세요. ({exc})",
                status_code=507,
            )

    try:
        load_result = load_template_workers(str(target_path))
    except Exception as exc:
        target_path.unlink(missing_ok=True)
        return _error_response(
            f"엑셀 파일은 맞지만 근무표 서식으로 분석하지 못했습니다. 기존 근무표 서식 파일을 선택해 주세요. ({exc})",
            status_code=422,
        )

    if not load_result.template_info.get("date_row"):
        target_path.unlink(missing_ok=True)
        detail = f" {load_result.warning}" if load_result.warning else ""
        return _error_response(
            f"엑셀 파일은 맞지만 근무표 서식으로 인식하지 못했습니다. 날짜/근무자 영역을 찾을 수 없어 업로드를 추가하지 않았습니다.{detail}",
            status_code=422,
        )

    if previous_file_id:
        _delete_upload(previous_file_id)

    return {
        "file_id": file_id,
        "filename": file.filename,
        "load_result": _to_dict(load_result),
        "shift_options": _shift_options({}),
    }


@app.delete("/api/upload-template/{file_id}")
def delete_template(file_id: str) -> dict[str, Any]:
    deleted = _delete_upload(file_id)
    return {"success": True, "deleted": deleted}


@app.post("/api/export-excel")
def export_excel(payload: dict[str, Any] = Body(...)):
    template_id = str(payload.get("template_id") or "").strip()
    if template_id:
        template_path = _find_upload(template_id)
        if template_path is None:
            return _error_response("업로드된 엑셀 서식 파일을 찾을 수 없습니다.", status_code=404)
    else:
        template_path = _find_default_template()
        if template_path is None:
            return _error_response("기본 엑셀 서식 파일을 찾을 수 없습니다.", status_code=404)

    try:
        result = _schedule_result_from_dict(payload.get("result") or {})
        filename = default_excel_filename(result.year, result.month)
        output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{filename}"
        export_schedule_result_to_excel(
            template_path=str(template_path),
            output_path=str(output_path),
            result=result,
            apply_shift_colors=bool(payload.get("apply_shift_colors", False)),
        )
    except Exception as exc:
        return _error_response(str(exc), status_code=500)

    return FileResponse(str(output_path), filename=filename)


def _schedule_request_from_dict(data: dict[str, Any]) -> ScheduleRequest:
    return ScheduleRequest(
        year=_to_int(data.get("year"), "year"),
        month=_to_int(data.get("month"), "month"),
        workers=[_worker_input_from_dict(item) for item in data.get("workers", [])],
        settings=_settings_from_dict(data.get("settings") or {}),
        random_seed=int(data.get("random_seed", 0) or 0),
    )


def _schedule_repair_request_from_dict(data: dict[str, Any]) -> ScheduleRepairRequest:
    return ScheduleRepairRequest(
        request=_schedule_request_from_dict(data.get("request") or {}),
        result=_schedule_result_from_dict(data.get("result") or {}),
        edits=[
            RepairEdit(
                worker_index=_to_int(item.get("worker_index"), "worker_index"),
                day=_to_int(item.get("day"), "day"),
                shift=str(item.get("shift") or ""),
            )
            for item in data.get("edits", [])
        ],
    )


def _worker_input_from_dict(data: dict[str, Any]) -> WorkerInput:
    fixed_shifts_raw = data.get("fixed_shifts") or {}
    fixed_shifts = {
        int(day): str(shift)
        for day, shift in fixed_shifts_raw.items()
        if str(shift).strip()
    }

    preference = data.get("preference")
    if preference not in ("day", "night", None, ""):
        raise ValueError(f"알 수 없는 선호도 값입니다: {preference}")

    dedicated_shift = str(data.get("dedicated_shift") or "").strip()
    legacy_day_only = bool(data.get("is_day_only", False))
    if not dedicated_shift and legacy_day_only:
        dedicated_shift = "day"
    if dedicated_shift not in ("", "day", "night"):
        raise ValueError(f"알 수 없는 전담 근무 값입니다: {dedicated_shift}")

    return WorkerInput(
        name=str(data.get("name") or ""),
        start_day=_optional_int(data.get("start_day"), 1) or 1,
        end_day=_optional_int(data.get("end_day"), None),
        is_day_only=dedicated_shift == "day",
        dedicated_shift=dedicated_shift or None,
        target_hours=_optional_int(data.get("target_hours"), None),
        preference=preference or None,
        prev_month_last_day_night=bool(data.get("prev_month_last_day_night", False)),
        fixed_shifts=fixed_shifts,
    )


def _settings_from_dict(data: dict[str, Any]) -> ScheduleSettings:
    special_shifts = {
        str(key): int(value)
        for key, value in (data.get("special_shifts") or {}).items()
        if str(key).strip()
    }

    return ScheduleSettings(
        target_day=_optional_int(data.get("target_day"), 1) or 0,
        target_night=_optional_int(data.get("target_night"), 2) or 0,
        min_day=_optional_int(data.get("min_day"), 1) or 0,
        max_day=_optional_int(data.get("max_day"), 2) or 0,
        min_night=_optional_int(data.get("min_night"), 1) or 0,
        max_night=_optional_int(data.get("max_night"), 2) or 0,
        use_emergency_range=bool(data.get("use_emergency_range", False)),
        emergency_min_day=_optional_int(data.get("emergency_min_day"), None),
        emergency_max_day=_optional_int(data.get("emergency_max_day"), None),
        emergency_min_night=_optional_int(data.get("emergency_min_night"), None),
        emergency_max_night=_optional_int(data.get("emergency_max_night"), None),
        use_preference=bool(data.get("use_preference", False)),
        allow_leave_after_off_night=bool(data.get("allow_leave_after_off_night", False)),
        allow_double_night_cycle=bool(data.get("allow_double_night_cycle", False)),
        use_advanced_settings=bool(data.get("use_advanced_settings", False)),
        max_consecutive_day=max(1, min(31, _optional_int(data.get("max_consecutive_day"), 5) or 5)),
        max_consecutive_rest=max(1, min(31, _optional_int(data.get("max_consecutive_rest"), 4) or 4)),
        allow_user_forced_rule_violations=bool(data.get("allow_user_forced_rule_violations", False)),
        penalty_order=[
            str(value)
            for value in (data.get("penalty_order") or [])
            if str(value).strip()
        ],
        special_shifts=special_shifts,
    )


def _schedule_result_from_dict(data: dict[str, Any]) -> ScheduleResult:
    return ScheduleResult(
        year=_to_int(data.get("year"), "year"),
        month=_to_int(data.get("month"), "month"),
        days_in_month=_to_int(data.get("days_in_month"), "days_in_month"),
        status=str(data.get("status") or ""),
        rows=[
            WorkerScheduleRow(
                name=str(row.get("name") or ""),
                days=[str(value) for value in row.get("days", [])],
                raw_days=list(row.get("raw_days", [])),
            )
            for row in data.get("rows", [])
        ],
        day_counts=[int(value) for value in data.get("day_counts", [])],
        night_counts=[int(value) for value in data.get("night_counts", [])],
        emergency_days=[int(value) for value in data.get("emergency_days", [])],
        double_night_cycle_days=list(data.get("double_night_cycle_days", [])),
        double_night_cycle_used=bool(data.get("double_night_cycle_used", False)),
        long_off_streak_fallback_used=bool(data.get("long_off_streak_fallback_used", False)),
        long_off_streaks=list(data.get("long_off_streaks", [])),
        repair_changed_count=int(data.get("repair_changed_count", 0) or 0),
        repair_changed_cells=list(data.get("repair_changed_cells", [])),
    )


def _to_int(value: Any, label: str) -> int:
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} 값이 올바른 숫자가 아닙니다.") from exc


def _optional_int(value: Any, default: int | None) -> int | None:
    if value in (None, ""):
        return default
    return int(value)


def _to_dict(value: Any) -> Any:
    if is_dataclass(value):
        return jsonable_encoder(asdict(value))
    if isinstance(value, list):
        return [_to_dict(item) for item in value]
    if isinstance(value, dict):
        return {key: _to_dict(item) for key, item in value.items()}
    return jsonable_encoder(value)


def _safe_filename(filename: str) -> str:
    name = Path(filename).name
    return re.sub(r"[^0-9A-Za-z가-힣_.-]+", "_", name) or "template.xlsx"


def _upload_candidates(file_id: str) -> list[Path]:
    if not file_id:
        return []
    return [UPLOAD_DIR / f"{file_id}.xlsx", *UPLOAD_DIR.glob(f"{file_id}_*")]


def _find_upload(file_id: str) -> Path | None:
    for path in _upload_candidates(file_id):
        if path.is_file():
            return path
    return None


def _find_default_template() -> Path | None:
    for path in DEFAULT_TEMPLATE_CANDIDATES:
        if path.is_file():
            return path
    return None


def _delete_upload(file_id: str) -> bool:
    deleted = False
    for path in _upload_candidates(file_id):
        if path.is_file():
            path.unlink(missing_ok=True)
            deleted = True
    return deleted


def _prune_uploads(keep_ids: list[str] | None = None) -> None:
    keep_ids = [item for item in (keep_ids or []) if item]
    for path in UPLOAD_DIR.glob("*"):
        if not path.is_file():
            continue
        if any(path.name.startswith(file_id) for file_id in keep_ids):
            continue
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


def _prune_outputs(max_files: int = 12) -> None:
    files = sorted(
        [path for path in OUTPUT_DIR.glob("*") if path.is_file()],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for path in files[max_files:]:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


def _write_solve_debug_log(payload: dict[str, Any], response: dict[str, Any]) -> None:
    if os.environ.get("WORK_SCHEDULER_DEBUG_LOG", "").strip().lower() not in {"1", "true", "yes"}:
        return
    debug_payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "request": payload,
        "response": response,
    }
    try:
        (DEBUG_DIR / "last_solve.json").write_text(
            json.dumps(jsonable_encoder(debug_payload), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _shift_options(special_shifts: dict[str, int]) -> list[dict[str, str]]:
    base_options = [
        ("", ""),
        ("day", "주"),
        ("night", "야"),
        ("off_night", "비"),
        ("off", "휴"),
        ("leave", "연"),
    ]
    options = [{"value": value, "label": label} for value, label in base_options]
    options.extend({"value": key, "label": key} for key in special_shifts)
    return options


def _error_response(message: str, status_code: int = 422) -> JSONResponse:
    return JSONResponse(
        status_code=status_code,
        content={"success": False, "error": message},
    )
