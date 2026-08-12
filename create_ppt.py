#!/usr/bin/env python3
"""
Feature-wise PPT report generator.

Supported usage:
  python create_ppt.py <run_dir> [output.pptx]
  python create_ppt.py <feature_wise_log> <screenshots_dir> [output.pptx]
"""

from __future__ import annotations

import re
import sys
import subprocess
import os
from collections import Counter
from datetime import datetime
from pathlib import Path


DEFAULT_RUN_DIR = Path("logs/run_20260521_180712")
DEFAULT_LOG = DEFAULT_RUN_DIR / "feature_wise_log.txt"
DEFAULT_SS_DIR = DEFAULT_RUN_DIR / "screenshots"


def esc(s: str) -> str:
    return str(s).replace("\\", "\\\\").replace("'", "\\'").replace("\n", " ").replace("\r", "")


def parse_dt(ts: str):
    if not ts:
        return None
    try:
        return datetime.strptime(ts.strip(), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def read_optional(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def parse_pipe_fields(line: str) -> dict[str, str]:
    body = re.sub(r"^\[[^\]]+\]\s*", "", line or "").strip()
    fields: dict[str, str] = {}
    for part in re.split(r"\s+\|\s+", body):
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == "'" and value[-1] == "'":
            value = value[1:-1]
        if key:
            fields[key] = value.strip()
    return fields


def combine_status(current: str, new: str) -> str:
    current = (current or "").upper()
    new = (new or "").upper()
    if not current:
        return new
    if "FAIL" in (current, new):
        return "FAIL"
    if new == "PASS":
        return "PASS"
    return current


def normalize_result_status(value: str) -> str:
    status = (value or "").strip().upper().replace("_", " ")
    if not status:
        return ""
    if status in {"Y", "YES", "TRUE", "1", "ENABLED"}:
        return "ALREADY ENABLED"
    if status in {"PASS", "SUCCESS", "ENABLED NOW", "ALREADY ENABLED", "AUTO ENABLED"}:
        return status
    if status in {"FAIL", "FAILED", "ERROR", "NOT FOUND"}:
        return "FAIL" if status in {"FAIL", "FAILED", "ERROR"} else "NOT FOUND"
    return status


def action_has_optin(action: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", " ", str(action or "").casefold())
    return bool(re.search(r"\bopt\s*in\b|\boptin\b", normalized))


def action_has_profile(action: str) -> bool:
    return "profile" in str(action or "").casefold()


def action_has_ess(action: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", " ", str(action or "").casefold())
    return bool(re.search(r"\bess\b|\bjob\b", normalized))


def combine_result_status(current: str, new: str) -> str:
    current = normalize_result_status(current)
    new = normalize_result_status(new)
    if not current:
        return new
    if new in {"FAIL", "NOT FOUND"}:
        return new
    if current in {"FAIL", "NOT FOUND"}:
        return current
    if new in {"ENABLED NOW", "ALREADY ENABLED", "PASS", "SUCCESS", "AUTO ENABLED"}:
        return new
    return current


def extract_profile_codes(value: str) -> list[str]:
    codes = []
    for m in re.finditer(r"\b([A-Z][A-Z0-9_]{4,})\b(?=\s*=)", value or ""):
        code = m.group(1).strip()
        if code not in codes:
            codes.append(code)
    for m in re.finditer(r"\b(ORA_[A-Z0-9_]+)\b", value or ""):
        code = m.group(1).strip()
        if code not in codes:
            codes.append(code)
    return codes


def parse_profile_statuses(run_dir: Path) -> dict[str, str]:
    text = read_optional(run_dir / "profile_options_log.txt")
    statuses: dict[str, str] = {}
    for line in text.splitlines():
        status_m = re.search(r"\bStatus\s*=\s*(PASS|FAIL)\b", line, re.IGNORECASE)
        if not status_m:
            continue
        status = status_m.group(1).upper()
        code_m = re.search(r"ProfileOptionCode='([^']+)'", line)
        if not code_m:
            code_m = re.search(r">\s*([A-Z][A-Z0-9_]{4,})\s*=", line)
        if not code_m:
            code_m = re.search(r"\b(ORA_[A-Z0-9_]+)\b", line)
        if code_m:
            code = code_m.group(1).strip()
            statuses[code] = combine_status(statuses.get(code, ""), status)
    return statuses


def parse_feature_statuses_from_log(run_dir: Path, filename: str) -> dict[str, str]:
    text = read_optional(run_dir / filename)
    statuses: dict[str, str] = {}
    for line in text.splitlines():
        fields = parse_pipe_fields(line)
        status = normalize_result_status(fields.get("Status", ""))
        if not status:
            continue
        feature = fields.get("Feature", "")
        if not feature:
            feature_m = re.search(r"Feature\s*:\s*(.+?)(?:\s+\|\s+|$)", line)
            feature = feature_m.group(1).strip() if feature_m else ""
        if feature:
            statuses[feature] = combine_result_status(statuses.get(feature, ""), status)
    return statuses


def extract_ess_keys(value: str) -> list[str]:
    keys = []
    for raw in re.split(r"\s*,\s*", value or ""):
        raw = raw.strip()
        if not raw:
            continue
        parts = [p.strip() for p in raw.split(">") if p.strip()]
        for candidate in reversed(parts):
            if candidate and candidate not in keys:
                keys.append(candidate)
                break
    return keys


def parse_ess_statuses(run_dir: Path) -> dict[str, str]:
    text = read_optional(run_dir / "ess_jobs_log.txt")
    statuses: dict[str, str] = {}
    for line in text.splitlines():
        status_m = re.search(r"\bStatus\s*=\s*(PASS|FAIL)\b", line, re.IGNORECASE)
        if not status_m:
            continue
        status = status_m.group(1).upper()
        keys = []
        for key in ("ProcessName", "Process", "Job", "ESSJob", "Parameter", "Value"):
            keys.extend(m.group(1).strip() for m in re.finditer(rf"{key}='([^']+)'", line))
        if not keys:
            left = line.split("|", 1)[0].strip()
            if left and not left.startswith("="):
                keys.append(left)
        for key in keys:
            statuses[key] = combine_status(statuses.get(key, ""), status)
    return statuses


def enrich_from_companion_logs(run_dir: Path, features: list[dict]) -> None:
    profile_statuses = parse_profile_statuses(run_dir)
    optin_statuses = parse_feature_statuses_from_log(run_dir, "optin_log.txt")
    auto_statuses = parse_feature_statuses_from_log(run_dir, "auto_log.txt")
    ess_statuses = parse_ess_statuses(run_dir)

    for f in features:
        action = (f.get("action") or "").upper()
        profile_detail = f.get("profile_detail") or f.get("profile_output") or ""
        ess_detail = f.get("ess_detail") or f.get("ess_output") or ""

        if "AUTO" in action:
            auto_status = auto_statuses.get(f["name"], "PASS")
            if auto_status in {"PASS", "SUCCESS", "AUTO ENABLED", "ENABLED NOW", "ALREADY ENABLED"}:
                f["auto"] = "AUTO ENABLED"
                if f["optin"] == "N/A":
                    f["optin"] = "Already Enabled*"
            elif auto_status in {"FAIL", "NOT FOUND"}:
                f["auto"] = "FAIL"

        if action_has_optin(action) and f["name"] in optin_statuses:
            optin_status = optin_statuses[f["name"]]
            # Normalize enabled-like values before falling back to N/A; opt-in logs vary by source.
            if optin_status in {"ALREADY ENABLED", "PASS", "SUCCESS"}:
                f["optin"] = "Already Enabled"
            elif optin_status == "ENABLED NOW":
                f["optin"] = "Enabled Now"
            elif optin_status == "NOT FOUND":
                f["optin"] = "Not Found"
            elif optin_status == "FAIL":
                f["optin"] = "FAIL"

        if profile_detail and f["profile"] == "N/A":
            codes = extract_profile_codes(profile_detail)
            matched = [profile_statuses[c] for c in codes if c in profile_statuses]
            if matched and len(matched) == len(codes):
                f["profile"] = "FAIL" if "FAIL" in matched else "PASS"
            elif matched:
                f["profile"] = "FAIL" if "FAIL" in matched else "PASS"

        if ess_detail and f["ess"] == "N/A":
            keys = extract_ess_keys(ess_detail)
            matched = []
            for key in keys:
                for logged_key, status in ess_statuses.items():
                    if key.lower() in logged_key.lower() or logged_key.lower() in key.lower():
                        matched.append(status)
                        break
            if matched and len(matched) == len(keys):
                f["ess"] = "FAIL" if "FAIL" in matched else "PASS"
            elif matched:
                f["ess"] = "FAIL" if "FAIL" in matched else "PASS"


def collect_run_datetimes(run_dir: Path) -> list:
    times = []
    for filename in (
        "feature_wise_log.txt",
        "optin_log.txt",
        "profile_options_log.txt",
        "ess_jobs_log.txt",
        "auto_log.txt",
        "summary_log.txt",
    ):
        text = read_optional(run_dir / filename)
        for raw_ts in re.findall(r"\[([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2})\]", text):
            dt = parse_dt(raw_ts)
            if dt:
                times.append(dt)
        for raw_ts in re.findall(r"^(?:Started At|Completed At)\s*:\s*(.+)$", text, re.MULTILINE):
            dt = parse_dt(raw_ts)
            if dt:
                times.append(dt)
    return times


def parse_feature_wise_log(path: Path) -> dict:
    text = path.read_text(encoding="utf-8", errors="replace")

    run_folder = ""
    started_at = ""
    m_run = re.search(r"^Run Folder:\s*(.+)$", text, re.MULTILINE)
    m_start = re.search(r"^Started At:\s*(.+)$", text, re.MULTILINE)
    if m_run:
        run_folder = m_run.group(1).strip()
    if m_start:
        started_at = m_start.group(1).strip()

    latest = {}
    feature_base = {}
    feature_keys_by_name: dict[str, list[str]] = {}
    all_events = []
    input_sequence = 0

    for line in text.splitlines():
        if "Feature='" not in line:
            continue

        def ex(key: str) -> str:
            return fields.get(key, "")

        ts_m = re.search(r"^\[([^\]]+)\]", line)
        timestamp = ts_m.group(1).strip() if ts_m else ""
        ts_dt = parse_dt(timestamp)
        fields = parse_pipe_fields(line)

        rec = {
            "timestamp": timestamp,
            "timestamp_dt": ts_dt,
            "source_row": ex("SourceRow") or ex("Row"),
            "feature": ex("Feature"),
            "product": ex("Product"),
            "module": ex("Module"),
            "action": ex("ActionRequired") or ex("OptIn"),
            "stage": ex("Stage").upper(),
            "status": ex("Status").upper(),
            "optin_output": ex("OptInOutput"),
            "profile_output": ex("ProfileOptionOutput") or ex("ProfileOption"),
            "ess_output": ex("ESSJobsOutput") or ex("ESSJobs"),
            "error": ex("Error"),
            "recommendation": ex("Recommendation"),
        }
        if not rec["feature"]:
            continue

        all_events.append(rec)
        if rec["stage"] == "INPUT_VALUES":
            input_sequence += 1
            feature_key = f"row:{rec['source_row']}" if rec["source_row"] else f"input:{input_sequence}"
        else:
            existing_keys = feature_keys_by_name.get(rec["feature"], [])
            feature_key = existing_keys[0] if existing_keys else f"event:{rec['feature']}"

        if feature_key not in feature_base:
            feature_base[feature_key] = {
                "name": rec["feature"],
                "product": rec["product"],
                "module": rec["module"],
                "action": rec["action"],
                "optin": "N/A",
                "profile": "N/A",
                "ess": "N/A",
                "auto": "",
                "profile_detail": rec["profile_output"],
                "ess_detail": rec["ess_output"],
                "recommendation": "",
                "error": "",
            }
            feature_keys_by_name.setdefault(rec["feature"], []).append(feature_key)

        if rec["action"] and not feature_base[feature_key]["action"]:
            feature_base[feature_key]["action"] = rec["action"]
        if rec["profile_output"] and not feature_base[feature_key]["profile_detail"]:
            feature_base[feature_key]["profile_detail"] = rec["profile_output"]
        if rec["ess_output"] and not feature_base[feature_key]["ess_detail"]:
            feature_base[feature_key]["ess_detail"] = rec["ess_output"]

        target_keys = feature_keys_by_name.get(rec["feature"], [feature_key])
        for target_key in target_keys:
            latest[(target_key, rec["stage"])] = rec

    features = []
    for feature_key, base in feature_base.items():
        f = dict(base)
        opt = latest.get((feature_key, "OPT_IN"))
        prof = latest.get((feature_key, "PROFILE_OPTIONS"))
        ess = latest.get((feature_key, "ESS_JOBS"))
        auto = latest.get((feature_key, "AUTO"))

        if opt:
            out = opt["optin_output"].lower()
            if "already enabled" in out:
                f["optin"] = "Already Enabled"
            elif "enabled now" in out:
                f["optin"] = "Enabled Now"
            elif "not found" in out:
                f["optin"] = "Not Found"
            elif opt["status"] == "PASS":
                f["optin"] = "Enabled Now"
            else:
                f["optin"] = "FAIL"
            f["error"] = opt["error"] or f["error"]
            f["recommendation"] = opt["recommendation"] or f["recommendation"]

        if prof:
            f["profile"] = "PASS" if prof["status"] == "PASS" else "FAIL"
            f["error"] = prof["error"] or f["error"]
            f["recommendation"] = prof["recommendation"] or f["recommendation"]

        if ess:
            f["ess"] = "PASS" if ess["status"] == "PASS" else "FAIL"

        if auto and auto["status"] == "PASS":
            f["auto"] = "AUTO ENABLED"
            # Notation requested: AUTO PASS implies feature is already opted in.
            if f["optin"] == "N/A":
                f["optin"] = "Already Enabled*"

        features.append(f)

    enrich_from_companion_logs(path.parent, features)
    features.sort(key=lambda x: (x["module"], x["name"]))

    valid_times = [e["timestamp_dt"] for e in all_events if e["timestamp_dt"]]
    valid_times.extend(collect_run_datetimes(path.parent))
    first_event = min(valid_times) if valid_times else None
    last_event = max(valid_times) if valid_times else None
    runtime_text = "N/A"
    if first_event and last_event:
        secs = int((last_event - first_event).total_seconds())
        mm, ss = divmod(secs, 60)
        hh, mm = divmod(mm, 60)
        runtime_text = f"{hh:02d}:{mm:02d}:{ss:02d}"

    return {
        "run_id": Path(run_folder).name if run_folder else path.parent.name,
        "run_folder": run_folder,
        "started_at": started_at,
        "first_event": first_event.strftime("%Y-%m-%d %H:%M:%S") if first_event else "",
        "last_event": last_event.strftime("%Y-%m-%d %H:%M:%S") if last_event else "",
        "runtime_text": runtime_text,
        "features": features,
    }


def compute_overall_status(f: dict) -> str:
    if f["auto"] == "AUTO ENABLED":
        return "AUTO"
    if f["optin"] in ("Not Found", "FAIL") or f["profile"] == "FAIL" or f["ess"] == "FAIL":
        return "PARTIAL"
    if f["optin"].startswith("Already Enabled") or f["optin"] == "Enabled Now" or f["profile"] == "PASS" or f["ess"] == "PASS":
        return "SUCCESS"
    return "N/A"


def collect_screenshots(ss_dir: Path) -> list:
    if not ss_dir.exists():
        return []
    rows = []
    for f in sorted(ss_dir.iterdir()):
        if f.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            continue
        name = f.name.lower()
        if "_optin_" in name:
            cat = "optin"
        elif "_profile_" in name:
            cat = "profile"
        elif "_ess_" in name:
            cat = "ess"
        else:
            cat = "other"

        if "already_enabled" in name:
            status = "already_enabled"
        elif "enabled_now" in name:
            status = "enabled_now"
        elif "notfound" in name or "not_found" in name or "error" in name:
            status = "not_found"
        elif "_fail" in name:
            status = "fail"
        elif "_pass" in name:
            status = "pass"
        else:
            status = "info"

        label = re.sub(r"^\d{6}_", "", f.stem).replace("_", " ")
        rows.append({"path": str(f), "name": f.name, "cat": cat, "status": status, "label": label[:95]})
    return rows


def build_remediation(features: list) -> list:
    rows = []
    for f in features:
        issues = []
        fixes = []
        action = (f.get("action") or "").upper()

        if action_has_optin(action):
            if f["optin"] in ("Not Found", "FAIL"):
                issues.append(f"Opt-in status is {f['optin']}")
                fixes.append("Setup and Maintenance → Offerings → New Features → search feature → open Allows Opt-In → Edit Features → Done.")

        if action_has_profile(action):
            if f["profile"] != "PASS":
                issues.append("Profile option update pending/failed")
                fixes.append("Setup and Maintenance → Manage Administrator Profile Values → search profile code → set Value=Yes at Site → Save.")

        if action_has_ess(action):
            if f["ess"] != "PASS":
                issues.append("ESS scheduling pending/failed")
                fixes.append("Run required ESS process from Scheduled Processes and validate completion status.")

        if "AUTO" in action:
            if f["auto"] != "AUTO ENABLED":
                issues.append("Auto-enabled feature requires verification")
                fixes.append("Verify feature in UI release notes and feature behavior in application pages.")

        if issues:
            fallback = f["recommendation"] or "Re-run automation and validate all dependent setup tasks manually."
            rows.append({
                "name": f["name"],
                "action": f.get("action", ""),
                "issue": "; ".join(issues),
                "fix": " | ".join(fixes) if fixes else fallback,
            })
    return rows


def summarize_profile_option_entries(features: list[dict]) -> dict[str, int]:
    raw_count = 0
    statuses: dict[str, str] = {}
    for f in features:
        if not action_has_profile(f.get("action", "")):
            continue
        for code in extract_profile_codes(f.get("profile_detail", "") or f.get("profile_output", "")):
            raw_count += 1
            current = statuses.get(code, "")
            new = f.get("profile", "")
            statuses[code] = combine_status(current, new)

    return {
        "raw": raw_count,
        "dedup": len(statuses),
        "pass": sum(1 for status in statuses.values() if status == "PASS"),
        "fail": sum(1 for status in statuses.values() if status == "FAIL"),
    }


def extract_ess_entry_keys(value: str) -> list[str]:
    keys = []
    for raw in re.split(r"\s*,\s*", value or ""):
        raw = raw.strip()
        if not raw:
            continue
        parts = [p.strip() for p in raw.split(">") if p.strip()]
        if len(parts) >= 3:
            key = " > ".join(parts[:3])
        else:
            key = raw
        key_norm = re.sub(r"\s+", " ", key).casefold()
        if key_norm and key_norm not in keys:
            keys.append(key_norm)
    return keys


def summarize_ess_job_entries(features: list[dict]) -> dict[str, int]:
    raw_count = 0
    statuses: dict[str, str] = {}
    for f in features:
        if not action_has_ess(f.get("action", "")):
            continue
        for key in extract_ess_entry_keys(f.get("ess_detail", "") or f.get("ess_output", "")):
            raw_count += 1
            current = statuses.get(key, "")
            new = f.get("ess", "")
            statuses[key] = combine_status(current, new)

    return {
        "raw": raw_count,
        "dedup": len(statuses),
        "pass": sum(1 for status in statuses.values() if status == "PASS"),
        "fail": sum(1 for status in statuses.values() if status == "FAIL"),
    }


def build_js(output_pptx: Path, data: dict, features: list, screenshots: list, remediation: list) -> str:
    run_id = data["run_id"]
    started_at = data["started_at"]
    first_event = data["first_event"]
    last_event = data["last_event"]
    runtime_text = data["runtime_text"]

    total = len(features)
    optin_ok = sum(1 for f in features if f["optin"].startswith("Already Enabled") or f["optin"] == "Enabled Now")
    optin_fail = sum(1 for f in features if f["optin"] in ("Not Found", "FAIL"))
    profile_counts = summarize_profile_option_entries(features)
    ess_counts = summarize_ess_job_entries(features)
    print(
        "[PPT TRACE] Profile Options count raw="
        f"{profile_counts['raw']} dedup={profile_counts['dedup']} "
        f"pass={profile_counts['pass']} fail={profile_counts['fail']}"
    )
    print(
        "[PPT TRACE] ESS Jobs count raw="
        f"{ess_counts['raw']} dedup={ess_counts['dedup']} "
        f"pass={ess_counts['pass']} fail={ess_counts['fail']}"
    )
    profile_pass = profile_counts["pass"]
    profile_fail = profile_counts["fail"]
    ess_pass = ess_counts["pass"]
    ess_fail = ess_counts["fail"]
    auto_count = sum(1 for f in features if f["auto"] == "AUTO ENABLED")
    partial = sum(1 for f in features if compute_overall_status(f) == "PARTIAL")

    action_counter = Counter((f.get("action") or "Unknown") for f in features)

    lines = []
    L = lines.append
    L("const pptxgen = require('pptxgenjs');")
    L("const fs = require('fs');")
    L("const pres = new pptxgen();")
    L("pres.layout = 'LAYOUT_16x9';")
    L(f"pres.title = 'Feature Enablement Report – {esc(run_id)}';")
    L("const C={navy:'0D2137',teal:'0891B2',white:'FFFFFF',offwhite:'F8FAFC',red:'DC2626',green:'059669',orange:'D97706',purple:'7C3AED',gray:'94A3B8',border:'E2E8F0'};")
    L("function hdr(s,t,bg){s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.6,fill:{color:bg||C.navy}});s.addText(t,{x:0.35,y:0,w:9.3,h:0.6,fontSize:15.5,bold:true,color:C.white,valign:'middle',margin:0});}")
    L("function badge(s,x,y,w,h,t,fill){s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,rectRadius:0.05,fill:{color:fill}});s.addText(t,{x,y,w,h,fontSize:7,bold:true,color:'FFFFFF',align:'center',valign:'middle',margin:0});}")

    # Cover slide improved
    L("{const s=pres.addSlide();s.background={color:C.navy};")
    L("s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.8,w:10,h:1.825,fill:{color:'0B1D31'}});")
    L("s.addText('Feature Enablement',{x:0.55,y:0.78,w:7.0,h:0.85,fontSize:42,bold:true,color:C.white});")
    L("s.addText('Execution Summary',{x:0.55,y:1.6,w:7.0,h:0.8,fontSize:38,bold:true,color:C.teal});")
    L("s.addText('Order Management · Redwood Rollout',{x:0.55,y:2.4,w:7.0,h:0.35,fontSize:12,color:'7FB3CF',italic:true});")
    L(f"s.addText('Run ID: {esc(run_id)}',{{x:0.55,y:4.0,w:3.2,h:0.32,fontSize:10,color:'7FB3CF'}});")
    L(f"s.addText('First Event: {esc(first_event or started_at)}',{{x:0.55,y:4.35,w:4.8,h:0.32,fontSize:9,color:'A7C7DB'}});")
    L(f"s.addText('Last Event: {esc(last_event)}',{{x:0.55,y:4.66,w:4.8,h:0.32,fontSize:9,color:'A7C7DB'}});")
    L(f"s.addText('Total Runtime: {esc(runtime_text)}',{{x:0.55,y:4.98,w:4.8,h:0.32,fontSize:10,bold:true,color:C.teal}});")
    L(f"s.addText('{total} Features',{{x:7.2,y:1.0,w:2.2,h:0.8,fontSize:32,bold:true,color:C.white,align:'center'}});")
    L(f"s.addText('{partial} Need Action',{{x:7.2,y:2.0,w:2.2,h:0.45,fontSize:14,bold:true,color:'FCA5A5',align:'center'}});")
    L(f"s.addText('{len(screenshots)} Screenshots',{{x:7.2,y:2.6,w:2.2,h:0.45,fontSize:14,bold:true,color:'93C5FD',align:'center'}});")
    L("}")

    # Better summary slide
    L("{const s=pres.addSlide();s.background={color:C.offwhite};hdr(s,'Executive Summary',C.navy);")
    cards = [
        (total, "Total Features", "1A56A0"),
        (optin_ok, "Opt-In OK", "059669"),
        (profile_fail, "Profile FAIL", "DC2626"),
        (auto_count, "Auto Enabled", "7C3AED"),
    ]
    for i, (val, label, color) in enumerate(cards):
        x = 0.35 + i * 2.38
        L(f"s.addShape(pres.shapes.RECTANGLE,{{x:{x},y:0.8,w:2.2,h:1.35,fill:{{color:'{color}'}}}});")
        L(f"s.addText('{val}',{{x:{x},y:0.9,w:2.2,h:0.65,fontSize:34,bold:true,color:'FFFFFF',align:'center'}});")
        L(f"s.addText('{label}',{{x:{x},y:1.58,w:2.2,h:0.25,fontSize:9,bold:true,color:'FFFFFF',align:'center'}});")

    L("const rows=[[{text:'Metric',options:{bold:true,color:'FFFFFF',fill:{color:C.navy}}},{text:'Value',options:{bold:true,color:'FFFFFF',fill:{color:C.navy}}},{text:'Remarks',options:{bold:true,color:'FFFFFF',fill:{color:C.navy}}}]];")
    L(f"rows.push(['Opt-In Success', '{optin_ok}', 'Failed/Not Found: {optin_fail}']);")
    L(f"rows.push(['Profile Options', 'PASS: {profile_pass}', 'FAIL: {profile_fail}']);")
    L(f"rows.push(['ESS Jobs', 'PASS: {ess_pass}', 'FAIL: {ess_fail}']);")
    L(f"rows.push(['Auto Enabled', '{auto_count}', 'Action type Auto']);")
    L(f"rows.push(['Execution Window', '{esc(first_event)} → {esc(last_event)}', 'Runtime: {esc(runtime_text)}']);")
    L("s.addTable(rows,{x:0.35,y:2.35,w:9.3,rowH:0.38,fontSize:9,border:{pt:0.4,color:C.border},colW:[2.2,1.4,5.7]});")

    action_text = ", ".join(f"{k}: {v}" for k, v in action_counter.items())
    L(f"s.addText('Action Required Mix: {esc(action_text)}',{{x:0.4,y:4.55,w:9.2,h:0.3,fontSize:9,color:'334155',italic:true}});")
    L("}")

    # Feature table pages
    per_page = 14
    pages = max(1, (len(features) + per_page - 1) // per_page)
    for p in range(pages):
        chunk = features[p * per_page : (p + 1) * per_page]
        L("{const s=pres.addSlide();s.background={color:C.offwhite};")
        L(f"hdr(s,'Feature Enablement Status ({p+1}/{pages})',C.navy);")
        L("s.addShape(pres.shapes.RECTANGLE,{x:0.15,y:0.75,w:9.7,h:0.32,fill:{color:'1E3A5F'}});")
        headers = [("Feature Name",0.22,4.7),("Action",4.93,1.35),("Opt-In",6.3,0.95),("Profile",7.28,0.95),("ESS",8.26,0.75),("Auto",9.03,0.75)]
        for txt, x, w in headers:
            L(f"s.addText('{txt}',{{x:{x},y:0.75,w:{w},h:0.32,fontSize:8.2,bold:true,color:'FFFFFF',align:'center',valign:'middle'}});")
        y = 1.08
        for i, f in enumerate(chunk):
            bg = "FFFFFF" if i % 2 == 0 else "F1F5F9"
            L(f"s.addShape(pres.shapes.RECTANGLE,{{x:0.15,y:{y:.3f},w:9.7,h:0.3,fill:{{color:'{bg}'}}}});")
            L(f"s.addText('{esc(f['name'])}',{{x:0.22,y:{y:.3f},w:4.65,h:0.3,fontSize:7.5,valign:'middle'}});")
            L(f"s.addText('{esc(f.get('action',''))}',{{x:4.95,y:{y:.3f},w:1.3,h:0.3,fontSize:7,align:'center',valign:'middle'}});")
            opt_fill = "059669" if (f["optin"].startswith("Already Enabled") or f["optin"] == "Enabled Now") else ("DC2626" if f["optin"] in ("Not Found", "FAIL") else "94A3B8")
            prof_fill = "059669" if f["profile"] == "PASS" else ("DC2626" if f["profile"] == "FAIL" else "94A3B8")
            ess_fill = "059669" if f["ess"] == "PASS" else ("DC2626" if f["ess"] == "FAIL" else "94A3B8")
            auto_fill = "7C3AED" if f["auto"] == "AUTO ENABLED" else ("DC2626" if f["auto"] == "FAIL" else "94A3B8")
            L(f"badge(s,6.35,{y+0.04:.3f},0.85,0.22,'{esc(f['optin'])}','{opt_fill}');")
            L(f"badge(s,7.33,{y+0.04:.3f},0.85,0.22,'{esc(f['profile'])}','{prof_fill}');")
            L(f"badge(s,8.31,{y+0.04:.3f},0.65,0.22,'{esc(f['ess'])}','{ess_fill}');")
            L(f"badge(s,9.08,{y+0.04:.3f},0.65,0.22,'{esc('Auto ✓' if f['auto'] == 'AUTO ENABLED' else (f['auto'] or 'N/A'))}','{auto_fill}');")
            y += 0.3
        L("}")

    # Comprehensive action-needed slides
    if remediation:
        per_r = 6
        rpages = (len(remediation) + per_r - 1) // per_r
        for rp in range(rpages):
            slice_r = remediation[rp * per_r : (rp + 1) * per_r]
            L("{const s=pres.addSlide();s.background={color:C.offwhite};")
            L(f"hdr(s,'⚠ Required Actions ({rp+1}/{rpages})','B91C1C');")
            y = 0.75
            for r in slice_r:
                L(f"s.addShape(pres.shapes.RECTANGLE,{{x:0.2,y:{y:.3f},w:9.6,h:0.7,fill:{{color:'FFFFFF'}}}});")
                L(f"s.addShape(pres.shapes.RECTANGLE,{{x:0.2,y:{y:.3f},w:0.06,h:0.7,fill:{{color:'DC2626'}}}});")
                L(f"s.addText('{esc(r['name'])}',{{x:0.32,y:{y+0.03:.3f},w:9.2,h:0.2,fontSize:8.8,bold:true,color:'0D2137'}});")
                L(f"s.addText('Action Type: {esc(r['action'])}',{{x:0.32,y:{y+0.22:.3f},w:9.2,h:0.16,fontSize:7.2,color:'334155',italic:true}});")
                L(f"s.addText('Issue: {esc(r['issue'])}',{{x:0.32,y:{y+0.38:.3f},w:9.2,h:0.14,fontSize:7.1,color:'B91C1C'}});")
                L(f"s.addText('Fix: {esc(r['fix'])}',{{x:0.32,y:{y+0.52:.3f},w:9.2,h:0.14,fontSize:7.1,color:'1E3A5F'}});")
                y += 0.77
            L("}")

    # ALL screenshots, paginated, mixed success/failure
    ss_per_slide = 6
    ss_pages = max(1, (len(screenshots) + ss_per_slide - 1) // ss_per_slide)
    for sp in range(ss_pages):
        chunk = screenshots[sp * ss_per_slide : (sp + 1) * ss_per_slide]
        L("{const s=pres.addSlide();s.background={color:C.offwhite};")
        L(f"hdr(s,'Screenshots (All) ({sp+1}/{ss_pages})','1A56A0');")
        x0, y0, w, h = 0.25, 0.85, 2.95, 1.72
        for i, it in enumerate(chunk):
            c, r = i % 3, i // 3
            x, y = x0 + c * 3.15, y0 + r * 2.25
            p = it["path"].replace("\\", "/")
            L(f"s.addShape(pres.shapes.RECTANGLE,{{x:{x-0.05:.2f},y:{y-0.04:.3f},w:3.05,h:1.82,fill:{{color:'FFFFFF'}}}});")
            L(f"if(fs.existsSync('{esc(p)}')){{s.addImage({{path:'{esc(p)}',x:{x:.2f},y:{y:.3f},w:{w},h:{h},sizing:{{type:'contain',w:{w},h:{h}}}}});}}")
            L(f"s.addText('{esc(it['label'])}',{{x:{x-0.04:.2f},y:{y+h+0.02:.3f},w:3.03,h:0.22,fontSize:6.8,align:'center',wrap:true}});")
            L(f"s.addText('{esc(it['name'])}',{{x:{x-0.04:.2f},y:{y+h+0.23:.3f},w:3.03,h:0.15,fontSize:5.8,color:'64748B',align:'center',wrap:true}});")
        L("}")

    # Final
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    L("{const s=pres.addSlide();s.background={color:C.navy};")
    L("s.addText('Run Complete',{x:0.6,y:0.9,w:8.8,h:0.8,fontSize:42,bold:true,color:'FFFFFF',align:'center'});")
    L(f"s.addText('{esc(run_id)}',{{x:0.6,y:1.7,w:8.8,h:0.4,fontSize:14,color:'7FB3CF',align:'center',italic:true}});")
    L(f"s.addText('Runtime: {esc(runtime_text)}  |  Window: {esc(first_event)} → {esc(last_event)}',{{x:0.6,y:2.2,w:8.8,h:0.35,fontSize:9.5,color:'93C5FD',align:'center'}});")
    L(f"s.addText('{total} features | {optin_ok} opt-in OK | {profile_pass} profile PASS | {profile_fail} profile FAIL | {auto_count} auto',{{x:0.6,y:2.6,w:8.8,h:0.35,fontSize:10,color:'0891B2',align:'center'}});")
    L(f"s.addText('Generated by Al Rajhi Automation Suite · {esc(gen_time)}',{{x:0.6,y:5.25,w:8.8,h:0.25,fontSize:8,color:'2D5A72',align:'center',italic:true}});")
    L("}")

    out = str(output_pptx).replace("\\", "/")
    L(f"pres.writeFile({{fileName:'{esc(out)}'}}).then(()=>console.log('PPT saved: {esc(out)}')).catch(err=>{{console.error(err);process.exit(1);}});")
    return "\n".join(lines)


def main() -> None:
    if len(sys.argv) > 1:
        first_arg = Path(sys.argv[1]).expanduser().resolve()
    else:
        first_arg = DEFAULT_RUN_DIR.resolve()

    if first_arg.is_dir():
        run_dir = first_arg
        feature_log = run_dir / "feature_wise_log.txt"
        ss_dir = run_dir / "screenshots"
        output = (
            Path(sys.argv[2]).expanduser().resolve()
            if len(sys.argv) > 2
            else run_dir / f"{run_dir.name}_featurewise_report.pptx"
        )
    else:
        feature_log = first_arg if len(sys.argv) > 1 else DEFAULT_LOG.resolve()
        ss_dir = (
            Path(sys.argv[2]).expanduser().resolve()
            if len(sys.argv) > 2
            else feature_log.parent / "screenshots"
        )
        output = (
            Path(sys.argv[3]).expanduser().resolve()
            if len(sys.argv) > 3
            else feature_log.parent / f"{feature_log.parent.name}_featurewise_report.pptx"
        )

    if not feature_log.exists():
        print(f"ERROR: feature_wise_log not found: {feature_log}")
        sys.exit(1)
    if not ss_dir.exists():
        print(f"ERROR: screenshots folder not found: {ss_dir}")
        sys.exit(1)

    parsed = parse_feature_wise_log(feature_log)
    features = parsed["features"]
    screenshots = collect_screenshots(ss_dir)
    remediation = build_remediation(features)

    js_code = build_js(output_pptx=output, data=parsed, features=features, screenshots=screenshots, remediation=remediation)

    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", output.stem).strip("_") or "report"
    js_file = feature_log.parent / f"_ppt_featurewise_tmp_{safe_stem}_{os.getpid()}.js"
    js_file.write_text(js_code, encoding="utf-8")

    result = subprocess.run(["node", str(js_file)], capture_output=True, text=True, cwd=str(feature_log.parent))
    js_file.unlink(missing_ok=True)

    if result.returncode != 0:
        print("❌ Node.js error:")
        print(result.stderr or result.stdout)
        sys.exit(1)

    print(result.stdout.strip())
    print(f"✅ Report generated: {output}")


if __name__ == "__main__":
    main()
