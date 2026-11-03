import os
import io
import uuid
import zipfile
import threading
import pandas as pd
import oracledb
from flask import Blueprint, request, jsonify, send_file
from werkzeug.utils import secure_filename

import controller.charge_utils  # Prevents circular import in db_fetch
from controller.db_fetch import FETCH_DB_ENV
from controller.ibc_db import BSEG_TABLE_REGISTRY, BILL_TABLE_REGISTRY
from controller.ibc_engine import IBCEngine, get_base_dir

ibc_bp = Blueprint("ibc_bp", __name__)

# Global in-memory store for active and completed IBC jobs
IBC_JOBS = {}
ibc_lock = threading.Lock()


# ==============================================================================
# 1. PRE-FLIGHT ORACLE CREDENTIAL CHECK
# ==============================================================================
def _verify_oracle_credentials(env: str, user: str, pwd: str):
    cfg = FETCH_DB_ENV.get(env.lower().strip())
    if not cfg:
        return False, f"Invalid Environment '{env}' selected."
    dsn = f"{cfg['host']}:{cfg['port']}/{cfg['service']}"
    try:
        with oracledb.connect(user=user, password=pwd, dsn=dsn, tcp_connect_timeout=5) as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1 FROM DUAL")
        return True, None
    except oracledb.DatabaseError as db_err:
        err_obj = db_err.args[0]
        if getattr(err_obj, "code", 0) == 1017:
            return False, "Invalid Oracle Username or Password. Run Aborted."
        elif getattr(err_obj, "code", 0) == 28000:
            return False, "Oracle Account is LOCKED. Please contact your DBA."
        return False, f"Database Error: {err_obj.message}"
    except Exception as e:
        return False, f"Unable to reach database: {str(e)}"


# ==============================================================================
# 2. BACKGROUND MULTI-THREADED BATCH ORCHESTRATOR
# ==============================================================================
def _run_ibc_job(task_id: str, df_input: pd.DataFrame, env: str, user: str, pwd: str, scope: str):
    job = IBC_JOBS[task_id]
    base_dir = get_base_dir()
    output_dir = os.path.join(base_dir, "output", f"ibc_run_{task_id}")
    os.makedirs(output_dir, exist_ok=True)

    try:
        engine = IBCEngine(env=env, user=user, password=pwd, scope=scope)

        # Stage 1: Auto-Discovery (Steps 1–4 from Teams)
        with ibc_lock:
            job["stage"] = "Step 1/4: Discovering & Pairing Bill Segments (CI_BSEG & CI_BSEG_CALC)..."
            job["percent"] = 15

        df_ibc_input, structured_pairs = engine.discover_and_pair(df_input)
        if not structured_pairs:
            with ibc_lock:
                job["status"] = "error"
                job["error"] = "No matching Bill or Bill Segment records found in CI_BSEG for the provided IDs."
            return

        with ibc_lock:
            job["total_bseg_pairs"] = len(structured_pairs)
            job["stage"] = f"Step 2/4: Extracting SQL Tables in parallel for {len(structured_pairs)} segment pairs..."
            job["percent"] = 45

        # Stage 2: Multi-threaded Chunked SQL Extraction
        raw_store = engine.extract_datasets(structured_pairs)

        # Stage 3: Order-Independent Table-to-Table Comparison
        with ibc_lock:
            job["stage"] = "Step 3/4: Running order-independent table comparisons & 25.4 exception rules..."
            job["percent"] = 75

        comp_results = engine.compare_all(structured_pairs, raw_store)

        # Stage 4: Generate Master & Individual Excel Workbooks + ZIP
        with ibc_lock:
            job["stage"] = "Step 4/4: Generating Master Report & Individual Bill/BSEG Excel files..."
            job["percent"] = 90

        master_path = engine.generate_all_reports(df_input, df_ibc_input, comp_results, output_dir)
        master_filename = os.path.basename(master_path)

        # Bundle all generated Excels into a single ZIP file
        zip_filename = f"IBC_All_Reports_{task_id}.zip"
        zip_path = os.path.join(output_dir, zip_filename)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(output_dir):
                for f in files:
                    if f != zip_filename and f.endswith(".xlsx"):
                        zf.write(os.path.join(root, f), arcname=f)

        # Compute Frontend KPIs & Chart.js Datasets
        bseg_sum = comp_results["bseg_summary"]
        bill_sum = comp_results["bill_summary"]

        total_bseg_match = sum(r["Total_row_validated_match"] for r in bseg_sum)
        total_bseg_mismatch = sum(r["Total_row_validated_mismatch"] for r in bseg_sum)
        total_bill_msg_match = sum(r.get("Total_Bill_row_validated_match", 0) for r in bill_sum)
        total_bill_msg_mismatch = sum(r.get("Total_Bill_row_validated_mismatch", 0) for r in bill_sum)

        grand_match = total_bseg_match + total_bill_msg_match
        grand_mismatch = total_bseg_mismatch + total_bill_msg_mismatch

        # Per-table match/mismatch breakdown for Chart.js Bar Chart
        table_chart_stats = {}
        if scope == "bill_and_bseg":
            for tbl in BILL_TABLE_REGISTRY.keys():
                table_chart_stats[tbl] = {
                    "match": sum(r.get(f"{tbl}_match", 0) for r in bill_sum),
                    "mismatch": sum(r.get(f"{tbl}_mismatch", 0) for r in bill_sum)
                }
        for tbl in BSEG_TABLE_REGISTRY.keys():
            table_chart_stats[tbl] = {
                "match": sum(r.get(f"{tbl}_match", 0) for r in bseg_sum),
                "mismatch": sum(r.get(f"{tbl}_mismatch", 0) for r in bseg_sum)
            }

        with ibc_lock:
            job["status"] = "completed"
            job["percent"] = 100
            job["stage"] = "Validation Completed Successfully!"
            job["kpis"] = {
                "total_bills": len({r["bill_pair_hash"] for r in structured_pairs}),
                "total_bsegs": len(bseg_sum),
                "matched_rows": grand_match,
                "mismatched_rows": grand_mismatch,
                "total_rows": grand_match + grand_mismatch,
                "approved_exceptions": comp_results.get("total_exceptions", 0)
            }
            job["table_chart_stats"] = table_chart_stats
            job["ibc_input"] = df_ibc_input.to_dict(orient="records")
            job["bseg_summary"] = bseg_sum
            job["bill_summary"] = bill_sum
            job["bseg_drilldowns"] = comp_results["bseg_drilldowns"]
            job["bill_drilldowns"] = comp_results["bill_drilldowns"]
            job["table_columns"] = comp_results.get("table_columns", {})
            job["master_download_url"] = f"/api/ibc/download/{task_id}/{master_filename}"
            job["zip_download_url"] = f"/api/ibc/download/{task_id}/{zip_filename}"

    except Exception as e:
        with ibc_lock:
            job["status"] = "error"
            job["error"] = str(e)


# ==============================================================================
# 3. FLASK API ROUTES
# ==============================================================================
@ibc_bp.route("/api/ibc/start", methods=["POST"])
def start_ibc_validation():
    """Starts an IBC validation job from an uploaded Excel file OR single pair inputs."""
    env = (request.form.get("dbEnv") or "").strip().lower()
    user = (request.form.get("dbUser") or "").strip()
    pwd = (request.form.get("dbPass") or "").strip()
    scope = (request.form.get("validationScope") or "bill_and_bseg").strip().lower()

    if not env or not user or not pwd:
        return jsonify({"error": "Database Environment, Username, and Password are required."}), 400

    # Pre-flight credential check
    ok, auth_err = _verify_oracle_credentials(env, user, pwd)
    if not ok:
        return jsonify({"error": auth_err}), 400

    # Read input from uploaded Excel OR manual single-pair fields
    uploaded_file = request.files.get("ibcFile")
    new_id_manual = (request.form.get("newId") or "").strip()
    old_id_manual = (request.form.get("oldId") or "").strip()

    if uploaded_file and uploaded_file.filename:
        base_dir = get_base_dir()
        uploads_dir = os.path.join(base_dir, "uploads")
        os.makedirs(uploads_dir, exist_ok=True)
        saved_path = os.path.join(uploads_dir, f"ibc_{uuid.uuid4().hex[:6]}_{secure_filename(uploaded_file.filename)}")
        uploaded_file.save(saved_path)
        try:
            df_input = pd.read_excel(saved_path, dtype=str)
        except Exception as e:
            return jsonify({"error": f"Invalid Excel file: {str(e)}"}), 400
    elif new_id_manual and old_id_manual:
        df_input = pd.DataFrame([{"New Bill id": new_id_manual, "cancelled Bill ID": old_id_manual}])
    else:
        return jsonify({"error": "Please upload a User_Input Excel file or enter New and Cancelled IDs."}), 400

    if df_input.empty:
        return jsonify({"error": "The provided input contains no rows."}), 400

    task_id = uuid.uuid4().hex[:10]
    IBC_JOBS[task_id] = {
        "task_id": task_id,
        "status": "processing",
        "scope": scope,
        "percent": 5,
        "stage": "Initializing IBC Engine...",
        "total_input_rows": len(df_input),
        "total_bseg_pairs": 0,
        "error": None
    }

    threading.Thread(
        target=_run_ibc_job,
        args=(task_id, df_input, env, user, pwd, scope),
        daemon=True
    ).start()

    return jsonify({"task_id": task_id, "total_input_rows": len(df_input)})


@ibc_bp.route("/api/ibc/status/<task_id>", methods=["GET"])
def get_ibc_status(task_id):
    """Returns job progress and summary tables (excludes heavy drilldown arrays for fast polling)."""
    job = IBC_JOBS.get(task_id)
    if not job:
        return jsonify({"status": "error", "error": "Task ID not found."}), 404

    with ibc_lock:
        response = {
            "task_id": job["task_id"],
            "status": job["status"],
            "scope": job["scope"],
            "percent": job.get("percent", 0),
            "stage": job.get("stage", ""),
            "error": job.get("error"),
            "kpis": job.get("kpis"),
            "table_chart_stats": job.get("table_chart_stats"),
            "ibc_input": job.get("ibc_input", []),
            "bseg_summary": job.get("bseg_summary", []),
            "bill_summary": job.get("bill_summary", []),
            "master_download_url": job.get("master_download_url"),
            "zip_download_url": job.get("zip_download_url")
        }
    return jsonify(response)


@ibc_bp.route("/api/ibc/drilldown/<task_id>/<dl_key>", methods=["GET"])
def get_ibc_drilldown(task_id, dl_key):
    """Lazy-loads side-by-side table rows when user clicks a Bill or BSEG row in the frontend."""
    job = IBC_JOBS.get(task_id)
    if not job or job.get("status") != "completed":
        return jsonify({"error": "Job not found or not completed."}), 404

    if dl_key.startswith("bill_"):
        data = job.get("bill_drilldowns", {}).get(dl_key)
    else:
        data = job.get("bseg_drilldowns", {}).get(dl_key)

    if not data:
        return jsonify({"error": f"Drilldown data for '{dl_key}' not found."}), 404

    return jsonify({
        "dl_key": dl_key,
        "tabs": data,
        "table_columns": job.get("table_columns", {})
    })


@ibc_bp.route("/api/ibc/download/<task_id>/<filename>", methods=["GET"])
def download_ibc_report(task_id, filename):
    """Downloads Master Excel, Individual BSEG/Bill Excel (`1112vs1113.xlsx`), or Master ZIP."""
    base_dir = get_base_dir()
    if not filename.endswith(".xlsx") and not filename.endswith(".zip"):
        filename = f"{filename}.xlsx"

    file_path = os.path.join(base_dir, "output", f"ibc_run_{task_id}", secure_filename(filename))
    if not os.path.exists(file_path):
        return jsonify({"error": f"Report file '{filename}' not found."}), 404

    mime = "application/zip" if filename.endswith(".zip") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return send_file(file_path, as_attachment=True, download_name=filename, mimetype=mime)


@ibc_bp.route("/api/ibc/template", methods=["GET"])
def download_ibc_template():
    """Generates and downloads the clean 2-column User_Input Excel template for testers."""
    bio = io.BytesIO()
    df_tmpl = pd.DataFrame([
        {"New Bill id": "800526010362", "cancelled Bill ID": "800362232269"}
    ])
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_tmpl.to_excel(writer, sheet_name="User_Input", index=False)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name="IBC_User_Input_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )