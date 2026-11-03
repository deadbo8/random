import os
import sys
import re
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from controller.ibc_db import (
    BILL_TABLE_REGISTRY,
    BSEG_TABLE_REGISTRY,
    IBCDataFetcher,
    norm_bill_id,
    norm_bseg_id,
    norm_sa_id,
    make_bill_hash,
    make_bseg_hash
)

# ==============================================================================
# 1. PATH & VALUE NORMALIZATION HELPERS
# ==============================================================================
def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.abspath(os.path.dirname(os.path.dirname(__file__)))

def norm_val(v):
    """Standardizes SQL cell values (floats, decimals, dates, strings) for 1:1 comparison."""
    if v is None or pd.isna(v):
        return ""
    if isinstance(v, (float, Decimal)):
        try:
            d = Decimal(str(v)).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)
            s = format(d.normalize(), "f")
            return "0" if s in ("-0", "-0.0", "0.0", "") else s
        except Exception:
            return str(round(float(v), 6))
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return " ".join(str(v).strip().split())


# ==============================================================================
# 2. CORE IBC RECONCILIATION & REPORTING ENGINE (ZERO SQL)
# ==============================================================================
class IBCEngine:
    def __init__(self, env: str, user: str, password: str, scope: str = "bill_and_bseg"):
        self.env = env.lower().strip()
        self.scope = scope.lower().strip()  # "bill_and_bseg" OR "bseg_only"
        self.db = IBCDataFetcher(self.env, user, password)
        self.base_dir = get_base_dir()
        self.exceptions = self._load_25_4_exceptions()

    def _load_25_4_exceptions(self) -> list[dict]:
        exc_file = os.path.join(self.base_dir, "storage", "compare_exceptions.xlsx")
        rules = []
        if not os.path.exists(exc_file):
            return rules
        try:
            xl = pd.ExcelFile(exc_file)
            if "IBC_25_4_EXCEPTIONS" in xl.sheet_names:
                df = pd.read_excel(exc_file, sheet_name="IBC_25_4_EXCEPTIONS")
                for _, r in df.iterrows():
                    pat = str(r.get("PATTERN", "")).strip()
                    if pat and pat.lower() != "nan":
                        rules.append({
                            "pattern": pat.lower(),
                            "match_type": str(r.get("MATCH_TYPE", "CONTAINS")).strip().upper(),
                            "zero_only": str(r.get("ZERO_AMT_ONLY", "Y")).strip().upper() in ("Y", "YES", "TRUE", "1")
                        })
            if "CCB_EXCEPTIONS" in xl.sheet_names:
                df_ccb = pd.read_excel(exc_file, sheet_name="CCB_EXCEPTIONS")
                for x in df_ccb.iloc[:, 0].dropna():
                    pat = str(x).strip().lower()
                    if pat:
                        rules.append({"pattern": pat, "match_type": "WILDCARD", "zero_only": True})
        except Exception as e:
            print(f"[IBC Exception Warning] {e}")
        return rules

    def _is_25_4_exception(self, row_dict: dict) -> bool:
        descr = str(row_dict.get("DESCR_ON_BILL", "")).strip().lower()
        if not descr:
            return False
        try:
            amt = float(row_dict.get("CALC_AMT") or 0.0)
        except Exception:
            amt = 0.0

        for r in self.exceptions:
            if r["zero_only"] and round(amt, 2) != 0.0:
                continue
            pat, m_type = r["pattern"], r["match_type"]
            if m_type == "EXACT" and descr == pat:
                return True
            elif m_type == "REGEX" and re.search(pat, descr, re.IGNORECASE):
                return True
            elif m_type == "WILDCARD" and "%a" in pat:
                regex = "^" + re.escape(pat).replace("%a", ".*") + "$"
                if re.match(regex, descr):
                    return True
            elif pat in descr:
                return True
        return False

    # ---------------- STEP 1 TO 4: AUTO-DISCOVERY ---------------- #
    def discover_and_pair(self, df_user_input: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
        cols = list(df_user_input.columns)
        if len(cols) < 2:
            raise ValueError("User_Input Excel must have at least 2 columns: 'New Bill id' and 'cancelled Bill ID'")

        col_map = {str(c).strip().lower().replace(" ", "_"): c for c in cols}
        new_bill_col = col_map.get("new_bill_id", cols[0])
        old_bill_col = col_map.get("cancelled_bill_id", col_map.get("canceled_bill_id", cols[1]))
        sa_col = col_map.get("sa_id")
        new_bseg_col = col_map.get("new_bseg", col_map.get("new_bseg_id"))
        old_bseg_col = col_map.get("old_bseg", col_map.get("old_bceg", col_map.get("old_bseg_id")))

        raw_rows = []
        for _, row in df_user_input.iterrows():
            nb = norm_bill_id(row.get(new_bill_col))
            ob = norm_bill_id(row.get(old_bill_col))
            sa = norm_sa_id(row.get(sa_col)) if sa_col else ""
            n_bs = norm_bseg_id(row.get(new_bseg_col)) if new_bseg_col else ""
            o_bs = norm_bseg_id(row.get(old_bseg_col)) if old_bseg_col else ""
            if nb or ob or n_bs or o_bs:
                raw_rows.append({"nb": nb, "ob": ob, "sa": sa, "n_bs": n_bs, "o_bs": o_bs})

        if not raw_rows:
            return pd.DataFrame(), []

        all_lookup_ids = list({x for r in raw_rows for x in (r["nb"], r["ob"], r["n_bs"], r["o_bs"]) if x})
        bseg_rows = self.db.query_ci_bseg_discovery(all_lookup_ids)

        all_bseg_ids = list({r["bseg_id"] for r in bseg_rows} | {r["n_bs"] for r in raw_rows if r["n_bs"]})
        rate_map = self.db.query_rate_schedules(all_bseg_ids)

        ibc_input_rows = []
        for r in raw_rows:
            new_id, old_id = r["nb"], r["ob"]

            if r["n_bs"] and r["o_bs"]:
                n_bs, o_bs = r["n_bs"], r["o_bs"]
                ns_meta = next((b for b in bseg_rows if b["bseg_id"] == n_bs), {})
                os_meta = next((b for b in bseg_rows if b["bseg_id"] == o_bs), {})
                n_bill = new_id or ns_meta.get("bill_id", "")
                o_bill = old_id or os_meta.get("bill_id", "")
                sa_id = r["sa"] or ns_meta.get("sa_id") or os_meta.get("sa_id", "")
                bp_hash = make_bill_hash(n_bill, o_bill)
                sp_hash = make_bseg_hash(bp_hash, n_bs, o_bs)
                ibc_input_rows.append({
                    "New Bill id": n_bill, "cancelled Bill ID": o_bill, "SA_ID": sa_id,
                    "New_bseg": n_bs, "old_bseg": o_bs,
                    "Rate_cd": rate_map.get(n_bs) or rate_map.get(o_bs) or "-",
                    "bill_pair_hash": bp_hash, "bseg_pair_hash": sp_hash
                })
                continue

            bp_hash = make_bill_hash(new_id, old_id)
            new_segs = [b for b in bseg_rows if b["bill_id"] == new_id or b["bseg_id"] == new_id]
            old_segs = [b for b in bseg_rows if b["bill_id"] == old_id or b["bseg_id"] == old_id]

            if new_id == old_id and len(new_segs) > 1:
                frozen = [b for b in new_segs if b["stat"] == "50"]
                canceled = [b for b in old_segs if b["stat"] == "60"]
                if frozen and canceled:
                    new_segs, old_segs = frozen, canceled

            pool_old = list(old_segs)
            paired = []
            for ns in new_segs:
                mo = next((o for o in pool_old if ns["can_bseg_id"] and o["bseg_id"] == ns["can_bseg_id"]), None)
                if not mo and ns["sa_id"]:
                    mo = next((o for o in pool_old if o["sa_id"] == ns["sa_id"] and o["bseg_id"] != ns["bseg_id"]), None)
                if not mo and ns["sa_id"]:
                    mo = next((o for o in pool_old if o["sa_id"] == ns["sa_id"]), None)
                if mo:
                    pool_old.remove(mo)
                    paired.append((ns, mo))
                else:
                    paired.append((ns, None))

            for rem_o in pool_old:
                paired.append((None, rem_o))

            for ns, mo in paired:
                n_bill = ns["bill_id"] if ns else new_id
                o_bill = mo["bill_id"] if mo else old_id
                n_bs = ns["bseg_id"] if ns else "MISSING"
                o_bs = mo["bseg_id"] if mo else "MISSING"
                sa_id = ns["sa_id"] if ns else (mo["sa_id"] if mo else "")
                rate_cd = rate_map.get(n_bs) or rate_map.get(o_bs) or "-"
                sp_hash = make_bseg_hash(bp_hash, n_bs, o_bs)

                ibc_input_rows.append({
                    "New Bill id": n_bill, "cancelled Bill ID": o_bill, "SA_ID": sa_id,
                    "New_bseg": n_bs, "old_bseg": o_bs, "Rate_cd": rate_cd,
                    "bill_pair_hash": bp_hash, "bseg_pair_hash": sp_hash
                })

        return pd.DataFrame(ibc_input_rows), ibc_input_rows

    def extract_datasets(self, pairs: list[dict]) -> dict:
        bills = list({p["New Bill id"] for p in pairs if p["New Bill id"]} |
                     {p["cancelled Bill ID"] for p in pairs if p["cancelled Bill ID"]})
        bsegs = list({p["New_bseg"] for p in pairs if p["New_bseg"] not in ("", "MISSING")} |
                     {p["old_bseg"] for p in pairs if p["old_bseg"] not in ("", "MISSING")})
        return self.db.fetch_all_registered_tables(bills, bsegs, self.scope)

    # ---------------- 3-PASS ORDER-INDEPENDENT TABLE COMPARER ---------------- #
    def _row_signature(self, row: dict, ignore_cols: set[str]) -> tuple:
        return tuple(
            (k, norm_val(v))
            for k, v in sorted(row.items())
            if k not in ignore_cols
        )

    def compare_table_rows(
        self,
        new_rows: list[dict],
        old_rows: list[dict],
        ignore_cols: set[str],
        align_keys: list[str],
        check_25_4: bool = False
    ) -> dict:
        match_cnt, mismatch_cnt, exc_cnt = 0, 0, 0
        side_by_side = []

        used_new = [False] * len(new_rows)
        used_old = [False] * len(old_rows)

        new_sigs = [self._row_signature(r, ignore_cols) for r in new_rows]
        old_sigs = [self._row_signature(r, ignore_cols) for r in old_rows]

        # PASS 1: 100% Content Match (Order-Independent)
        for o_idx, o_row in enumerate(old_rows):
            o_sig = old_sigs[o_idx]
            for n_idx, n_row in enumerate(new_rows):
                if not used_new[n_idx] and new_sigs[n_idx] == o_sig:
                    used_old[o_idx] = True
                    used_new[n_idx] = True
                    match_cnt += 1
                    side_by_side.append({"new": n_row, "old": o_row, "status": "match"})
                    break

        # PASS 2: Business Key Alignment for Modified Rows
        for o_idx, o_row in enumerate(old_rows):
            if used_old[o_idx]:
                continue
            for n_idx, n_row in enumerate(new_rows):
                if not used_new[n_idx] and all(norm_val(o_row.get(k)) == norm_val(n_row.get(k)) for k in align_keys):
                    used_old[o_idx] = True
                    used_new[n_idx] = True
                    mismatch_cnt += 1
                    side_by_side.append({"new": n_row, "old": o_row, "status": "mismatch"})
                    break

        # PASS 3A: Unmatched Old Rows
        for o_idx, o_row in enumerate(old_rows):
            if not used_old[o_idx]:
                mismatch_cnt += 1
                side_by_side.append({"new": {}, "old": o_row, "status": "mismatch"})

        # PASS 3B: Unmatched New Rows
        for n_idx, n_row in enumerate(new_rows):
            if not used_new[n_idx]:
                if check_25_4 and self._is_25_4_exception(n_row):
                    match_cnt += 1
                    exc_cnt += 1
                    side_by_side.append({"new": n_row, "old": {}, "status": "match (25.4 exception)"})
                else:
                    mismatch_cnt += 1
                    side_by_side.append({"new": n_row, "old": {}, "status": "mismatch"})

        return {"match": match_cnt, "mismatch": mismatch_cnt, "exceptions": exc_cnt, "rows": side_by_side}

    # ---------------- DYNAMIC REGISTRY-DRIVEN RECONCILIATION ---------------- #
    def compare_all(self, pairs: list[dict], store: dict) -> dict:
        bseg_summary, bill_summary = [], []
        bseg_drilldowns, bill_drilldowns = {}, {}
        bill_groups = defaultdict(list)
        total_exceptions = 0

        for p in pairs:
            bill_groups[(p["New Bill id"], p["cancelled Bill ID"], p["bill_pair_hash"])].append(p)
            nb, ob = p["New_bseg"], p["old_bseg"]

            tot_m, tot_mm, pair_exc = 0, 0, 0
            row_metrics = {}
            drill_tabs = {}

            # Dynamically loop over all registered BSEG tables
            for tbl_name, meta in BSEG_TABLE_REGISTRY.items():
                res = self.compare_table_rows(
                    store[tbl_name][nb],
                    store[tbl_name][ob],
                    meta["ignore_cols"],
                    meta["align_keys"],
                    check_25_4=meta["check_25_4"]
                )
                row_metrics[f"{tbl_name}_match"] = res["match"]
                row_metrics[f"{tbl_name}_mismatch"] = res["mismatch"]
                tot_m += res["match"]
                tot_mm += res["mismatch"]
                pair_exc += res["exceptions"]
                drill_tabs[tbl_name] = res["rows"]

            total_exceptions += pair_exc
            dl_key = f"{nb}vs{ob}"
            bseg_summary.append({
                **p,
                **row_metrics,
                "Total_row_validated_match": tot_m,
                "Total_row_validated_mismatch": tot_mm,
                "Total_row_validated_total": tot_m + tot_mm,
                "exceptions_count": pair_exc,
                "download": dl_key
            })
            bseg_drilldowns[dl_key] = drill_tabs

        if self.scope == "bill_and_bseg":
            for (n_bill, o_bill, bp_hash), group in bill_groups.items():
                bill_m, bill_mm = 0, 0
                bill_metrics = {}
                bill_tabs = {}

                # Dynamically loop over all registered Bill tables
                for tbl_name, meta in BILL_TABLE_REGISTRY.items():
                    res = self.compare_table_rows(
                        store[tbl_name][n_bill],
                        store[tbl_name][o_bill],
                        meta["ignore_cols"],
                        meta["align_keys"],
                        check_25_4=meta["check_25_4"]
                    )
                    bill_metrics[f"{tbl_name}_match"] = res["match"]
                    bill_metrics[f"{tbl_name}_mismatch"] = res["mismatch"]
                    bill_m += res["match"]
                    bill_mm += res["mismatch"]
                    bill_tabs[tbl_name] = res["rows"]

                sp_hashes = {g["bseg_pair_hash"] for g in group}
                child_bsegs = [b for b in bseg_summary if b["bseg_pair_hash"] in sp_hashes]

                bs_m = sum(b["Total_row_validated_match"] for b in child_bsegs)
                bs_mm = sum(b["Total_row_validated_mismatch"] for b in child_bsegs)
                g_m = bill_m + bs_m
                g_mm = bill_mm + bs_mm

                dl_bill_key = f"bill_{n_bill}vs{o_bill}"
                bill_summary.append({
                    "New Bill id": n_bill,
                    "cancelled Bill ID": o_bill,
                    **bill_metrics,
                    "Total_Bill_row_validated_match": bill_m,
                    "Total_Bill_row_validated_mismatch": bill_mm,
                    "total_bseg_ID_validate": len(group),
                    "Total_Bseg_row_validated_match": bs_m,
                    "Total_Bseg_row_validated_mismatch": bs_mm,
                    "Total_row_validated_match": g_m,
                    "Total_row_validated_mismatch": g_mm,
                    "Total_row_validated_total": g_m + g_mm,
                    "download": dl_bill_key,
                    "bill_pair_hash": bp_hash
                })
                bill_drilldowns[dl_bill_key] = {
                    **bill_tabs,
                    "child_bseg_keys": [f"{g['New_bseg']}vs{g['old_bseg']}" for g in group]
                }

        return {
            "scope": self.scope,
            "bseg_summary": bseg_summary,
            "bill_summary": bill_summary,
            "bseg_drilldowns": bseg_drilldowns,
            "bill_drilldowns": bill_drilldowns,
            "total_exceptions": total_exceptions,
            "table_columns": store.get("_columns", {})
        }

    # ---------------- DYNAMIC EXCEL REPORT GENERATORS ---------------- #
    def generate_all_reports(self, user_input_df: pd.DataFrame, ibc_input_df: pd.DataFrame, results: dict, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        task_id = os.path.basename(output_dir).replace("ibc_run_", "")
        master_path = os.path.join(output_dir, f"IBC_Master_Report_{task_id}.xlsx")

        wb_master = openpyxl.Workbook()

        # Tab 1: User_Input
        ws_user = wb_master.active
        ws_user.title = "User_Input"
        ws_user.append(list(user_input_df.columns))
        for _, r in user_input_df.iterrows():
            ws_user.append([str(x) if pd.notna(x) else "" for x in r])

        # Tab 2: IBC_Input
        ws_ibc_in = wb_master.create_sheet(title="IBC_Input")
        cols_in = ["New Bill id", "cancelled Bill ID", "SA_ID", "New_bseg", "old_bseg", "Rate_cd"]
        ws_ibc_in.append(cols_in)
        if not ibc_input_df.empty:
            for _, r in ibc_input_df[cols_in].iterrows():
                ws_ibc_in.append(list(r))

        # Tab 3: IBC_Output_Bseg_level (Dynamically builds headers from BSEG_TABLE_REGISTRY)
        ws_bseg_out = wb_master.create_sheet(title="IBC_Output_Bseg_level")
        bseg_tbl_names = list(BSEG_TABLE_REGISTRY.keys())
        r1_bseg = ["New Bill id", "cancelled Bill ID", "SA_ID", "New_bseg", "old_bseg", "Rate_cd"]
        r2_bseg = ["", "", "", "", "", ""]
        for t_name in bseg_tbl_names:
            r1_bseg.extend([t_name, ""])
            r2_bseg.extend(["match", "mismatch"])
        r1_bseg.extend(["Total row validated", ""])
        r2_bseg.extend(["match", "mismatch"])
        ws_bseg_out.append(r1_bseg)
        ws_bseg_out.append(r2_bseg)

        for b in results["bseg_summary"]:
            row_vals = [b["New Bill id"], b["cancelled Bill ID"], b["SA_ID"], b["New_bseg"], b["old_bseg"], b["Rate_cd"]]
            for t_name in bseg_tbl_names:
                row_vals.extend([b.get(f"{t_name}_match", 0), b.get(f"{t_name}_mismatch", 0)])
            row_vals.extend([b["Total_row_validated_match"], b["Total_row_validated_mismatch"]])
            ws_bseg_out.append(row_vals)

        # Tab 4 & 5: Bill Level Tabs (When scope == "bill_and_bseg")
        if self.scope == "bill_and_bseg":
            ws_bill_out = wb_master.create_sheet(title="IBC_Output_Bill_level")
            bill_tbl_names = list(BILL_TABLE_REGISTRY.keys())
            r1_bill = ["New Bill id", "cancelled Bill ID"]
            r2_bill = ["", ""]
            for t_name in bill_tbl_names:
                r1_bill.extend([t_name, ""])
                r2_bill.extend(["match", "mismatch"])
            r1_bill.extend(["Total Bill row validated", "", "total bseg ID validate", "Total Bseg row validated", "", "Total row validated", "", ""])
            r2_bill.extend(["match", "mismatch", "", "match", "mismatch", "match", "mismatch", "total"])
            ws_bill_out.append(r1_bill)
            ws_bill_out.append(r2_bill)

            for bi in results["bill_summary"]:
                row_vals = [bi["New Bill id"], bi["cancelled Bill ID"]]
                for t_name in bill_tbl_names:
                    row_vals.extend([bi.get(f"{t_name}_match", 0), bi.get(f"{t_name}_mismatch", 0)])
                row_vals.extend([
                    bi["Total_Bill_row_validated_match"], bi["Total_Bill_row_validated_mismatch"],
                    bi["total_bseg_ID_validate"],
                    bi["Total_Bseg_row_validated_match"], bi["Total_Bseg_row_validated_mismatch"],
                    bi["Total_row_validated_match"], bi["Total_row_validated_mismatch"],
                    bi["Total_row_validated_total"]
                ])
                ws_bill_out.append(row_vals)

            ws_fe_bill = wb_master.create_sheet(title="Ouptut front end_bill level")
            ws_fe_bill.append(["New Bill id", "cancelled Bill ID", "match", "mismatch", "total", "download"])
            for bi in results["bill_summary"]:
                ws_fe_bill.append([
                    bi["New Bill id"], bi["cancelled Bill ID"],
                    bi["Total_row_validated_match"], bi["Total_row_validated_mismatch"],
                    bi["Total_row_validated_total"], bi["download"]
                ])

        # Tab 6: Ouptut front end_bseg level
        ws_fe_bseg = wb_master.create_sheet(title="Ouptut front end_bseg level")
        ws_fe_bseg.append(["New Bill id", "cancelled Bill ID", "SA_ID", "New_bseg", "old_bseg", "Rate_cd", "match", "mismatch", "download"])
        for b in results["bseg_summary"]:
            ws_fe_bseg.append([
                b["New Bill id"], b["cancelled Bill ID"], b["SA_ID"], b["New_bseg"], b["old_bseg"], b["Rate_cd"],
                b["Total_row_validated_match"], b["Total_row_validated_mismatch"], b["download"]
            ])

        self._style_and_autofit(wb_master)
        wb_master.save(master_path)

        # Individual BSEG Workbooks
        for bseg_key, tabs_dict in results["bseg_drilldowns"].items():
            bseg_file = os.path.join(output_dir, f"{bseg_key}.xlsx")
            wb_bs = openpyxl.Workbook()
            first = True
            n_id, o_id = bseg_key.split("vs")
            for tab_name, rows in tabs_dict.items():
                ws = wb_bs.active if first else wb_bs.create_sheet(title=tab_name)
                ws.title = tab_name
                first = False
                self._write_side_by_side_sheet(ws, n_id, o_id, rows, results.get("table_columns", {}).get(tab_name, []))
            self._style_and_autofit(wb_bs)
            wb_bs.save(bseg_file)

        # Individual Bill Workbooks
        if self.scope == "bill_and_bseg":
            for bill_key, b_data in results["bill_drilldowns"].items():
                bill_file = os.path.join(output_dir, f"{bill_key}.xlsx")
                wb_bi = openpyxl.Workbook()
                first = True
                raw_pair = bill_key.replace("bill_", "").split("vs")
                n_b, o_b = raw_pair[0], raw_pair[1]
                for tbl_name in BILL_TABLE_REGISTRY.keys():
                    ws = wb_bi.active if first else wb_bi.create_sheet(title=tbl_name)
                    ws.title = tbl_name
                    first = False
                    self._write_side_by_side_sheet(ws, n_b, o_b, b_data.get(tbl_name, []), results.get("table_columns", {}).get(tbl_name, []))

                ws_bs_sum = wb_bi.create_sheet(title="BSEG_Summary")
                ws_bs_sum.append(["SA_ID", "New_bseg", "old_bseg", "Rate_cd", "match", "mismatch"])
                for ck in b_data.get("child_bseg_keys", []):
                    match_b = next((x for x in results["bseg_summary"] if x["download"] == ck), None)
                    if match_b:
                        ws_bs_sum.append([match_b["SA_ID"], match_b["New_bseg"], match_b["old_bseg"], match_b["Rate_cd"],
                                          match_b["Total_row_validated_match"], match_b["Total_row_validated_mismatch"]])
                self._style_and_autofit(wb_bi)
                wb_bi.save(bill_file)

        return master_path

    def _write_side_by_side_sheet(self, ws, new_id: str, old_id: str, rows: list[dict], default_cols: list[str]):
        GREEN = PatternFill("solid", fgColor="D4EDDA")
        RED = PatternFill("solid", fgColor="F8D7DA")
        BLUE = PatternFill("solid", fgColor="D1ECF1")
        BOLD = Font(name="Calibri", size=10, bold=True)

        cols = list(default_cols)
        if not cols:
            for r in rows:
                if r["new"]:
                    cols = list(r["new"].keys())
                    break
                if r["old"]:
                    cols = list(r["old"].keys())
                    break
        if not cols:
            cols = ["NO_DATA"]

        num_cols = len(cols)
        old_start_col = num_cols + 3
        status_col = old_start_col + num_cols

        ws.cell(row=1, column=2, value=f"New: {new_id}").font = BOLD
        ws.cell(row=1, column=old_start_col, value=f"Old: {old_id}").font = BOLD
        ws.cell(row=1, column=status_col, value="status").font = BOLD

        for idx, c_name in enumerate(cols):
            ws.cell(row=2, column=2 + idx, value=c_name).font = BOLD
            ws.cell(row=2, column=old_start_col + idx, value=c_name).font = BOLD

        for r_i, item in enumerate(rows, start=3):
            n_dict, o_dict, st = item["new"], item["old"], item["status"]
            for idx, c_name in enumerate(cols):
                if n_dict:
                    ws.cell(row=r_i, column=2 + idx, value=norm_val(n_dict.get(c_name)))
                if o_dict:
                    ws.cell(row=r_i, column=old_start_col + idx, value=norm_val(o_dict.get(c_name)))

            st_cell = ws.cell(row=r_i, column=status_col, value=st)
            st_cell.font = BOLD
            st_cell.fill = GREEN if st == "match" else (BLUE if "exception" in st else RED)

    def _style_and_autofit(self, wb):
        NAVY = PatternFill("solid", fgColor="003366")
        WHITE_BOLD = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        for ws in wb.worksheets:
            ws.views.sheetView[0].showGridLines = True
            if ws.max_row >= 1:
                for cell in ws[1]:
                    if cell.value:
                        cell.fill = NAVY
                        cell.font = WHITE_BOLD
            for col in ws.columns:
                max_len = max(len(str(c.value or "")) for c in col[:200])
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)