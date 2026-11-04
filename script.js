/* ==========================================================================
       IBC (INTER-BILL & BSEG COMPARISON) FRONTEND CONTROLLER
    ========================================================================== */
    let _ibcJobData = null;
    let _ibcActiveTab = "bill";
    let _ibcActiveBillHash = "ALL";
    let _ibcDrilldownCache = null;
    let _ibcActiveDrillTab = null;

    function toggleIbcInputMode() {
        const mode = document.getElementById("ibcInputMode").value;
        const batchBox = document.getElementById("ibcBatchUploadBox");
        const singleBox = document.getElementById("ibcSingleEntryBox");
        if (mode === "single") {
            batchBox.style.display = "none";
            singleBox.style.display = "grid";
        } else {
            batchBox.style.display = "block";
            singleBox.style.display = "none";
        }
    }

    function updateIbcStepper(percent) {
        const steps = [
            { id: "ibcStep1", threshold: 15 },
            { id: "ibcStep2", threshold: 45 },
            { id: "ibcStep3", threshold: 75 },
            { id: "ibcStep4", threshold: 90 }
        ];
        steps.forEach(s => {
            const el = document.getElementById(s.id);
            if (!el) return;
            if (percent >= 100 || percent > s.threshold) {
                el.style.background = "#dcfce7";
                el.style.color = "#166534";
            } else if (percent === s.threshold) {
                el.style.background = "#dbeafe";
                el.style.color = "#1e40af";
            } else {
                el.style.background = "#edf2f7";
                el.style.color = "#4a5568";
            }
        });
    }

    async function startIbcRun() {
        const scope = document.getElementById("ibcScope").value;
        const inputMode = document.getElementById("ibcInputMode").value;
        const env = document.getElementById("ibcEnv").value.trim().toLowerCase();
        const user = document.getElementById("ibcUser").value.trim();
        const pass = document.getElementById("ibcPass").value.trim();
        const fileInput = document.getElementById("ibcFile");
        const newId = document.getElementById("ibcNewId").value.trim();
        const oldId = document.getElementById("ibcOldId").value.trim();

        if (!env || !user || !pass) {
            alert("Please enter Database Environment, Username, and Password.");
            return;
        }
        if (inputMode === "batch" && (!fileInput.files || !fileInput.files[0])) {
            alert("Please select a User_Input Excel file to upload.");
            return;
        }
        if (inputMode === "single" && (!newId || !oldId)) {
            alert("Please enter both New and Cancelled Bill / BSEG IDs.");
            return;
        }

        const runBtn = document.getElementById("runIbcBtn");
        const progBox = document.getElementById("ibcProgressContainer");
        const statusTxt = document.getElementById("ibcStatusText");
        const pctTxt = document.getElementById("ibcPercentText");
        const progBar = document.getElementById("ibcProgressBar");
        const dashArea = document.getElementById("ibcDashboardArea");

        runBtn.disabled = true;
        runBtn.textContent = "Processing IBC Validation...";
        progBox.style.display = "block";
        dashArea.style.display = "none";
        progBar.style.width = "5%";
        pctTxt.textContent = "5%";
        statusTxt.textContent = "Verifying Oracle credentials & starting job...";
        document.getElementById("ibcStatInput").textContent = "0";
        document.getElementById("ibcStatBsegs").textContent = "0";
        updateIbcStepper(5);

        const fd = new FormData();
        fd.append("validationScope", scope);
        fd.append("dbEnv", env);
        fd.append("dbUser", user);
        fd.append("dbPass", pass);
        if (inputMode === "batch") {
            fd.append("ibcFile", fileInput.files[0]);
        } else {
            fd.append("newId", newId);
            fd.append("oldId", oldId);
        }

        try {
            const res = await fetch("/api/ibc/start", { method: "POST", body: fd });
            const data = await res.json();
            if (!res.ok || data.error) {
                alert(`IBC Error:\n\n${data.error || "Failed to start validation."}`);
                runBtn.disabled = false;
                runBtn.textContent = "Run IBC Validation";
                return;
            }

            document.getElementById("ibcStatInput").textContent = data.total_input_rows || 1;
            const taskId = data.task_id;

            const pollTimer = setInterval(async () => {
                try {
                    const sRes = await fetch(`/api/ibc/status/${taskId}`);
                    const sData = await sRes.json();

                    if (!sRes.ok || sData.status === "error") {
                        clearInterval(pollTimer);
                        alert(`IBC Execution Error:\n\n${sData.error || "Validation failed."}`);
                        statusTxt.textContent = `Error: ${sData.error}`;
                        runBtn.disabled = false;
                        runBtn.textContent = "Run IBC Validation";
                        return;
                    }

                    const pct = sData.percent || 15;
                    progBar.style.width = `${pct}%`;
                    pctTxt.textContent = `${pct}%`;
                    statusTxt.textContent = sData.stage || "Processing...";
                    updateIbcStepper(pct);

                    if (sData.ibc_input && sData.ibc_input.length) {
                        document.getElementById("ibcStatBsegs").textContent = sData.ibc_input.length;
                    }

                    if (sData.status === "completed") {
                        clearInterval(pollTimer);
                        runBtn.disabled = false;
                        runBtn.textContent = "Run IBC Validation";
                        _ibcJobData = sData;
                        _ibcActiveBillHash = "ALL";
                        document.getElementById("ibcStatBsegs").textContent = (sData.bseg_summary || []).length;
                        initIbcDashboard(sData);
                    }
                } catch (pollErr) {
                    console.error("IBC Poll Error:", pollErr);
                }
            }, 800);
        } catch (err) {
            alert(`Network Error: ${err.message}`);
            runBtn.disabled = false;
            runBtn.textContent = "Run IBC Validation";
        }
    }

    function initIbcDashboard(data) {
        document.getElementById("ibcDashboardArea").style.display = "block";
        document.getElementById("ibcDrilldownInspector").style.display = "none";

        document.getElementById("ibcMasterDownloadBtn").href = data.master_download_url || "#";
        document.getElementById("ibcZipDownloadBtn").href = data.zip_download_url || "#";

        const k = data.kpis || {};
        document.getElementById("ibcKpiPairs").textContent = `${k.total_bills || 0} / ${k.total_bsegs || 0}`;
        document.getElementById("ibcKpiMatch").textContent = (k.matched_rows || 0).toLocaleString();
        document.getElementById("ibcKpiMismatch").textContent = (k.mismatched_rows || 0).toLocaleString();
        document.getElementById("ibcKpiExceptions").textContent = (k.approved_exceptions || 0).toLocaleString();

        const rateSelect = document.getElementById("ibcRateFilter");
        const uniqueRates = Array.from(new Set((data.bseg_summary || []).map(r => r.Rate_cd).filter(Boolean))).sort();
        rateSelect.innerHTML = `<option value="ALL">All Rates (${uniqueRates.length})</option>` +
            uniqueRates.map(r => `<option value="${r}">${r}</option>`).join("");

        renderIbcCharts(data.charts || {});

        const billTabBtn = document.getElementById("tabBtnBillLevel");
        if (data.scope === "bseg_only") {
            billTabBtn.style.display = "none";
            switchIbcTableTab("bseg");
        } else {
            billTabBtn.style.display = "inline-block";
            switchIbcTableTab("bill");
        }

        renderIbcTables();
    }

    function renderIbcCharts(charts) {
        const img1 = document.getElementById("ibcOverallChartImg");
        const img2 = document.getElementById("ibcTableBarChartImg");
        if (img1 && charts.overall_b64) img1.src = charts.overall_b64;
        if (img2 && charts.table_b64) img2.src = charts.table_b64;
    }

    function switchIbcTableTab(tab) {
        _ibcActiveTab = tab;
        document.getElementById("ibcBillTableContainer").style.display = tab === "bill" ? "block" : "none";
        document.getElementById("ibcBsegTableContainer").style.display = tab === "bseg" ? "block" : "none";
        document.getElementById("ibcInputTableContainer").style.display = tab === "input" ? "block" : "none";

        const bBill = document.getElementById("tabBtnBillLevel");
        const bBseg = document.getElementById("tabBtnBsegLevel");
        const bInput = document.getElementById("tabBtnIbcInput");

        bBill.style.background = tab === "bill" ? "#007bff" : "#e2e8f0";
        bBill.style.color = tab === "bill" ? "#fff" : "#334155";
        bBseg.style.background = tab === "bseg" ? "#007bff" : "#e2e8f0";
        bBseg.style.color = tab === "bseg" ? "#fff" : "#334155";
        bInput.style.background = tab === "input" ? "#007bff" : "#e2e8f0";
        bInput.style.color = tab === "input" ? "#fff" : "#334155";
    }

    function filterByBillPairHash(bpHash, labelText) {
        _ibcActiveBillHash = bpHash;
        document.getElementById("ibcHashFilterText").textContent = labelText;
        document.getElementById("ibcHashFilterBadge").style.display = "inline-block";
        switchIbcTableTab("bseg");
        renderIbcTables();
    }

    function clearIbcBillHashFilter() {
        _ibcActiveBillHash = "ALL";
        document.getElementById("ibcHashFilterBadge").style.display = "none";
        renderIbcTables();
    }

    function renderIbcTables() {
        if (!_ibcJobData) return;
        const taskId = _ibcJobData.task_id;
        const q = (document.getElementById("ibcSearchInput").value || "").trim().toLowerCase();
        const rateFilter = document.getElementById("ibcRateFilter").value;
        const statusFacet = document.getElementById("ibcStatusFacet").value;

        // 1. Bill-Level Table
        const billBody = document.getElementById("ibcBillTableBody");
        const bills = (_ibcJobData.bill_summary || []).filter(r => {
            const matchQ = !q || r["New Bill id"].toLowerCase().includes(q) || r["cancelled Bill ID"].toLowerCase().includes(q);
            const matchSt = statusFacet === "ALL" ||
                (statusFacet === "MISMATCH_ONLY" && r.Total_row_validated_mismatch > 0) ||
                (statusFacet === "MATCH_ONLY" && r.Total_row_validated_mismatch === 0);
            return matchQ && matchSt;
        });

        billBody.innerHTML = bills.map(r => `
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px 10px; font-weight: 600;">${r["New Bill id"]}</td>
                <td style="padding: 8px 10px; font-weight: 600;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 8px 10px; text-align: center;">${r.total_bseg_ID_validate}</td>
                <td style="padding: 8px 10px; color: #15803d; font-weight: 700;">${r.Total_row_validated_match}</td>
                <td style="padding: 8px 10px; color: #b91c1c; font-weight: 700;">${r.Total_row_validated_mismatch}</td>
                <td style="padding: 8px 10px; font-weight: 600;">${r.Total_row_validated_total}</td>
                <td style="padding: 8px 10px; text-align: center;">
                    <a href="/api/ibc/download/${taskId}/${r.download}.xlsx" download style="color: #2563eb; font-weight: 600; text-decoration: underline;">Download</a>
                </td>
                <td style="padding: 8px 10px; text-align: center;">
                    <button type="button" onclick="filterByBillPairHash('${r.bill_pair_hash}', '${r["New Bill id"]}')" style="background: #007bff; color: #fff; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        BSEGs (${r.total_bseg_ID_validate})
                    </button>
                </td>
            </tr>
        `).join("");

        // 2. BSEG-Level Table
        const bsegBody = document.getElementById("ibcBsegTableBody");
        const bsegs = (_ibcJobData.bseg_summary || []).filter(r => {
            const matchHash = _ibcActiveBillHash === "ALL" || r.bill_pair_hash === _ibcActiveBillHash;
            const matchRate = rateFilter === "ALL" || r.Rate_cd === rateFilter;
            const matchQ = !q ||
                r["New Bill id"].toLowerCase().includes(q) ||
                r["cancelled Bill ID"].toLowerCase().includes(q) ||
                r.SA_ID.toLowerCase().includes(q) ||
                r.New_bseg.toLowerCase().includes(q) ||
                r.old_bseg.toLowerCase().includes(q);
            const matchSt = statusFacet === "ALL" ||
                (statusFacet === "MISMATCH_ONLY" && r.Total_row_validated_mismatch > 0) ||
                (statusFacet === "MATCH_ONLY" && r.Total_row_validated_mismatch === 0);
            return matchHash && matchRate && matchQ && matchSt;
        });

        bsegBody.innerHTML = bsegs.map(r => `
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px;">${r["New Bill id"]}</td>
                <td style="padding: 8px;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 8px; font-weight: 600;">${r.SA_ID}</td>
                <td style="padding: 8px; font-weight: 700;">${r.New_bseg}</td>
                <td style="padding: 8px; font-weight: 700;">${r.old_bseg}</td>
                <td style="padding: 8px; font-weight: 600;">${r.Rate_cd}</td>
                <td style="padding: 8px; color: #15803d; font-weight: 700;">${r.Total_row_validated_match}</td>
                <td style="padding: 8px; color: #b91c1c; font-weight: 700;">${r.Total_row_validated_mismatch}</td>
                <td style="padding: 8px; text-align: center;">
                    <a href="/api/ibc/download/${taskId}/${r.download}.xlsx" download style="color: #2563eb; font-weight: 600; text-decoration: underline;">${r.download}</a>
                </td>
                <td style="padding: 8px; text-align: center;">
                    <button type="button" onclick="openIbcDrilldown('${taskId}', '${r.download}')" style="background: #007bff; color: #fff; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        Inspect
                    </button>
                </td>
            </tr>
        `).join("");

        // 3. Discovered IBC_Input Table
        const inputBody = document.getElementById("ibcInputTableBody");
        inputBody.innerHTML = (_ibcJobData.ibc_input || []).map(r => `
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px 10px;">${r["New Bill id"]}</td>
                <td style="padding: 8px 10px;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 8px 10px;">${r.SA_ID}</td>
                <td style="padding: 8px 10px; font-weight: 600;">${r.New_bseg}</td>
                <td style="padding: 8px 10px; font-weight: 600;">${r.old_bseg}</td>
                <td style="padding: 8px 10px; font-weight: 600;">${r.Rate_cd}</td>
            </tr>
        `).join("");
    }

    async function openIbcDrilldown(taskId, dlKey) {
        const inspector = document.getElementById("ibcDrilldownInspector");
        const titleEl = document.getElementById("ibcInspectorTitle");
        const tabsContainer = document.getElementById("ibcInspectorTabs");
        const bodyContainer = document.getElementById("ibcInspectorSheetBody");

        inspector.style.display = "block";
        titleEl.textContent = `${dlKey} (Loading...)`;
        bodyContainer.innerHTML = `<div style="padding: 15px; text-align: center; color: #475569;">Loading table comparison...</div>`;

        try {
            const res = await fetch(`/api/ibc/drilldown/${taskId}/${dlKey}`);
            const data = await res.json();
            if (!res.ok || data.error) {
                bodyContainer.innerHTML = `<div style="color: #dc2626; padding: 12px;">${data.error || "Failed to load."}</div>`;
                return;
            }

            _ibcDrilldownCache = data;
            titleEl.textContent = dlKey;
            const tabNames = Object.keys(data.tabs || {}).filter(k => k !== "child_bseg_keys");
            _ibcActiveDrillTab = tabNames[0];

            tabsContainer.innerHTML = tabNames.map(tName => {
                const rows = data.tabs[tName] || [];
                const mmCount = rows.filter(r => r.status === "mismatch").length;
                return `
                    <button type="button" class="ibc-drill-tab-btn" data-tab="${tName}" onclick="selectInspectorTab('${tName}')"
                        style="padding: 5px 10px; border: 1px solid #cbd5e0; border-radius: 4px; font-weight: 600; font-size: 12px; cursor: pointer; background: ${tName === _ibcActiveDrillTab ? '#007bff' : '#f8fafc'}; color: ${tName === _ibcActiveDrillTab ? '#fff' : '#1e293b'};">
                        ${tName} (${rows.length}) ${mmCount > 0 ? `<span style="background:#dc2626;color:#fff;padding:1px 5px;border-radius:6px;font-size:10px;margin-left:3px;">${mmCount}</span>` : ''}
                    </button>
                `;
            }).join("");

            renderInspectorActiveSheet();
            inspector.scrollIntoView({ behavior: "smooth" });
        } catch (e) {
            bodyContainer.innerHTML = `<div style="color: #dc2626; padding: 12px;">Error: ${e.message}</div>`;
        }
    }

    function selectInspectorTab(tName) {
        _ibcActiveDrillTab = tName;
        document.querySelectorAll(".ibc-drill-tab-btn").forEach(btn => {
            const active = btn.getAttribute("data-tab") === tName;
            btn.style.background = active ? "#007bff" : "#f8fafc";
            btn.style.color = active ? "#fff" : "#1e293b";
        });
        renderInspectorActiveSheet();
    }

    function renderInspectorActiveSheet() {
        if (!_ibcDrilldownCache || !_ibcActiveDrillTab) return;
        const rows = _ibcDrilldownCache.tabs[_ibcActiveDrillTab] || [];
        const filterVal = document.getElementById("ibcInspectorStatusFilter").value;
        const bodyContainer = document.getElementById("ibcInspectorSheetBody");

        const filtered = rows.filter(r => {
            if (filterVal === "ALL") return true;
            if (filterVal === "mismatch") return r.status === "mismatch";
            return r.status.startsWith("match");
        });

        if (filtered.length === 0) {
            bodyContainer.innerHTML = `<div style="padding: 15px; color: #64748b; text-align: center;">No rows match the selected filter.</div>`;
            return;
        }

        let cols = (_ibcDrilldownCache.table_columns && _ibcDrilldownCache.table_columns[_ibcActiveDrillTab]) || [];
        if (!cols.length) {
            for (const r of rows) {
                if (r.new && Object.keys(r.new).length) { cols = Object.keys(r.new); break; }
                if (r.old && Object.keys(r.old).length) { cols = Object.keys(r.old); break; }
            }
        }

        const rawKey = _ibcDrilldownCache.dl_key.replace("bill_", "").split("vs");
        const nId = rawKey[0], oId = rawKey[1];

        let html = `<table style="width: 100%; border-collapse: collapse; font-size: 11px; color: #1a202c;">
            <thead>
                <tr style="background: #1e293b; color: #fff;">
                    <th colspan="${cols.length}" style="padding: 6px; border-right: 2px solid #fff;">New: ${nId}</th>
                    <th colspan="${cols.length}" style="padding: 6px; border-right: 2px solid #fff;">Old: ${oId}</th>
                    <th style="padding: 6px;">status</th>
                </tr>
                <tr style="background: #f1f5f9; color: #1e293b; border-bottom: 2px solid #cbd5e0;">
                    ${cols.map(c => `<th style="padding: 5px;">${c}</th>`).join("")}
                    ${cols.map(c => `<th style="padding: 5px;">${c}</th>`).join("")}
                    <th style="padding: 5px;">status</th>
                </tr>
            </thead>
            <tbody>`;

        filtered.forEach(item => {
            const n = item.new || {};
            const o = item.old || {};
            const st = item.status;
            const bg = st === "match" ? "#f0fdf4" : (st.includes("exception") ? "#ecfeff" : "#fef2f2");
            const stColor = st === "match" ? "#15803d" : (st.includes("exception") ? "#0284c7" : "#dc2626");

            html += `<tr style="background: ${bg}; border-bottom: 1px solid #e2e8f0;">
                ${cols.map(c => `<td style="padding: 5px;">${n[c] !== undefined && n[c] !== null ? n[c] : ""}</td>`).join("")}
                ${cols.map(c => `<td style="padding: 5px;">${o[c] !== undefined && o[c] !== null ? o[c] : ""}</td>`).join("")}
                <td style="padding: 5px; font-weight: 700; color: ${stColor};">${st}</td>
            </tr>`;
        });

        html += `</tbody></table>`;
        bodyContainer.innerHTML = html;
    }

    window.toggleIbcInputMode = toggleIbcInputMode;
    window.startIbcRun = startIbcRun;
    window.switchIbcTableTab = switchIbcTableTab;
    window.filterByBillPairHash = filterByBillPairHash;
    window.clearIbcBillHashFilter = clearIbcBillHashFilter;
    window.renderIbcTables = renderIbcTables;
    window.openIbcDrilldown = openIbcDrilldown;
    window.selectInspectorTab = selectInspectorTab;
    window.renderInspectorActiveSheet = renderInspectorActiveSheet;