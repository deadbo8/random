/* ==========================================================================
       IBC (INTER-BILL & BSEG COMPARISON) FRONTEND CONTROLLER
    ========================================================================== */
    let _ibcJobData = null;
    let _ibcActiveTab = "bill";
    let _ibcActiveBillHash = "ALL";
    let _ibcOverallChartObj = null;
    let _ibcTableChartObj = null;
    let _ibcDrilldownCache = null;
    let _ibcActiveDrillTab = null;

    async function startIbcRun() {
        const scope = document.getElementById("ibcScope").value;
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
        if ((!fileInput.files || !fileInput.files[0]) && (!newId || !oldId)) {
            alert("Please upload a User_Input Excel file OR enter both New and Cancelled IDs.");
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
        statusTxt.textContent = "Verifying Oracle credentials & uploading input...";

        const fd = new FormData();
        fd.append("validationScope", scope);
        fd.append("dbEnv", env);
        fd.append("dbUser", user);
        fd.append("dbPass", pass);
        if (fileInput.files && fileInput.files[0]) {
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
                runBtn.textContent = "▶ Run IBC Table-to-Table Validation";
                return;
            }

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
                        runBtn.textContent = "▶ Run IBC Table-to-Table Validation";
                        return;
                    }

                    const pct = sData.percent || 10;
                    progBar.style.width = `${pct}%`;
                    pctTxt.textContent = `${pct}%`;
                    statusTxt.textContent = sData.stage || "Processing...";

                    if (sData.status === "completed") {
                        clearInterval(pollTimer);
                        runBtn.disabled = false;
                        runBtn.textContent = "▶ Run IBC Table-to-Table Validation";
                        _ibcJobData = sData;
                        _ibcActiveBillHash = "ALL";
                        initIbcDashboard(sData);
                    }
                } catch (pollErr) {
                    console.error("IBC Poll Error:", pollErr);
                }
            }, 1000);
        } catch (err) {
            alert(`Network Error: ${err.message}`);
            runBtn.disabled = false;
            runBtn.textContent = "▶ Run IBC Table-to-Table Validation";
        }
    }

    function initIbcDashboard(data) {
        document.getElementById("ibcDashboardArea").style.display = "block";
        document.getElementById("ibcDrilldownInspector").style.display = "none";

        // Set Download Links
        document.getElementById("ibcMasterDownloadBtn").href = data.master_download_url || "#";
        document.getElementById("ibcZipDownloadBtn").href = data.zip_download_url || "#";

        // Populate KPIs
        const k = data.kpis || {};
        document.getElementById("ibcKpiPairs").textContent = `${k.total_bills || 0} / ${k.total_bsegs || 0}`;
        document.getElementById("ibcKpiMatch").textContent = (k.matched_rows || 0).toLocaleString();
        document.getElementById("ibcKpiMismatch").textContent = (k.mismatched_rows || 0).toLocaleString();
        document.getElementById("ibcKpiExceptions").textContent = (k.approved_exceptions || 0).toLocaleString();
        document.getElementById("ibcKpiTotalRows").textContent = (k.total_rows || 0).toLocaleString();

        // Populate Rate Schedule E-Commerce Filter
        const rateSelect = document.getElementById("ibcRateFilter");
        const uniqueRates = Array.from(new Set((data.bseg_summary || []).map(r => r.Rate_cd).filter(Boolean))).sort();
        rateSelect.innerHTML = `<option value="ALL">All Rate Schedules (${uniqueRates.length})</option>` +
            uniqueRates.map(r => `<option value="${r}">${r}</option>`).join("");

        // Render Charts
        renderIbcCharts(k, data.table_chart_stats || {});

        // Show/Hide Bill-Level Tab based on validationScope
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

    function renderIbcCharts(kpis, tableStats) {
        const ctx1 = document.getElementById("ibcOverallChart").getContext("2d");
        if (_ibcOverallChartObj) _ibcOverallChartObj.destroy();
        _ibcOverallChartObj = new Chart(ctx1, {
            type: "doughnut",
            data: {
                labels: ["Matched Rows", "Mismatched Rows", "25.4 Exceptions"],
                datasets: [{
                    data: [kpis.matched_rows || 0, kpis.mismatched_rows || 0, kpis.approved_exceptions || 0],
                    backgroundColor: ["#28a745", "#dc3545", "#17a2b8"]
                }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
        });

        const ctx2 = document.getElementById("ibcTableBarChart").getContext("2d");
        if (_ibcTableChartObj) _ibcTableChartObj.destroy();
        const tblNames = Object.keys(tableStats);
        _ibcTableChartObj = new Chart(ctx2, {
            type: "bar",
            data: {
                labels: tblNames,
                datasets: [
                    { label: "Match", data: tblNames.map(t => tableStats[t].match), backgroundColor: "#28a745" },
                    { label: "Mismatch", data: tblNames.map(t => tableStats[t].mismatch), backgroundColor: "#dc3545" }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: { x: { stacked: false }, y: { beginAtZero: true } }
            }
        });
    }

    function switchIbcTableTab(tab) {
        _ibcActiveTab = tab;
        const cBill = document.getElementById("ibcBillTableContainer");
        const cBseg = document.getElementById("ibcBsegTableContainer");
        const cInput = document.getElementById("ibcInputTableContainer");
        const bBill = document.getElementById("tabBtnBillLevel");
        const bBseg = document.getElementById("tabBtnBsegLevel");
        const bInput = document.getElementById("tabBtnIbcInput");

        cBill.style.display = tab === "bill" ? "block" : "none";
        cBseg.style.display = tab === "bseg" ? "block" : "none";
        cInput.style.display = tab === "input" ? "block" : "none";

        bBill.style.background = tab === "bill" ? "#007bff" : "#e2e8f0";
        bBill.style.color = tab === "bill" ? "#fff" : "#2d3748";
        bBseg.style.background = tab === "bseg" ? "#007bff" : "#e2e8f0";
        bBseg.style.color = tab === "bseg" ? "#fff" : "#2d3748";
        bInput.style.background = tab === "input" ? "#007bff" : "#e2e8f0";
        bInput.style.color = tab === "input" ? "#fff" : "#2d3748";
    }

    function filterByBillPairHash(bpHash, labelText) {
        _ibcActiveBillHash = bpHash;
        const badge = document.getElementById("ibcHashFilterBadge");
        document.getElementById("ibcHashFilterText").textContent = `(${labelText})`;
        badge.style.display = "inline-block";
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

        // 1. Render Bill-Level Table
        const billBody = document.getElementById("ibcBillTableBody");
        const bills = (_ibcJobData.bill_summary || []).filter(r => {
            const matchQ = !q || r["New Bill id"].toLowerCase().includes(q) || r["cancelled Bill ID"].toLowerCase().includes(q);
            const matchSt = statusFacet === "ALL" ||
                (statusFacet === "MISMATCH_ONLY" && r.Total_row_validated_mismatch > 0) ||
                (statusFacet === "MATCH_ONLY" && r.Total_row_validated_mismatch === 0);
            return matchQ && matchSt;
        });

        billBody.innerHTML = bills.map(r => `
            <tr style="border-bottom: 1px solid #edf2f7;">
                <td style="padding: 9px; font-weight: 700;">${r["New Bill id"]}</td>
                <td style="padding: 9px; font-weight: 700;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 9px; text-align: center;">${r.total_bseg_ID_validate}</td>
                <td style="padding: 9px; text-align: right; color: #15803d; font-weight: 700;">${r.Total_row_validated_match}</td>
                <td style="padding: 9px; text-align: right; color: ${r.Total_row_validated_mismatch > 0 ? '#b91c1c' : '#64748b'}; font-weight: 700;">${r.Total_row_validated_mismatch}</td>
                <td style="padding: 9px; text-align: right; font-weight: 700;">${r.Total_row_validated_total}</td>
                <td style="padding: 9px; text-align: center;">
                    <a href="/api/ibc/download/${taskId}/${r.download}.xlsx" download style="color: #007bff; font-weight: 600; text-decoration: underline;">${r.download}</a>
                </td>
                <td style="padding: 9px; text-align: center;">
                    <button type="button" onclick="filterByBillPairHash('${r.bill_pair_hash}', '${r["New Bill id"]}')" style="background: #003366; color: #fff; border: none; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        🔗 View BSEGs (${r.total_bseg_ID_validate})
                    </button>
                    <button type="button" onclick="openIbcDrilldown('${taskId}', '${r.download}')" style="background: #17a2b8; color: #fff; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px; margin-left: 4px;">
                        🔍 Msgs
                    </button>
                </td>
            </tr>
        `).join("");

        // 2. Render BSEG-Level Table
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

        const fmtCell = (m, mm) => `<span style="color:#15803d;font-weight:700;">${m}</span> / <span style="color:${mm > 0 ? '#b91c1c' : '#94a3b8'};font-weight:700;">${mm}</span>`;

        bsegBody.innerHTML = bsegs.map(r => `
            <tr style="border-bottom: 1px solid #edf2f7;">
                <td style="padding: 8px;">${r["New Bill id"]}</td>
                <td style="padding: 8px;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 8px; font-family: Consolas, monospace;">${r.SA_ID}</td>
                <td style="padding: 8px; font-weight: 700; font-family: Consolas, monospace;">${r.New_bseg}</td>
                <td style="padding: 8px; font-weight: 700; font-family: Consolas, monospace;">${r.old_bseg}</td>
                <td style="padding: 8px;"><span style="background: #edf2f7; padding: 2px 6px; border-radius: 4px; font-weight: 700;">${r.Rate_cd}</span></td>
                <td style="padding: 8px; text-align: center;">${fmtCell(r.CI_BSEG_MSG_match || 0, r.CI_BSEG_MSG_mismatch || 0)}</td>
                <td style="padding: 8px; text-align: center;">${fmtCell(r.CI_BSEG_CALC_match || 0, r.CI_BSEG_CALC_mismatch || 0)}</td>
                <td style="padding: 8px; text-align: center;">${fmtCell(r.CI_BSEG_READ_match || 0, r.CI_BSEG_READ_mismatch || 0)}</td>
                <td style="padding: 8px; text-align: center;">${fmtCell(r.CI_BSEG_SQ_match || 0, r.CI_BSEG_SQ_mismatch || 0)}</td>
                <td style="padding: 8px; text-align: center;">${fmtCell(r.CI_BSEG_CALC_LN_match || 0, r.CI_BSEG_CALC_LN_mismatch || 0)}</td>
                <td style="padding: 8px; text-align: right; color: #15803d; font-weight: 700;">${r.Total_row_validated_match}</td>
                <td style="padding: 8px; text-align: right; color: ${r.Total_row_validated_mismatch > 0 ? '#b91c1c' : '#64748b'}; font-weight: 700;">${r.Total_row_validated_mismatch}</td>
                <td style="padding: 8px; text-align: center;">
                    <a href="/api/ibc/download/${taskId}/${r.download}.xlsx" download style="color: #007bff; font-weight: 600; text-decoration: underline;">${r.download}</a>
                </td>
                <td style="padding: 8px; text-align: center;">
                    <button type="button" onclick="openIbcDrilldown('${taskId}', '${r.download}')" style="background: #007bff; color: #fff; border: none; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        🔍 Inspect
                    </button>
                </td>
            </tr>
        `).join("");

        // 3. Render Discovered IBC_Input Table
        const inputBody = document.getElementById("ibcInputTableBody");
        inputBody.innerHTML = (_ibcJobData.ibc_input || []).map(r => `
            <tr style="border-bottom: 1px solid #edf2f7;">
                <td style="padding: 8px;">${r["New Bill id"]}</td>
                <td style="padding: 8px;">${r["cancelled Bill ID"]}</td>
                <td style="padding: 8px; font-family: Consolas, monospace;">${r.SA_ID}</td>
                <td style="padding: 8px; font-family: Consolas, monospace;">${r.New_bseg}</td>
                <td style="padding: 8px; font-family: Consolas, monospace;">${r.old_bseg}</td>
                <td style="padding: 8px; font-weight: 700;">${r.Rate_cd}</td>
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
        bodyContainer.innerHTML = `<div style="padding: 20px; text-align: center;">Loading side-by-side table comparison...</div>`;

        try {
            const res = await fetch(`/api/ibc/drilldown/${taskId}/${dlKey}`);
            const data = await res.json();
            if (!res.ok || data.error) {
                bodyContainer.innerHTML = `<div style="color: red; padding: 15px;">${data.error || "Failed to load drilldown."}</div>`;
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
                        style="padding: 6px 12px; border: 1px solid #cbd5e0; border-radius: 4px; font-weight: 700; cursor: pointer; background: ${tName === _ibcActiveDrillTab ? '#003366' : '#f8fafc'}; color: ${tName === _ibcActiveDrillTab ? '#fff' : '#2d3748'};">
                        ${tName} (${rows.length}) ${mmCount > 0 ? `<span style="background:#dc3545;color:#fff;padding:1px 5px;border-radius:8px;font-size:10px;margin-left:4px;">${mmCount} diff</span>` : ''}
                    </button>
                `;
            }).join("");

            renderInspectorActiveSheet();
            inspector.scrollIntoView({ behavior: "smooth" });
        } catch (e) {
            bodyContainer.innerHTML = `<div style="color: red; padding: 15px;">Error: ${e.message}</div>`;
        }
    }

    function selectInspectorTab(tName) {
        _ibcActiveDrillTab = tName;
        document.querySelectorAll(".ibc-drill-tab-btn").forEach(btn => {
            const active = btn.getAttribute("data-tab") === tName;
            btn.style.background = active ? "#003366" : "#f8fafc";
            btn.style.color = active ? "#fff" : "#2d3748";
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
            bodyContainer.innerHTML = `<div style="padding: 20px; color: #718096; text-align: center;">No rows match the selected filter in ${_ibcActiveDrillTab}.</div>`;
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

        let html = `<table style="width: 100%; border-collapse: collapse; font-size: 11px; font-family: Consolas, monospace;">
            <thead>
                <tr style="background: #003366; color: #fff;">
                    <th colspan="${cols.length}" style="padding: 6px; border-right: 3px solid #fff;">New ID: ${nId}</th>
                    <th colspan="${cols.length}" style="padding: 6px; border-right: 3px solid #fff;">Old ID: ${oId}</th>
                    <th style="padding: 6px;">status</th>
                </tr>
                <tr style="background: #edf2f7; color: #1a202c; border-bottom: 2px solid #cbd5e0;">
                    ${cols.map(c => `<th style="padding: 6px;">${c}</th>`).join("")}
                    ${cols.map(c => `<th style="padding: 6px;">${c}</th>`).join("")}
                    <th style="padding: 6px;">status</th>
                </tr>
            </thead>
            <tbody>`;

        filtered.forEach(item => {
            const n = item.new || {};
            const o = item.old || {};
            const st = item.status;
            const bg = st === "match" ? "#f0fdf4" : (st.includes("exception") ? "#ecfeff" : "#fef2f2");
            const stColor = st === "match" ? "#15803d" : (st.includes("exception") ? "#0e7490" : "#b91c1c");

            html += `<tr style="background: ${bg}; border-bottom: 1px solid #e2e8f0;">
                ${cols.map(c => `<td style="padding: 5px;">${n[c] !== undefined && n[c] !== null ? n[c] : ""}</td>`).join("")}
                ${cols.map(c => `<td style="padding: 5px;">${o[c] !== undefined && o[c] !== null ? o[c] : ""}</td>`).join("")}
                <td style="padding: 5px; font-weight: bold; color: ${stColor};">${st}</td>
            </tr>`;
        });

        html += `</tbody></table>`;
        bodyContainer.innerHTML = html;
    }

    window.startIbcRun = startIbcRun;
    window.switchIbcTableTab = switchIbcTableTab;
    window.filterByBillPairHash = filterByBillPairHash;
    window.clearIbcBillHashFilter = clearIbcBillHashFilter;
    window.renderIbcTables = renderIbcTables;
    window.openIbcDrilldown = openIbcDrilldown;
    window.selectInspectorTab = selectInspectorTab;
    window.renderInspectorActiveSheet = renderInspectorActiveSheet;