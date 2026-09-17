// app.js - Wipro Task Extractor Multi-Page Interactive Engine
document.addEventListener("DOMContentLoaded", () => {
    
    // Application State with sessionStorage persistence
    const state = {
        filename: sessionStorage.getItem("filename") || "abhinav-input-Jun'26dump.xlsx",
        filepath: sessionStorage.getItem("filepath") || null,
        assignees: [],
        selectedAssignee: sessionStorage.getItem("selectedAssignee") || null,
        currentTasks: [],
        currentMonthPeriods: [],
        filterMode: sessionStorage.getItem("filterMode") || "all", // 'all', 'single', 'range'
        singleMonth: sessionStorage.getItem("singleMonth") || null,
        startMonth: sessionStorage.getItem("startMonth") || null,
        endMonth: sessionStorage.getItem("endMonth") || null,
        filteredTasks: []
    };

    // DOM Elements across pages
    const elements = {
        hdrFileName: document.getElementById("hdrFileName"),
        hdrAssigneeName: document.getElementById("hdrAssigneeName"),

        // Dataset Page Elements
        currentFileName: document.getElementById("currentFileName"),
        dropzone: document.getElementById("dropzone"),
        fileInput: document.getElementById("fileInput"),
        searchAssignee: document.getElementById("searchAssignee"),
        assigneeList: document.getElementById("assigneeList"),
        assigneeCount: document.getElementById("assigneeCount"),
        proceedToExtractorBtn: document.getElementById("proceedToExtractorBtn"),
        
        // Extractor Page Elements
        selectedAssigneeName: document.getElementById("selectedAssigneeName"),
        targetTaskCount: document.getElementById("targetTaskCount"),
        targetPointsCount: document.getElementById("targetPointsCount"),
        
        filterTabs: document.querySelectorAll(".filter-tab"),
        singleMonthPanel: document.getElementById("singleMonthPanel"),
        dateRangePanel: document.getElementById("dateRangePanel"),
        monthBadges: document.getElementById("monthBadges"),
        startMonthInput: document.getElementById("startMonthInput"),
        endMonthInput: document.getElementById("endMonthInput"),
        applyRangeBtn: document.getElementById("applyRangeBtn"),
        
        previewTaskCount: document.getElementById("previewTaskCount"),
        previewFilterTag: document.getElementById("previewFilterTag"),
        previewTableBody: document.getElementById("previewTableBody"),
        dumpBtn: document.getElementById("dumpBtn"),
        
        // Reports & Activity Page Elements
        downloadBtn: document.getElementById("downloadBtn"),
        clearLogBtn: document.getElementById("clearLogBtn"),
        terminalConsole: document.getElementById("terminalConsole"),
        cliInput: document.getElementById("cliInput")
    };

    // Update Persistent Context Header Bar
    function updateContextBar() {
        if (elements.hdrFileName) elements.hdrFileName.textContent = state.filename;
        if (elements.hdrAssigneeName) {
            elements.hdrAssigneeName.textContent = state.selectedAssignee || "None Selected";
            elements.hdrAssigneeName.classList.toggle("highlight", !!state.selectedAssignee);
        }
        if (elements.currentFileName) elements.currentFileName.textContent = state.filename;
        if (elements.selectedAssigneeName) elements.selectedAssigneeName.textContent = state.selectedAssignee || "Select an Assignee";
    }

    // Save State to sessionStorage
    function saveState() {
        sessionStorage.setItem("filename", state.filename);
        if (state.filepath) sessionStorage.setItem("filepath", state.filepath);
        if (state.selectedAssignee) sessionStorage.setItem("selectedAssignee", state.selectedAssignee);
        sessionStorage.setItem("filterMode", state.filterMode);
        if (state.singleMonth) sessionStorage.setItem("singleMonth", state.singleMonth);
        updateContextBar();
    }

    // Logging & Activity helper
    function logActivity(message, type = "info") {
        if (!elements.terminalConsole) return;
        const line = document.createElement("div");
        line.className = `log-line ${type}`;
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        line.innerHTML = `<span class="activity-time">[${time}]</span> <span class="activity-msg">${message}</span>`;
        elements.terminalConsole.appendChild(line);
        elements.terminalConsole.scrollTop = elements.terminalConsole.scrollHeight;
    }

    // 1. Initial Data Fetching
    async function loadWorkbook() {
        try {
            logActivity("Connecting to Task Extractor engine...", "system");
            const res = await fetch("/api/load");
            const data = await res.json();
            
            if (data.success) {
                state.filename = data.filename;
                state.assignees = data.assignees;
                saveState();
                
                if (elements.assigneeCount) elements.assigneeCount.textContent = data.total_assignees;
                renderAssigneeList(data.assignees);
                
                logActivity(`Loaded dataset '${data.filename}' with ${data.total_assignees} assignees (${data.total_tasks} total tasks).`, "success");

                // Auto-load details if an assignee was previously selected
                if (state.selectedAssignee) {
                    fetchAssigneeDetails(state.selectedAssignee);
                }
            } else {
                logActivity(`Error loading dataset: ${data.error}`, "error");
            }
        } catch (err) {
            logActivity(`Failed to communicate with server: ${err.message}`, "error");
        }
    }

    // Render Assignee Directory List
    function renderAssigneeList(list) {
        if (!elements.assigneeList) return;
        elements.assigneeList.innerHTML = "";
        if (!list || list.length === 0) {
            elements.assigneeList.innerHTML = '<div class="loading-state">No assignees found matching search filter.</div>';
            return;
        }

        list.forEach((item, idx) => {
            const card = document.createElement("div");
            card.className = `assignee-card ${state.selectedAssignee === item.name ? "selected" : ""}`;
            card.innerHTML = `
                <span class="card-name">${idx + 1}. ${item.name}</span>
                <span class="card-cnt">${item.count} tasks</span>
            `;
            card.addEventListener("click", () => selectAssignee(item.name));
            elements.assigneeList.appendChild(card);
        });
    }

    // Search Assignee Filter
    if (elements.searchAssignee) {
        elements.searchAssignee.addEventListener("input", (e) => {
            const q = e.target.value.toLowerCase().trim();
            const filtered = state.assignees.filter(a => a.name.toLowerCase().includes(q));
            renderAssigneeList(filtered);
        });
    }

    // Select Assignee
    function selectAssignee(name) {
        state.selectedAssignee = name;
        saveState();
        renderAssigneeList(state.assignees);

        if (elements.proceedToExtractorBtn) {
            elements.proceedToExtractorBtn.disabled = false;
        }

        logActivity(`Selected target assignee '${name}'`, "cmd");
        fetchAssigneeDetails(name);
    }

    // Fetch Assignee Task Details
    async function fetchAssigneeDetails(name) {
        try {
            const res = await fetch("/api/assignee-details", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ name, filepath: state.filepath })
            });
            const data = await res.json();

            if (data.success) {
                state.currentTasks = data.tasks;
                state.currentMonthPeriods = data.month_periods;
                
                if (elements.targetTaskCount) elements.targetTaskCount.textContent = data.total_tasks;
                if (elements.targetPointsCount) elements.targetPointsCount.textContent = data.total_points.toFixed(2);
                
                renderMonthBadges(data.month_periods);
                applyFilter();
                
                if (elements.dumpBtn) elements.dumpBtn.disabled = false;
                logActivity(`Loaded ${data.total_tasks} tasks for '${name}' (${data.total_points.toFixed(2)} pts).`, "success");
            } else {
                logActivity(`Error loading details for ${name}: ${data.error}`, "error");
            }
        } catch (err) {
            logActivity(`API Error: ${err.message}`, "error");
        }
    }

    // Render Month Badges
    function renderMonthBadges(periods) {
        if (!elements.monthBadges) return;
        elements.monthBadges.innerHTML = "";
        periods.forEach(p => {
            const badge = document.createElement("button");
            badge.className = `month-badge ${state.singleMonth === p.month_year ? "selected" : ""}`;
            badge.textContent = `${p.month_year} (${p.count})`;
            badge.addEventListener("click", () => {
                state.singleMonth = p.month_year;
                saveState();
                renderMonthBadges(periods);
                applyFilter();
            });
            elements.monthBadges.appendChild(badge);
        });
    }

    // Filter Mode Tab Switcher
    if (elements.filterTabs) {
        elements.filterTabs.forEach(tab => {
            tab.addEventListener("click", () => {
                elements.filterTabs.forEach(t => t.classList.remove("active"));
                tab.classList.add("active");
                state.filterMode = tab.dataset.mode;
                saveState();

                if (elements.singleMonthPanel) elements.singleMonthPanel.classList.toggle("hidden", state.filterMode !== "single");
                if (elements.dateRangePanel) elements.dateRangePanel.classList.toggle("hidden", state.filterMode !== "range");

                logActivity(`Filter mode set to [${state.filterMode.toUpperCase()}]`, "info");
                applyFilter();
            });
        });
    }

    if (elements.applyRangeBtn) {
        elements.applyRangeBtn.addEventListener("click", () => {
            state.startMonth = elements.startMonthInput.value.trim();
            state.endMonth = elements.endMonthInput.value.trim();
            saveState();
            applyFilter();
        });
    }

    // Apply Filter to Preview Table
    function applyFilter() {
        if (!state.currentTasks || state.currentTasks.length === 0) {
            renderPreviewTable([]);
            return;
        }

        let result = state.currentTasks;
        let tag = "ALL TASKS";

        if (state.filterMode === "single" && state.singleMonth) {
            const q = state.singleMonth.toLowerCase();
            result = state.currentTasks.filter(t => (t.month_year || "").toLowerCase().includes(q));
            tag = `SINGLE MONTH: ${state.singleMonth}`;
        } else if (state.filterMode === "range" && (state.startMonth || (elements.startMonthInput && elements.startMonthInput.value))) {
            const s = (state.startMonth || elements.startMonthInput.value).trim();
            const e = (state.endMonth || (elements.endMonthInput ? elements.endMonthInput.value : "")).trim() || s;
            tag = `RANGE: ${s} to ${e}`;
            result = state.currentTasks.filter(t => t.month_year && t.month_year !== "Unknown Date");
        }

        state.filteredTasks = result;
        if (elements.previewTaskCount) elements.previewTaskCount.textContent = result.length;
        if (elements.previewFilterTag) elements.previewFilterTag.textContent = `Filter: ${tag}`;
        renderPreviewTable(result);
    }

    // Render Preview Table
    function renderPreviewTable(tasks) {
        if (!elements.previewTableBody) return;
        elements.previewTableBody.innerHTML = "";
        if (!tasks || tasks.length === 0) {
            elements.previewTableBody.innerHTML = `<tr><td colspan="4" class="empty-state">No tasks matching current filter criteria.</td></tr>`;
            return;
        }

        tasks.forEach(t => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><strong>${t.task_id}</strong></td>
                <td>${t.points.toFixed(2)}</td>
                <td><span class="ms-pill">${t.milestone}</span></td>
                <td>${t.month_year || '-'}</td>
            `;
            elements.previewTableBody.appendChild(tr);
        });
    }

    // Proceed to Task Extractor Button (on Dataset Page)
    if (elements.proceedToExtractorBtn) {
        if (state.selectedAssignee) elements.proceedToExtractorBtn.disabled = false;
        elements.proceedToExtractorBtn.addEventListener("click", () => {
            if (!state.selectedAssignee) return;
            window.location.href = "/extractor";
        });
    }

    // Dump & Sync Action
    if (elements.dumpBtn) {
        if (state.selectedAssignee) elements.dumpBtn.disabled = false;
        elements.dumpBtn.addEventListener("click", async () => {
            if (!state.selectedAssignee) {
                logActivity("Please select an assignee first.", "warning");
                return;
            }

            logActivity(`Syncing tasks for '${state.selectedAssignee}' to Master Excel...`, "cmd");
            elements.dumpBtn.disabled = true;

            try {
                const payload = {
                    name: state.selectedAssignee,
                    filepath: state.filepath,
                    filter_type: state.filterMode,
                    single_month: state.singleMonth,
                    start_month: elements.startMonthInput ? elements.startMonthInput.value.trim() : "",
                    end_month: elements.endMonthInput ? elements.endMonthInput.value.trim() : ""
                };

                const res = await fetch("/api/process-dump", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(payload)
                });
                const data = await res.json();

                if (data.success) {
                    logActivity(`SUCCESS: Synced ${data.task_count} tasks (${data.total_points} pts) into '${data.master_file}'.`, "success");
                    logActivity(`Filter Applied: ${data.filter_applied}`, "info");
                } else {
                    logActivity(`Dump failed: ${data.error}`, "error");
                }
            } catch (err) {
                logActivity(`Server error: ${err.message}`, "error");
            } finally {
                elements.dumpBtn.disabled = false;
            }
        });
    }

    // Download Master Excel
    if (elements.downloadBtn) {
        elements.downloadBtn.addEventListener("click", () => {
            logActivity("Downloading Billing_Extracted_Report.xlsx...", "cmd");
            window.location.href = "/api/download-master";
        });
    }

    // Clear Activity Log
    if (elements.clearLogBtn) {
        elements.clearLogBtn.addEventListener("click", () => {
            if (elements.terminalConsole) elements.terminalConsole.innerHTML = "";
        });
    }

    // Quick Command Console Parser
    if (elements.cliInput) {
        elements.cliInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                const raw = elements.cliInput.value.trim();
                if (!raw) return;
                elements.cliInput.value = "";

                logActivity(`> ${raw}`, "cmd");
                parseCliCommand(raw);
            }
        });
    }

    function parseCliCommand(cmdStr) {
        const parts = cmdStr.split(/\s+/);
        const cmd = parts[0].toLowerCase();
        const args = parts.slice(1);

        switch (cmd) {
            case "help":
                logActivity("--- AVAILABLE QUICK COMMANDS ---", "system");
                logActivity("  list                      : List all available assignees", "info");
                logActivity("  select <num_or_name>      : Select an assignee by number or name", "info");
                logActivity("  filter all                : Dump all tasks without date filter", "info");
                logActivity("  filter single <month>     : Filter by single month (e.g. filter single Jan 2025)", "info");
                logActivity("  filter range <start> <end>: Filter by range (e.g. filter range Jan 2025 Jun 2026)", "info");
                logActivity("  dump / sync               : Sync filtered data to Master Excel", "info");
                logActivity("  download                  : Download Billing_Extracted_Report.xlsx", "info");
                logActivity("  clear                     : Clear activity log feed", "info");
                break;

            case "clear":
                if (elements.terminalConsole) elements.terminalConsole.innerHTML = "";
                break;

            case "list":
                state.assignees.forEach((a, i) => logActivity(`  ${i + 1}. ${a.name} (${a.count} tasks)`, "info"));
                break;

            case "select":
                if (args.length === 0) {
                    logActivity("Usage: select <number_or_name>", "warning");
                    return;
                }
                const target = args.join(" ");
                const num = parseInt(target);
                if (!isNaN(num) && num >= 1 && num <= state.assignees.length) {
                    selectAssignee(state.assignees[num - 1].name);
                } else {
                    const match = state.assignees.find(a => a.name.toLowerCase().includes(target.toLowerCase()));
                    if (match) selectAssignee(match.name);
                    else logActivity(`No match found for assignee '${target}'`, "error");
                }
                break;

            case "dump":
            case "sync":
                if (elements.dumpBtn) elements.dumpBtn.click();
                break;

            case "download":
                if (elements.downloadBtn) elements.downloadBtn.click();
                break;

            default:
                logActivity(`Unknown command '${cmd}'. Type 'help' for available actions.`, "error");
        }
    }

    // Drag & Drop File Upload
    if (elements.dropzone) {
        elements.dropzone.addEventListener("click", () => {
            if (elements.fileInput) elements.fileInput.click();
        });

        elements.dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
            elements.dropzone.classList.add("dragover");
        });

        elements.dropzone.addEventListener("dragleave", () => elements.dropzone.classList.remove("dragover"));

        elements.dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
            elements.dropzone.classList.remove("dragover");
            if (e.dataTransfer.files.length > 0) {
                handleFileUpload(e.dataTransfer.files[0]);
            }
        });
    }

    if (elements.fileInput) {
        elements.fileInput.addEventListener("change", () => {
            if (elements.fileInput.files.length > 0) {
                handleFileUpload(elements.fileInput.files[0]);
            }
        });
    }

    async function handleFileUpload(file) {
        logActivity(`Uploading dataset '${file.name}'...`, "cmd");
        const formData = new FormData();
        formData.append("file", file);

        try {
            const res = await fetch("/api/upload", {
                method: "POST",
                body: formData
            });
            const data = await res.json();

            if (data.success) {
                state.filename = data.filename;
                state.filepath = data.filepath;
                state.assignees = data.assignees;
                saveState();

                if (elements.assigneeCount) elements.assigneeCount.textContent = data.total_assignees;
                renderAssigneeList(data.assignees);
                logActivity(`Successfully uploaded dataset '${data.filename}' (${data.total_assignees} assignees).`, "success");
            } else {
                logActivity(`Upload error: ${data.error}`, "error");
            }
        } catch (err) {
            logActivity(`Upload failed: ${err.message}`, "error");
        }
    }

    // Initialize App
    updateContextBar();
    loadWorkbook();
});
