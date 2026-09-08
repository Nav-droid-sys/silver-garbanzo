// app.js - Cyberpunk Terminal UI Interactive Engine
document.addEventListener("DOMContentLoaded", () => {
    
    // Application State
    const state = {
        filename: "abhinav-input-Jun'26dump.xlsx",
        filepath: null,
        assignees: [],
        selectedAssignee: null,
        currentTasks: [],
        currentMonthPeriods: [],
        filterMode: "all", // 'all', 'single', 'range'
        singleMonth: null,
        startMonth: null,
        endMonth: null,
        filteredTasks: []
    };

    // DOM Elements
    const elements = {
        currentFileName: document.getElementById("currentFileName"),
        dropzone: document.getElementById("dropzone"),
        fileInput: document.getElementById("fileInput"),
        searchAssignee: document.getElementById("searchAssignee"),
        assigneeList: document.getElementById("assigneeList"),
        assigneeCount: document.getElementById("assigneeCount"),
        
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
        
        terminalConsole: document.getElementById("terminalConsole"),
        cliInput: document.getElementById("cliInput"),
        
        previewTaskCount: document.getElementById("previewTaskCount"),
        previewFilterTag: document.getElementById("previewFilterTag"),
        previewTableBody: document.getElementById("previewTableBody"),
        
        dumpBtn: document.getElementById("dumpBtn"),
        downloadBtn: document.getElementById("downloadBtn"),
        
        themeBtns: document.querySelectorAll(".theme-btn"),
        toggleScanlines: document.getElementById("toggleScanlines"),
        scanlines: document.getElementById("scanlines")
    };

    // Logging helper
    function logTerminal(message, type = "info") {
        const line = document.createElement("div");
        line.className = `log-line ${type}`;
        const time = new Date().toLocaleTimeString();
        line.textContent = `[${time}] ${message}`;
        elements.terminalConsole.appendChild(line);
        elements.terminalConsole.scrollTop = elements.terminalConsole.scrollHeight;
    }

    // 1. Theme & Scanline Controls
    elements.themeBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            elements.themeBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            document.body.className = document.body.className.replace(/theme-\w+/g, "");
            const theme = btn.dataset.theme;
            document.body.classList.add(`theme-${theme}`);
            logTerminal(`Switched theme to [${theme.toUpperCase()}]`, "system");
        });
    });

    elements.toggleScanlines.addEventListener("click", () => {
        const active = document.body.classList.toggle("crt-overlay");
        elements.toggleScanlines.classList.toggle("active", active);
        elements.toggleScanlines.textContent = active ? "[SCANLINES: ON]" : "[SCANLINES: OFF]";
        logTerminal(`CRT Scanlines ${active ? "ENABLED" : "DISABLED"}`, "info");
    });

    // 2. Load Initial Workbook
    async function loadWorkbook() {
        try {
            logTerminal("Connecting to Task Extractor engine...", "system");
            const res = await fetch("/api/load");
            const data = await res.json();
            
            if (data.success) {
                state.filename = data.filename;
                state.assignees = data.assignees;
                elements.currentFileName.textContent = data.filename;
                elements.assigneeCount.textContent = data.total_assignees;
                
                renderAssigneeList(data.assignees);
                logTerminal(`Loaded dataset '${data.filename}' with ${data.total_assignees} assignees (${data.total_tasks} total tasks).`, "success");
            } else {
                logTerminal(`Error loading data: ${data.error}`, "error");
            }
        } catch (err) {
            logTerminal(`Failed to communicate with server: ${err.message}`, "error");
        }
    }

    // Render Assignee Directory List
    function renderAssigneeList(list) {
        elements.assigneeList.innerHTML = "";
        if (!list || list.length === 0) {
            elements.assigneeList.innerHTML = '<div class="empty-msg">No assignees found.</div>';
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
    elements.searchAssignee.addEventListener("input", (e) => {
        const q = e.target.value.toLowerCase().trim();
        const filtered = state.assignees.filter(a => a.name.toLowerCase().includes(q));
        renderAssigneeList(filtered);
    });

    // 3. Select Assignee
    async function selectAssignee(name) {
        state.selectedAssignee = name;
        elements.selectedAssigneeName.textContent = name;
        renderAssigneeList(state.assignees); // re-highlight

        logTerminal(`Selecting assignee '${name}'...`, "cmd");

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
                
                elements.targetTaskCount.textContent = data.total_tasks;
                elements.targetPointsCount.textContent = data.total_points.toFixed(2);
                
                renderMonthBadges(data.month_periods);
                applyFilter();
                
                elements.dumpBtn.disabled = false;
                logTerminal(`Extracted ${data.total_tasks} tasks for '${name}' (${data.total_points.toFixed(2)} pts).`, "success");
            } else {
                logTerminal(`Error fetching details for ${name}: ${data.error}`, "error");
            }
        } catch (err) {
            logTerminal(`API Error: ${err.message}`, "error");
        }
    }

    // Render Month Badges
    function renderMonthBadges(periods) {
        elements.monthBadges.innerHTML = "";
        periods.forEach(p => {
            const badge = document.createElement("button");
            badge.className = `month-badge ${state.singleMonth === p.month_year ? "selected" : ""}`;
            badge.textContent = `${p.month_year} (${p.count})`;
            badge.addEventListener("click", () => {
                state.singleMonth = p.month_year;
                renderMonthBadges(periods);
                applyFilter();
            });
            elements.monthBadges.appendChild(badge);
        });
    }

    // 4. Filter Modes (Tab Switcher)
    elements.filterTabs.forEach(tab => {
        tab.addEventListener("click", () => {
            elements.filterTabs.forEach(t => t.classList.remove("active"));
            tab.classList.add("active");
            state.filterMode = tab.dataset.mode;

            elements.singleMonthPanel.classList.toggle("hidden", state.filterMode !== "single");
            elements.dateRangePanel.classList.toggle("hidden", state.filterMode !== "range");

            logTerminal(`Filter mode set to [${state.filterMode.toUpperCase()}]`, "info");
            applyFilter();
        });
    });

    elements.applyRangeBtn.addEventListener("click", () => {
        state.startMonth = elements.startMonthInput.value.trim();
        state.endMonth = elements.endMonthInput.value.trim();
        applyFilter();
    });

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
        } else if (state.filterMode === "range" && (state.startMonth || elements.startMonthInput.value)) {
            const s = (state.startMonth || elements.startMonthInput.value).trim();
            const e = (state.endMonth || elements.endMonthInput.value).trim() || s;
            tag = `RANGE: ${s} to ${e}`;
            result = state.currentTasks.filter(t => t.month_year && t.month_year !== "Unknown Date");
        }

        state.filteredTasks = result;
        elements.previewTaskCount.textContent = result.length;
        elements.previewFilterTag.textContent = `Filter: ${tag}`;
        renderPreviewTable(result);
    }

    // Render Preview Table
    function renderPreviewTable(tasks) {
        elements.previewTableBody.innerHTML = "";
        if (!tasks || tasks.length === 0) {
            elements.previewTableBody.innerHTML = `<tr><td colspan="4" class="empty-msg">No tasks matching current filter criteria.</td></tr>`;
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

    // 5. Dump & Sync Action
    elements.dumpBtn.addEventListener("click", async () => {
        if (!state.selectedAssignee) {
            logTerminal("Please select an assignee first.", "warning");
            return;
        }

        logTerminal(`Syncing tasks for '${state.selectedAssignee}' to Master Excel...`, "cmd");
        elements.dumpBtn.disabled = true;

        try {
            const payload = {
                name: state.selectedAssignee,
                filepath: state.filepath,
                filter_type: state.filterMode,
                single_month: state.singleMonth,
                start_month: elements.startMonthInput.value.trim(),
                end_month: elements.endMonthInput.value.trim()
            };

            const res = await fetch("/api/process-dump", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });
            const data = await res.json();

            if (data.success) {
                logTerminal(`SUCCESS: Dumped ${data.task_count} tasks (${data.total_points} pts) into '${data.master_file}'.`, "success");
                logTerminal(`Filter Applied: ${data.filter_applied}`, "info");
                elements.downloadBtn.classList.add("pulse");
            } else {
                logTerminal(`Dump failed: ${data.error}`, "error");
            }
        } catch (err) {
            logTerminal(`Server error: ${err.message}`, "error");
        } finally {
            elements.dumpBtn.disabled = false;
        }
    });

    // Download Master Excel
    elements.downloadBtn.addEventListener("click", () => {
        logTerminal("Initiating Master Excel download...", "cmd");
        window.location.href = "/api/download-master";
    });

    // 6. Interactive CLI Terminal Console Emulator
    elements.cliInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            const raw = elements.cliInput.value.trim();
            if (!raw) return;
            elements.cliInput.value = "";

            logTerminal(`CYBERPUNK> ${raw}`, "cmd");
            parseCliCommand(raw);
        }
    });

    function parseCliCommand(cmdStr) {
        const parts = cmdStr.split(/\s+/);
        const cmd = parts[0].toLowerCase();
        const args = parts.slice(1);

        switch (cmd) {
            case "help":
                logTerminal("--- AVAILABLE CLI COMMANDS ---", "system");
                logTerminal("  list                      : List all available assignees", "info");
                logTerminal("  select <num_or_name>      : Select an assignee by number or name", "info");
                logTerminal("  filter all                : Dump all tasks without date filter", "info");
                logTerminal("  filter single <month>     : Filter by single month (e.g. filter single Jan 2025)", "info");
                logTerminal("  filter range <start> <end>: Filter by range (e.g. filter range Jan 2025 Jun 2026)", "info");
                logTerminal("  dump / sync               : Sync filtered data to Master Excel", "info");
                logTerminal("  download                  : Download Billing_Extracted_Report.xlsx", "info");
                logTerminal("  theme <neon|amber|matrix> : Switch UI visual theme", "info");
                logTerminal("  clear                     : Clear terminal log screen", "info");
                break;

            case "clear":
                elements.terminalConsole.innerHTML = "";
                break;

            case "list":
                state.assignees.forEach((a, i) => logTerminal(`  ${i + 1}. ${a.name} (${a.count} tasks)`, "info"));
                break;

            case "select":
                if (args.length === 0) {
                    logTerminal("Usage: select <number_or_name>", "warning");
                    return;
                }
                const target = args.join(" ");
                const num = parseInt(target);
                if (!isNaN(num) && num >= 1 && num <= state.assignees.length) {
                    selectAssignee(state.assignees[num - 1].name);
                } else {
                    const match = state.assignees.find(a => a.name.toLowerCase().includes(target.toLowerCase()));
                    if (match) selectAssignee(match.name);
                    else logTerminal(`No match for assignee '${target}'`, "error");
                }
                break;

            case "filter":
                if (args[0] === "all") {
                    document.querySelector('.filter-tab[data-mode="all"]').click();
                } else if (args[0] === "single" && args[1]) {
                    document.querySelector('.filter-tab[data-mode="single"]').click();
                    state.singleMonth = args.slice(1).join(" ");
                    applyFilter();
                } else if (args[0] === "range" && args[1]) {
                    document.querySelector('.filter-tab[data-mode="range"]').click();
                    elements.startMonthInput.value = args[1] + (args[2] ? " " + args[2] : "");
                    elements.endMonthInput.value = args[3] ? (args[3] + (args[4] ? " " + args[4] : "")) : elements.startMonthInput.value;
                    state.startMonth = elements.startMonthInput.value;
                    state.endMonth = elements.endMonthInput.value;
                    applyFilter();
                } else {
                    logTerminal("Usage: filter all | filter single <Month YYYY> | filter range <Start> <End>", "warning");
                }
                break;

            case "dump":
            case "sync":
                elements.dumpBtn.click();
                break;

            case "download":
                elements.downloadBtn.click();
                break;

            case "theme":
                if (["neon", "amber", "matrix"].includes(args[0])) {
                    document.querySelector(`.theme-btn[data-theme="${args[0]}"]`).click();
                } else {
                    logTerminal("Usage: theme neon | theme amber | theme matrix", "warning");
                }
                break;

            default:
                logTerminal(`Unknown command '${cmd}'. Type 'help' for available commands.`, "error");
        }
    }

    // 7. Drag & Drop File Upload
    elements.dropzone.addEventListener("click", () => elements.fileInput.click());

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

    elements.fileInput.addEventListener("change", () => {
        if (elements.fileInput.files.length > 0) {
            handleFileUpload(elements.fileInput.files[0]);
        }
    });

    async function handleFileUpload(file) {
        logTerminal(`Uploading file '${file.name}'...`, "cmd");
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
                elements.currentFileName.textContent = data.filename;
                elements.assigneeCount.textContent = data.total_assignees;

                renderAssigneeList(data.assignees);
                logTerminal(`Successfully uploaded and parsed dataset '${data.filename}' (${data.total_assignees} assignees).`, "success");
            } else {
                logTerminal(`Upload error: ${data.error}`, "error");
            }
        } catch (err) {
            logTerminal(`Upload failed: ${err.message}`, "error");
        }
    }

    // Initialize App
    loadWorkbook();
});
