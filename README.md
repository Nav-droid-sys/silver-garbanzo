# ⚡ Cyberpunk Task Extractor & Billing Sync HUD

A powerful, interactive **Retro Cyberpunk Terminal UI Web Application** and CLI tool to extract Task IDs, Estimator Points, Milestone Status, and Date Periods from Jira Excel dumps and sync them directly into a Master Excel Report (`Billing_Extracted_Report.xlsx`).

---

## 📁 Repository Folder Structure

```text
Billing Test/
├── backend/
│   ├── app.py                   # Flask REST API Web Server
│   └── task_extractor.py        # Core extraction & Excel sync engine
├── frontend/
│   ├── static/
│   │   ├── css/
│   │   │   └── style.css        # Retro Cyberpunk / CRT HUD styling
│   │   └── js/
│   │       └── app.js           # Interactive Terminal UI & CLI emulator
│   └── templates/
│       └── index.html           # HTML HUD Dashboard Structure
├── run.py                       # Root launcher for Web Server
├── requirements.txt             # Python dependencies
├── README.md                    # Setup & Usage Guide
└── abhinav-input-Jun'26dump.xlsx # Default input dataset
```

---

## 🚀 Quick Start Guide

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/billing-task-extractor.git
cd billing-task-extractor
```

### 2. Install Dependencies
Make sure you have Python 3.8+ installed. Install the required Python packages:
```bash
pip install -r requirements.txt
```

---

## 🖥️ Running the Application

### Option A: Web Application (Cyberpunk Terminal UI) - Recommended
Run the root server launcher:
```bash
python run.py
```
Or run the backend server directly:
```bash
python backend/app.py
```
Open your web browser and navigate to:
👉 **`http://127.0.0.1:5000`**

#### Web Features:
- **Drag & Drop Upload**: Upload any Excel dump file (`.xlsx`).
- **Assignee Directory**: Dynamic search bar and interactive resource cards.
- **Filter Modes**:
  - `[1] DUMP ALL`: Extract all tasks for chosen resource.
  - `[2] SINGLE MONTH`: Pick specific Month & Year (e.g. `Jan 2025`).
  - `[3] DATE RANGE`: Filter by range (e.g. `Jan 2025` to `Jun 2026`).
- **Interactive CLI Line (`>_`)**: Type commands in the web terminal (`select Karthick`, `filter range Jan 2025 Jun 2026`, `dump`, `download`, `help`).
- **Master Excel Download**: One-click download button for `Billing_Extracted_Report.xlsx`.
- **Deduplication**: Automatically skips tasks already dumped in the Master report to prevent repeated duplicate rows.

---

### Option B: Terminal CLI Mode (Console Only)
If you prefer running directly inside your terminal window:
```bash
python backend/task_extractor.py
```
Follow the interactive numbered prompt in your terminal to select resources, filter by Month/Year or Date Range, and save to Master Excel.

---

## 📜 Dependencies
- `Python 3.8+`
- `openpyxl >= 3.1.0`
- `Flask >= 3.0.0`
- `Werkzeug >= 3.0.0`
