from pathlib import Path
from copy import copy
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "input sheet.xlsx"
TEMPLATE_FILE = BASE_DIR / "Output.xlsx"
OUTPUT_FILE = BASE_DIR / "Automated_Output1.xlsx"

JIRA_SHEET = "Jira Dump"
RESOURCE_SHEET = "Resources List"
OUTPUT_SHEET = "Mobily billing actuals"

JIRA_HEADER_ROW = 7
RESOURCE_HEADER_ROW = 7

DEFAULT_FACTOR = 1.30

# Jira -> output milestone split from the demo
MILESTONE_SPLIT = {
    "UAT": 0.60,
    "RFS": 0.20,
    "PAC": 0.20,
}

# Output columns
C = {
    "SNO": 1,
    "RESOURCE": 2,
    "DATE": 3,
    "TASK": 4,
    "POINTS": 5,
    "MILESTONE": 6,
    "CALENDAR": 7,
    "INVOICE": 8,
    "COMMENTS": 9,
    "DUE": 10,
    "EARNED": 11,
    "BILLED": 12,
    "BALANCE": 13,
    "UNBILLED_INR": 14,
    "UNBILLED_POINTS": 15,
    "UNBILLED_AMOUNT": 16,
    "FACTOR": 17,
    "EMPLOYEE": 18,
    "RATE": 19,
    "JAN": 20,
    "DEC": 31,
    "TOTAL": 32,
}


# ============================================================
# HELPERS
# ============================================================

def norm(value):
    if value is None:
        return ""
    return str(value).strip().upper()


def clean(value):
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def header_map(ws, row):
    result = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row, col).value
        if value is not None:
            key = norm(value)
            if key and key not in result:
                result[key] = col
    return result


def find_col(headers, *names):
    for name in names:
        if norm(name) in headers:
            return headers[norm(name)]
    return None


def month_label(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%b '%y")
    return str(value).strip()


def parse_month_year(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.year, value.month
    if hasattr(value, "year") and hasattr(value, "month"):
        try:
            return int(value.year), int(value.month)
        except Exception:
            pass
    text = str(value).strip()
    if not text:
        return None

    # Comments contain labels such as "UAT - Oct '24". Extract the
    # month/year portion wherever it appears in the text.
    import re
    match = re.search(
        r"\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s*[-,']?\s*(\d{2}|\d{4})\b",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        months = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4,
            "may": 5, "jun": 6, "jul": 7, "aug": 8,
            "sep": 9, "oct": 10, "nov": 11, "dec": 12,
        }
        year = int(match.group(2))
        if year < 100:
            year += 2000
        return year, months[match.group(1)[:3].lower()]

    formats = ["%b-%y", "%b %y", "%b'%y", "%b '%y", "%B-%y", "%B %y", "%B'%y", "%B '%y", "%m-%Y", "%m/%Y", "%Y-%m", "%Y/%m", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y"]
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.year, parsed.month
        except ValueError:
            pass
    normalized = text.replace("-", " ").replace("/", " ").replace("'", " ")
    parts = normalized.split()
    if len(parts) >= 2:
        months = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        m = parts[0][:3].lower()
        if m in months:
            try:
                year = int(parts[1])
                if year < 100:
                    year += 2000
                return year, months[m]
            except ValueError:
                pass
    return None

def add_months(year, month, offset):
    total = year * 12 + (month - 1) + offset
    return total // 12, total % 12 + 1

def format_month_year(year, month):
    return datetime(year, month, 1).strftime("%b '%y")

def calculate_invoice_months(dac_month, rfs_date=None):
    """
    Build the three Comments sub-rows from Jira Dump -> DAC Month.

    Required business rule:
      UAT = DAC Month
      RFS = one month after DAC Month
      PAC = two months after DAC Month

    The Jira RFS date is intentionally NOT used here, because the Comments
    column is driven by the DAC Month rule shown in the required output.
    """
    uat = parse_month_year(dac_month)
    if not uat:
        return {"uat": "", "rfs": "", "pac": ""}

    uy, um = uat
    ry, rm = add_months(uy, um, 1)
    py, pm = add_months(uy, um, 2)

    return {
        "uat": format_month_year(uy, um),
        "rfs": format_month_year(ry, rm),
        "pac": format_month_year(py, pm),
    }


def copy_style(source, target):
    if source.has_style:
        target._style = copy(source._style)
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format
    target.protection = copy(source.protection)


def copy_row_style(template_ws, template_row, ws, output_row):
    if template_row in template_ws.row_dimensions:
        ws.row_dimensions[output_row].height = (
            template_ws.row_dimensions[template_row].height
        )

    for col in range(1, template_ws.max_column + 1):
        copy_style(
            template_ws.cell(template_row, col),
            ws.cell(output_row, col)
        )


# ============================================================
# RESOURCE MASTER
# ============================================================

def read_resources(wb):
    ws = wb[RESOURCE_SHEET]
    headers = header_map(ws, RESOURCE_HEADER_ROW)

    name_col = find_col(headers, "Name as per Jira")
    employee_col = find_col(headers, "Emp.Name")
    doj_col = find_col(headers, "DOJ")
    factor_col = find_col(headers, "Factor")
    rate_col = find_col(headers, "PPP as per Contract")

    if not name_col:
        raise ValueError("Could not find 'Name as per Jira' in Resources List.")

    resources = {}

    for row in range(RESOURCE_HEADER_ROW + 1, ws.max_row + 1):
        jira_name = clean(ws.cell(row, name_col).value)
        if not jira_name:
            continue

        factor = (
            ws.cell(row, factor_col).value
            if factor_col else DEFAULT_FACTOR
        )
        if not isinstance(factor, (int, float)):
            factor = DEFAULT_FACTOR

        employee = (
            clean(ws.cell(row, employee_col).value)
            if employee_col else jira_name
        )

        rate = (
            ws.cell(row, rate_col).value
            if rate_col else None
        )

        # Resources List monthly values are read from T:AE.
        monthly = {}
        for col in range(20, 32):
            header_value = ws.cell(RESOURCE_HEADER_ROW, col).value
            if header_value is not None:
                monthly[norm(header_value)] = ws.cell(row, col).value

        resources[norm(jira_name)] = {
            "jira_name": jira_name,
            "employee": employee or jira_name,
            "date": ws.cell(row, doj_col).value if doj_col else None,
            "factor": factor,
            "rate": rate,
            "monthly": monthly,
        }

    return resources


# ============================================================
# JIRA DATA
# ============================================================

def read_jira(wb_values, wb_formula):
    ws = wb_values[JIRA_SHEET]
    ws_formula = wb_formula[JIRA_SHEET]

    headers = header_map(ws_formula, JIRA_HEADER_ROW)

    task_col = find_col(headers, "Task ID")
    resource_col = find_col(
        headers,
        "Current Assignee: Name",
        "Current Assignee:  Name"
    )
    points_col = find_col(headers, "Project Total Points")
    milestone_col = find_col(headers, "Milestone Reached")
    price_col = find_col(headers, "Price per point")
    remarks_col = find_col(headers, "Remarks")
    dac_month_col = find_col(headers, "DAC Month")

    # In the supplied Jira Dump layout these are:
    # N = UAT date, O = RFS date, P = PAC date
    uat_date_col = 14
    rfs_date_col = 15
    pac_date_col = 16

    # In the supplied Jira Dump layout, row 6/7 shows
    # columns N:P (14:16) under the "Billed Month" group:
    # N = UAT billed month, O = RFS billed month, P = PAC billed month.
    billed_uat_col = 14
    billed_rfs_col = 15
    billed_pac_col = 16

    # In the supplied Jira Dump layout:
    # V:X = UAT/RFS/PAC billable points
    billable_uat_col = 22
    billable_rfs_col = 23
    billable_pac_col = 24

    if not task_col:
        raise ValueError("Could not find 'Task ID' in Jira Dump.")
    if not resource_col:
        raise ValueError("Could not find Current Assignee in Jira Dump.")
    if not points_col:
        raise ValueError("Could not find Project Total Points in Jira Dump.")

    records = {}

    milestone_rank = {
        "": 0,
        "RFSIT": 1,
        "RFS": 2,
        "PAC": 3,
    }

    for row in range(JIRA_HEADER_ROW + 1, ws.max_row + 1):
        task_id = clean(ws.cell(row, task_col).value)
        if not task_id:
            continue

        resource = clean(ws.cell(row, resource_col).value)

        points = ws.cell(row, points_col).value or 0

        uat_date = ws.cell(row, uat_date_col).value
        rfs_date = ws.cell(row, rfs_date_col).value
        pac_date = ws.cell(row, pac_date_col).value

        dac_month = ws.cell(row, dac_month_col).value if dac_month_col else None
        invoice_months = calculate_invoice_months(dac_month, rfs_date)

        milestone = (
            clean(ws.cell(row, milestone_col).value)
            if milestone_col else None
        )
        milestone = norm(milestone)

        # Important fallback: openpyxl does not calculate formulas.
        # If Excel's cached Milestone Reached is blank, calculate it.
        if not milestone:
            if uat_date and rfs_date and pac_date:
                milestone = "PAC"
            elif uat_date and rfs_date:
                milestone = "RFS"
            elif uat_date:
                milestone = "RFSIT"
            else:
                milestone = ""

        record = {
            "task_id": str(task_id).strip(),
            "resource": resource or "Unassigned",
            "points": points,
            "milestone": milestone,
            "uat_date": uat_date,
            "rfs_date": rfs_date,
            "pac_date": pac_date,
            "dac_month": dac_month,
            "invoice_months": invoice_months,
            "billed_uat": clean(ws.cell(row, billed_uat_col).value),
            "billed_rfs": clean(ws.cell(row, billed_rfs_col).value),
            "billed_pac": clean(ws.cell(row, billed_pac_col).value),
            "billable_uat": ws.cell(row, billable_uat_col).value or 0,
            "billable_rfs": ws.cell(row, billable_rfs_col).value or 0,
            "billable_pac": ws.cell(row, billable_pac_col).value or 0,
            "price": (
                ws.cell(row, price_col).value
                if price_col else None
            ),
            "remarks": (
                ws.cell(row, remarks_col).value
                if remarks_col else None
            ),
            "source_row": row,
        }

        key = norm(task_id)

        # If duplicate Task IDs exist, keep the record with
        # the furthest milestone reached.
        if key not in records:
            records[key] = record
        else:
            old = records[key]
            if milestone_rank.get(record["milestone"], 0) > milestone_rank.get(
                old["milestone"], 0
            ):
                records[key] = record

    return records


# ============================================================
# GROUP BY RESOURCE
# ============================================================

def group_tasks(records):
    groups = {}

    for record in records.values():
        key = norm(record["resource"])

        if key not in groups:
            groups[key] = {
                "resource": record["resource"],
                "tasks": [],
            }

        groups[key]["tasks"].append(record)

    for group in groups.values():
        group["tasks"].sort(
            key=lambda x: norm(x["task_id"])
        )

    return dict(sorted(groups.items()))


# ============================================================
# COLORS / TEMPLATE FORMATTING
# ============================================================

def apply_header_colors(ws):
    # We preserve the template's actual style first, then reinforce
    # the known header color groups from the uploaded Output.xlsx.

    for col in range(1, 38):
        cell = ws.cell(1, col)

        # Keep template font/alignment.
        cell.alignment = copy(cell.alignment)

        if 1 <= col <= 6:
            # Blue
            cell.fill = PatternFill(
                "solid",
                fgColor="BDD7EE"
            )

        elif 7 <= col <= 14:
            # Light yellow
            cell.fill = PatternFill(
                "solid",
                fgColor="FFF2CB"
            )

        elif 15 <= col <= 16:
            # Gray
            cell.fill = PatternFill(
                "solid",
                fgColor="D8D8D8"
            )

        elif 18 <= col <= 32:
            # Gold
            cell.fill = PatternFill(
                "solid",
                fgColor="FFC000"
            )


# ============================================================
# TASK ROW COLOR LOGIC
#
# Three fills are used:
#
#   WHITE = the milestone's data belongs to year 2025 (expected month
#           and/or actual month falls in 2025). This overrides GREEN/RED.
#   GREEN = Invoice Due (actual completion month) is on or before
#           the Comments (expected) month.
#   RED   = Invoice Due is after the expected month, OR Invoice Due
#           is blank (not completed yet).
#
# Each UAT / RFS / PAC row is evaluated independently using its own
# Comments cell and its own Invoice Due cell - Current Milestone is
# NOT used to decide color.
# ============================================================

DUE_FILL = "FF0000"          # Red = not completed / completed late
COMPLETED_FILL = "00B050"    # Green = completed on or before expected month
WHITE_FILL = "FFFFFF"        # White = 2025 entries (colorless, overrides green/red)


def apply_task_colors(ws, row, milestone, current_milestone):
    """Color each milestone row from expected month vs actual completion month.

    Business rule:
      - Comments contains the expected milestone month.
      - Invoice Due contains the actual completion month from Jira.
      - If the expected or actual month falls in 2025: WHITE (colorless),
        regardless of on-time/late/blank status.
      - Otherwise, if actual completion month <= expected month: GREEN.
      - Otherwise (actual completion month is after expected month, or
        Invoice Due is missing): RED.

    The current milestone is NOT used to decide the row color. Each UAT/RFS/PAC
    row is evaluated independently.
    """
    expected = parse_month_year(ws.cell(row, C["COMMENTS"]).value)
    actual = parse_month_year(ws.cell(row, C["DUE"]).value)

    is_2025 = (expected is not None and expected[0] == 2025) or \
              (actual is not None and actual[0] == 2025)

    if is_2025:
        fill = PatternFill("solid", fgColor=WHITE_FILL)
    else:
        # Missing actual completion means the work has not been completed.
        if expected is not None and actual is not None:
            actual_key = actual[0] * 12 + actual[1]
            expected_key = expected[0] * 12 + expected[1]
            completed_on_time = actual_key <= expected_key
        else:
            completed_on_time = False

        fill = PatternFill(
            "solid",
            fgColor=COMPLETED_FILL if completed_on_time else DUE_FILL
        )

    # Invoicing Calendar -> Invoice Due (G:J)
    for col in range(C["CALENDAR"], C["DUE"] + 1):
        ws.cell(row, col).fill = copy(fill)


# ============================================================
# BUILD ONE RESOURCE SECTION
# ============================================================

def build_resource_section(
    ws,
    template_ws,
    resource_name,
    tasks,
    resource_info,
    start_row,
    sno_start
):
    employee = resource_info.get("employee", resource_name)
    onboard_date = resource_info.get("date")
    factor = resource_info.get("factor", DEFAULT_FACTOR)
    rate = resource_info.get("rate")
    monthly = resource_info.get("monthly", {})

    # If Resources List does not have the rate, use Jira's rate.
    if rate is None and tasks:
        rate = tasks[0].get("price")

    resource_start = start_row
    current_row = start_row
    sno = sno_start

    task_ranges = []

    # --------------------------------------------------------
    # TASK BLOCKS
    # --------------------------------------------------------

    for task in tasks:
        task_start = current_row
        task_end = task_start + 2

        task_ranges.append(
            (task_start, task_end, task)
        )

        # Copy the exact visual style of the sample's
        # UAT / RFS / PAC rows.
        copy_row_style(template_ws, 2, ws, task_start)
        copy_row_style(template_ws, 3, ws, task_start + 1)
        copy_row_style(template_ws, 4, ws, task_start + 2)

        # Task-level information.
        ws.cell(task_start, C["SNO"]).value = sno
        ws.cell(task_start, C["RESOURCE"]).value = resource_name
        ws.cell(task_start, C["DATE"]).value = onboard_date
        ws.cell(task_start, C["TASK"]).value = task["task_id"]
        ws.cell(task_start, C["POINTS"]).value = task["points"]

        # CRITICAL BUSINESS RULE:
        # Jira PAC -> output PAC.
        ws.cell(task_start, C["MILESTONE"]).value = task["milestone"]

        # Resource information.
        ws.cell(task_start, C["FACTOR"]).value = factor
        ws.cell(task_start, C["EMPLOYEE"]).value = employee
        ws.cell(task_start, C["RATE"]).value = rate

        # Monthly values are kept at the resource level,
        # exactly as in the sample.
        month_order = [
            "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
            "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
        ]

        for i, month in enumerate(month_order):
            value = None

            for key, monthly_value in monthly.items():
                if norm(key).startswith(month):
                    value = monthly_value
                    break

            ws.cell(
                task_start,
                20 + i
            ).value = value

        # ----------------------------------------------------
        # Three milestone rows
        # ----------------------------------------------------

        milestone_rows = [
            (task_start, "UAT", 0.60, task["billed_uat"]),
            (task_start + 1, "RFS", 0.20, task["billed_rfs"]),
            (task_start + 2, "PAC", 0.20, task["billed_pac"]),
        ]

        for row, milestone_name, percentage, billed_month in milestone_rows:
            ws.cell(row, C["CALENDAR"]).value = (
                f"={percentage}*E{task_start}"
            )

            ws.cell(row, C["INVOICE"]).value = (
                f"=(G{row}*$S${task_start})"
            )

            # Comments continue to use the DAC Month rule:
            # UAT = DAC, RFS = DAC + 1 month, PAC = DAC + 2 months.
            invoice_months = task.get(
                "invoice_months",
                {"uat": "", "rfs": "", "pac": ""}
            )
            comment_month = invoice_months.get(milestone_name.lower(), "")
            ws.cell(row, C["COMMENTS"]).value = (
                f"{milestone_name} - {comment_month}"
            )

            # Invoice Due MUST come from Jira Dump -> Billed Month
            # columns N: P, in UAT / RFS / PAC order.
            due_value = billed_month
            ws.cell(row, C["DUE"]).value = month_label(due_value)

            apply_task_colors(
                ws,
                row,
                milestone_name,
                task.get("milestone", "")
            )

        # Merge Task ID / Points / Current Milestone.
        for col in [
            C["TASK"],
            C["POINTS"],
            C["MILESTONE"],
        ]:
            ws.merge_cells(
                start_row=task_start,
                start_column=col,
                end_row=task_end,
                end_column=col
            )

        current_row += 3
        sno += 1

    resource_end = current_row - 1

    # --------------------------------------------------------
    # RESOURCE-LEVEL MERGES
    #
    # A:C and K:P span the complete employee section.
    # --------------------------------------------------------

    for col in [
        C["SNO"],
        C["RESOURCE"],
        C["DATE"],
        C["EARNED"],
        C["BILLED"],
        C["BALANCE"],
        C["UNBILLED_INR"],
        C["UNBILLED_POINTS"],
        C["UNBILLED_AMOUNT"],
    ]:
        ws.merge_cells(
            start_row=resource_start,
            start_column=col,
            end_row=resource_end,
            end_column=col
        )

    # --------------------------------------------------------
    # Resource-level information
    # --------------------------------------------------------

    ws.cell(resource_start, C["SNO"]).value = sno_start
    ws.cell(resource_start, C["RESOURCE"]).value = resource_name
    ws.cell(resource_start, C["DATE"]).value = onboard_date

    ws.cell(resource_start, C["FACTOR"]).value = factor
    ws.cell(resource_start, C["EMPLOYEE"]).value = employee
    ws.cell(resource_start, C["RATE"]).value = rate

    # Yellow resource information exactly like sample.
    for col in [
        C["FACTOR"],
        C["EMPLOYEE"],
        C["RATE"],
    ]:
        ws.cell(resource_start, col).fill = PatternFill(
            "solid",
            fgColor="FFFF00"
        )

    # --------------------------------------------------------
    # Resource-level formulas
    # --------------------------------------------------------

    # Earned SP
    ws.cell(resource_start, C["EARNED"]).value = (
        f"=AF{resource_start}*Q{resource_start}"
    )

    # Billed SP:
    # use the PAC milestone points for each task, following
    # the structure represented in the supplied sample.
    pac_refs = [
        f"G{task_start + 2}"
        for task_start, _, _ in task_ranges
    ]

    ws.cell(resource_start, C["BILLED"]).value = (
        "=SUM(" + ",".join(pac_refs) + ")"
        if pac_refs else "=0"
    )

    # Balance
    ws.cell(resource_start, C["BALANCE"]).value = (
        f"=K{resource_start}-L{resource_start}"
    )

    # Unbilled INR
    ws.cell(resource_start, C["UNBILLED_INR"]).value = (
        f"=M{resource_start}*S{resource_start}"
    )

    # Unbilled points
    ws.cell(resource_start, C["UNBILLED_POINTS"]).value = (
        f"=SUM(G{resource_start + 2}:"
        f"G{resource_end})"
    )

    # Unbilled amount
    ws.cell(resource_start, C["UNBILLED_AMOUNT"]).value = (
        f"=O{resource_start}*S{resource_start}"
    )

    # Total Jan-Dec
    ws.cell(resource_start, C["TOTAL"]).value = (
        f"=SUM(T{resource_start}:AE{resource_start})"
    )

    ws.cell(resource_start, C["TOTAL"]).fill = PatternFill(
        "solid",
        fgColor="A5A5A5"
    )

    # --------------------------------------------------------
    # Resource Total Row
    # --------------------------------------------------------

    total_row = current_row

    ws.cell(
        total_row,
        C["RESOURCE"]
    ).value = f"{resource_name} Total"

    # K:P gray total cells.
    for col in range(C["EARNED"], C["UNBILLED_AMOUNT"] + 1):
        letter = get_column_letter(col)

        ws.cell(total_row, col).value = (
            f"=SUM({letter}{resource_start}:"
            f"{letter}{resource_end})"
        )

        ws.cell(total_row, col).fill = PatternFill(
            "solid",
            fgColor="BFBFBF"
        )

        ws.cell(total_row, col).font = Font(
            bold=True
        )

    # Monthly totals.
    for col in range(C["JAN"], C["DEC"] + 1):
        letter = get_column_letter(col)

        ws.cell(total_row, col).value = (
            f"=SUM({letter}{resource_start}:"
            f"{letter}{resource_end})"
        )

    # AF total.
    ws.cell(total_row, C["TOTAL"]).value = (
        f"=SUM(T{total_row}:AE{total_row})"
    )

    ws.cell(total_row, C["TOTAL"]).fill = PatternFill(
        "solid",
        fgColor="A5A5A5"
    )

    ws.cell(total_row, C["TOTAL"]).font = Font(
        bold=True
    )

    # --------------------------------------------------------
    # Formatting
    # --------------------------------------------------------

    for row in range(resource_start, total_row + 1):
        for col in range(1, 33):
            ws.cell(row, col).alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    return total_row + 2, sno


# ============================================================
# MAIN REPORT CREATION
# ============================================================

def create_report(records, resources):

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input workbook not found:\n{INPUT_FILE}"
        )

    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(
            f"Template workbook not found:\n{TEMPLATE_FILE}\n\n"
            "Put your uploaded Output.xlsx in the same folder "
            "as automation.py."
        )

    # --------------------------------------------------------
    # Load the template.
    #
    # This is the key difference from the previous version:
    # the actual uploaded Output.xlsx supplies the formatting.
    # --------------------------------------------------------

    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[OUTPUT_SHEET]

    # Keep a clean copy of the template sheet for copying
    # its original styles.
    template_wb = load_workbook(TEMPLATE_FILE)
    template_ws = template_wb[OUTPUT_SHEET]

    # --------------------------------------------------------
    # Remove template merges BEFORE deleting rows.
    #
    # openpyxl can raise KeyError if rows containing merged
    # cells are deleted first and the merges are unmerged
    # afterwards.  The template contains many merged cells,
    # so the order here is important.
    # --------------------------------------------------------

    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))

    # --------------------------------------------------------
    # Remove sample data rows, but keep the header.
    # --------------------------------------------------------

    if ws.max_row > 1:
        ws.delete_rows(
            2,
            ws.max_row - 1
        )

    # --------------------------------------------------------
    # Restore header styles.
    # --------------------------------------------------------

    apply_header_colors(ws)

    # --------------------------------------------------------
    # Group tasks.
    # --------------------------------------------------------

    groups = group_tasks(records)

    current_row = 2
    sno = 1

    for _, group in groups.items():
        resource_name = group["resource"]

        resource_info = resources.get(
            norm(resource_name),
            {}
        )

        current_row, sno = build_resource_section(
            ws=ws,
            template_ws=template_ws,
            resource_name=resource_name,
            tasks=group["tasks"],
            resource_info=resource_info,
            start_row=current_row,
            sno_start=sno
        )

    # --------------------------------------------------------
    # Legend at bottom.
    # --------------------------------------------------------

    ws.cell(current_row, 2).value = "invoiced"
    ws.cell(current_row, 2).fill = PatternFill(
        "solid",
        fgColor="00B050"
    )

    ws.cell(current_row + 1, 2).value = "uninvoiced"
    ws.cell(current_row + 1, 2).fill = PatternFill(
        "solid",
        fgColor="FF0000"
    )

    # --------------------------------------------------------
    # Freeze panes.
    # --------------------------------------------------------

    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # Save.
    # --------------------------------------------------------

    wb.save(OUTPUT_FILE)

    return groups


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 60)
    print("MOBILY BILLING EXCEL AUTOMATION")
    print("=" * 60)

    print(f"Input:    {INPUT_FILE}")
    print(f"Template: {TEMPLATE_FILE}")
    print(f"Output:   {OUTPUT_FILE}")

    # Formula workbook.
    wb_formula = load_workbook(
        INPUT_FILE,
        data_only=False
    )

    # Cached-value workbook.
    wb_values = load_workbook(
        INPUT_FILE,
        data_only=True
    )

    print("\nReading Resources List...")
    resources = read_resources(wb_formula)
    print(f"Resources found: {len(resources)}")

    print("\nReading Jira Dump...")
    records = read_jira(
        wb_values,
        wb_formula
    )
    print(f"Unique tasks found: {len(records)}")
    dac_count = sum(1 for record in records.values() if record.get("dac_month") is not None)
    print(f"Tasks with DAC Month: {dac_count}")

    # Milestone summary.
    milestone_counts = {
        "PAC": 0,
        "RFS": 0,
        "RFSIT": 0,
        "": 0,
    }

    for record in records.values():
        milestone = record["milestone"]
        milestone_counts[milestone] = (
            milestone_counts.get(milestone, 0) + 1
        )

    print("\nMilestone distribution:")
    print(f"PAC   : {milestone_counts.get('PAC', 0)}")
    print(f"RFS   : {milestone_counts.get('RFS', 0)}")
    print(f"RFSIT : {milestone_counts.get('RFSIT', 0)}")
    print(f"Blank : {milestone_counts.get('', 0)}")

    print("\nCreating formatted report...")
    groups = create_report(
        records,
        resources
    )

    print("\nResource sections:")
    for group in groups.values():
        print(
            f"  {group['resource']}: "
            f"{len(group['tasks'])} tasks"
        )

    print("\n" + "=" * 60)
    print("DONE")
    print("=" * 60)
    print(f"Created: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()