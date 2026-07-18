"""
Build iskolar-tracker.xlsx — Scholarship Portfolio Tracker
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from scholarship_engine import S, compute_stats, determine_eligibility, Profile

HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=11)
TITLE_FONT    = Font(name="Arial", bold=True, size=14, color="1F4E79")
SUBTITLE_FONT = Font(name="Arial", bold=True, size=12, color="1F4E79")
BODY_FONT     = Font(name="Arial", size=10)
BOLD_FONT     = Font(name="Arial", size=10, bold=True)
LINK_FONT     = Font(name="Arial", size=10, color="0563C1", underline="single")
ELIGIBLE_FILL = PatternFill("solid", fgColor="C6EFCE")
INELIG_FILL   = PatternFill("solid", fgColor="FFC7CE")
COND_FILL     = PatternFill("solid", fgColor="FFEB9C")
ALT_ROW_FILL  = PatternFill("solid", fgColor="F2F7FB")
THIN_BORDER   = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

def style_header(ws, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

def style_body(ws, max_row, max_col, alt=True):
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if alt and r % 2 == 0:
                if not cell.fill or cell.fill.fgColor.rgb == "00000000":
                    cell.fill = ALT_ROW_FILL

def auto_width(ws, min_w=8, max_w=45):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            if cell.value:
                lengths.append(min(len(str(cell.value)), max_w))
        best = max(lengths) + 3 if lengths else min_w
        ws.column_dimensions[col_letter].width = max(min_w, min(best, max_w))

def style_elig_cell(cell, value):
    if value == "Eligible":
        cell.fill = ELIGIBLE_FILL
        cell.font = Font(name="Arial", size=10, bold=True, color="006100")
    elif value == "Ineligible":
        cell.fill = INELIG_FILL
        cell.font = Font(name="Arial", size=10, bold=True, color="9C0006")
    elif value == "Conditional":
        cell.fill = COND_FILL
        cell.font = Font(name="Arial", size=10, bold=True, color="9C6500")

def write_link(cell, url):
    if url and str(url).startswith("http"):
        cell.font = LINK_FONT
        cell.hyperlink = str(url)


def generate_workbook(profile, output_path):
    wb = Workbook()

    results = []
    for s in S:
        elig = determine_eligibility(s, profile)
        results.append({**s, "your_eligible": elig})

    stats = compute_stats(results)

    open_count = stats["open"]
    eligible = stats["eligible"]
    conditional = stats["conditional"]
    ineligible = stats["ineligible"]
    high_match = stats["high_match"]
    total = stats["total"]

    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash.sheet_properties.tabColor = "1F4E79"

    ws_dash.merge_cells("A1:H1")
    ws_dash["A1"] = "ISKOLAR TRACKER \u2014 Scholarship Portfolio"
    ws_dash["A1"].font = Font(name="Arial", bold=True, size=18, color="1F4E79")
    ws_dash["A1"].alignment = Alignment(horizontal="left")

    ws_dash.merge_cells("A2:H2")
    ws_dash["A2"] = f"{profile.course} @ {profile.school} \u00b7 {profile.year} \u00b7 {profile.municipality}"
    ws_dash["A2"].font = Font(name="Arial", size=11, color="555555")
    ws_dash["A2"].alignment = Alignment(horizontal="left")

    ws_dash.merge_cells("A3:H3")
    ws_dash["A3"] = f"Generated: {datetime.date.today().strftime('%B %d, %Y')}  |  Profile: {profile.name}"
    ws_dash["A3"].font = Font(name="Arial", size=10, italic=True, color="999999")

    ws_dash.merge_cells("A5:C5")
    ws_dash["A5"] = "CURRENTLY HELD AWARDS"
    ws_dash["A5"].font = SUBTITLE_FONT

    for i, award in enumerate(profile.held_awards):
        r = 6 + i
        ws_dash[f"A{r}"] = award
        ws_dash[f"A{r}"].font = BOLD_FONT
        ws_dash[f"B{r}"] = "Various"
        ws_dash[f"B{r}"].font = BODY_FONT
        ws_dash[f"C{r}"] = "Awarded"
        ws_dash[f"C{r}"].font = Font(name="Arial", size=10, bold=True, color="006100")
        ws_dash[f"C{r}"].fill = PatternFill("solid", fgColor="C6EFCE")

    stat_start = 6 + len(profile.held_awards) + 1
    ws_dash.merge_cells(f"A{stat_start}:C{stat_start}")
    ws_dash[f"A{stat_start}"] = "PORTFOLIO SUMMARY"
    ws_dash[f"A{stat_start}"].font = SUBTITLE_FONT

    stat_items = [
        ("Total Scholarships Evaluated", total, None),
        ("Currently Open for Application", open_count, "E2EFDA"),
        ("Eligible for You (based on profile)", eligible, "C6EFCE"),
        ("Conditional Eligibility", conditional, "FFEB9C"),
        ("Ineligible", ineligible, "FFC7CE"),
        ("Top Priority Targets (Score \u2265 7 + Eligible)", high_match, None),
        ("Currently Holding (Active)", len(profile.held_awards), None),
    ]
    for i, (label, val, color) in enumerate(stat_items):
        r = stat_start + 1 + i
        ws_dash[f"A{r}"] = label
        ws_dash[f"A{r}"].font = BODY_FONT
        ws_dash[f"B{r}"] = val
        ws_dash[f"B{r}"].font = Font(name="Arial", size=12, bold=True, color="1F4E79")
        ws_dash[f"B{r}"].alignment = Alignment(horizontal="center")
        if color:
            ws_dash[f"A{r}"].fill = PatternFill("solid", fgColor=color)
            ws_dash[f"B{r}"].fill = PatternFill("solid", fgColor=color)

    deadline_start = stat_start + 1 + len(stat_items) + 1
    ws_dash.merge_cells(f"A{deadline_start}:D{deadline_start}")
    ws_dash[f"A{deadline_start}"] = "UPCOMING DEADLINES (Eligible / Conditional)"
    ws_dash[f"A{deadline_start}"].font = SUBTITLE_FONT

    hd = deadline_start + 1
    for c, h in enumerate(["Scholarship", "Deadline", "Status", "Priority"], 1):
        cell = ws_dash.cell(row=hd, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER

    deadlines = []
    for r in results:
        if r["your_eligible"] in ("Eligible", "Conditional") and r["status"] == "Open":
            priority = "HIGH" if r["match_score"] >= 7 else "MEDIUM" if r["match_score"] >= 4 else "LOW"
            deadlines.append((r["name"], r["deadline"], r["status"], priority))
    deadlines.sort(key=lambda x: x[1] if x[1] != "Rolling" else "ZZZZ")

    for i, (name, dl, st, pri) in enumerate(deadlines):
        r = hd + 1 + i
        ws_dash[f"A{r}"] = name
        ws_dash[f"A{r}"].font = BODY_FONT
        ws_dash[f"B{r}"] = dl
        ws_dash[f"B{r}"].font = BODY_FONT
        ws_dash[f"C{r}"] = st
        ws_dash[f"C{r}"].font = BODY_FONT
        ws_dash[f"D{r}"] = pri
        ws_dash[f"D{r}"].font = Font(name="Arial", size=10, bold=True,
                                      color="C00000" if pri == "HIGH" else "BF8F00" if pri == "MEDIUM" else "555555")
        ws_dash[f"D{r}"].alignment = Alignment(horizontal="center")

    leg_r = hd + 1 + len(deadlines) + 1
    ws_dash.merge_cells(f"A{leg_r}:D{leg_r}")
    ws_dash[f"A{leg_r}"] = "ELIGIBILITY LEGEND"
    ws_dash[f"A{leg_r}"].font = SUBTITLE_FONT

    for i, (label, color) in enumerate([
        ("Eligible \u2014 ready to apply", "C6EFCE"),
        ("Conditional \u2014 depends on specific criteria", "FFEB9C"),
        ("Ineligible \u2014 does not match your profile", "FFC7CE"),
    ]):
        r = leg_r + 1 + i
        ws_dash[f"A{r}"] = label
        ws_dash[f"A{r}"].fill = PatternFill("solid", fgColor=color)

    ws_dash.column_dimensions["A"].width = 50
    ws_dash.column_dimensions["B"].width = 35
    ws_dash.column_dimensions["C"].width = 20
    ws_dash.column_dimensions["D"].width = 15
    for c in "EFGH":
        ws_dash.column_dimensions[c].width = 12

    headers_master = [
        "#", "Scholarship Name", "Organization", "Status", "Open Date", "Deadline",
        "Eligible 3rd Yr", "Eligible BSIT", "Region Restriction", "Benefits",
        "Income Requirement", "GPA Requirement", "Exam", "Interview",
        "Competitiveness", "Match Score", "Your Eligibility",
        "ELAP Compatible", "Provincial Compatible", "Application Link", "Notes",
    ]

    ws_master = wb.create_sheet("Master List")
    ws_master.sheet_properties.tabColor = "2E75B6"
    for c, h in enumerate(headers_master, 1):
        ws_master.cell(row=1, column=c, value=h)
    style_header(ws_master, len(headers_master))

    for i, r in enumerate(results):
        row = i + 2
        vals = [
            i + 1, r["name"], r["org"], r["status"], r["open_date"], r["deadline"],
            r["yr3"], r["bsit"], r["region"], r["benefits"],
            r["income"], r["gpa"], r["exam"], r["interview"],
            r["comp"], r["match_score"], r["your_eligible"],
            r["elap_ok"], r["prov_ok"], r["link"], r["notes"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_master.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 20:
                write_link(cell, v)
        style_elig_cell(ws_master.cell(row=row, column=17), r["your_eligible"])

    ws_master.auto_filter.ref = f"A1:{get_column_letter(len(headers_master))}{len(results)+1}"
    ws_master.freeze_panes = "A2"
    auto_width(ws_master, max_w=50)

    eligible_only = [r for r in results if r["your_eligible"] == "Eligible"]
    eligible_only.sort(key=lambda r: -r["match_score"])

    ws_eligible = wb.create_sheet("My Eligible Scholarships")
    ws_eligible.sheet_properties.tabColor = "548235"
    for c, h in enumerate(headers_master, 1):
        ws_eligible.cell(row=1, column=c, value=h)
    style_header(ws_eligible, len(headers_master))

    for i, r in enumerate(eligible_only):
        row = i + 2
        vals = [
            i + 1, r["name"], r["org"], r["status"], r["open_date"], r["deadline"],
            r["yr3"], r["bsit"], r["region"], r["benefits"],
            r["income"], r["gpa"], r["exam"], r["interview"],
            r["comp"], r["match_score"], r["your_eligible"],
            r["elap_ok"], r["prov_ok"], r["link"], r["notes"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_eligible.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 20:
                write_link(cell, v)
        ws_eligible.cell(row=row, column=16).font = Font(name="Arial", size=10, bold=True, color="1F4E79")

    ws_eligible.auto_filter.ref = f"A1:{get_column_letter(len(headers_master))}{len(eligible_only)+1}"
    ws_eligible.freeze_panes = "A2"
    auto_width(ws_eligible, max_w=50)

    ws_tracker = wb.create_sheet("Application Tracker")
    ws_tracker.sheet_properties.tabColor = "ED7D31"

    track_headers = [
        "Scholarship", "Status", "Deadline", "Link",
        "Docs Needed", "Docs Ready?", "Priority", "Notes / Next Step",
    ]
    for c, h in enumerate(track_headers, 1):
        ws_tracker.cell(row=1, column=c, value=h)
    style_header(ws_tracker, len(track_headers))

    track_items = [r for r in results if r["your_eligible"] in ("Eligible", "Conditional")]
    track_items.sort(key=lambda r: (-r["match_score"], r["status"] != "Open"))

    track_data = []
    for a in profile.held_awards:
        track_data.append((a, "Awarded", "N/A", "N/A", "N/A", "Yes", "\u2014", "Currently active. Monitor renewal requirements."))

    for r in track_items:
        if r["status"] == "Open":
            default_status = "Preparing"
        elif r["status"] == "Closed" and r["your_eligible"] == "Conditional":
            default_status = "Monitor"
        else:
            default_status = "Monitor (Next Cycle)"
        priority = "High" if r["match_score"] >= 7 else "Medium" if r["match_score"] >= 4 else "Low"
        note = r["notes"][:120] + ("..." if len(r["notes"]) > 120 else "")
        track_data.append((
            r["name"], default_status, r["deadline"],
            r["link"] if str(r["link"]).startswith("http") else "",
            "", "No", priority, note,
        ))

    status_fills = {
        "Awarded": PatternFill("solid", fgColor="C6EFCE"),
        "Preparing": PatternFill("solid", fgColor="FFEB9C"),
        "Applied": PatternFill("solid", fgColor="BDD7EE"),
        "Interviewing": PatternFill("solid", fgColor="DDEBF7"),
        "Monitor": PatternFill("solid", fgColor="F2F2F2"),
        "Monitor (Next Cycle)": PatternFill("solid", fgColor="F2F2F2"),
    }
    status_fonts = {
        "Awarded": Font(name="Arial", size=10, bold=True, color="006100"),
        "Preparing": Font(name="Arial", size=10, bold=True, color="9C6500"),
        "Monitor": Font(name="Arial", size=10, color="888888"),
        "Monitor (Next Cycle)": Font(name="Arial", size=10, color="888888"),
    }

    for i, row_data in enumerate(track_data):
        r = i + 2
        for c, v in enumerate(row_data, 1):
            cell = ws_tracker.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 4 and v and str(v).startswith("http"):
                cell.font = LINK_FONT
                cell.hyperlink = str(v)
        status_cell = ws_tracker.cell(row=r, column=2)
        st = str(row_data[1])
        if st in status_fills:
            status_cell.fill = status_fills[st]
        if st in status_fonts:
            status_cell.font = status_fonts[st]
        pri_cell = ws_tracker.cell(row=r, column=7)
        p = str(row_data[6])
        if p == "High":
            pri_cell.font = Font(name="Arial", size=10, bold=True, color="C00000")
        elif p == "Medium":
            pri_cell.font = Font(name="Arial", size=10, bold=True, color="BF8F00")

    ws_tracker.auto_filter.ref = f"A1:{get_column_letter(len(track_headers))}{len(track_data)+1}"
    ws_tracker.freeze_panes = "A2"
    auto_width(ws_tracker, max_w=55)
    ws_tracker.column_dimensions["D"].width = 40
    ws_tracker.column_dimensions["E"].width = 45
    ws_tracker.column_dimensions["H"].width = 55

    ws_rank = wb.create_sheet("Ranked Prioritization")
    ws_rank.sheet_properties.tabColor = "C00000"

    rank_pool = [r for r in results if r["your_eligible"] in ("Eligible", "Conditional")]
    rank_pool.sort(key=lambda r: (-r["match_score"], r["status"] != "Open"))

    rank_headers = ["Rank", "Scholarship", "Organization", "Status", "Deadline",
                    "Match Score", "Benefits", "Your Eligibility", "Why"]
    for c, h in enumerate(rank_headers, 1):
        ws_rank.cell(row=1, column=c, value=h)
    style_header(ws_rank, len(rank_headers))

    for i, r in enumerate(rank_pool):
        row = i + 2
        why = r["notes"][:150] + ("..." if len(r["notes"]) > 150 else "")
        vals = [i + 1, r["name"], r["org"], r["status"], r["deadline"],
                r["match_score"], r["benefits"], r["your_eligible"], why]
        for c, v in enumerate(vals, 1):
            cell = ws_rank.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        rank_cell = ws_rank.cell(row=row, column=1)
        rank_cell.font = Font(name="Arial", size=11, bold=True, color="1F4E79")
        rank_cell.alignment = Alignment(horizontal="center")
        ms = ws_rank.cell(row=row, column=6)
        ms.font = Font(name="Arial", size=11, bold=True, color="1F4E79")
        ms.alignment = Alignment(horizontal="center")
        ec = ws_rank.cell(row=row, column=8)
        if r["your_eligible"] == "Eligible":
            ec.fill = ELIGIBLE_FILL
        elif r["your_eligible"] == "Conditional":
            ec.fill = COND_FILL

    ws_rank.freeze_panes = "A2"
    auto_width(ws_rank, max_w=55)

    ws_comp = wb.create_sheet("Comparison Matrix")
    ws_comp.sheet_properties.tabColor = "7030A0"

    comp_headers = ["#", "Scholarship", "Org", "Status", "Deadline", "Match",
                    "Your Eligibility", "Benefits (Short)", "Exam", "Interview"]
    for c, h in enumerate(comp_headers, 1):
        ws_comp.cell(row=1, column=c, value=h)
    style_header(ws_comp, len(comp_headers))

    for i, r in enumerate(results):
        row = i + 2
        short_ben = r["benefits"][:100] + ("..." if len(r["benefits"]) > 100 else "")
        vals = [i + 1, r["name"], r["org"], r["status"], r["deadline"],
                r["match_score"], r["your_eligible"], short_ben, r["exam"], r["interview"]]
        for c, v in enumerate(vals, 1):
            cell = ws_comp.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ec2 = ws_comp.cell(row=row, column=7)
        style_elig_cell(ec2, r["your_eligible"])
        ws_comp.cell(row=row, column=6).alignment = Alignment(horizontal="center")
        ws_comp.cell(row=row, column=6).font = Font(name="Arial", size=10, bold=True, color="1F4E79")

    ws_comp.auto_filter.ref = f"A1:{get_column_letter(len(comp_headers))}{len(results)+1}"
    ws_comp.freeze_panes = "A2"
    auto_width(ws_comp, max_w=50)

    ws_tact = wb.create_sheet("Tactical Manuals")
    ws_tact.sheet_properties.tabColor = "00B050"

    tact_headers = ["Scholarship", "Match Score", "Your Eligibility", "Justification",
                    "Execution Issues", "Documents to Prepare", "Strategic Tips"]
    for c, h in enumerate(tact_headers, 1):
        ws_tact.cell(row=1, column=c, value=h)
    style_header(ws_tact, len(tact_headers))

    tacticals = [
        {"name": "DOST JLSS (RA 7687 & Merit)", "score": "10/10", "elig": "Eligible",
         "just": "Your 1.26\u20131.50 GWA far exceeds the 83% threshold. BSIT is a DOST priority course. BatStateU is a premier SUC.",
         "issues": "Strict 'no double-dipping' rule for national merit grants. May need to forfeit CHED-TDP if awarded. Qualifying exam covers advanced physics/chem/calculus outside BSIT curriculum.",
         "docs": "Certified True Copy of Grades (1st & 2nd yr), Form C (Good Moral), Form D (Program Cert), Parent BIR Form 2316 or Barangay Indigency",
         "tips": "Focus study on logical reasoning and quantitative math for JLSS exam. Monitor DOST-SEI portal for Sept 2026 qualifier announcements."},
        {"name": "GBF STEM College Scholarship", "score": "9/10", "elig": "Eligible",
         "just": "Designed for STEM students with financial need. Direct pipeline to JG Summit tech roles. Aligns with SWE career goal.",
         "issues": "Priority given to those without other active scholarships despite allowing existing grants. Mandatory return service at Gokongwei Group post-graduation.",
         "docs": "Certified TCG (sophomore semesters), 2025-2026 registration form, utility bill, parent Certificate of Employment with compensation",
         "tips": "Build a clean digital portfolio (GitHub repos, web apps, DB scripts). During panel interview, emphasize interest in automation, cloud, or data platforms within Gokongwei Group."},
        {"name": "Presidential Scholars (B BBM)", "score": "9/10", "elig": "Eligible",
         "just": "Province-wide Batangas rollout. 5 scholars per barangay. Santa Teresita qualifies. Fully compatible with ELAP & Provincial.",
         "issues": "Managed through local barangay leagues; timelines vary by barangay. Highly localized selection.",
         "docs": "Barangay Certificate of Residency, Barangay Indigency, Certificate of Registration (COR) from BatStateU confirming BSIT enrollment",
         "tips": "Visit Barangay Hall in Calayaan immediately. Connect with Barangay Captain. Submit certified grades + residency docs on the first day of enrollment period."},
        {"name": "CHED Tulong Dunong (TDP)", "score": "9/10", "elig": "Eligible",
         "just": "Direct financial assistance for BatStateU students. Can be used as stipend. Compatible with existing awards.",
         "issues": "Congressional allocations are highly competitive. Cannot be combined with full TES.",
         "docs": "Wet-signature COR from BatStateU Registrar, Barangay Indigency, clear student ID copy",
         "tips": "Contact your congressional representative's staff (Santa Teresita district). Secure district application code. Submit physical docs during the designated 3-4 day enrollment window."},
        {"name": "CHED Tertiary Education Subsidy (TES)", "score": "9/10", "elig": "Eligible",
         "just": "Up to \u20b140k/yr cash subsidy under Free Higher Education law. For SUC students like BatStateU.",
         "issues": "Non-Listahanan students have lower priority. National budget limits apply.",
         "docs": "Updated COR, Certified True Copy of Grades from BatStateU Registrar",
         "tips": "Coordinate with BatStateU Scholarship Office during enrollment. Ensure registrar encoded correct GWA. Non-Listahanan students should emphasize financial need in application."},
        {"name": "Fuji Haya Electric Scholarship", "score": "8/10", "elig": "Eligible",
         "just": "Direct BatStateU corporate partner. Targets IT/engineering. Includes OJT internships.",
         "issues": "Prohibits holding other private corporate scholarships. ELAP and Provincial LGU are ok.",
         "docs": "BatStateU internal scholarship form, TCG (sophomore yr), Dean's recommendation from College of Informatics",
         "tips": "Visit Scholarship Office at Alangilan Campus. During interview, highlight interest in industrial automation, IoT, or power system software."},
        {"name": "BPI-DOST Innovation Awards", "score": "7/10", "elig": "Eligible",
         "just": "Prestigious competition for junior/senior applied science students. Group project format.",
         "issues": "Requires 3-member team and dean-endorsed lab-validated prototype (TRL 4). Not a traditional stipend.",
         "docs": "Contestant form, certified grades, non-disciplinary cert from Dean, 3-min video pitch",
         "tips": "Team up with classmates. Design a software solution (IoT, mobile health, agri-data platform) that solves a local problem. Make a compelling video pitch."},
        {"name": "DICT STEM/ICT Scholarship", "score": "7/10", "elig": "Eligible",
         "just": "Targets ICT talent. BSIT directly aligns with DICT's strategic goals.",
         "issues": "Timelines and disbursements subject to national coordination. May have delays.",
         "docs": "Certified COR, student ID, freshman & sophomore grades transcript",
         "tips": "Highlight any tech certifications (AWS, Cisco, Google IT). Emphasize goal of becoming a locally-based SWE to align with DICT's digital workforce mandate."},
        {"name": "Accenture Internship & Sponsorship", "score": "7/10", "elig": "Eligible",
         "just": "Direct industry pipeline. ABAP/SAP upskilling for IT students. Aligns with SWE goals.",
         "issues": "Career-development program, not traditional cash stipend. Focus on upskilling rather than monthly allowance.",
         "docs": "Updated technical resume, Letter of Intent, Certified TCG from BatStateU",
         "tips": "Highlight proficiency in Python, Java, or SQL. Show interest in enterprise software (SAP/ABAP, Cloud Computing) during interview."},
        {"name": "DBP INSPIRE Scholarship", "score": "4/10", "elig": "Conditional",
         "just": "Semestral allowance for underprivileged SUC students. Good GWA match.",
         "issues": "Program details and slots managed internally via university. Requires verification with BatStateU Scholarship Office.",
         "docs": "True copy of grades, parent ITR, enrollment verification certificate",
         "tips": "Proactively visit BatStateU Scholarship Office to inquire about active DBP INSPIRE slots for the current cycle."},
        {"name": "Huawei Seeds for the Future", "score": "5/10", "elig": "Eligible",
         "just": "High-quality tech training and certification vouchers. Excellent for upskilling.",
         "issues": "Short-term training program, not ongoing cash stipend. Focus on competitions and global networking.",
         "docs": "Resume, personal statement, digital coding portfolio",
         "tips": "Showcase strong interest in emerging tech (Cloud, AI). High GPA helps. This complements other cash stipend scholarships."},
        {"name": "CHED CoScho", "score": "3/10", "elig": "Conditional",
         "just": "Up to \u20b140k/sem if parent is a registered coconut farmer.",
         "issues": "Requires NCFRS registration. Must confirm if parent is listed.",
         "docs": "Parent NCFRS certificate, birth certificate, latest college grades",
         "tips": "Check with parents if they are registered in the National Coconut Farmers Registry. If yes, apply through CHED regional portal."},
        {"name": "CSC Gawad PASUC", "score": "5/10", "elig": "Conditional",
         "just": "Educational subsidies at any PASUC SUC. BatStateU qualifies.",
         "issues": "Restricted to Gawad Lingkod Bayani awardees and their immediate kin.",
         "docs": "CSC nomination/award certificate, PDS form, Good Moral, enrollment verification",
         "tips": "Check if any parent/guardian is a Gawad awardee. If eligible, coordinate with CSC Central Office for endorsement."},
    ]

    for i, t in enumerate(tacticals):
        r = i + 2
        vals = [t["name"], t["score"], t["elig"], t["just"], t["issues"], t["docs"], t["tips"]]
        for c, v in enumerate(vals, 1):
            cell = ws_tact.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ec3 = ws_tact.cell(row=r, column=3)
        if t["elig"] == "Eligible":
            ec3.fill = ELIGIBLE_FILL
            ec3.font = Font(name="Arial", size=10, bold=True, color="006100")
        elif t["elig"] == "Conditional":
            ec3.fill = COND_FILL
            ec3.font = Font(name="Arial", size=10, bold=True, color="9C6500")
        ws_tact.cell(row=r, column=2).font = Font(name="Arial", size=11, bold=True, color="1F4E79")
        ws_tact.cell(row=r, column=2).alignment = Alignment(horizontal="center")

    ws_tact.freeze_panes = "A2"
    ws_tact.column_dimensions["A"].width = 38
    ws_tact.column_dimensions["B"].width = 10
    ws_tact.column_dimensions["C"].width = 14
    ws_tact.column_dimensions["D"].width = 55
    ws_tact.column_dimensions["E"].width = 50
    ws_tact.column_dimensions["F"].width = 50
    ws_tact.column_dimensions["G"].width = 55

    wb.save(output_path)
    return stats


if __name__ == "__main__":
    profile = Profile(
        name="Demo Scholar",
        course="BSIT",
        year="Incoming 3rd Year",
        school="Batangas State University",
        municipality="Batangas Province",
        is_female=False,
        held_awards=["LGU Educational Grant", "Provincial Merit Scholarship"],
    )
    output_path = r"D:\ISKORE\iskolar-tracker.xlsx"
    stats = generate_workbook(profile, output_path)
    print(f"Workbook saved to {output_path}")
    print(f"Total scholarships: {stats['total']}")
    print(f"  Eligible: {stats['eligible']}")
    print(f"  Conditional: {stats['conditional']}")
    print(f"  Ineligible: {stats['ineligible']}")
    print(f"  Open: {stats['open']}")
