from pydantic import BaseModel, Field
from typing import Optional

class Profile(BaseModel):
    name: str = "Demo Scholar"
    course: str = "BSIT"
    year: str = "Incoming 3rd Year"
    school: str = "Batangas State University"
    municipality: str = "Batangas Province"
    is_female: bool = False
    held_awards: list[str] = []

S = []

def add(name, org, status, open_date, deadline, yr3, bsit, region,
        benefits, income, gpa, exam, interview, comp, match_score,
        elap_ok, prov_ok, link, notes, your_eligible):
    S.append(dict(
        name=name, org=org, status=status, open_date=open_date,
        deadline=deadline, yr3=yr3, bsit=bsit, region=region,
        benefits=benefits, income=income, gpa=gpa, exam=exam,
        interview=interview, comp=comp, match_score=match_score,
        elap_ok=elap_ok, prov_ok=prov_ok, link=link, notes=notes,
        your_eligible=your_eligible,
    ))

add("DOST JLSS \u2014 RA 7687 Track", "DOST-SEI", "Closed", "Apr 13, 2026", "May 15, 2026",
    "Yes", "Yes", "None (National)", "\u20b18,000/mo living allowance, \u20b110k/yr books, \u20b110k thesis, health insurance, \u20b11k graduation, travel allowance",
    "Disadvantaged families", "GWA \u2265 83%", "Yes", "No", "Very High", 10,
    "Not stated", "Not stated",
    "https://jlss.science-scholarships.ph",
    "Top priority. Requires relinquishing other nat'l merit grants if awarded. Exam covers advanced physics/chem/calculus.",
    "Eligible")

add("DOST JLSS \u2014 Merit Track", "DOST-SEI", "Closed", "Apr 13, 2026", "May 15, 2026",
    "Yes", "Yes", "None (National)", "\u20b18,000/mo living allowance, \u20b110k/yr books, \u20b110k thesis, health insurance",
    "None (merit-based)", "GWA \u2265 83%", "Yes", "No", "Very High", 10,
    "Not stated", "Not stated",
    "https://jlss.science-scholarships.ph",
    "Perfect for strong GWA. No income barrier. Same exam as RA 7687.",
    "Eligible")

add("DOST JLSS \u2014 RA 10612 Track", "DOST-SEI", "Closed", "Apr 13, 2026", "May 15, 2026",
    "Yes", "Yes", "None (National)", "\u20b18,000/mo living allowance, \u20b110k/yr books, guaranteed teaching job placement",
    "None", "GWA \u2265 83%", "Yes", "No", "High", 9,
    "Not stated", "Not stated",
    "https://jlss.science-scholarships.ph",
    "Mandatory teaching service may divert from SWE goals. Slightly lower match.",
    "Eligible")

add("GBF STEM-College Scholarship", "Gokongwei Brothers Foundation", "Closed", "Mar/Apr 2026", "May 31, 2026",
    "Yes", "Yes", "None (National)", "Up to \u20b1120k/yr, leadership training, internship/employment at Gokongwei Group",
    "Financial need", "GWA \u2265 85% / 2.0", "Yes", "Yes", "High", 9,
    "Yes (case-by-case)", "Yes",
    "https://bit.ly/GBFSTEMCollege",
    "Strong fit for SWE goals. Return service at Gokongwei Group. Portfolio matters.",
    "Eligible")

add("Presidential Scholars Program (B BBM)", "Office of the President / Batangas LGU", "Open", "Apr 24, 2026", "Rolling (barangay allocation)",
    "Yes", "Yes", "Batangas Province (5 per barangay)", "\u20b120,000 educational assistance",
    "Low-income priority", "Passing standing (not explicitly stated)", "No", "No", "Medium", 9,
    "Yes", "Yes",
    "Coordinated at Barangay Hall",
    "Active rolling. Visit your Barangay Captain. Submit early.",
    "Eligible")

add("Batangas City EBD Scholarship", "Batangas City Government", "Open", "Jul 2026", "Aug 2026",
    "Yes", "Yes", "Batangas City resident \u22655 yrs", "\u20b13,000/sem allowance (SUC scholars)",
    "Marginally poor", "GWA \u2265 2.5 (80%)", "No", "Yes (phone)", "Medium-High", 9,
    "Not stated (generally allowed)", "Not stated",
    "ebd_scholarship@batangascity.gov.ph",
    "\u26a0 INELIGIBLE \u2014 resident must be Batangas City.",
    "Ineligible")

add("CHED Tulong Dunong (TDP)", "CHED", "Open", "Jul 13, 2026", "Mid-late Jul 2026 (3-4 day window)",
    "Yes", "Yes", "Legislative district of sponsoring rep", "\u20b115,000/yr",
    "Marginally disadvantaged", "Passing average", "No", "No", "High", 9,
    "Not stated", "Not stated",
    "https://ched.gov.ph",
    "Facilitated through congressional district office. Very short submission window.",
    "Eligible")

add("CHED Tertiary Education Subsidy (TES)", "CHED / UniFAST", "Open", "Jun/Jul 2026", "Determined by SUC enrollment",
    "Yes", "Yes", "None (enrolled in SUC)", "\u20b120,000\u2013\u20b140,000/yr",
    "DSWD Listahanan / 4Ps priority", "Satisfactory standing", "No", "No", "Extremely High", 9,
    "Not stated", "Not stated",
    "https://unifast.gov.ph",
    "Internal via university scholarship office. Non-Listahanan = lower priority.",
    "Eligible")

add("Fuji Haya Electric Scholarship", "Fuji Haya Electric / BatStateU", "Open", "Enrollment cycle", "Semestral deadlines",
    "Yes", "Yes", "BatStateU student", "Semestral grant, seminars, OJT, potential employment",
    "Financial need", "High standing", "No", "Yes", "Medium-High", 8,
    "Not stated", "Not stated",
    "https://global.batstateu.edu.ph",
    "Directly at campus. Prohibits other private corporate awards but LGU ok.",
    "Eligible")

add("Generation Google (APAC) \u2014 Women in CS", "Google Inc.", "Open", "Jul 1, 2026", "Aug 14, 2026",
    "Yes", "Yes", "None (APAC region)", "$2,500 USD (~\u20b1140,000) pure stipend",
    "Financial need considered", "GWA \u2265 85%", "Yes (Online Challenge)", "Yes", "Extremely High", 8,
    "Yes", "Yes",
    "https://buildyourfuture.withgoogle.com",
    "\u26a0 INELIGIBLE \u2014 female-identifying applicants only.",
    "Ineligible")

add("Megaworld Foundation Scholarship", "Megaworld Foundation", "Closed", "Feb 11, 2026", "Jul 31, 2026",
    "Yes", "Yes", "Partner schools (BatStateU is partner)", "Full tuition waiver, monthly living allowance, seminars, career placement",
    "Income \u2264 \u20b1400k/yr", "GWA \u2265 85%, no subject <80%", "Yes", "Yes", "High", 8,
    "No", "No",
    "https://www.megaworldfoundation.com/scholarship/apply/form",
    "Strict 'no other scholarship' clause conflicts with held awards.",
    "Ineligible")

add("DBP INSPIRE Scholarship", "Development Bank of the Philippines", "Open", "Mid-year 2026", "Not explicitly stated",
    "Yes", "Yes", "None (PH-wide SUC network)", "Semestral allowance + learning materials subsidy",
    "Marginally disadvantaged", "Passing standing", "No", "No", "High", 4,
    "Not stated", "Not stated",
    "https://www.dbp.ph",
    "Limited info on SUC quotas. Requires verification with university scholarship office.",
    "Conditional")

add("DICT STEM/ICT Scholarship", "DICT", "Open", "Rolling", "Aligns with enrollment",
    "Yes", "Yes", "None", "Specialized training, global certifications, upskilling",
    "Disadvantaged families", "Passing marks", "No", "No", "Medium-High", 7,
    "Not stated", "Not stated",
    "https://dict.gov.ph",
    "High alignment with SWE goals. Certification-focused rather than cash stipend.",
    "Eligible")

add("Accenture Internship & Sponsorship", "Accenture Philippines", "Open", "Rolling", "Rolling",
    "Yes", "Yes", "None", "Upskilling (ABAP, SAP, Salesforce), internship stipend, priority hiring",
    "None", "Satisfactory progress", "Yes (Cognitive)", "Yes", "Medium", 7,
    "Yes", "Yes",
    "https://www.accenture.com/ph-en/careers/form/phstudents",
    "Career-development program, not traditional cash stipend. Excellent for SWE pipeline.",
    "Eligible")

add("San Juan Iskolar ng Bayan", "Municipal Gov of San Juan, Batangas", "Open", "Jul/Aug 2026", "Municipal registration deadline",
    "Yes", "Yes", "San Juan, Batangas resident", "Load assistance + semestral financial support",
    "Financially needy", "Passing grades", "No", "No", "Low-Medium", 6,
    "Yes", "Yes",
    "San Juan LGU Mayor's Office",
    "\u26a0 INELIGIBLE \u2014 resident must be San Juan.",
    "Ineligible")

add("CHED Merit Scholarship (CMSP)", "CHED", "Closed", "Jun 22, 2026", "Jul 31, 2026",
    "No (freshman only)", "Yes", "None", "\u20b140,000\u2013\u20b180,000/yr",
    "Income \u2264 \u20b1500k", "Grade 12 GWA \u2265 93%", "No", "No", "Extremely High", 1,
    "No", "No",
    "https://ched.gov.ph",
    "INELIGIBLE \u2014 restricted to incoming freshmen.",
    "Ineligible")

add("CHED Bagong Pilipinas Merit (BPMSP)", "CHED", "Closed", "Jun 2026", "Jul 2026",
    "No (freshman only)", "Yes", "None", "Full tuition support (private HEIs), stipend, living allowance",
    "Income \u2264 \u20b12M", "Grade 12 GWA \u2265 95%", "No", "No", "Extremely High", 1,
    "No", "No",
    "https://bpms.ched.gov.ph",
    "INELIGIBLE \u2014 restricted to first-time college entrants.",
    "Ineligible")

add("CHED Estatistikolar", "CHED", "Closed", "Jun 22, 2026", "Aug 15, 2026",
    "Yes", "No (Statistics only)", "None", "\u20b135,000/sem stipend + \u20b15k books",
    "Income \u2264 \u20b1500k", "Passing", "No", "No", "High", 4,
    "Not stated", "Not stated",
    "https://ched.gov.ph",
    "INELIGIBLE \u2014 restricted to BS Statistics / Applied Statistics.",
    "Ineligible")

add("CHED Coconut Farmers (CoScho)", "CHED", "Open", "Jun 22, 2026", "Jul 31, 2026",
    "Yes", "Yes", "None", "Up to \u20b140,000/sem allowance",
    "Registered coconut farmer/dependent", "Satisfactory standing", "No", "No", "High", 3,
    "Not stated", "Not stated",
    "https://ched.gov.ph",
    "Requires NCFRS registration. Conditional on parent being a registered coconut farmer.",
    "Conditional")

add("CHED Medical Scholarship (MSRS)", "CHED", "Open", "Jun/Jul 2026", "Not stated",
    "No (MD students)", "No (Medicine only)", "Partner HEIs", "Full tuition, living/book/transport allowance",
    "Indigent", "NMAT + med school standards", "Yes (NMAT)", "Yes", "Extremely High", 1,
    "No", "No",
    "https://ched.gov.ph",
    "INELIGIBLE \u2014 Doctor of Medicine program only.",
    "Ineligible")

add("CHED AHEAD Grant", "CHED", "Open", "Jul 2026", "School-specific",
    "Yes", "No (Nursing/Allied Health only)", "None", "One-time \u20b125,000 clinical training assistance",
    "Financial need", "Satisfactory standing", "No", "No", "Medium", 1,
    "No", "No",
    "https://ched.gov.ph",
    "INELIGIBLE \u2014 restricted to Nursing and Allied Health students.",
    "Ineligible")

add("CHED SIDA-SGP", "CHED", "Open", "Jun 2026", "Regional schedules",
    "No (freshman entry)", "Yes", "Special development areas", "Semestral cash subsidy, tuition, book support",
    "Low-income", "Passing marks", "No", "No", "High", 1,
    "No", "No",
    "https://ched.gov.ph",
    "INELIGIBLE \u2014 restricted to specific regions and incoming freshmen.",
    "Ineligible")

add("Aboitiz Future Leaders Scholarship", "Aboitiz Foundation", "Closed", "Aug 1, 2025/26", "Sep 1, 2025/26",
    "No (incoming sophomores)", "Yes", "Partner universities", "Full tuition, monthly stipend, 400-hr internship",
    "Financial need", "GWA \u2265 88%", "Yes", "Yes", "High", 1,
    "No", "No",
    "https://aboitiz.com",
    "INELIGIBLE \u2014 open only to incoming 2nd year students.",
    "Ineligible")

add("Aboitiz Dualtech Scholarship", "Aboitiz / Dualtech", "Open", "Rolling", "Rolling",
    "No (2-yr tech-voc)", "No (Electromechanics)", "None", "Subsidized training, OJT, tools",
    "Indigent young men", "Basic passing", "Yes", "Yes", "Medium", 1,
    "No", "No",
    "https://www.dualtech.org.ph",
    "INELIGIBLE \u2014 2-year tech-voc program, not BSIT.",
    "Ineligible")

add("Lao Foundation Dualtech Scholarship", "Lao Foundation / Dualtech", "Open", "Rolling", "Rolling",
    "No (2-yr tech-voc)", "No (Electromechanics)", "None", "Subsidized training fees, industrial placement",
    "Indigent youth", "Basic passing", "Yes", "Yes", "Medium", 1,
    "No", "No",
    "https://www.dualtech.org.ph",
    "INELIGIBLE \u2014 2-year tech-voc program, not BSIT.",
    "Ineligible")

add("SM Foundation College Scholarship", "SM Foundation", "Closed", "Dec-Feb", "Not stated",
    "No (Grade 12 only)", "Yes", "Partner schools", "Full tuition, monthly allowance, part-time jobs",
    "Income \u2264 \u20b1250k", "Grade 12 GWA \u2265 92%", "Yes", "Yes", "Extremely High", 1,
    "No", "No",
    "https://www.sm-foundation.org",
    "INELIGIBLE \u2014 restricted to graduating Grade 12 students.",
    "Ineligible")

add("Security Bank External College Scholarship", "Security Bank Foundation", "Closed", "May 28, 2025/26", "Jun 30, 2025/26",
    "Yes (limited slots)", "Yes", "8 partner schools only (ADMU, DLSU, FEU, PLM, PUP, UPD, UST, CKSC)", "\u20b1100,000/yr stipend (SUC scholars)",
    "Underprivileged", "GWA \u2265 93%, no subject <86%", "No", "No (validation call)", "Very High", 1,
    "No", "No",
    "https://www.securitybank.com/foundation",
    "INELIGIBLE \u2014 BatStateU is not a partner school.",
    "Ineligible")

add("Security Bank RMKK Agency Personnel Scholarship", "Security Bank Foundation", "Open", "Not stated", "Not stated",
    "Yes", "Yes", "None", "Up to \u20b130,000/yr",
    "SB agency personnel only", "GWA \u2265 83%", "No", "No", "Medium", 1,
    "No", "No",
    "https://www.securitybank.com",
    "INELIGIBLE \u2014 restricted to Security Bank agency employees.",
    "Ineligible")

add("Security Bank Employee Dependent Internal Scholarship", "Security Bank Foundation", "Open", "Jun 2026", "Jun 30, 2026",
    "Yes", "Yes", "None", "Direct tuition and stipend assistance",
    "SB permanent employees \u22652 yrs", "GWA \u2265 83%, min 83% per subject", "No", "No", "Medium", 1,
    "No", "No",
    "https://www.securitybank.com",
    "INELIGIBLE \u2014 restricted to children of SB employees.",
    "Ineligible")

add("BPI Foundation Pagpupugay Scholarship", "BPI Foundation", "Closed", "Jun 2026", "Closed",
    "Yes", "Yes", "None", "Up to \u20b1100,000/yr",
    "Next of kin of COVID frontliner", "GWA \u2265 85%", "No", "No", "High", 2,
    "Not stated", "Not stated",
    "https://www.bpifoundation.org",
    "INELIGIBLE \u2014 requires being next of kin of a medical frontliner.",
    "Ineligible")

add("BPI-DOST Innovation Awards", "BPI Foundation & DOST-STII", "Closed", "Jan 2026", "Jun 30, 2026",
    "Yes", "Yes", "None", "\u20b160k\u2013\u20b1150k cash prizes (team competition)",
    "None", "Regular enrollment, no failing grades", "No", "Yes (presentation)", "Extremely High", 7,
    "Yes", "Yes",
    "https://www.bpifoundation.org",
    "Group project competition. Requires dean endorsement and 3-member team.",
    "Eligible")

add("InLife Gold Eagle College Scholarship", "Insular Life Foundation", "Closed", "Feb 6, 2026", "Jun 30, 2026",
    "No (freshman only)", "Yes", "Partner SUCs (UPD, ASCOT, BSU, BISU, CMU, DORSU)", "Semestral stipends, book/transport allowances",
    "Income \u2264 \u20b1250k (provincial SUCs)", "Grade 11-12 GWA \u2265 85%", "No", "No", "High", 1,
    "No", "No",
    "https://www.insularfoundation.com.ph",
    "INELIGIBLE \u2014 restricted to incoming 1st year college students.",
    "Ineligible")

add("PHINMA National Scholarship (PNS)", "PHINMA Foundation", "Closed", "Jun 12, 2026", "Jul 15, 2026",
    "No (incoming 2nd yr)", "No (specific Eng'g/Accountancy)", "5 partner schools (UPD, PUP, PNU, TUP, UPang)", "Monthly stipend, leadership training, mentorship",
    "Income \u2264 \u20b1300k", "GWA \u2265 80%", "No", "Yes", "Very High", 1,
    "No", "No",
    "https://www.phinmafoundation.org",
    "INELIGIBLE \u2014 wrong year level, wrong course, wrong SUC.",
    "Ineligible")

add("Huawei Seeds for the Future", "Huawei Technologies", "Open", "Not stated", "Program-specific",
    "Yes", "Yes", "None", "All-expenses-paid tech training, global competition, cert vouchers",
    "None (merit-based)", "Strong academic records", "No", "Yes", "High", 5,
    "Yes", "Yes",
    "https://www.huawei.com/minisite/seeds-for-the-future",
    "Short-term training program, not ongoing cash stipend. Great for upskilling.",
    "Eligible")

add("Huawei-DLSU Engineering & Tech Scholarship", "Huawei / DLSU", "Closed", "Academic year start", "Semestral cycle",
    "Yes", "Yes", "DLSU Manila student", "Semestral tuition + cash allowance",
    "Financial need + merit", "DLSU standards", "No", "Yes", "High", 1,
    "No", "No",
    "https://www.dlsu.edu.ph",
    "INELIGIBLE \u2014 restricted to DLSU students.",
    "Ineligible")

add("Huawei-Ateneo Science & Engineering Scholarship", "Huawei / Ateneo", "Open", "Enrollment aligned", "Ateneo schedule",
    "Yes", "No (CS, ECE, CpE at Ateneo)", "Ateneo de Manila student", "Tuition waiver + monthly living support",
    "Merit + financial need", "High average", "No", "Yes", "High", 1,
    "No", "No",
    "https://www.ateneo.edu",
    "INELIGIBLE \u2014 restricted to Ateneo students in specific programs.",
    "Ineligible")

add("Accenture Analytics Career Path Scholarship", "Accenture / UPC", "Closed", "Not stated", "Closed",
    "Yes", "No (Statistics/OR)", "UPC Barcelona, Spain", "Enrollment credits + admin fees",
    "None", "High standing", "No", "No", "Very High", 1,
    "No", "No",
    "https://www.wemakescholars.com",
    "INELIGIBLE \u2014 restricted to UPC Barcelona and Statistics/OR majors.",
    "Ineligible")

add("Study in Saudi Arabia Program", "Saudi Ministry of Education", "Open", "International schedule", "University-dependent",
    "No (freshmen abroad)", "Yes", "Must move to Saudi Arabia", "Full tuition waiver, living allowance, housing, medical",
    "None", "Excellent standing", "No", "No", "High", 1,
    "No", "No",
    "https://studyinsaudi.moe.gov.sa",
    "INELIGIBLE \u2014 requires relocating abroad as a freshman.",
    "Ineligible")

add("AMEXCID Merit Scholarship", "Mexican Agency for International Cooperation", "Open", "Academic cycle", "Varies",
    "No (mobility program)", "Yes", "Must move to Mexico", "Living allowance, tuition, health insurance, travel",
    "None", "GWA \u2265 85%", "No", "No", "Very High", 1,
    "No", "No",
    "https://www.gob.mx/amexcid",
    "INELIGIBLE \u2014 requires relocating to Mexico.",
    "Ineligible")

add("Gawad Lingkod Bayani PASUC", "CSC & PASUC", "Open", "May 25, 2026", "Not stated",
    "Yes", "Yes", "PASUC-member SUCs (incl. BatStateU)", "Educational subsidies at any PASUC SUC",
    "Not stated", "Not stated", "No", "No", "Low-Medium", 5,
    "Yes", "Yes",
    "https://www.csc.gov.ph",
    "Restricted to Gawad awardees and their immediate kin. Conditional on family status.",
    "Conditional")

add("Metrobank ACCESS Scholarship", "Metrobank Foundation", "Closed", "Enrollment cycle", "Partner timeline",
    "No (freshman only)", "Yes", "Partner PAASCU Level II/III schools", "Tuition + fixed semestral allowances + workshops",
    "Income \u2264 \u20b1500k", "GWA \u2265 85%", "Yes", "Yes", "Extremely High", 1,
    "No", "No",
    "https://www.mbfoundation.org.ph",
    "INELIGIBLE \u2014 restricted to incoming freshmen.",
    "Ineligible")

add("Metrobank Excellence in Teaching (ETP)", "Metrobank Foundation", "Closed", "Academic year start", "Enrollment season",
    "Yes", "No (Education majors only)", "COE SUCs", "Stipends, tuition support, development programs",
    "Poorest regions", "GWA \u2265 85%", "No", "Yes", "High", 1,
    "No", "No",
    "https://www.mbfoundation.org.ph",
    "INELIGIBLE \u2014 restricted to Education majors.",
    "Ineligible")

add("Metrobank GRACE Grant", "Metrobank Foundation", "Closed", "Final year start", "School-specific",
    "No (graduating seniors)", "Yes", "Partner Level II/III schools", "Semestral cash subsidy",
    "High drop-out risk", "GWA \u2265 85%", "No", "Yes", "High", 1,
    "No", "No",
    "https://www.mbfoundation.org.ph",
    "INELIGIBLE \u2014 restricted to graduating seniors.",
    "Ineligible")

add("Metrobank Scholarship Program (MSP)", "Metrobank Foundation", "Closed", "Academic year start", "Enrollment season",
    "No (freshman only)", "Yes", "Partner schools", "Semestral cash stipends + tuition waivers",
    "Marginalized households", "Satisfactory progress", "Yes", "Yes", "Very High", 1,
    "No", "No",
    "https://www.mbfoundation.org.ph",
    "INELIGIBLE \u2014 restricted to incoming freshmen.",
    "Ineligible")

add("Metrobank-Boysen Paint Scholarship", "Metrobank Foundation & Boysen", "Open", "Enrollment aligned", "Not stated",
    "Yes", "No (Architecture only)", "Architecture partner schools", "Semestral cash stipends, design project support, books",
    "Underprivileged", "GWA \u2265 85%", "No", "Yes", "Very High", 1,
    "No", "No",
    "https://www.mbfoundation.org.ph",
    "INELIGIBLE \u2014 restricted to Architecture majors.",
    "Ineligible")

add("Ayala Foundation U-Go Scholar Grant", "Ayala Foundation / U-Go Global", "Closed", "May 5, 2026", "Jun 6, 2026",
    "Yes", "Yes", "Public/state universities", "\u20b140,000/yr + leadership training + professional development",
    "Low-income", "GWA \u2265 85%", "No", "Yes", "Very High", 1,
    "No", "No",
    "https://ayalafoundation.org/programs/scholarships",
    "INELIGIBLE \u2014 female-prioritized AND prohibits holding other active grants.",
    "Ineligible")

add("GSIS Subsidy for STEM (GSSP)", "GSIS", "Closed", "Jun 30, 2026", "Aug 31, 2026",
    "No (freshman only)", "Yes", "None", "\u20b115,000/yr + Latin honors incentives up to \u20b150k",
    "Parent GSIS member SG \u226415", "Grade 12 GWA \u2265 90%", "No", "No", "High", 1,
    "No", "No",
    "https://www.gsis.gov.ph",
    "INELIGIBLE \u2014 restricted to incoming first-time college freshmen.",
    "Ineligible")


def determine_eligibility(s, profile):
    course_ok = s["bsit"] == "Yes" and profile.course == "BSIT"
    yr_ok = s["yr3"] == "Yes" and "3rd" in profile.year
    region = s["region"]
    mun = profile.municipality
    is_female = profile.is_female
    held = profile.held_awards

    if s["bsit"] == "No (Statistics only)":
        return "Ineligible"
    if s["bsit"] == "No (Medicine only)":
        return "Ineligible"
    if s["bsit"] == "No (Nursing/Allied Health only)":
        return "Ineligible"
    if s["bsit"] == "No (Electromechanics)":
        return "Ineligible"
    if s["bsit"] == "No (specific Eng'g/Accountancy)":
        return "Ineligible"
    if s["bsit"] == "No (CS, ECE, CpE at Ateneo)":
        return "Ineligible"
    if s["bsit"] == "No (Statistics/OR)":
        return "Ineligible"
    if s["bsit"] == "No (Education majors only)":
        return "Ineligible"
    if s["bsit"] == "No (Architecture only)":
        return "Ineligible"

    if not course_ok:
        return "Ineligible"

    if s["yr3"] in ("No (freshman only)", "No (incoming sophomores)",
                     "No (2-yr tech-voc)", "No (freshmen abroad)",
                     "No (mobility program)", "No (incoming 2nd yr)",
                     "No (Grade 12 only)", "No (graduating seniors)",
                     "No (MD students)", "No (freshman entry)"):
        return "Ineligible"
    if not yr_ok and s["yr3"] == "No (freshman only)":
        return "Ineligible"

    if is_female and "female" in s["notes"].lower() and "ineligible" in s["notes"].lower():
        return "Ineligible"

    region_lower = region.lower()
    mun_lower = mun.lower()
    if "batangas city" in region_lower and "batangas city" not in mun_lower:
        return "Ineligible"
    if "san juan" in region_lower and "san juan" not in mun_lower:
        return "Ineligible"
    if "must move to" in region_lower:
        return "Ineligible"
    if "partner schools" in region_lower or "partner" in region_lower:
        schools = region_lower.replace("partner schools only (", "").replace(")", "").replace("partner schools (", "").replace(")", "")
        if profile.school.lower() not in schools:
            return "Ineligible"

    if "no double-dipping" in s["notes"].lower() or "no other scholarship" in s["notes"].lower() or "prohibits holding" in s["notes"].lower():
        if held:
            return "Ineligible"

    if s["name"] == "DBP INSPIRE Scholarship":
        return "Conditional"
    if "coconut farmer" in s["notes"].lower():
        return "Conditional"
    if "gawad" in s["notes"].lower():
        return "Conditional"

    return "Eligible"


def compute_all(profile):
    results = []
    for s in S:
        elig = determine_eligibility(s, profile)
        results.append({**s, "your_eligible": elig})
    return results

def compute_stats(results):
    total = len(results)
    open_count = sum(1 for r in results if r["status"] == "Open")
    eligible = sum(1 for r in results if r["your_eligible"] == "Eligible")
    conditional = sum(1 for r in results if r["your_eligible"] == "Conditional")
    ineligible = sum(1 for r in results if r["your_eligible"] == "Ineligible")
    high_match = sum(1 for r in results if r["match_score"] >= 7 and r["your_eligible"] == "Eligible")
    return {
        "total": total,
        "open": open_count,
        "eligible": eligible,
        "conditional": conditional,
        "ineligible": ineligible,
        "high_match": high_match,
    }

KNOWN_AWARDS = [
    "LGU Educational Grant",
    "Provincial Merit Scholarship",
    "OWWA ELAP",
    "CHED Tulong Dunong (TDP)",
    "CHED TES",
    "DOST JLSS (any track)",
    "GBF STEM-College",
    "Fuji Haya Electric",
    "BPI-DOST Innovation",
    "DICT STEM/ICT",
    "Huawei Seeds",
    "Accenture Internship",
    "Presidential BBM",
    "Megaworld Foundation",
    "SM Foundation",
    "Security Bank",
    "Ayala U-Go",
]
