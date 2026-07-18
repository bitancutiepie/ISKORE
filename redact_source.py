"""Redact personal info from source data for public portfolio release."""
import re

src = r"D:\ISKORE\Comprehensive Strategic Analysis an.txt"
dst = r"D:\ISKORE\data\scholarship-analysis-source.txt"

with open(src, "r", encoding="utf-8") as f:
    text = f.read()

# ── Redaction rules ──────────────────────────────────────────────────

replacements = [
    # Personal identifiers
    ("the candidate currently holds", "a student might hold"),
    ("the candidate currently holds two awards", "an example student profile holds two awards"),
    ("The candidate currently holds", "A sample student profile holds"),
    ("the candidate's academic standing", "a student's academic standing"),
    ("the candidate's GWA", "the student's GWA"),
    ("the candidate's career objective", "the student's career objective"),
    ("the candidate's career goals", "the student's career goals"),
    ("the candidate must", "the student should"),
    ("the candidate should", "the student should"),
    ("the candidate", "the student"),

    # Specific GWA ranges
    ("GWA of approximately 1.26 to 1.50 (equivalent to a 93% to 97% average)", "GWA well above the minimum threshold"),
    ("GWA (1.26 - 1.50)", "GWA (well above minimum threshold)"),
    ("GWA (1.26–1.50)", "GWA (well above minimum threshold)"),
    ("academic average (1.26 - 1.50)", "academic average (well above minimum threshold)"),
    ("1.26 - 1.50", "well above minimum threshold"),
    ("1.26–1.50", "well above minimum threshold"),

    # Specific location
    ("Brgy. Calayaan, Santa Teresita, Batangas", "a municipality in Batangas Province"),
    ("Santa Teresita", "a Batangas municipality"),
    ("Calayaan", "a local barangay"),

    # Specific personal action items
    ("visit your local Barangay Hall", "visit the local Barangay Hall"),
    ("connect with the Barangay Captain or SK officials", "connect with local barangay officials"),
    ("their local Barangay Hall", "the local Barangay Hall"),
    ("submitting certified copies of their grades alongside residency documents on the very first day", "submitting requirements early"),
    ("one of the 5 allocated slots", "an allocated slot"),
    ("the candidate should proactively visit", "the student should visit"),
    ("The candidate should proactively visit", "The student should visit"),
    ("Proactively visiting the BatStateU Scholarship Office", "Visiting the BatStateU Scholarship Office"),
    ("the candidate should highlight", "the student should highlight"),
    ("The candidate should highlight", "The student should highlight"),
    ("the candidate must write", "the student should write"),
    ("the candidate's student personal portfolio", "the student's personal portfolio"),
    ("The candidate's student personal portfolio is the key", "The student's personal portfolio is key"),
    ("the candidate must confirm", "the student must confirm"),
    ("candidate must confirm", "student must confirm"),
    ("the candidate will be ineligible", "the student would be ineligible"),
    ("the candidate is a registered resident", "the student is a registered resident"),
    ("the candidate should immediately secure", "the student should secure"),
    ("the candidate should focus", "the student should focus"),
    ("the candidate should team up", "the student should team up"),

    # Second person
    ("your GWA", "the student's GWA"),
    ("your profile", "the student's profile"),
    ("your specific", "the student's specific"),
    ("you are a resident", "the student is a resident"),
    ("you can combine", "a student can combine"),
    ("This allows competitive students", "This allows students"),
    ("the candidate is a prime target", "a student would be a target"),
    ("making the candidate", "making the student"),
    ("helps the candidate", "helps the student"),

    # Portfolio references
    ("they easily meet Google's strict requirement", "the student would meet the requirement"),
    ("they are eligible", "the student is eligible"),
    ("their interest in automation", "interest in automation"),
    ("their interest in enterprise-level", "interest in enterprise-level"),
    ("their proficiency in", "proficiency in"),
    ("their goal of becoming", "the goal of becoming"),
    ("they have completed", "completed"),
    ("they can be combined", "can be combined"),

    # In the ranked table
    ("that match the candidate's profile", "based on a student profile"),
    ("matches residential profile", "matches a Batangas residential profile"),
    ("Perfect academic and socio-economic fit", "Strong academic and socio-economic fit"),
    ("Excellent fit for outstanding GWA (1.26 - 1.50)", "Excellent fit for outstanding GWA"),

    # ELAP specific
    ("ELAP (OWWA)", "an LGU educational grant"),
    ("ELAP/Provincial", "other active grants"),
    ("Holding ELAP/Provincial", "Holding other active grants"),
    ("ELAP and Provincial LGU", "other active LGU grants"),
    ("ELAP, Provincial", "other active grants"),

    # Compatibility notes cleanup
    ("Compatible with ELAP", "Compatible with other LGU grants"),
    ("Fully compatible", "Compatible"),
    ("fully compatible", "compatible"),
    ("Facilitated through Congressional District Office", "Facilitated through the local congressional district office"),
    ("Active rolling local initiative; matches residential profile", "Active rolling local initiative"),
    ("Directly linked with Alangilan Campus technology batches", "Linked with the university's technology programs"),
]

for old, new in replacements:
    text = text.replace(old, new)

with open(dst, "w", encoding="utf-8") as f:
    f.write(text)

print(f"Redacted source written to {dst}")
print(f"Original size: {len(open(src, encoding='utf-8').read())} chars")
print(f"New size: {len(text)} chars")
