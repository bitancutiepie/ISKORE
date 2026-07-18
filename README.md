# Iskolar Tracker — Scholarship Portfolio Manager

A data-driven Python tool that parses a comprehensive scholarship analysis into a structured, color-coded Excel tracker with automated eligibility scoring — built for Filipino SUC students navigating the complex tertiary funding ecosystem.

![Dashboard](screenshots/01-dashboard.png)
![Eligible Scholarships](screenshots/02-eligible-list.png)

## Overview

The Philippine tertiary funding landscape involves 40+ national, local government, and corporate scholarship programs — each with different deadlines, eligibility rules, and compatibility constraints. This project transforms a raw strategic analysis text into a living portfolio tracker:

- **Parses** 47 scholarship entries from a structured analysis document
- **Scores** each against a user profile (course, year level, residency, gender, existing awards)
- **Generates** a 7-sheet Excel workbook with filters, color-coding, and cross-sheet references
- **Renders** clean table screenshots for portfolio documentation

## Features

| Sheet | Contents |
|---|---|
| **Dashboard** | Profile summary, currently held awards, portfolio stats, upcoming deadlines |
| **Master List** | All 47 scholarships with 21 fields + personal eligibility column |
| **My Eligible Scholarships** | Filtered view: only what you can actually apply for |
| **Application Tracker** | Status tracking: Preparing → Applied → Interviewing → Awarded |
| **Ranked Prioritization** | Eligible scholarships sorted by match score and probability |
| **Comparison Matrix** | Side-by-side comparison of all programs |
| **Tactical Manuals** | Strategic guides: doc prep, interview tips, execution issues |

![Application Tracker](screenshots/03-application-tracker.png)
![Ranked Prioritization](screenshots/04-ranked-priorities.png)

## Tech Stack

- **Python 3** — data processing and workbook generation
- **openpyxl** — Excel creation with formatting, filters, and hyperlinks
- **Pillow** — programmatic table rendering for screenshots
- **Data-driven design** — eligibility engine parametrized by user profile

## Quick Start

```bash
# Clone the repo
git clone https://github.com/yourusername/iskolar-tracker.git
cd iskolar-tracker

# Install dependencies
pip install -r requirements.txt

# Generate the tracker (edit PROFILE in build_tracker.py first)
python build_tracker.py

# (Optional) Regenerate screenshots
python capture_screenshots.py
```

## Profile Customization

The eligibility engine is configured through a single dictionary at the top of `build_tracker.py`:

```python
PROFILE = {
    "name": "Your Name",
    "course": "BSIT",
    "year": "Incoming 3rd Year",
    "school": "Batangas State University",
    "municipality": "Your Municipality, Province",
    "is_female": False,
    "held_awards": ["LGU Educational Grant", "Provincial Merit Scholarship"],
}
```

The engine automatically evaluates each scholarship against:
- **Year level eligibility** — freshman-only vs. continuing students
- **Course/program fit** — BSIT vs. course-restricted programs
- **Residency restrictions** — national vs. LGU-specific programs
- **Gender-specific grants** — female-only scholarships
- **Award compatibility** — no-conflict clauses with existing grants

## Project Structure

```
iskolar-tracker/
├── build_tracker.py                   # Main generator script
├── capture_screenshots.py             # Screenshot renderer
├── requirements.txt                   # Python dependencies
├── iskolar-tracker.xlsx               # Generated workbook (sample)
├── data/
│   └── scholarship-analysis-source.txt # Source analysis (redacted)
├── screenshots/
│   ├── 01-dashboard.png
│   ├── 02-eligible-list.png
│   ├── 03-application-tracker.png
│   └── 04-ranked-priorities.png
└── README.md
```

## Data Source

The scholarship data was compiled from official government and corporate scholarship portals including DOST-SEI, CHED, UniFAST, GBF, DICT, and various LGU and private foundation programs for Academic Year 2026–2027. All monetary values are in Philippine Pesos (₱) unless noted.

## License

MIT
