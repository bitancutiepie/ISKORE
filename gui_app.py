import sys, os
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QLineEdit, QComboBox, QRadioButton, QButtonGroup,
    QListWidget, QListWidgetItem, QPushButton, QLabel, QFileDialog,
    QGroupBox, QMessageBox, QStatusBar, QCheckBox, QScrollArea,
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QFont, QIcon

from scholarship_engine import Profile, compute_all, compute_stats, KNOWN_AWARDS
from build_tracker import generate_workbook

YEAR_LEVELS = [
    "Incoming 3rd Year",
    "Incoming 2nd Year",
    "Freshman",
    "Graduating Senior",
    "Incoming 1st Year",
]

SCHOOLS = [
    "Batangas State University",
    "University of the Philippines Diliman",
    "Polytechnic University of the Philippines",
    "De La Salle University",
    "Ateneo de Manila University",
    "FEU Institute of Technology",
    "University of Santo Tomas",
    "Pamantasan ng Lungsod ng Maynila",
    "Technological University of the Philippines",
    "Philippine Normal University",
    "University of Pangasinan",
    "Bicol State University",
    "Central Mindanao University",
    "Davao Oriental State University",
]

MUNICIPALITIES = [
    "Batangas Province",
    "Batangas City",
    "San Juan, Batangas",
    "Santa Teresita, Batangas",
    "Lipa City",
    "Tanauan City",
    "Nasugbu, Batangas",
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Iskolar Tracker")
        self.setMinimumSize(680, 720)
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #bbb;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 18px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #1F4E79;
            }
            QPushButton#generateBtn {
                background-color: #1F4E79;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 12px 32px;
                border: none;
                border-radius: 6px;
            }
            QPushButton#generateBtn:hover {
                background-color: #2B6CB0;
            }
            QPushButton#browseBtn {
                padding: 6px 16px;
                border: 1px solid #999;
                border-radius: 4px;
            }
            QLineEdit, QComboBox {
                padding: 6px;
                border: 1px solid #bbb;
                border-radius: 4px;
            }
            QLabel#summaryLabel {
                font-size: 13px;
                padding: 10px;
                background: #E8F0FE;
                border: 1px solid #B6D4FE;
                border-radius: 6px;
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 16, 24, 16)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll_content = QWidget()
        form_layout = QVBoxLayout(scroll_content)
        form_layout.setSpacing(8)

        logo_path = Path(__file__).parent / "header.png"
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            logo_label = QLabel()
            logo_label.setPixmap(pixmap.scaledToWidth(600, Qt.TransformationMode.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            logo_label.setObjectName("logoLabel")
            form_layout.addWidget(logo_label)

        title = QLabel("Generate Your Scholarship Workbook")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #1F4E79;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        form_layout.addWidget(title)

        profile_group = QGroupBox("Your Profile")
        profile_form = QFormLayout(profile_group)
        profile_form.setSpacing(8)
        profile_form.setContentsMargins(16, 20, 16, 12)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("e.g. Juan Dela Cruz")
        profile_form.addRow("Full Name:", self.name_input)

        self.course_input = QComboBox()
        self.course_input.addItems(["BSIT"])
        self.course_input.setEditable(True)
        profile_form.addRow("Course / Program:", self.course_input)

        self.year_input = QComboBox()
        self.year_input.addItems(YEAR_LEVELS)
        profile_form.addRow("Year Level:", self.year_input)

        self.school_input = QComboBox()
        self.school_input.addItems(SCHOOLS)
        self.school_input.setEditable(True)
        profile_form.addRow("School / University:", self.school_input)

        self.municipality_input = QComboBox()
        self.municipality_input.addItems(MUNICIPALITIES)
        self.municipality_input.setEditable(True)
        profile_form.addRow("City / Municipality:", self.municipality_input)

        gender_widget = QWidget()
        gender_layout = QHBoxLayout(gender_widget)
        gender_layout.setContentsMargins(0, 0, 0, 0)
        self.gender_group = QButtonGroup(self)
        self.male_radio = QRadioButton("Male")
        self.female_radio = QRadioButton("Female")
        self.gender_group.addButton(self.male_radio)
        self.gender_group.addButton(self.female_radio)
        self.male_radio.setChecked(True)
        gender_layout.addWidget(self.male_radio)
        gender_layout.addWidget(self.female_radio)
        gender_layout.addStretch()
        profile_form.addRow("Sex:", gender_widget)

        form_layout.addWidget(profile_group)

        awards_group = QGroupBox("Held Awards")
        awards_layout = QVBoxLayout(awards_group)
        awards_layout.setContentsMargins(16, 20, 16, 12)
        self.award_checks = []
        for award in KNOWN_AWARDS:
            cb = QCheckBox(award)
            awards_layout.addWidget(cb)
            self.award_checks.append(cb)
        form_layout.addWidget(awards_group)

        output_group = QGroupBox("Output")
        output_layout = QHBoxLayout(output_group)
        output_layout.setContentsMargins(16, 20, 16, 12)
        self.output_path_input = QLineEdit()
        self.output_path_input.setText(str(Path.cwd() / "iskolar-tracker.xlsx"))
        browse_btn = QPushButton("Browse")
        browse_btn.setObjectName("browseBtn")
        browse_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(self.output_path_input)
        output_layout.addWidget(browse_btn)
        form_layout.addWidget(output_group)

        self.generate_btn = QPushButton("GENERATE WORKBOOK")
        self.generate_btn.setObjectName("generateBtn")
        self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate)
        form_layout.addWidget(self.generate_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.summary_label = QLabel()
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setWordWrap(True)
        self.summary_label.setVisible(False)
        form_layout.addWidget(self.summary_label)

        form_layout.addStretch()

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Workbook", str(Path.cwd() / "iskolar-tracker.xlsx"),
            "Excel Files (*.xlsx)"
        )
        if path:
            self.output_path_input.setText(path)

    def get_selected_awards(self):
        return [cb.text() for cb in self.award_checks if cb.isChecked()]

    def generate(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing Field", "Please enter your full name.")
            return

        profile = Profile(
            name=name,
            course=self.course_input.currentText().strip(),
            year=self.year_input.currentText().strip(),
            school=self.school_input.currentText().strip(),
            municipality=self.municipality_input.currentText().strip(),
            is_female=self.female_radio.isChecked(),
            held_awards=self.get_selected_awards(),
        )

        output_path = self.output_path_input.text().strip()
        if not output_path:
            QMessageBox.warning(self, "Missing Output", "Please specify where to save the workbook.")
            return
        if not output_path.endswith(".xlsx"):
            output_path += ".xlsx"

        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("Generating...")
        self.status_bar.showMessage("Generating workbook...")
        QApplication.processEvents()

        try:
            stats = generate_workbook(profile, output_path)
            elig = stats["eligible"]
            cond = stats["conditional"]
            inelig = stats["ineligible"]
            open_c = stats["open"]
            high = stats["high_match"]
            total = stats["total"]

            self.summary_label.setText(
                f"Workbook saved to:\n{output_path}\n\n"
                f"Total: {total} scholarships evaluated\n"
                f"Eligible: {elig}  |  Conditional: {cond}  |  Ineligible: {inelig}\n"
                f"Open: {open_c}  |  Top Targets (Score\u22657): {high}"
            )
            self.summary_label.setVisible(True)
            self.status_bar.showMessage(
                f"Done! {elig} Eligible, {cond} Conditional, {inelig} Ineligible", 5000
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate workbook:\n{e}")
            self.status_bar.showMessage("Generation failed", 5000)
        finally:
            self.generate_btn.setEnabled(True)
            self.generate_btn.setText("GENERATE WORKBOOK")


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon())
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
