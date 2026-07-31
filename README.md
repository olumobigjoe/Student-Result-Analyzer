# 📊 Student Results Analyzer

A Streamlit web app that lets course lecturers upload student results (CSV/Excel), instantly get a detailed statistical performance analysis, and download a full PDF report — no setup beyond Python required.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-App-red)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- **Flexible file upload** — accepts `.csv`, `.xlsx`, and `.xls`
- **Two grading layouts supported** — CA + Practical + Exam + Total, or CA + Exam + Total (no practical component)
- **Smart column detection** — auto-detects Name, CA, Practical, Exam, Total, and Course columns, with manual override if it guesses wrong
- **Multi-course support** — filter one file down to a single course if it contains several
- **Detailed statistics** — mean, median, std dev, min/max per component
- **Pass/fail summary** — configurable pass mark
- **Grade distribution** — configurable grade boundaries (A/B/C/D/F)
- **Top & bottom performers** — quickly spot high achievers and students who may need support
- **Correlation analysis** — see how CA, practical, and exam scores relate to each other
- **Component contribution breakdown** — see what % of the final total each component contributes
- **Automatic written summary** — a plain-language narrative generated directly from the computed statistics (no external AI service required)
- **Downloadable PDF report** — combines every chart, table, and the narrative into one polished PDF

## 🖥️ Demo

Upload a file like this:

| Student Name | Course | CA Score | Practical Score | Exam Score | Total |
|---|---|---|---|---|---|
| Adaobi Okafor | CSC301 | 18 | 17 | 52 | 87 |
| Chinedu Eze | CSC301 | 14 | 12 | 38 | 64 |

...and get instant charts, tables, and a downloadable PDF report.

A ready-to-use `sample_data.csv` is included so you can try it immediately.

## 🚀 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/student-results-analyzer.git
cd student-results-analyzer
```

### 2. Create a virtual environment (recommended)
```bash
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the app
```bash
streamlit run student_results_app.py
```

Open the URL Streamlit prints (usually `http://localhost:8501`).

## 📂 Project Structure
```
student-results-analyzer/
├── student_results_app.py   # Full app: UI, data processing, analysis, PDF export
├── requirements.txt         # Python dependencies
├── sample_data.csv          # Example dataset to try the app with
└── README.md
```

## 📋 Expected Column Names

The app auto-detects columns using keyword matching (case-insensitive), so your headers don't need to match exactly:

| Role | Recognized header examples |
|---|---|
| Name/ID | `Student Name`, `Name`, `Matric No`, `Student ID` |
| Course (optional) | `Course`, `Course Code`, `Subject` |
| CA | `CA Score`, `CA`, `Continuous Assessment` |
| Practical (optional) | `Practical Score`, `Practical`, `Lab Score` |
| Exam | `Exam Score`, `Exam`, `Final Exam` |
| Total (optional) | `Total`, `Grand Total` — computed automatically if missing |

## 🛠️ Built With
- [Streamlit](https://streamlit.io/) — web app framework
- [pandas](https://pandas.pydata.org/) / [NumPy](https://numpy.org/) — data processing
- [Matplotlib](https://matplotlib.org/) — charts
- [ReportLab](https://www.reportlab.com/) — PDF generation

## 🗺️ Roadmap Ideas
- [ ] Batch report generation across all courses in one upload
- [ ] Historical trend tracking across semesters
- [ ] Export to Excel in addition to PDF
- [ ] Optional local-LLM narrative generation (e.g. via Ollama) as an add-on module

## 🤝 Contributing
Issues and pull requests are welcome. If you spot a bug or have a feature idea, please open an issue first to discuss what you'd like to change.

## 📄 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
