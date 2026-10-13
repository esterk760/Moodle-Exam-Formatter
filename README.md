# Moodle Question Bank Builder

A Python desktop application that converts Word (`.docx`) documents containing multiple question banks into separate Moodle-compatible XML files for import into Moodle.

## Features

- **Word document processing:** Reads question banks from a `.docx` document.
- **Automatic question bank detection:** Identifies separate question pools using markers in the source document.
- **Moodle XML generation:** Creates individual XML files for each detected question bank.
- **Answer configuration:** Supports the project's convention that the first answer option is correct.
- **Shuffled answer choices:** Configures Moodle questions to shuffle answer choices.
- **Question bank organization:** Uses configurable question pool names for generated files and question names.
- **Batch processing:** Processes multiple question banks in one run.
- **ZIP output:** Packages generated XML files for convenient storage and transfer.
- **Desktop interface:** Provides a graphical interface for selecting a Word document and running the conversion.
- **Windows executable support:** Can be packaged as a standalone application using PyInstaller.

## Requirements

- Windows
- Python 3.10 or later
- The Python packages listed in `requirements.txt`

## Project Structure

```text
MoodleExamFormatter/
├── app.py              # Desktop application interface
├── examformatter.py    # Document parsing and Moodle XML generation
├── requirements.txt    # Python dependencies
├── README.md           # Project documentation
└── .gitignore          # Files excluded from Git
```

Additional files and folders may be created during conversion, including generated XML files, ZIP archives, and build artifacts.

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/Moodle-Exam-Formatter.git
cd MoodleExamFormatter
```

Replace `YOUR-USERNAME` with the appropriate GitHub username.

Alternatively, download the repository as a ZIP file from GitHub and extract it.

### 2. Create a virtual environment

On Windows, open PowerShell in the project folder and run:

```powershell
py -m venv .venv
```

Activate the environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, you can run the Python executable directly from `.venv\Scripts\python.exe` instead.

### 3. Install dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Usage

### Run the desktop application

From the activated virtual environment, run:

```powershell
python app.py
```

### Convert a Word document

1. Open the application.
2. Select or drag and drop the source `.docx` file.
3. Start the conversion.
4. Review the conversion status and any warnings.
5. Open the output folder to access the generated Moodle XML files and ZIP archive, where available.

The converter detects question pools and processes them into separate XML files. Output details depend on the question formats and markers present in the source document.

## Source Document Format

The converter expects a Word document containing recognizable question pool markers and question/answer content.

For reliable results:

- Keep question pool markers consistent.
- Use a consistent layout for questions and answer choices.
- Ensure the intended correct answer follows the project's expected answer-order convention.
- Review any automatically detected matching questions.
- Check the generated XML before importing a large batch into Moodle.


## Development Notes

The application is intended to automate repetitive question-bank conversion while allowing users to inspect results before importing them into Moodle.

Future enhancements may include:

- A dedicated question validation screen.
- Improved reporting and error handling.
- Direct Moodle Web Services integration.
- Automated creation of question-bank categories.
- Automated creation of a comprehensive quiz for question-bank checking.
- Verification of imported question counts.

These are potential enhancements and should not be considered implemented unless they are present in the current codebase.

