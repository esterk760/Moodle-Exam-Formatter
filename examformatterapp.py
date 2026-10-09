import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QFileDialog
)

from PySide6.QtCore import Qt

from examformatter import convert_document


class DropArea(QLabel):

    def __init__(self, parent=None):

        super().__init__(parent)

        self.setAcceptDrops(True)

        self.setAlignment(
            Qt.AlignCenter
        )

        self.setText(
            "Drag & Drop a Word document here\n\n"
            "or click to select a .DOCX file"
        )

        self.setMinimumHeight(180)

        self.setStyleSheet("""
            QLabel {
                border: 3px dashed #888;
                border-radius: 15px;
                font-size: 18px;
                padding: 30px;
            }
        """)

        self.file_path = None


    def dragEnterEvent(self, event):

        if event.mimeData().hasUrls():

            urls = event.mimeData().urls()

            if urls:

                file_path = urls[0].toLocalFile()

                if file_path.lower().endswith(
                    ".docx"
                ):

                    event.acceptProposedAction()


    def dropEvent(self, event):

        urls = event.mimeData().urls()

        if not urls:
            return

        file_path = urls[0].toLocalFile()

        if file_path.lower().endswith(
            ".docx"
        ):

            self.file_path = file_path

            self.setText(
                f"Selected:\n\n"
                f"{Path(file_path).name}"
            )

            window = self.window()

            window.selected_file = file_path

            window.convert_button.setEnabled(
                True
            )


class MainWindow(QWidget):

    def __init__(self):

        super().__init__()

        self.selected_file = None

        self.setWindowTitle(
            "Moodle Question Bank Converter"
        )

        self.setMinimumSize(
            650,
            550
        )

        layout = QVBoxLayout()

        # --------------------------------------------------
        # TITLE
        # --------------------------------------------------

        title = QLabel(
            "Moodle Question Bank Converter"
        )

        title.setAlignment(
            Qt.AlignCenter
        )

        title.setStyleSheet("""
            QLabel {
                font-size: 26px;
                font-weight: bold;
                padding: 15px;
            }
        """)

        layout.addWidget(title)


        # --------------------------------------------------
        # DROP AREA
        # --------------------------------------------------

        self.drop_area = DropArea()

        layout.addWidget(
            self.drop_area
        )


        # --------------------------------------------------
        # SELECT BUTTON
        # --------------------------------------------------

        select_button = QPushButton(
            "Choose Word Document"
        )

        select_button.clicked.connect(
            self.choose_file
        )

        layout.addWidget(
            select_button
        )


        # --------------------------------------------------
        # CONVERT BUTTON
        # --------------------------------------------------

        self.convert_button = QPushButton(
            "Convert to Moodle Question Banks"
        )

        self.convert_button.setEnabled(
            False
        )

        self.convert_button.clicked.connect(
            self.convert
        )

        layout.addWidget(
            self.convert_button
        )


        # --------------------------------------------------
        # STATUS
        # --------------------------------------------------

        self.status = QTextEdit()

        self.status.setReadOnly(
            True
        )

        layout.addWidget(
            self.status
        )


        # --------------------------------------------------
        # OPEN OUTPUT
        # --------------------------------------------------

        self.open_button = QPushButton(
            "Open Output Folder"
        )

        self.open_button.setEnabled(
            False
        )

        self.open_button.clicked.connect(
            self.open_output
        )

        layout.addWidget(
            self.open_button
        )


        self.setLayout(
            layout
        )


    def choose_file(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Word Document",
            "",
            "Word Documents (*.docx)"
        )

        if file_path:

            self.selected_file = file_path

            self.drop_area.file_path = file_path

            self.drop_area.setText(
                f"Selected:\n\n"
                f"{Path(file_path).name}"
            )

            self.convert_button.setEnabled(
                True
            )


    def convert(self):

        if not self.selected_file:
            return

        input_file = Path(
            self.selected_file
        )

        # Create output folder next to the DOCX
        output_dir = (
            input_file.parent /
            f"{input_file.stem}_Moodle_Output"
        )

        self.status.clear()

        self.status.append(
            "Starting conversion..."
        )

        QApplication.processEvents()

        try:

            result = convert_document(
                input_file,
                output_dir
            )

            self.status.append(
                ""
            )

            self.status.append(
                "CONVERSION COMPLETE"
            )

            self.status.append(
                "-------------------"
            )

            self.status.append(
                f"Question pools: "
                f"{result['pools']}"
            )

            self.status.append(
                f"XML files: "
                f"{result['files']}"
            )

            if result["warnings"]:

                self.status.append(
                    ""
                )

                self.status.append(
                    "Warnings:"
                )

                for warning in result[
                    "warnings"
                ]:

                    self.status.append(
                        f"• {warning}"
                    )

            self.status.append(
                ""
            )

            self.status.append(
                f"ZIP file:\n"
                f"{result['zip_file']}"
            )

            self.open_button.setEnabled(
                True
            )

            self.output_dir = output_dir

        except Exception as error:

            self.status.append(
                ""
            )

            self.status.append(
                "ERROR:"
            )

            self.status.append(
                str(error)
            )


    def open_output(self):

        if not hasattr(
            self,
            "output_dir"
        ):
            return

        import os

        os.startfile(
            self.output_dir
        )


# ----------------------------------------------------------
# START APPLICATION
# ----------------------------------------------------------

app = QApplication(sys.argv)

window = MainWindow()

window.show()

sys.exit(
    app.exec()
)