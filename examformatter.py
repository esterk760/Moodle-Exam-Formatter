from docx import Document
from pathlib import Path
import re
import html
import zipfile
import json
from docx.oxml.ns import qn

def convert_document(input_file, output_dir):

    input_file = Path(input_file)
    output_dir = Path(output_dir)

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    ZIP_FILE = Path("Connected_question_banks_automated.zip")


    # ============================================================
    # READ WORD DOCUMENT
    # ============================================================

    doc = Document(input_file)

    # Extract all non-empty paragraphs
   

    paragraphs = [
        "".join(
            node.text or ""
            for node in p._p.iter(qn("w:t"))
        ).strip()
        for p in doc.paragraphs
    ]

    print(f"Read {len(paragraphs)} non-empty paragraphs.")


    # ============================================================
    # FIND QUESTION POOL MARKERS
    # ============================================================

    def parse_pool_marker(text):

        pattern = re.compile(
            r"^\[(?:call|Call) this Question Pool"
            r"(?::| something like:)\s*(.*?)\]"
            r"(?:\s*\[(.*?)\])?$",
            re.IGNORECASE
        )

        match = pattern.match(text)

        if not match:
            return None

        pool_name = match.group(1).strip()
        directive = (match.group(2) or "").strip()

        return pool_name, directive


    def get_selection_information(directive):

        directive_lower = directive.lower()

        number = None

        match = re.search(
            r"(?:pull|randomly pull)\s+(\d+)",
            directive_lower
        )

        if match:
            number = int(match.group(1))

        pull_all = "pull all" in directive_lower

        random_selection = (
            "randomly pull" in directive_lower
        )

        return number, pull_all, random_selection


    # ============================================================
    # SPLIT DOCUMENT INTO QUESTION POOLS
    # ============================================================

    pools = []

    current_pool = None

    for line in paragraphs:

        marker = parse_pool_marker(line)

        if marker:

            # Save previous pool
            if current_pool:
                pools.append(current_pool)

            pool_name, directive = marker

            number, pull_all, random_selection = (
                get_selection_information(directive)
            )

            current_pool = {
                "name": pool_name,
                "directive": directive,
                "selection_number": number,
                "pull_all": pull_all,
                "random_selection": random_selection,
                "lines": []
            }

        elif current_pool:

            current_pool["lines"].append(line)


    # Save final pool
    if current_pool:
        pools.append(current_pool)


    print(f"Found {len(pools)} question pools.")


    # ============================================================
    # IDENTIFY NON-QUESTION CONTENT
    # ============================================================

    def is_noise(text):

        # Exam headings
        if re.match(
            r"^Exam \d+ Objective$",
            text,
            re.IGNORECASE
        ):
            return True

        # Matching image instructions
        if text.startswith("Matching: Use image "):
            return True

        return False


    # ============================================================
    # PARSE STANDARD MULTIPLE-CHOICE QUESTIONS
    # ============================================================

    # def parse_standard_questions(lines):

    #     questions = []

    #     i = 0

    #     while i < len(lines):

    #         line = lines[i]

    #         # A normal question was identified by ending in ?
    #         if line.endswith("?"):

    #             question_text = line

    #             answers = []

    #             j = i + 1

    #             while j < len(lines):

    #                 candidate = lines[j]

    #                 # Next question
    #                 if candidate.endswith("?"):
    #                     break
    #                 # Ignore blank lines in between questions
    #                 if not candidate.strip():
    #                     j += 1 
    #                     continue
    #                 # Ignore bracketed instructions
    #                 if candidate.startswith("["):
    #                     j += 1
    #                     continue

    #                 # Ignore known noise
    #                 if is_noise(candidate):
    #                     j += 1
    #                     continue

    #                 answers.append(candidate)

    #                 j += 1

    #             # Keep anything with at least two answer choices
    #             if len(answers) >= 2:

    #                 questions.append({
    #                     "question": question_text,
    #                     "answers": answers
    #                 })

    #                 i = j
    #                 continue

    #         i += 1

    #     return questions


    # ============================================================
    # PARSE "NO QUESTION MARK" SECTIONS
    # ============================================================

    def parse_questions(lines):

        questions = []

        current_block = []

        for line in lines:

            # Blank line = end of current question
            if not line.strip():

                if len(current_block) >= 2:

                    question = current_block[0]

                    # Second line is the correct answer
                    correct_answer = current_block[1]

                    # Everything after the second line is an option
                    options = current_block[1:]

                    questions.append({
                        "question": question,
                        "answers": options
                    })

                # Start a new question
                current_block = []

                continue

            # Skip instructions/noise
            if line.startswith("["):
                continue

            if is_noise(line):
                continue

            current_block.append(line)

        # Process final block if there is no blank line at the end
        if len(current_block) >= 2:

            question = current_block[0]

            options = current_block[1:]

            questions.append({
                "question": question,
                "answers": options
            })

        return questions


    # ============================================================
    # PARSE EACH QUESTION POOL
    # ============================================================

    parsed_pools = []

    warnings = []


    for pool in pools:

        lines = pool["lines"]

        # Determine whether this appears to be a matching section
        matching_detected = any(
            "Matching Question" in line
            or line.startswith("Matching:")
            for line in lines
        )

        # # Determine whether this is a no-question-mark section
        # no_question_mark = any(
        #     "No question mark" in line
        #     for line in lines[:3]
        # )

        # Remove the special instruction line
        cleaned_lines = [
            line
            for line in lines
            if "No question mark" not in line
        ]


        questions = parse_questions(cleaned_lines)
        # # Select parser
        # if no_question_mark:

        #     questions = parse_no_question_mark_questions(
        #         cleaned_lines
        #     )

        # else:

        #     questions = parse_standard_questions(
        #         cleaned_lines
        #     )

        # Matching questions are not safely converted by this
        # first-pass parser.
        if matching_detected:

            warnings.append(
                f"{pool['name']}: matching content detected; "
                f"retained for manual review."
            )

        parsed_pools.append({
            **pool,
            "questions": questions,
            "matching_detected": matching_detected
        })


    # ============================================================
    # CREATE MOODLE XML
    # ============================================================

    def create_moodle_xml(pool):

        xml = []

        xml.append(
            '<?xml version="1.0" encoding="UTF-8"?>'
        )

        xml.append("<quiz>")

        questions = pool["questions"]

        for number, question in enumerate(
            questions,
            start=1
        ):

            question_name = (
                f"{pool['name']} q{number}"
            )

            xml.append(
                '  <question type="multichoice">'
            )

            # Question name
            xml.append(
                "    <name>"
                f"<text>{html.escape(question_name)}</text>"
                "</name>"
            )

            # Question text
            xml.append(
                '    <questiontext format="html">'
                f"<text><![CDATA[{html.escape(question['question'])}]]></text>"
                "</questiontext>"
            )

            # Standard Moodle settings
            xml.append(
                "    <generalfeedback format=\"html\">"
                "<text></text>"
                "</generalfeedback>"
            )

            xml.append(
                "    <defaultgrade>1.0000000</defaultgrade>"
            )

            xml.append(
                "    <penalty>0.3333333</penalty>"
            )

            xml.append(
                "    <hidden>0</hidden>"
            )

            # Shuffle answers
            xml.append(
                "    <shuffleanswers>1</shuffleanswers>"
            )

            # Do not number answers
            xml.append(
                "    <answernumbering>none</answernumbering>"
            )

            # One correct answer
            xml.append(
                "    <single>true</single>"
            )

            # ----------------------------------------------------
            # ANSWERS
            # ----------------------------------------------------

            for index, answer in enumerate(
                question["answers"]
            ):

                # First answer is correct
                if index == 0:

                    fraction = "100"

                else:

                    fraction = "0"

                xml.append(
                    f'    <answer fraction="{fraction}" format="html">'
                )

                xml.append(
                    f"      <text><![CDATA[{html.escape(answer)}]]></text>"
                )

                xml.append(
                    '      <feedback format="html">'
                    '<text></text>'
                    '</feedback>'
                )

                xml.append(
                    "    </answer>"
                )

            xml.append(
                "  </question>"
            )

        xml.append("</quiz>")

        return "\n".join(xml)


    # ============================================================
    # CREATE XML FILES
    # ============================================================

    created_files = []


    for pool in parsed_pools:

        questions = pool["questions"]

        # Don't create an empty XML file
        if not questions:
            continue

        xml = create_moodle_xml(pool)

        # Make a safe filename
        filename = re.sub(
            r'[\\/:*?"<>|]+',
            "_",
            pool["name"]
        ).strip()

        filename += ".xml"

        output_file = output_dir / filename

        output_file.write_text(
            xml,
            encoding="utf-8"
        )

        created_files.append(
            output_file
        )

        print(
            f"Created: {output_file.name} "
            f"({len(questions)} questions)"
        )


    # ============================================================
    # CREATE CONVERSION REPORT
    # ============================================================

    report = {

        "source_document": input_file.name,

        "paragraphs_extracted": len(paragraphs),

        "question_pools_detected": len(pools),

        "xml_files_created": len(created_files),

        "pools": [

            {
                "name": pool["name"],
                "directive": pool["directive"],
                "questions_detected": len(
                    pool["questions"]
                ),
                "matching_detected":
                    pool["matching_detected"]
            }

            for pool in parsed_pools
        ],

        "warnings": warnings
    }


    report_file = output_dir / "conversion_report.json"

    report_file.write_text(
        json.dumps(
            report,
            indent=2
        ),
        encoding="utf-8"
    )


    # ============================================================
    # CREATE ZIP FILE
    # ============================================================

    with zipfile.ZipFile(
        ZIP_FILE,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_archive:

        for file in created_files:

            zip_archive.write(
                file,
                arcname=file.name
            )

        zip_archive.write(
            report_file,
            arcname="conversion_report.json"
        )


    # ============================================================
    # SUMMARY
    # ============================================================

    print()
    print("=" * 60)
    print("CONVERSION COMPLETE")
    print("=" * 60)

    print(
        f"Question pools detected: {len(pools)}"
    )

    print(
        f"XML files created: {len(created_files)}"
    )

    print(
        f"ZIP created: {ZIP_FILE}"
    )

    print()

    if warnings:

        print("WARNINGS:")
        
        for warning in warnings:
            print(" -", warning)

    else:

        print("No warnings.")
        
    return {
        "pools": len(pools),
        "files": len(created_files),
        "warnings": warnings,
        "zip_file": str(ZIP_FILE)
    }