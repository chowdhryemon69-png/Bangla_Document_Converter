from pathlib import Path
from docx import Document


INPUT_FILE = Path(
    r"D:\Bangla_Document_Converter\samples\Note_TEC.docx"
)

SOURCE_FONT = "SutonnyMJ"


def main():

    document = Document(INPUT_FILE)

    print("=" * 80)
    print("DEBUG: ORIGINAL SUTONNYMJ RUNS")
    print("=" * 80)

    for p_index, paragraph in enumerate(document.paragraphs, start=1):

        if p_index not in (1, 3, 6, 8, 10, 12, 13, 14, 15):
            continue

        print()
        print("-" * 80)
        print(f"PARAGRAPH {p_index}")
        print("-" * 80)

        for r_index, run in enumerate(paragraph.runs, start=1):

            if run.font.name == SOURCE_FONT and run.text:

                print(
                    f"Run {r_index:03d} | "
                    f"Text={run.text!r}"
                )


if __name__ == "__main__":
    main()