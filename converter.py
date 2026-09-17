
from pathlib import Path
import json
import subprocess
import sys
import re
from copy import deepcopy

from docx import Document


if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_FILE = Path(
    r"D:\Bangla_Document_Converter\samples\Note_TEC.docx"
)

OUTPUT_DIR = Path(
    r"D:\Bangla_Document_Converter\output"
)

OUTPUT_FILE = OUTPUT_DIR / "Note_TEC_Unicode_Nikosh.docx"

BRIDGE_FILE = Path(
    r"D:\Bangla_Document_Converter\bijoy_bridge.mjs"
)

SOURCE_FONT = "SutonnyMJ"
TARGET_FONT = "Nikosh"
LATIN_TARGET_FONT = "Times New Roman"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# English tokens that must remain unchanged.
#
# IMPORTANT:
# We do NOT use a generic A-Z detector because Bijoy Bangla
# itself contains ASCII characters such as:
#
#   wefvM
#   Bs
#   G
#   KwgwUi
#
# Those are Bangla, not English.
#
PROTECTED_ENGLISH_TOKENS = {
    "TOR-1",
    "TOR-2",
}


# ============================================================
# BIJOY CONVERTER
# ============================================================

class BijoyConverter:

    def __init__(self, bridge_file: Path):
        self.bridge_file = bridge_file
        self.process = None
        self.counter = 0

    def start(self):

        if not self.bridge_file.exists():
            raise FileNotFoundError(
                f"Node bridge not found:\n{self.bridge_file}"
            )

        self.process = subprocess.Popen(
            ["node", str(self.bridge_file)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="strict",
            bufsize=1,
        )

    def convert(self, text: str) -> str:

        if not text:
            return text

        if self.process is None:
            self.start()

        self.counter += 1

        request_id = self.counter

        request = {
            "id": request_id,
            "text": text,
        }

        line = (
            json.dumps(
                request,
                ensure_ascii=False,
            )
            + "\n"
        )

        try:

            self.process.stdin.write(line)

            self.process.stdin.flush()

            response_line = self.process.stdout.readline()

            if not response_line:

                stderr_output = ""

                if self.process.stderr:
                    try:
                        stderr_output = self.process.stderr.read()
                    except Exception:
                        stderr_output = ""

                raise RuntimeError(
                    "Node bridge stopped unexpectedly.\n"
                    + stderr_output
                )

            response = json.loads(response_line)

        except Exception as exc:

            raise RuntimeError(
                f"Bijoy conversion failed for text:\n"
                f"{text!r}\n\n"
                f"Error: {exc}"
            ) from exc

        if "error" in response:

            raise RuntimeError(
                "Node converter error:\n"
                + str(response["error"])
            )

        if response.get("id") != request_id:

            raise RuntimeError(
                "Node bridge returned an unexpected response ID."
            )

        return response.get("text", "")

    def close(self):

        if self.process is None:
            return

        try:

            if self.process.stdin:
                self.process.stdin.close()

        except Exception:
            pass

        try:

            self.process.wait(timeout=5)

        except subprocess.TimeoutExpired:

            self.process.kill()

        self.process = None


# ============================================================
# FONT HANDLING
# ============================================================

def set_run_font(run, font_name: str):

    """
    Set all relevant Word font attributes.

    This is important for Bangla because Word may use
    the complex-script font rather than only ascii/hAnsi.
    """

    run.font.name = font_name

    rPr = run._element.get_or_add_rPr()

    rFonts = rPr.rFonts

    if rFonts is None:

        from docx.oxml import OxmlElement

        rFonts = OxmlElement("w:rFonts")

        rPr.insert(0, rFonts)

    rFonts.set(
        f"{{{W_NS}}}ascii",
        font_name,
    )

    rFonts.set(
        f"{{{W_NS}}}hAnsi",
        font_name,
    )

    rFonts.set(
        f"{{{W_NS}}}eastAsia",
        font_name,
    )

    rFonts.set(
        f"{{{W_NS}}}cs",
        font_name,
    )


LATIN_PATTERN = re.compile(r"[A-Za-z0-9]")
BENGALI_PATTERN = re.compile(r"[\u0980-\u09FF]")
MIXED_RUN_SEGMENT_RE = re.compile(
    r"[A-Za-z0-9]+(?:[-_/][A-Za-z0-9]+)*|[\u0980-\u09FF]+|[^A-Za-z0-9\u0980-\u09FF]+"
)


def contains_latin_text(text: str) -> bool:
    return bool(text and LATIN_PATTERN.search(text))


def split_run_for_latin_font(run):
    """Split a run so Bangla and Latin segments can keep their assigned fonts."""
    text = run.text or ""
    if not text or not contains_latin_text(text):
        return False

    parent = run._r.getparent()
    if parent is None:
        return False

    segments = []
    last_index = 0
    for match in MIXED_RUN_SEGMENT_RE.finditer(text):
        segment_text = match.group(0)
        if not segment_text:
            continue
        if segment_text and any(ch.isascii() and (ch.isalpha() or ch.isdigit()) for ch in segment_text):
            segments.append((segment_text, "latin"))
        elif BENGALI_PATTERN.search(segment_text):
            segments.append((segment_text, "bangla"))
        elif segment_text.strip():
            segments.append((segment_text, "other"))
        last_index = match.end()

    if not segments:
        set_run_font(run, LATIN_TARGET_FONT)
        return True

    if len(segments) == 1 and segments[0][1] == "latin":
        set_run_font(run, LATIN_TARGET_FONT)
        return True

    original_r = run._r
    original_index = parent.index(original_r)
    all_new_runs = []

    for segment_text, segment_kind in segments:
        if not segment_text:
            continue

        new_run = create_run_from_original(run, segment_text)
        if segment_kind == "latin":
            set_run_font(new_run, LATIN_TARGET_FONT)
        elif segment_kind == "bangla":
            set_run_font(new_run, TARGET_FONT)
        else:
            # Keep punctuation/spacing attached to the neighboring content's
            # default style in a way that does not alter the run text.
            if BENGALI_PATTERN.search(segment_text):
                set_run_font(new_run, TARGET_FONT)
            else:
                set_run_font(new_run, LATIN_TARGET_FONT)
        all_new_runs.append(new_run._r)

    for offset, new_r in enumerate(all_new_runs):
        parent.insert(original_index + offset, new_r)
    parent.remove(original_r)
    return True


def normalize_latin_runs_in_document(document):
    """Ensure that Latin/English text uses Times New Roman while Bangla stays Nikosh."""
    for paragraph in document.paragraphs:
        for run in list(paragraph.runs):
            if run.text and contains_latin_text(run.text):
                if BENGALI_PATTERN.search(run.text):
                    split_run_for_latin_font(run)
                else:
                    set_run_font(run, LATIN_TARGET_FONT)

    for table in document.tables:
        def normalize_table(tbl):
            for row in tbl.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in list(paragraph.runs):
                            if run.text and contains_latin_text(run.text):
                                if BENGALI_PATTERN.search(run.text):
                                    split_run_for_latin_font(run)
                                else:
                                    set_run_font(run, LATIN_TARGET_FONT)
                    for nested_table in cell.tables:
                        normalize_table(nested_table)

        normalize_table(table)

    return document


# ============================================================
# PROTECTED TOKEN HANDLING
# ============================================================

def split_protected_tokens(text: str):

    """
    Split text around explicitly known English tokens.

    Example:

        Bangla TOR-1 Bangla TOR-2

    becomes:

        Bangla
        TOR-1
        Bangla
        TOR-2

    We intentionally do NOT use generic English detection.
    """

    if not text:
        return []

    tokens = sorted(
        PROTECTED_ENGLISH_TOKENS,
        key=len,
        reverse=True,
    )

    pattern = "(" + "|".join(
        re.escape(token)
        for token in tokens
    ) + ")"

    return re.split(pattern, text)


# ============================================================
# SIMPLE BIJOY CONVERSION
# ============================================================

def convert_bijoy_text(
    text: str,
    converter: BijoyConverter,
) -> str:

    """
    Convert text while leaving explicitly protected English
    tokens untouched.
    """

    if not text:
        return text

    parts = split_protected_tokens(text)

    output_parts = []

    for part in parts:

        if not part:
            continue

        if part in PROTECTED_ENGLISH_TOKENS:

            # English token is copied exactly.
            output_parts.append(part)

        else:

            output_parts.append(
                converter.convert(part)
            )

    return "".join(output_parts)


# ============================================================
# RUN COPYING
# ============================================================

def create_run_from_original(
    original_run,
    text: str,
):

    """
    Create a new Word run by copying the complete XML of
    the original run.

    This preserves run-level formatting such as:

    - bold
    - italic
    - underline
    - font size
    - color
    - highlight
    - character properties
    - etc.

    The caller can then change only the text/font required.
    """

    new_r = deepcopy(
        original_run._r
    )

    # Remove existing text nodes.
    #
    # Other run properties remain untouched.
    for child in list(new_r):

        if child.tag.endswith("}t"):

            new_r.remove(child)

    # python-docx requires both the XML element and the parent story object.
    # Using the original run's parent keeps the cloned run attached to the
    # document structure without disturbing the existing formatting tree.
    new_run = original_run.__class__(
        new_r,
        original_run._parent,
    )

    new_run.text = text

    return new_run


# Evidence-based compatibility forms found in the source DOCX.
SUTONNY_COMPATIBILITY_NORMALIZATIONS = {
    "AvnŸvb": "Avn&evb",
    "wb¤œiƒc": "wbgœiƒc",
    "me©wb¤œ": "me©wbgœ",
}


def normalize_sutonny_compatibility(text: str) -> str:
    """Normalize only known SutonnyMJ legacy variants before Bijoy conversion."""
    for source_text, canonical_text in SUTONNY_COMPATIBILITY_NORMALIZATIONS.items():
        text = text.replace(source_text, canonical_text)
    return text


# ============================================================
# MIXED RUN PROCESSING
# ============================================================

def process_run(
    run,
    converter: BijoyConverter,
):

    """
    Process one Word run.

    Only runs explicitly using SutonnyMJ are converted.

    If the run contains protected English tokens, it is split
    into multiple runs so the English token can retain its
    original formatting/font.
    """

    # --------------------------------------------------------
    # Only process SutonnyMJ runs.
    # --------------------------------------------------------

    if run.font.name != SOURCE_FONT:
        return False

    if not run.text:
        return False

    original_text = normalize_sutonny_compatibility(run.text)
    run.text = original_text

    # --------------------------------------------------------
    # Split around protected English tokens.
    # --------------------------------------------------------

    parts = split_protected_tokens(
        original_text
    )

    has_protected_token = any(
        part in PROTECTED_ENGLISH_TOKENS
        for part in parts
    )

    # --------------------------------------------------------
    # Case 1:
    # Pure SutonnyMJ / no protected English.
    # --------------------------------------------------------

    if not has_protected_token:

        converted_text = converter.convert(
            original_text
        )

        run.text = converted_text

        set_run_font(
            run,
            TARGET_FONT,
        )

        return original_text != converted_text

    # --------------------------------------------------------
    # Case 2:
    # Mixed SutonnyMJ + protected English.
    #
    # We create separate runs.
    # --------------------------------------------------------

    parent = run._r.getparent()

    original_r = run._r

    new_elements = []

    for part in parts:

        if not part:
            continue

        # ----------------------------------------------------
        # English protected token
        # ----------------------------------------------------

        if part in PROTECTED_ENGLISH_TOKENS:

            new_run = create_run_from_original(
                run,
                part,
            )

            # IMPORTANT:
            #
            # Keep the original run's formatting.
            #
            # We do NOT set Nikosh here.
            #
            new_elements.append(
                new_run._r
            )

        # ----------------------------------------------------
        # Bijoy Bangla / surrounding text
        # ----------------------------------------------------

        else:

            converted = converter.convert(
                part
            )

            new_run = create_run_from_original(
                run,
                converted,
            )

            set_run_font(
                new_run,
                TARGET_FONT,
            )

            new_elements.append(
                new_run._r
            )

    # --------------------------------------------------------
    # Insert new runs at the exact position of the original.
    # --------------------------------------------------------

    original_index = parent.index(
        original_r
    )

    for offset, new_r in enumerate(
        new_elements
    ):

        parent.insert(
            original_index + offset,
            new_r,
        )

    # Remove original mixed run.
    parent.remove(
        original_r
    )

    return True


# ============================================================
# RUN CONSOLIDATION PRE-PROCESSING
# ============================================================

def colors_equal(c1, c2) -> bool:
    """
    Compare two Font.color objects for genuine visual equivalence.
    """
    if c1 is None and c2 is None:
        return True
    if (c1 is None) != (c2 is None):
        return False
    if c1.type != c2.type:
        return False
    if c1.rgb != c2.rgb:
        return False
    if getattr(c1, "theme_color", None) != getattr(c2, "theme_color", None):
        return False
    return True


def runs_have_identical_formatting(r1, r2) -> bool:
    """
    Check if two runs have genuinely identical run-level formatting.
    Only runs with identical formatting can be consolidated.
    """
    if r1.font.name != SOURCE_FONT or r2.font.name != SOURCE_FONT:
        return False

    if r1.bold != r2.bold:
        return False
    if r1.italic != r2.italic:
        return False
    if r1.underline != r2.underline:
        return False
    if r1.font.size != r2.font.size:
        return False
    if r1.font.strike != r2.font.strike:
        return False
    if r1.font.subscript != r2.font.subscript:
        return False
    if r1.font.superscript != r2.font.superscript:
        return False
    if r1.font.highlight_color != r2.font.highlight_color:
        return False

    if not colors_equal(r1.font.color, r2.font.color):
        return False

    if r1.style != r2.style:
        return False

    return True


def can_merge_runs(r1, r2) -> bool:
    """
    Check if two adjacent runs can be safely merged prior to conversion.

    Requirements:
    1. Both runs must use SOURCE_FONT (SutonnyMJ).
    2. Neither run contains any protected English token.
    3. Both runs have genuinely identical run-level formatting.
    """
    if r1.font.name != SOURCE_FONT or r2.font.name != SOURCE_FONT:
        return False

    r1_text = r1.text or ""
    r2_text = r2.text or ""

    if any(tok in r1_text for tok in PROTECTED_ENGLISH_TOKENS):
        return False
    if any(tok in r2_text for tok in PROTECTED_ENGLISH_TOKENS):
        return False

    return runs_have_identical_formatting(r1, r2)


def consolidate_runs_in_paragraph(paragraph) -> int:
    """
    Consolidate adjacent SutonnyMJ runs in a paragraph that share genuinely
    identical formatting.

    This prevents Word run fragmentation from breaking Bijoy pre-base vowel
    signs (e.g. e-kar † / ‡) or post-base modifiers (e.g. reph ©) away from
    their consonants when passed to the converter.
    """
    runs = list(paragraph.runs)
    if len(runs) <= 1:
        return 0

    merged_count = 0
    i = 0
    while i < len(runs) - 1:
        r1 = runs[i]
        r2 = runs[i + 1]

        if can_merge_runs(r1, r2):
            r1.text = (r1.text or "") + (r2.text or "")
            parent = r2._r.getparent()
            if parent is not None:
                parent.remove(r2._r)
            runs.pop(i + 1)
            merged_count += 1
        else:
            i += 1

    return merged_count


# ============================================================
# PARAGRAPH PROCESSING
# ============================================================

def process_paragraph(
    paragraph,
    converter: BijoyConverter,
):

    changed = 0

    # Consolidate adjacent SutonnyMJ runs with identical formatting
    # before conversion so Bijoy character sequences are intact.
    consolidate_runs_in_paragraph(paragraph)

    # Make a snapshot because process_run() may replace runs.
    runs = list(
        paragraph.runs
    )

    for run in runs:

        if process_run(
            run,
            converter,
        ):

            changed += 1

    return changed


# ============================================================
# TABLE PROCESSING
# ============================================================

def process_table(
    table,
    converter: BijoyConverter,
):

    changed = 0

    for row in table.rows:

        for cell in row.cells:

            # ------------------------------------------------
            # Normal paragraphs
            # ------------------------------------------------

            for paragraph in cell.paragraphs:

                changed += process_paragraph(
                    paragraph,
                    converter,
                )

            # ------------------------------------------------
            # Nested tables
            # ------------------------------------------------

            for nested_table in cell.tables:

                changed += process_table(
                    nested_table,
                    converter,
                )

    return changed


# ============================================================
# DOCUMENT PROCESSING
# ============================================================

def process_document(
    document,
    converter: BijoyConverter,
):

    changed_runs = 0

    # --------------------------------------------------------
    # Body paragraphs
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        changed_runs += process_paragraph(
            paragraph,
            converter,
        )

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    for table in document.tables:

        changed_runs += process_table(
            table,
            converter,
        )

    return changed_runs


# ============================================================
# COUNT SUTONNY RUNS
# ============================================================

def count_sutonny_runs(
    document,
):

    body_count = 0
    table_count = 0

    # --------------------------------------------------------
    # Body
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        for run in paragraph.runs:

            if run.font.name == SOURCE_FONT:

                body_count += 1

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    def count_table(table):

        count = 0

        for row in table.rows:

            for cell in row.cells:

                for paragraph in cell.paragraphs:

                    for run in paragraph.runs:

                        if run.font.name == SOURCE_FONT:

                            count += 1

                for nested_table in cell.tables:

                    count += count_table(
                        nested_table
                    )

        return count

    for table in document.tables:

        table_count += count_table(
            table
        )

    return body_count, table_count


# ============================================================
# MAIN
# VALIDATION UTILITY
# ============================================================

def main():
    pass

'''
def validate_converted_docx(output_file: Path) -> dict:
    """
    Perform essential integrity validation on the generated DOCX.
    Verifies that:
    1. SutonnyMJ runs are completely eliminated (count == 0).
    2. Nikosh font is actively applied to converted Bangla text.
    3. English/Times New Roman runs are preserved.
    4. Protected TOR-1 / TOR-2 tokens are preserved without corruption.
    5. No 'ঞঙজ' Bijoy misconversion artefact exists.
    """
    output_file = Path(output_file)
    if not output_file.exists():
        return {
            "valid": False,
            "details": f"File does not exist: {output_file}",
        }

    print("=" * 65)
    doc = Document(output_file)

    print(
        "BANGLA DOCUMENT CONVERTER"
    )
    nikosh_runs = 0
    sutonny_runs = 0
    times_runs = 0
    other_runs = 0
    tor_found = False
    corrupt_tor_found = False

    print("=" * 65)
    def check_paragraph(p):
        nonlocal nikosh_runs, sutonny_runs, times_runs, other_runs, tor_found, corrupt_tor_found
        text = p.text or ""
        if "TOR-1" in text or "TOR-2" in text:
            tor_found = True
        if "ঞঙজ" in text:
            corrupt_tor_found = True

    print(
        f"Input : {INPUT_FILE}"
    )
        for r in p.runs:
            fn = r.font.name
            if fn == TARGET_FONT:
                nikosh_runs += 1
            elif fn == SOURCE_FONT:
                sutonny_runs += 1
            elif fn == "Times New Roman":
                times_runs += 1
            elif r.text:
                other_runs += 1

    print(
        f"Output: {OUTPUT_FILE}"
    )
    for p in doc.paragraphs:
        check_paragraph(p)

    print()
    def check_table(tbl):
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    check_paragraph(p)
                for nested in cell.tables:
                    check_table(nested)

    # --------------------------------------------------------
    # Safety checks
    # --------------------------------------------------------
    for t in doc.tables:
        check_table(t)

    if not INPUT_FILE.exists():
    if sutonny_runs > 0:
        return {
            "valid": False,
            "nikosh_runs": nikosh_runs,
            "sutonny_runs": sutonny_runs,
            "times_runs": times_runs,
            "details": f"Failed: {sutonny_runs} unconverted SutonnyMJ runs found.",
        }

        raise FileNotFoundError(
            f"Input document not found:\n"
            f"{INPUT_FILE}"
        )
    if nikosh_runs == 0:
        return {
            "valid": False,
            "nikosh_runs": nikosh_runs,
            "sutonny_runs": sutonny_runs,
            "times_runs": times_runs,
            "details": "Failed: No Nikosh font runs found in document.",
        }

    if not BRIDGE_FILE.exists():
    if corrupt_tor_found:
        return {
            "valid": False,
            "nikosh_runs": nikosh_runs,
            "sutonny_runs": sutonny_runs,
            "times_runs": times_runs,
            "details": "Failed: Corrupted TOR token ('ঞঙজ') detected.",
        }

        raise FileNotFoundError(
            f"Node bridge not found:\n"
            f"{BRIDGE_FILE}"
        )
    return {
        "valid": True,
        "nikosh_runs": nikosh_runs,
        "sutonny_runs": sutonny_runs,
        "times_runs": times_runs,
        "other_runs": other_runs,
        "tor_preserved": tor_found,
        "details": "PASS",
    }

    if (
        INPUT_FILE.resolve()
        == OUTPUT_FILE.resolve()
    ):

        raise RuntimeError(
            "SAFETY ERROR: "
            "Input and output files are identical."
        )
# ============================================================
# LOGGING
# ============================================================

    # --------------------------------------------------------
    # Create output directory.
    # --------------------------------------------------------
def log_conversion_result(
    source_file: str,
    output_file: str,
    success: bool,
    validation_status: str,
    error_message: str = None,
    log_dir: Path = None,
):
    """
    Append a concise conversion record to logs/converter.log.
    Never logs document text content.
    """
    if log_dir is None:
        log_dir = Path(__file__).resolve().parent / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "converter.log"

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # --------------------------------------------------------
    # Load original document.
    #
    # This is read-only until save().
    # --------------------------------------------------------
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}]\n")
        f.write(f"  Source     : {source_file}\n")
        f.write(f"  Output     : {output_file}\n")
        f.write(f"  Result     : {'SUCCESS' if success else 'FAILED'}\n")
        f.write(f"  Validation : {validation_status}\n")
        if error_message:
            f.write(f"  Error      : {error_message}\n")
        f.write("\n")

    document = Document(
        INPUT_FILE
    )

    # --------------------------------------------------------
    # Count source runs.
    # --------------------------------------------------------
# ============================================================
# FULL CONVERSION PIPELINE
# ============================================================

    body_count, table_count = (
        count_sutonny_runs(
            document
        )
    )
def convert_docx_file(
    input_file: Path,
    output_file: Path,
    bridge_file: Path = BRIDGE_FILE,
    progress_callback = None,
) -> dict:
    """
    High-level conversion entry point suitable for CLI and GUI.

    print(
        f"SutonnyMJ body runs  : {body_count}"
    )
    Pipeline stages:
    1. Safety & parameter checks.
    2. Loading document (read-only from source).
    3. Collecting all paragraphs across body and nested tables.
    4. Consolidating adjacent compatible SutonnyMJ runs.
    5. Converting complete Bijoy sequences through Node bridge.
    6. Applying Nikosh font.
    7. Saving strictly to output_file.
    8. Validating output integrity.
    9. Logging result.
    """
    input_path = Path(input_file)
    output_path = Path(output_file)
    bridge_path = Path(bridge_file)

    print(
        f"SutonnyMJ table runs : {table_count}"
    )
    # Safety checks
    if not input_path.exists():
        raise FileNotFoundError(f"Input document not found:\n{input_path}")
    if input_path.suffix.lower() != ".docx":
        raise ValueError(f"Only .docx documents are supported. Received: {input_path.name}")
    if not bridge_path.exists():
        raise FileNotFoundError(f"Node bridge not found:\n{bridge_path}")
    if input_path.resolve() == output_path.resolve():
        raise RuntimeError("SAFETY ERROR: Input and output files are identical.")

    print()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # --------------------------------------------------------
    # Start Node converter.
    # --------------------------------------------------------
    def report(status: str, percent: int):
        if progress_callback:
            try:
                progress_callback(status, percent)
            except Exception:
                pass

    converter = BijoyConverter(
        BRIDGE_FILE
    )

    try:
        report("Loading document...", 5)
        document = Document(input_path)

        converter.start()
        report("Scanning document...", 10)
        all_paragraphs = list(document.paragraphs)
        for table in document.tables:
            def collect_table(tbl):
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            all_paragraphs.append(p)
                        for nested in cell.tables:
                            collect_table(nested)
            collect_table(table)

        # ----------------------------------------------------
        # Convert document.
        # ----------------------------------------------------
        report("Consolidating Bangla runs...", 20)
        for p in all_paragraphs:
            consolidate_runs_in_paragraph(p)

        changed_runs = process_document(
            document,
            converter,
        # Collect eligible SutonnyMJ runs
        eligible_runs = []
        for p in all_paragraphs:
            for r in p.runs:
                if r.font.name == SOURCE_FONT and r.text:
                    eligible_runs.append(r)

        total_runs = len(eligible_runs)
        report(f"Starting converter (0/{total_runs})...", 30)

        converter = BijoyConverter(bridge_path)
        changed_runs = 0
        try:
            converter.start()
            for idx, run in enumerate(eligible_runs, start=1):
                if process_run(run, converter):
                    changed_runs += 1
                if total_runs > 0:
                    pct = 30 + int(50 * idx / total_runs)
                    report(f"Converting Bangla... {idx}/{total_runs}", pct)
        finally:
            converter.close()

        report("Applying Nikosh...", 85)

        report("Saving document...", 90)
        document.save(output_path)

        report("Validating output...", 95)
        val_result = validate_converted_docx(output_path)
        if not val_result.get("valid", False):
            raise RuntimeError(f"Output validation failed: {val_result.get('details', 'Unknown validation error')}")

        report("Complete", 100)

        log_conversion_result(
            source_file=str(input_path),
            output_file=str(output_path),
            success=True,
            validation_status=val_result.get("details", "PASS"),
        )

    finally:
        return {
            "success": True,
            "converted_runs": changed_runs,
            "validation": val_result,
            "output_file": str(output_path),
            "source_file": str(input_path),
        }

        converter.close()
    except Exception as exc:
        log_conversion_result(
            source_file=str(input_path),
            output_file=str(output_path),
            success=False,
            validation_status="FAILED",
            error_message=str(exc),
        )
        raise

    # --------------------------------------------------------
    # Save ONLY to output file.
    #
    # Original Note_TEC.docx remains untouched.
    # --------------------------------------------------------

    document.save(
        OUTPUT_FILE
    )
# ============================================================
# MAIN
# ============================================================

    print(
        f"Converted runs       : {changed_runs}"
    )
def main():

    print("=" * 65)
    print("BANGLA DOCUMENT CONVERTER")
    print("=" * 65)
    print(f"Input : {INPUT_FILE}")
    print(f"Output: {OUTPUT_FILE}")
    print()

    print(
        "Conversion completed successfully."
    )
    def console_progress(msg, pct):
        print(f"[{pct:3d}%] {msg}")

    print(
        f"Output saved to:\n"
        f"{OUTPUT_FILE}"
    result = convert_docx_file(
        INPUT_FILE,
        OUTPUT_FILE,
        BRIDGE_FILE,
        progress_callback=console_progress,
    )

    print()

    print(
        "Original document was NOT modified."
    )

    print(f"Converted runs       : {result['converted_runs']}")
    print(f"Validation status    : {result['validation']['details']}")
    print()
    print("Conversion completed successfully.")
    print(f"Output saved to:\n{OUTPUT_FILE}")
    print()
    print("Original document was NOT modified.")
    print("=" * 65)


# ============================================================
# PROGRAM ENTRY
# ============================================================

'''

# ============================================================
# GUI/CLI CONVERSION API
# ============================================================

def validate_converted_docx(output_file: Path) -> dict:
    """Run the essential output checks without logging document contents."""
    output_path = Path(output_file)
    if not output_path.exists():
        return {"valid": False, "details": f"File does not exist: {output_path}"}

    document = Document(output_path)
    nikosh_runs = 0
    sutonny_runs = 0
    times_runs = 0
    tor_found = False
    corrupt_tor_found = False

    def inspect_paragraph(paragraph):
        nonlocal nikosh_runs, sutonny_runs, times_runs, tor_found, corrupt_tor_found
        text = paragraph.text or ""
        tor_found = tor_found or "TOR-1" in text or "TOR-2" in text
        corrupt_tor_found = corrupt_tor_found or "ঞঙজ" in text
        for run in paragraph.runs:
            if run.font.name == TARGET_FONT:
                nikosh_runs += 1
            elif run.font.name == SOURCE_FONT:
                sutonny_runs += 1
            elif run.font.name == "Times New Roman":
                times_runs += 1

    def inspect_table(table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    inspect_paragraph(paragraph)
                for nested_table in cell.tables:
                    inspect_table(nested_table)

    for paragraph in document.paragraphs:
        inspect_paragraph(paragraph)
    for table in document.tables:
        inspect_table(table)

    if sutonny_runs:
        details = f"Failed: {sutonny_runs} unconverted SutonnyMJ runs found."
    elif not nikosh_runs:
        details = "Failed: No Nikosh font runs found in document."
    elif corrupt_tor_found:
        details = "Failed: Corrupted TOR token detected."
    else:
        details = "PASS"

    return {
        "valid": details == "PASS",
        "nikosh_runs": nikosh_runs,
        "sutonny_runs": sutonny_runs,
        "times_runs": times_runs,
        "tor_preserved": tor_found,
        "details": details,
    }


def log_conversion_result(
    source_file: str,
    output_file: str,
    success: bool,
    validation_status: str,
    error_message: str = None,
    log_dir: Path = None,
):
    """Append concise conversion metadata without document contents."""
    import datetime

    directory = Path(log_dir) if log_dir is not None else Path(__file__).resolve().parent / "logs"
    directory.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with (directory / "converter.log").open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}]\n")
        log_file.write(f"  Source     : {source_file}\n")
        log_file.write(f"  Output     : {output_file}\n")
        log_file.write(f"  Result     : {'SUCCESS' if success else 'FAILED'}\n")
        log_file.write(f"  Validation : {validation_status}\n")
        if error_message:
            log_file.write(f"  Error      : {error_message}\n")
        log_file.write("\n")


def convert_docx_file(
    input_file: Path,
    output_file: Path,
    bridge_file: Path = BRIDGE_FILE,
    progress_callback=None,
) -> dict:
    """Convert one DOCX using one persistent Node bridge process."""
    input_path = Path(input_file)
    output_path = Path(output_file)
    bridge_path = Path(bridge_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input document not found:\n{input_path}")
    if input_path.suffix.lower() != ".docx":
        raise ValueError(f"Only .docx documents are supported. Received: {input_path.name}")
    if not bridge_path.exists():
        raise FileNotFoundError(f"Node bridge not found:\n{bridge_path}")
    if input_path.resolve() == output_path.resolve():
        raise RuntimeError("SAFETY ERROR: Input and output files are identical.")

    def report(status, percent):
        if progress_callback:
            progress_callback(status, percent)

    changed_runs = 0
    converter = BijoyConverter(bridge_path)
    try:
        report("Loading document...", 5)
        document = Document(input_path)
        report("Scanning document...", 10)
        paragraphs = list(document.paragraphs)

        def collect_table(table):
            for row in table.rows:
                for cell in row.cells:
                    paragraphs.extend(cell.paragraphs)
                    for nested_table in cell.tables:
                        collect_table(nested_table)

        for table in document.tables:
            collect_table(table)

        report("Consolidating Bangla runs...", 20)
        for paragraph in paragraphs:
            consolidate_runs_in_paragraph(paragraph)

        eligible_runs = [
            run for paragraph in paragraphs for run in paragraph.runs
            if run.font.name == SOURCE_FONT and run.text
        ]
        total_runs = len(eligible_runs)
        report(f"Converting Bangla... 0/{total_runs}", 30)
        converter.start()
        for index, run in enumerate(eligible_runs, start=1):
            if process_run(run, converter):
                changed_runs += 1
            percent = 30 + int(50 * index / total_runs) if total_runs else 80
            report(f"Converting Bangla... {index}/{total_runs}", percent)

        report("Applying Nikosh...", 85)
        normalize_latin_runs_in_document(document)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        report("Saving document...", 90)
        document.save(output_path)
        report("Validating output...", 95)
        validation = validate_converted_docx(output_path)
        if not validation["valid"]:
            raise RuntimeError(f"Output validation failed: {validation['details']}")
        report("Complete", 100)
        log_conversion_result(str(input_path), str(output_path), True, validation["details"])
        return {
            "success": True,
            "converted_runs": changed_runs,
            "validation": validation,
            "output_file": str(output_path),
            "source_file": str(input_path),
        }
    except Exception as exc:
        log_conversion_result(str(input_path), str(output_path), False, "FAILED", str(exc))
        raise
    finally:
        converter.close()


if __name__ == "__main__":

    try:

        main()

    except Exception as exc:

        print()

        print(
            "[ERROR]"
        )

        print(
            str(exc)
        )

        sys.exit(1)