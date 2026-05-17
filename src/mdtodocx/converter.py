"""
Pandoc-based Markdown to DOCX conversion pipeline.

Orchestrates: pandoc (with Lua filter) → DOCX → post-processing → final DOCX.
"""

import os
import pypandoc

from .config import DEFAULT_REF, LUA_FILTER
from .postprocess import postprocess_docx


def md_to_docx(
    input_md: str,
    output_docx: str,
    reference_doc: str = DEFAULT_REF,
) -> None:
    """
    Convert a Markdown file to DOCX using the academic style filter.

    Pipeline:
        1. pandoc converts MD → DOCX (using Lua filter for style mapping)
        2. postprocess_docx fixes tables, lists, numbering in the XML

    Args:
        input_md: Path to input .md file.
        output_docx: Path to output .docx file.
        reference_doc: Path to reference .docx template (default: docx/ref.docx).
    """
    for path, label in [
        (input_md, "Input file"),
        (reference_doc, "Reference docx"),
        (LUA_FILTER, "Lua filter"),
    ]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} not found: {path}")

    os.makedirs(os.path.dirname(os.path.abspath(output_docx)), exist_ok=True)

    extra_args = [
        f"--reference-doc={reference_doc}",
        f"--lua-filter={LUA_FILTER}",
        "--wrap=none",
    ]

    pypandoc.convert_file(
        input_md,
        "docx",
        outputfile=output_docx,
        extra_args=extra_args,
    )

    postprocess_docx(output_docx)
    print(f"OK  {input_md}  ->  {output_docx}")
