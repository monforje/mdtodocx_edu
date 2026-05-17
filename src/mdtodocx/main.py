"""
mdToDocx — Markdown to academic DOCX converter.

Usage:
    python -m mdtodocx input.md output.docx
    python -m mdtodocx input.md output.docx --ref docx/ref.docx
"""

import argparse

from .config import DEFAULT_REF
from .converter import md_to_docx


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Markdown to academic DOCX")
    parser.add_argument("input", help="Input .md file")
    parser.add_argument("output", help="Output .docx file")
    parser.add_argument(
        "--ref",
        default=DEFAULT_REF,
        help=f"Reference .docx template (default: {DEFAULT_REF})",
    )
    args = parser.parse_args()
    md_to_docx(args.input, args.output, args.ref)


if __name__ == "__main__":
    main()
