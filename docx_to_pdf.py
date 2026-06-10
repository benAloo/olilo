from __future__ import annotations

import argparse
import logging
import shutil
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

try:
    from docx2pdf import convert
except ImportError:
    sys.exit(
        "ERROR: 'docx2pdf' is not installed.\n"
        " Install it with: pip install docx2pdf\n"
    )


def _configure_logging(verbose: bool = False) -> logging.Logger:
    """Return a module-level logger with console output."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        format="%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        level=level,
        stream=sys.stdout,
    )

    if not verbose:
        for noisy in ("docx2pdf", "fonttools", "PIL", "pythoncom"):
            logging.getLogger(noisy).setLevel(logging.ERROR)

    return logging.getLogger(__name__)


@dataclass
class ConversionSummary:
    converted: list[Path] = field(default_factory=list)
    skipped_existing: list[Path] = field(default_factory=list)
    skipped_corrupted: list[Path] = field(default_factory=list)
    failed: list[tuple[Path, str]] = field(default_factory=list)

    def total_found(self) -> int:
        return (
            len(self.converted)
            + len(self.skipped_existing)
            + len(self.skipped_corrupted)
            + len(self.failed)
        )

    def print_report(self, logger: logging.Logger) -> None:
        logger.info("=" * 60)
        logger.info("CONVERSION SUMMARY")
        logger.info("=" * 60)
        logger.info("  DOCX found      : %d", self.total_found())
        logger.info("  Converted       : %d", len(self.converted))
        logger.info("  Skipped (exist) : %d", len(self.skipped_existing))
        logger.info("  Skipped (corr.) : %d", len(self.skipped_corrupted))
        logger.info("  Failed          : %d", len(self.failed))

        if self.skipped_existing:
            logger.info("\n --Skipped (already exist) --")
            for p in self.skipped_existing:
                logger.info("   %s", p.name)

        if self.skipped_corrupted:
            logger.info("\n --Skipped (corrupted / unreadable) --")
            for p in self.skipped_corrupted:
                logger.info("   %s", p.name)

        if self.failed:
            logger.info("\n --Failed conversions --")
            for p, reason in self.failed:
                logger.info("   %s -> %s", p.name, reason)

        logger.info("=" * 60)


def _check_docx(docx_path: Path, logger: logging.Logger) -> str | None:
    """
    Perform lightweight checks on a DOCX file before conversion.

    Returns:
        None if the file looks fine.
        'corrupted' if the file cannot be opened as a valid DOCX.
    """
    if docx_path.suffix.lower() != ".docx":
        logger.debug("%s does not have a .docx suffix.", docx_path.name)
        return "corrupted"

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_docx = Path(temp_dir) / "health_check.docx"
            shutil.copy2(docx_path, temp_docx)
            if not temp_docx.exists() or temp_docx.stat().st_size == 0:
                return "corrupted"
    except Exception as exc:
        logger.debug("Unexpected pre-flight error on %s: %s", docx_path.name, exc)
        return "corrupted"

    return None


def _convert_one(
    docx_path: Path,
    pdf_path: Path,
    logger: logging.Logger,
) -> bool:
    """
    Convert a single DOCX -> PDF using docx2pdf.

    Writes to a temporary file first, then atomically replaces the target.
    Returns True on success, False on failure.
    """
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        suffix=".pdf",
        dir=str(pdf_path.parent),
        delete=False,
    ) as temp_file:
        temp_pdf_path = Path(temp_file.name)

    try:
        convert(str(docx_path), str(temp_pdf_path))

        if not temp_pdf_path.exists() or temp_pdf_path.stat().st_size == 0:
            raise RuntimeError("Conversion produced no output file.")

        temp_pdf_path.replace(pdf_path)
        return True
    except Exception as exc:
        if temp_pdf_path.exists():
            try:
                temp_pdf_path.unlink()
            except OSError:
                pass

        if pdf_path.exists():
            try:
                pdf_path.unlink()
            except OSError:
                pass

        logger.error("Conversion error for %s: %s", docx_path.name, exc)
        return False


def batch_convert(
    input_dir: Path,
    output_dir: Path,
    *,
    recursive: bool = False,
    overwrite: bool = False,
    dry_run: bool = False,
    mirror_structure: bool = False,
    logger: logging.Logger,
) -> ConversionSummary:
    summary = ConversionSummary()
    pattern = "**/*.docx" if recursive else "*.docx"
    docx_files = sorted(input_dir.glob(pattern))

    if not docx_files:
        logger.warning("No DOCX files found in '%s'.", input_dir)
        return summary

    logger.info("Found %d DOCX file(s) in '%s'.", len(docx_files), input_dir)

    for docx_path in docx_files:
        if mirror_structure and recursive:
            rel = docx_path.relative_to(input_dir)
            pdf_path = output_dir / rel.with_suffix(".pdf")
        else:
            pdf_path = output_dir / (docx_path.stem + ".pdf")

        logger.info("Processing: %s", docx_path.name)

        if pdf_path.exists() and not overwrite:
            logger.info(
                " SKIP - '%s' already exists (use --overwrite to replace).",
                pdf_path.name,
            )
            summary.skipped_existing.append(docx_path)
            continue

        issue = _check_docx(docx_path, logger)
        if issue == "corrupted":
            logger.warning(
                " SKIP - '%s' appears corrupted or unreadable.", docx_path.name
            )
            summary.skipped_corrupted.append(docx_path)
            continue

        if dry_run:
            logger.info(" DRY-RUN - would convert to '%s'.", pdf_path)
            summary.converted.append(docx_path)
            continue

        start = time.perf_counter()
        success = _convert_one(docx_path, pdf_path, logger)
        elapsed = time.perf_counter() - start

        if success:
            logger.info(" OK - saved '%s' (%.1fs)", pdf_path.name, elapsed)
            summary.converted.append(docx_path)
        else:
            reason = "Conversion failed - check logs above for details."
            summary.failed.append((docx_path, reason))

    return summary


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Batch-convert DOCX files to PDF format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    %(prog)s --input ./docx --output ./pdf_output
    %(prog)s --input ./archive --output ./converted --recursive --overwrite
    %(prog)s --input ./inbox --dry-run
""",
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        type=Path,
        metavar="DIR",
        help="Directory containing DOCX files to convert.",
    )

    parser.add_argument(
        "--output", "-o",
        type=Path,
        metavar="DIR",
        default=None,
        help=(
            "Directory to write PDF files. "
            "Defaults to the same directory as --input."
        ),
    )

    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        default=False,
        help="Search for DOCX files recursively in subdirectories.",
    )

    parser.add_argument(
        "--mirror", "-m",
        action="store_true",
        default=False,
        help=(
            "Mirror the source sub-directory structure under --output "
            "(only relevant with --recursive)."
        ),
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        default=False,
        help="Overwrite existing PDF files. Skips them by default.",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Show what would be converted without writing any files.",
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        default=False,
        help="Enable debug-level logging.",
    )

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    logger = _configure_logging(verbose=args.verbose)

    input_dir: Path = args.input.resolve()
    if not input_dir.is_dir():
        logger.error("--input '%s' is not a directory or does not exist.", input_dir)
        return 1

    output_dir: Path = (args.output or input_dir).resolve()
    if not args.dry_run:
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            logger.error("Cannot create output directory '%s': %s", output_dir, exc)
            return 1

    if args.dry_run:
        logger.info("DRY-RUN mode - no files will be written.")

    logger.info("Input  : %s", input_dir)
    logger.info("Output : %s", output_dir)
    logger.info("Options: recursive=%s overwrite=%s mirror=%s", args.recursive, args.overwrite, args.mirror)

    t0 = time.perf_counter()
    summary = batch_convert(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=args.recursive,
        overwrite=args.overwrite,
        dry_run=args.dry_run,
        mirror_structure=args.mirror,
        logger=logger,
    )
    total_time = time.perf_counter() - t0

    summary.print_report(logger)
    logger.info("Total elapsed: %.1fs", total_time)

    return 1 if summary.failed else 0


if __name__ == "__main__":
    sys.exit(main())
