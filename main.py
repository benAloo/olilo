from __future__ import annotations

import argparse
import logging
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

try:
    from pdf2docx import Converter
except ImportError:
    sys.exit(
        "ERROR: 'pdf2docx' is not installed.\n"
        " Install it with: pip install pdf2docx\n"
    )

try:
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError
except ImportError:
    sys.exit(
        "ERROR: 'pypdf' is not installed.\n"
        " Install it with: pip install pypdf\n"
    )

def _configure_logging(verbose: bool = False) -> logging.Logger:
    """Return a module-level logger with console output"""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        format="%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        level=level,
        stream=sys.stdout,
    )

    if not verbose:
        for noisy in ("pdfminer", "pdf2docx", "fonttools"):
            logging.getLogger(noisy).setLevel(logging.ERROR)
    
    return logging.getLogger(__name__)

@dataclass
class ConversionSummary:
    converted: list[Path] = field(default_factory=list)
    skipped_existing: list[Path] = field(default_factory=list)
    skipped_encrypted: list[Path] = field(default_factory=list)
    skipped_corrupted: list[Path] = field(default_factory=list)
    failed: list[tuple[Path, str]] = field(default_factory=list)

    def total_found(self) -> int:
        return (
            len(self.converted)
            + len(self.skipped_existing)
            + len(self.skipped_encrypted)
            + len(self.skipped_corrupted)
            + len(self.failed)
        )
    
    def print_report(self, logger: logging.Logger) -> None:
        logger.info("=" * 60)
        logger.info("CONVERSION SUMMARY")
        logger.info("=" * 60)
        logger.info("  PDFs found     : %d", self.total_found())
        logger.info("  Converted      : %d", len(self.converted))
        logger.info("  Skipped (exist): %d", len(self.skipped_existing))
        logger.info("  Skipped (enc.) : %d", len(self.skipped_encrypted))
        logger.info("  Skipped (corr.): %d", len(self.skipped_corrupted))
        logger.info("  Failed         : %d", len(self.failed))

        if self.skipped_existing:
            logger.info("\n --Skipped (already exist) --")
            for p in self.skipped_existing:
                logger.info("   %s", p.name)
        
        if self.skipped_encrypted:
            logger.info("\n --Skipped (password-protected) --")
            for p in self.skipped_encrypted:
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

def _check_pdf(pdf_path: Path, logger: logging.Logger) -> str | None:
    """
    Perform lightweight checks on a PDF before attempting conversion.

    Returns:
        None if the file looks fine.
        'encrypted' if it is password-protected
        'corrupted' if it cannot be parsed at all
    """
    try:
        reader = PdfReader(str(pdf_path))

        if reader.is_encrypted:
            try:
                reader.decrypt("")
                if reader.is_encrypted:
                    return "encrypted"
            except Exception:
                return "encrypted"
        
        if len(reader.pages) == 0:
            logger.debug("%s has 0 pages - treating as corrupted.", pdf_path.name)
            return "corrupted"
        
        _ = reader.pages[0]
    except PdfReadError:
        return "corrupted"
    except Exception as exc:
        logger.debug("Unexpected pre-flight error on %s: %s", pdf_path.name, exc)
        return "corrupted"

    return None

def _convert_one(
        pdf_path: Path,
        docx_path: Path,
        logger: logging.Logger,
) -> bool:
    """
    Convert a single PDF -> DOCX using pdf2docx.

    pdf2docx streams pages individually, so large files don't blow out RAM.

    Returns True on succes, False on failure
    """
    docx_path.parent.mkdir(parents=True, exist_ok=True)

    cv = Converter(str(pdf_path))

    try:
        cv.convert(str(docx_path))
        return True
    except Exception as exc:
        if docx_path.exists():
            try:
                docx_path.unlink()
            except OSError:
                pass
        
        logger.error("Conversion error for %s: %s", pdf_path.name, exc)
        return False
    finally:
        cv.close()

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
    pattern = "**/*.pdf" if recursive else "*.pdf"
    pdf_files = sorted(input_dir.glob(pattern))

    if not pdf_files:
        logger.warning("No PDF files found in '%s'.", input_dir)
        return summary
    
    logger.info("Found %d PDF file(s) in '%s'.", len(pdf_files), input_dir)

    for pdf_path in pdf_files:
        if mirror_structure and recursive:
            rel = pdf_path.relative_to(input_dir)
            docx_path = output_dir / rel.with_suffix(".docx")
        else:
            docx_path = output_dir / (pdf_path.stem + ".docx")
        
        logger.info("Processing: %s", pdf_path.name)

        if docx_path.exists() and not overwrite:
            logger.info(
                " SKIP - '%s' already exists (use --overwrite to replace).",
                docx_path.name,
            )
            summary.skipped_existing.append(pdf_path)
            continue

        issue = _check_pdf(pdf_path, logger)
        if issue == "encrypted":
            logger.warning(
                " SKIP - '%s' is password-protected or encrypted.", pdf_path.name
            )
            summary.skipped_encrypted.append(pdf_path)
            continue

        if issue == "corrupted":
            logger.warning(
                " SKIP - '%s' appears corrupted or unreadable.", pdf_path.name
            )
            summary.skipped_corrupted.append(pdf_path)
            continue

        if dry_run:
            logger.info(
                " DRY-RUN - would convert to '%s'.", docx_path
            )
            summary.converted.append(pdf_path)
            continue

        start = time.perf_counter()
        success = _convert_one(pdf_path, docx_path, logger)
        elapsed = time.perf_counter() - start

        if success:
            logger.info(
                " OK - saved '%s' (%.1fs)", docx_path.name, elapsed
            )
            summary.converted.append(pdf_path)
        else:
            reason = "Conversion failed - check logs above for details."
            summary.failed.append((pdf_path, reason))

    return summary

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Batch-convert PDF files to DOCX format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    %(prog)s --input ./reports --output ./docx_output
    %(prog)s --input ./archive --output ./converted --recursive --overwrite
    %(prog)s --input ./inbox --dry-run
""",
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        type=Path,
        metavar="DIR",
        help="Directory containing PDF files to convert.",
    )

    parser.add_argument(
        "--output", "-o",
        type=Path,
        metavar="DIR",
        default=None,
        help=(
            "Directory containing PDF files to convert. "
            "Defaults to the same directory as --input."
        ),
    )

    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        default=False,
        help="Search for PDF files recursively in subdirectories.",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        default=False,
        help="Overwrite existing DOCX files. Skips them by default.",
    )

    parser.add_argument(
        "--mirror",
        action="store_true",
        default=False,
        help=(
            "Mirror the source sub-directory structure under --output "
            "(only relevant with --recursive)."
        ),
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

        
