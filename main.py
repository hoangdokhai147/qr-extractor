import argparse
import sys
import logging
from datetime import datetime
from pathlib import Path

from src.utils.logger import setup_logging
from src.io.file_utils import get_targets, get_image_files
from src.io.excel_writer import write_excel
from src.qrcode.scanner import QRScanner

logger = logging.getLogger(__name__)

def resolve_output_path(output_arg: str | None, root_path: Path) -> Path:
    """Determine the final output Excel file path."""
    if output_arg is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(f"qr_results_{timestamp}.xlsx")
    else:
        output_path = Path(output_arg)

    if output_path.exists():
        answer = input(
            f"\nOutput file '{output_path}' already exists.\n"
            "  [O] Overwrite  [T] Create timestamped copy  [Q] Quit\n"
            "Your choice: "
        ).strip().upper()

        if answer == "O":
            pass  # overwrite
        elif answer == "T":
            stem = output_path.stem
            suffix = output_path.suffix
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = output_path.with_name(f"{stem}_{timestamp}{suffix}")
            print(f"Will save to: {output_path}")
        else:
            print("Aborted by user.")
            sys.exit(0)

    return output_path

def process_folder(folder_path: Path, scanner: QRScanner) -> dict:
    """Process all images in *folder_path* and collect QR decoding results."""
    folder_name = folder_path.name
    reading_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    image_files = get_image_files(folder_path)
    total_files = len(image_files)

    if total_files == 0:
        logger.warning("No image files found in folder: %s – skipping.", folder_path)

    details = []
    success_count = 0
    fail_count = 0

    for img_path in image_files:
        logger.info("  Decoding: %s", img_path.name)
        qr_content, status = scanner.scan_image(img_path)

        if status == "success":
            success_count += 1
            logger.info("    -> SUCCESS")
        else:
            fail_count += 1
            logger.info("    -> FAIL")

        details.append(
            {
                "folder": folder_name,
                "file_name": img_path.name,
                "qr_content": qr_content,
                "status": status,
            }
        )

    return {
        "folder_name": folder_name,
        "reading_date": reading_date,
        "total_files": total_files,
        "success_count": success_count,
        "fail_count": fail_count,
        "details": details,
    }

def main():
    parser = argparse.ArgumentParser(
        description="Extract QR code information using advanced AI detector.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--root",
        required=True,
        metavar="DIR",
        help="Root directory with subfolders nested or flat images.",
    )
    parser.add_argument(
        "--output",
        metavar="FILE",
        default=None,
        help="Output Excel file path.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print detailed progress.",
    )

    args = parser.parse_args()

    setup_logging(args.verbose)

    root_path = Path(args.root).resolve()
    if not root_path.exists() or not root_path.is_dir():
        print(f"ERROR: Invalid root directory: {root_path}", file=sys.stderr)
        sys.exit(1)

    targets, layout = get_targets(root_path)
    if layout == "empty":
        print(f"ERROR: No subfolders or image files found in '{root_path}'.", file=sys.stderr)
        sys.exit(1)
    
    print(f"\nLayout detected: {layout} ({len(targets)} group(s))")

    scanner = QRScanner()
    folders_data = []

    for target in targets:
        print(f"\nProcessing folder: {target.name}")
        result = process_folder(target, scanner)
        folders_data.append(result)
        print(
            f"  Files: {result['total_files']}, "
            f"Success: {result['success_count']}, "
            f"Fail: {result['fail_count']}"
        )

    output_path = resolve_output_path(args.output, root_path)
    
    try:
        write_excel(output_path, folders_data)
        print(f"\nResults written to: {output_path.resolve()}")
    except Exception as exc:
        print(f"ERROR: Failed to write Excel file: {exc}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
