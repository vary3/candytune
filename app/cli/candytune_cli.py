import argparse
import os
from pathlib import Path
import sys

from app.candytune.core.converter import convert_to_pdf, ConversionError
from app.candytune.ui.banner import show_banner
from app.candytune.ui.progress import ConversionProgress


SUPPORTED_SUFFIXES = {
    ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".xlsm", ".csv",
    ".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp",
    ".pdf",
}


def iter_target_files(base_dir: Path):
    for path in base_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES:
            yield path


def _unique_output_path(base_dir: Path, desired_name: str) -> Path:
    candidate = base_dir / desired_name
    if not candidate.exists():
        return candidate
    stem, suffix = Path(desired_name).stem, Path(desired_name).suffix
    i = 1
    while True:
        cand = base_dir / f"{stem} ({i}){suffix}"
        if not cand.exists():
            return cand
        i += 1


def convert_tree(input_dir: Path, output_dir: Path, image_dpi: int, flatten: bool) -> int:
    ui = ConversionProgress()
    errors = []
    converted = 0
    
    # ファイルリストを収集
    files = list(iter_target_files(input_dir))
    total = len(files)
    
    if total == 0:
        ui.console.print("[yellow]⚠ No files to convert[/yellow]")
        return 0
    
    # Execute conversion with progress bar
    with ui.create_progress_bar(total) as progress:
        task = progress.add_task(
            "[cyan]Converting files...", 
            total=total
        )
        
        for src in files:
            rel = src.relative_to(input_dir)
            out_dir = (output_dir if flatten else (output_dir / rel.parent))
            out_dir.mkdir(parents=True, exist_ok=True)
            out_pdf = out_dir / (src.stem + ".pdf")
            if flatten:
                out_pdf = _unique_output_path(out_dir, out_pdf.name)
            
            try:
                produced = convert_to_pdf(src, workdir=out_dir, image_dpi=image_dpi)
                if produced != out_pdf:
                    produced.replace(out_pdf)
                ui.print_converting(src, out_pdf)
                converted += 1
            except ConversionError as e:
                errors.append((src, str(e)))
                ui.print_error(src, str(e))
            
            progress.advance(task)
    
    # サマリー表示
    ui.show_summary(total, converted, errors)
    
    return 1 if errors else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="candytune-cli",
        description="Convert all supported files under input dir to PDF under output dir",
    )
    default_input = os.environ.get("CANDYTUNE_INPUT", "input")
    default_output = os.environ.get("CANDYTUNE_OUTPUT", "output")
    parser.add_argument(
        "--input",
        default=default_input,
        help="Input directory (default: env CANDYTUNE_INPUT or 'input')",
    )
    parser.add_argument(
        "--output",
        default=default_output,
        help="Output directory (default: env CANDYTUNE_OUTPUT or 'output')",
    )
    parser.add_argument(
        "--flatten",
        action="store_true",
        help="Do not preserve directory structure; save all PDFs directly under output",
    )
    parser.add_argument(
        "--image-dpi",
        type=int,
        default=200,
        help="DPI for image to PDF conversion (default: 200)",
    )
    return parser


def main(argv=None) -> int:
    # バナー表示
    show_banner()
    
    parser = build_parser()
    args = parser.parse_args(argv)
    input_dir = Path(args.input)
    output_dir = Path(args.output)

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory not found: {input_dir}", file=sys.stderr)
        return 2
    
    # 設定表示
    ui = ConversionProgress()
    ui.print_config(input_dir, output_dir, args.image_dpi, args.flatten)

    return convert_tree(input_dir, output_dir, image_dpi=args.image_dpi, flatten=args.flatten)


if __name__ == "__main__":
    raise SystemExit(main())


