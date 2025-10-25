# Candytune CLI — Turn Anything into PDFs. Fast. In Docker.

[![Python](https://img.shields.io/badge/python-3.11-blue.svg)](https://www.python.org/)
[![Docker](https://img.shields.io/badge/docker-ready-2496ED.svg)](https://www.docker.com/)
[![License](https://img.shields.io/badge/license-Apache--2.0-green.svg)](./LICENSE)

Candytune is a zero‑setup CLI that batch‑converts Office documents, images, and PDFs into clean PDFs — all inside Docker. It preserves your folder hierarchy by default and lets you flatten outputs when needed.

## Highlights
- Inputs: Office (.doc/.docx/.ppt/.pptx/.xls/.xlsx/.xlsm/.csv), Images (.jpg/.jpeg/.png/.webp/.tif/.tiff/.bmp), PDF
- Output: PDF (preserve hierarchy by default; `--flatten` for a single folder)
- Engines: LibreOffice (Office), ImageMagick (images)
- Zero host deps: runs entirely in Docker

## Quickstart (Docker)
1) Build the image
```bash
docker compose build
```
2) Prepare I/O directories
```
./workspace/input  # put your files here
./workspace/output # PDFs will be written here
```
3) Run the CLI (defaults provided via docker-compose)
```bash
docker compose run --rm candytune-cli
```
4) Flatten outputs (no directory structure)
```bash
docker compose run --rm candytune-cli --flatten
```

## CLI Usage
Candytune scans the input directory recursively and emits PDFs into the output directory.

```bash
python -m app.cli.candytune_cli \
  [--input <dir>] \
  [--output <dir>] \
  [--flatten] \
  [--image-dpi <int>]  # default: 200
```

- `--input`: input directory (default: env `CANDYTUNE_INPUT` or `input`)
- `--output`: output directory (default: env `CANDYTUNE_OUTPUT` or `output`)
- `--flatten`: do not preserve directory structure; put all PDFs directly under output (auto‑dedup names)
- `--image-dpi`: DPI for image→PDF conversion (default 200)

Defaults in `docker-compose.yml`:
```
CANDYTUNE_INPUT=/app/workspace/input
CANDYTUNE_OUTPUT=/app/workspace/output
```

## Resource Tuning (Memory/CPU)
Large Excel or Office conversions may require more resources.

### Compose settings
Edit `docker-compose.yml` under the `candytune-cli` service:

```yaml
mem_limit: "8g"   # container memory limit
cpus: "4"         # CPU quota (4 cores)
shm_size: "2g"    # enlarge shared memory (/dev/shm)
tmpfs:
  - /tmp:size=2g,exec  # keep temp files in RAM
```

After editing, restart the service:

```bash
docker compose down && docker compose up -d --build
```

### Docker Desktop (macOS) resources
Ensure Docker Desktop → Settings → Resources provide enough host resources:
- Memory: 12–16 GB or more
- Swap: 2–4 GB or more

Notes:
- `deploy.resources.*` is for Swarm and is ignored by normal `docker compose`.
- If conversions still stall, consider increasing `/dev/shm` and temp space or converting very large Excel files to CSV first.
 - `/dev/shm` is an in‑memory filesystem used by many libraries; increasing `shm_size` prevents "No space left on device" or hangs during large conversions.

## Verify Conversion
Quick sanity run:
```bash
docker compose run --rm candytune-cli -- --image-dpi 200
```
Check:
- exit code 0
- PDFs open correctly under `workspace/output`

Recommended golden set checks:
- 10–20 representative files
- page counts, garbling, missing figures
- Excel: confirm 1‑sheet‑per‑page as intended

## Architecture at a Glance
- Host → Docker (`candytune-cli`) → PDF output
- Office: LibreOffice (UNO where needed; 1 sheet = 1 page for Excel)
- Images: ImageMagick (1 image = 1 page; DPI adjustable)

## Design Notes
- Scope: inputs under `workspace/input`, outputs under `workspace/output` (all conversion inside the container)
- Directory structure: preserved by default; `--flatten` aggregates to the root (auto numbering on conflicts)
- Error handling: continues on failures and summarizes at the end; returns exit code 1 if any failed

## Roadmap (short)
- Distribution packaging for CLI
- Improved Excel pagination & conversion quality
- Basic logging/metrics (counts, failure rate)
- Optional future: cloud integrations

## License
Apache License 2.0 — see `LICENSE`.

Third‑party licenses bundled in the Docker image: see `THIRD_PARTY_NOTICES.md`.
