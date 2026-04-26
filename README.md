# extract-metadata-from-images

A command-line tool that scans a folder of images, extracts EXIF metadata, deduplicates by file hash, and produces:

- a **CSV** with one row per image (filename, dimensions, datetime, GPS coordinates, camera make/model, MD5 hash)
- a **Word document** with all images embedded in chronological order, each captioned with its date and time

## Supported formats

| Format | Datetime | GPS |
|--------|----------|-----|
| JPEG / JPG | yes | yes |
| TIFF / TIF | yes | yes |
| HEIC | yes | yes |
| WebP | yes | yes |
| PNG | — | — |
| BMP / GIF | — | — |

## Requirements

- Python 3.13+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

## Installation

```bash
git clone https://github.com/MatteoGaetzner/image-exif-export.git
cd image-exif-export
uv sync
```

Or with pip:

```bash
pip install pillow pillow-heif pandas python-docx
```

## Usage

```bash
python main.py <path/to/images> [output_base_name]
```

| Argument | Required | Description |
|----------|----------|-------------|
| `path/to/images` | yes | Directory to scan (searched recursively) |
| `output_base_name` | no | Base name for output files (default: `image_metadata`) |

**Examples:**

```bash
# Outputs image_metadata.csv and image_metadata.docx
python main.py ~/Photos

# Outputs trip_2024.csv and trip_2024.docx
python main.py ~/Photos/trip trip_2024
```

## Output

### CSV columns

| Column | Description |
|--------|-------------|
| `filename` | File name |
| `filepath` | Absolute path |
| `extension` | File extension |
| `size_bytes` | File size in bytes |
| `file_hash` | MD5 hash (used for deduplication) |
| `width` / `height` | Dimensions in pixels |
| `datetime` | EXIF `DateTimeOriginal` or `DateTime` |
| `latitude` / `longitude` | GPS decimal degrees (negative = S/W) |
| `make` / `model` | Camera manufacturer and model |
| `error` | Error message if extraction failed |

### Word document

Images are sorted chronologically. Each image is captioned `YYYY-MM-DD  HH:MM:SS`; if the datetime is missing the filename is shown instead.
