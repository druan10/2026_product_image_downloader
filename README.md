# Image Downloader to Dropbox Share Link

Downloads images from URLs, converts them to JPEG, uploads to Dropbox, and outputs an Excel file with direct share links.
Automatically compresses images for best performance.

## Setup

Create a `.env` file in the same folder as `app.exe`:

```
DROPBOX_ACCESS_TOKEN=your_token_here
DROPBOX_IMAGE_FOLDER=/your/dropbox/upload/folder/
TEMP_RETENTION_DAYS=7
```

| Variable | Required | Description |
|----------|----------|-------------|
| `DROPBOX_ACCESS_TOKEN` | Yes | Dropbox API token |
| `DROPBOX_IMAGE_FOLDER` | Yes | Dropbox folder path to upload images into |
| `TEMP_RETENTION_DAYS` | No | Days to keep local temp files (default: `7`). Set to `0` to clear on every run. Dropbox files are never deleted. |

To get a token: go to the [Dropbox Developer Console](https://www.dropbox.com/developers/apps), create an app with **Full Dropbox** access, then generate an access token under the **Settings** tab.

## Input file format

Provide a `.csv`, `.xls`, or `.xlsx` file with these columns:

| id | image_1 | image_2 | image_3 |
|----|---------|---------|---------|
| SKU001 | https://... | https://... | |
| SKU002 | https://... | | |

- `id` column is required.
- Add as many `image_N` columns as needed.
- Empty cells are skipped.

## Usage

**GUI (default)**

Double-click `app.exe`. Click **Browse** to select your input file, then **Download, Convert & Upload**.

If an image fails, a dialog will ask whether to continue with the remaining images or stop.

**CLI**

```
app.exe --input data.xlsx [--output result.xlsx]
```

| Flag | Short | Description |
|------|-------|-------------|
| `--input` | `-i` | Input CSV or Excel file (required) |
| `--output` | `-o` | Output XLSX path (default: `<name>_formatted_<timestamp>.xlsx` next to input) |

Failed images are skipped automatically and logged; processing always continues.

## Output

An Excel file with two columns:

| id | images |
|----|--------|
| SKU001 | ["https://dl.dropboxusercontent.com/...","https://..."] |

Images are uploaded to Dropbox under the folder set in your env file. 

## Pyinstaller Build Instructions

pyinstaller --onefile --add-data ".env;." app.py