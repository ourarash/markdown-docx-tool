#!/usr/bin/env bash
set -euo pipefail

# pandoc_md_to_docx.sh
#
# Generic markdown-to-DOCX converter for standalone markdown document repos
# and installable Codex skills.
#
# Behavior:
# - Accepts absolute paths, repo-relative paths, or current-directory paths.
# - Defaults output to the input file's directory with a `.docx` extension.
# - Uses the input file's directory as the Pandoc resource path so relative
#   images and other linked assets resolve the same way they do in markdown.
# - Reuses the bundled `reference.docx` for Word-native styling. That file is
#   the preferred place to manage code, table, and other visual formatting.
# - Applies one post-processing fix after Pandoc runs: switch Word tables from
#   fixed layout to auto-fit so wide markdown tables render more naturally.
#
# Example usage:
#   ./scripts/pandoc_md_to_docx.sh samples/showcase.md
#   ./scripts/pandoc_md_to_docx.sh README.md /tmp/repo_readme.docx
#   ./scripts/pandoc_md_to_docx.sh /absolute/path/to/file.md

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DEFAULT_REFERENCE_DOC="$SCRIPT_DIR/reference.docx"
REFERENCE_DOC="$DEFAULT_REFERENCE_DOC"
METADATA_FILE=""
ENABLE_TOC=0
OUTPUT_DIR=""

usage() {
  cat <<EOF
Usage: $(basename "$0") [options] <input.md> [output.docx]

Examples:
  $(basename "$0") samples/showcase.md
  $(basename "$0") README.md /tmp/repo_readme.docx
  $(basename "$0") --toc --output-dir output samples/showcase.md
  $(basename "$0") /absolute/path/to/file.md

Notes:
  - If output is omitted, the script writes <input-basename>.docx next to the
    markdown file.
  - Use --output-dir to keep the default filename but write it elsewhere.
  - Use --reference-doc to override the bundled reference template.
  - Use --metadata-file to pass a Pandoc metadata file.
  - Use --toc to include a table of contents.
  - Styling comes from the bundled reference.docx.
  - The script keeps one XML post-processing step to force Word table auto-fit.
EOF
}

require_command() {
  local cmd="$1"

  if ! command -v "$cmd" >/dev/null 2>&1; then
    echo "Error: required command not found: $cmd" >&2
    exit 1
  fi
}

abs_path() {
  local path="$1"
  local dir
  local base

  dir="$(cd "$(dirname "$path")" && pwd)"
  base="$(basename "$path")"
  printf '%s/%s\n' "$dir" "$base"
}

resolve_existing_path() {
  local candidate="$1"

  if [[ "$candidate" = /* ]]; then
    [[ -e "$candidate" ]] || return 1
    printf '%s\n' "$candidate"
    return 0
  fi

  if [[ -e "$candidate" ]]; then
    abs_path "$candidate"
    return 0
  fi

  if [[ -e "$REPO_ROOT/$candidate" ]]; then
    abs_path "$REPO_ROOT/$candidate"
    return 0
  fi

  return 1
}

resolve_output_path() {
  local candidate="$1"
  local dir
  local base

  if [[ "$candidate" = /* ]]; then
    dir="$(dirname "$candidate")"
    base="$(basename "$candidate")"
    printf '%s/%s\n' "$dir" "$base"
    return 0
  fi

  dir="$PWD/$(dirname "$candidate")"
  base="$(basename "$candidate")"
  printf '%s/%s\n' "$dir" "$base"
}

autofit_docx_tables() {
  local docx_path="$1"
  local tmp_dir

  tmp_dir="$(mktemp -d)"
  unzip -qq "$docx_path" -d "$tmp_dir"

  # Pandoc's DOCX writer emits fixed-layout tables by default. For narrow
  # tables that is fine, but for wide task/status tables it tends to produce
  # cramped columns in Word. Remove the fixed-layout marker and restore
  # Word's auto-fit width behavior.
  perl -0pi -e 's#<w:tblLayout w:type="fixed"\s*/>##g; s#<w:tblW w:type="pct" w:w="5000"\s*/>#<w:tblW w:type="auto" w:w="0"/>#g' \
    "$tmp_dir/word/document.xml"

  (
    cd "$tmp_dir"
    zip -X -qr "$docx_path" .
  )

  rm -rf "$tmp_dir"
}

if [[ "${1:-}" = "-h" || "${1:-}" = "--help" ]]; then
  usage
  exit 0
fi

POSITIONAL_ARGS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    -h|--help)
      usage
      exit 0
      ;;
    --reference-doc)
      [[ $# -ge 2 ]] || { echo "Error: --reference-doc requires a path" >&2; exit 1; }
      REFERENCE_DOC="$(resolve_existing_path "$2")" || {
        echo "Error: reference document not found: $2" >&2
        exit 1
      }
      shift 2
      ;;
    --metadata-file)
      [[ $# -ge 2 ]] || { echo "Error: --metadata-file requires a path" >&2; exit 1; }
      METADATA_FILE="$(resolve_existing_path "$2")" || {
        echo "Error: metadata file not found: $2" >&2
        exit 1
      }
      shift 2
      ;;
    --output-dir)
      [[ $# -ge 2 ]] || { echo "Error: --output-dir requires a path" >&2; exit 1; }
      OUTPUT_DIR="$(resolve_output_path "$2")"
      shift 2
      ;;
    --toc)
      ENABLE_TOC=1
      shift
      ;;
    --)
      shift
      while [[ $# -gt 0 ]]; do
        POSITIONAL_ARGS+=("$1")
        shift
      done
      ;;
    -*)
      echo "Error: unknown option: $1" >&2
      usage >&2
      exit 1
      ;;
    *)
      POSITIONAL_ARGS+=("$1")
      shift
      ;;
  esac
done

set -- "${POSITIONAL_ARGS[@]}"

if [[ $# -lt 1 || $# -gt 2 ]]; then
  usage >&2
  exit 1
fi

require_command pandoc
require_command perl
require_command unzip
require_command zip

INPUT_MD="$(resolve_existing_path "$1")" || {
  echo "Error: markdown file not found: $1" >&2
  exit 1
}

if [[ ! -f "$INPUT_MD" ]]; then
  echo "Error: input is not a file: $INPUT_MD" >&2
  exit 1
fi

if [[ $# -eq 2 && -n "$OUTPUT_DIR" ]]; then
  echo "Error: use either an explicit output path or --output-dir, not both" >&2
  exit 1
fi

if [[ $# -eq 2 ]]; then
  OUTPUT_DOCX="$(resolve_output_path "$2")"
elif [[ -n "$OUTPUT_DIR" ]]; then
  OUTPUT_DOCX="$OUTPUT_DIR/$(basename "${INPUT_MD%.*}").docx"
else
  OUTPUT_DOCX="${INPUT_MD%.*}.docx"
fi

INPUT_DIR="$(dirname "$INPUT_MD")"
OUTPUT_DIR="$(dirname "$OUTPUT_DOCX")"
OUTPUT_BASE="$(basename "${OUTPUT_DOCX%.*}")"
MEDIA_DIR="$OUTPUT_DIR/${OUTPUT_BASE}_media"

mkdir -p "$OUTPUT_DIR"

cd "$REPO_ROOT"

PANDOC_ARGS=(
  "$INPUT_MD"
  --from=markdown+smart
  --to=docx
  --wrap=none
  "--resource-path=$INPUT_DIR:$REPO_ROOT"
  "--extract-media=$MEDIA_DIR"
  --dpi=300
  "--reference-doc=$REFERENCE_DOC"
)

if [[ $ENABLE_TOC -eq 1 ]]; then
  PANDOC_ARGS+=(--toc)
fi

if [[ -n "$METADATA_FILE" ]]; then
  PANDOC_ARGS+=("--metadata-file=$METADATA_FILE")
fi

PANDOC_ARGS+=(-o "$OUTPUT_DOCX")

pandoc "${PANDOC_ARGS[@]}"

autofit_docx_tables "$OUTPUT_DOCX"

echo "Wrote: $OUTPUT_DOCX"
