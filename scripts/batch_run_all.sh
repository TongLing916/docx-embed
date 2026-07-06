#!/usr/bin/env bash
# Batch-run every DOCX in data/ through parser/pipeline methods, including
# separate MinerU backend variants for pipeline, vlm-engine, and hybrid-engine.
# Resumable: skips a (file,method) whose DONE marker exists.
set -u

REPO="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO"

DATA_DIR="data"
OUT_ROOT="output_batch"
LOG_DIR="$OUT_ROOT/_logs"
mkdir -p "$LOG_DIR"

SUMMARY="$LOG_DIR/summary.tsv"
if [ ! -f "$SUMMARY" ]; then
  printf "file\tmethod\tstatus\tduration_s\tout_dir\n" > "$SUMMARY"
fi

run_one() {
  local method="$1" docx="$2"
  local stem; stem="$(basename "$docx" .docx)"
  local out="$OUT_ROOT/$stem/$method"
  local done="$out/__done__"
  if [ -f "$done" ]; then
    return 0
  fi
  rm -rf "$out"
  mkdir -p "$out"
  local logf="$LOG_DIR/${stem}__${method}.log"
  local start; start=$(date +%s)
  case "$method" in
    pipeline-markitdown)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser markitdown --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-mineru-pipeline)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser mineru-pipeline --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-mineru-vlm-engine)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser mineru-vlm-engine --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-mineru-hybrid-engine)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser mineru-hybrid-engine --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-docling)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser docling --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-pandoc)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser pandoc --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pipeline-ragflow)
      uv run examples/run_pipeline.py "$docx" "$out" --main-parser ragflow --unsafe-unwrap-embedded >"$logf" 2>&1 ;;
    pure-markitdown)
      uv run examples/run_parser.py markitdown "$docx" "$out" >"$logf" 2>&1 ;;
    pure-mineru-pipeline)
      uv run examples/run_parser.py mineru-pipeline "$docx" "$out" >"$logf" 2>&1 ;;
    pure-mineru-vlm-engine)
      uv run examples/run_parser.py mineru-vlm-engine "$docx" "$out" >"$logf" 2>&1 ;;
    pure-mineru-hybrid-engine)
      uv run examples/run_parser.py mineru-hybrid-engine "$docx" "$out" >"$logf" 2>&1 ;;
    pure-docling)
      uv run examples/run_parser.py docling "$docx" "$out" >"$logf" 2>&1 ;;
    pure-pandoc)
      uv run examples/run_parser.py pandoc "$docx" "$out" >"$logf" 2>&1 ;;
    pure-ragflow)
      uv run examples/run_parser.py ragflow "$docx" "$out" >"$logf" 2>&1 ;;
    *) echo "unknown method $method" >"$logf"; return 1 ;;
  esac
  local rc=$?
  local end; end=$(date +%s)
  local dur=$((end - start))
  local status
  if [ $rc -eq 0 ]; then status="ok"; elif [ $rc -eq 2 ]; then status="warn"; else status="fail"; fi
  printf '%s\t%s\t%s\t%s\t%s\n' "$stem" "$method" "$status" "$dur" "$out" >> "$SUMMARY"
  touch "$done"
  return 0
}

METHODS=(pipeline-markitdown pipeline-mineru-pipeline pipeline-mineru-vlm-engine pipeline-mineru-hybrid-engine pipeline-docling pipeline-pandoc pipeline-ragflow pure-markitdown pure-mineru-pipeline pure-mineru-vlm-engine pure-mineru-hybrid-engine pure-docling pure-pandoc pure-ragflow)

shopt -s nullglob
docs=("$DATA_DIR"/*.docx)
total=$((${#docs[@]} * ${#METHODS[@]}))
echo "Total runs: $total (${#docs[@]} files x ${#METHODS[@]} methods)"
i=0
for docx in "${docs[@]}"; do
  for method in "${METHODS[@]}"; do
    i=$((i + 1))
    printf '[%d/%d] %s :: %s\n' "$i" "$total" "$(basename "$docx")" "$method"
    run_one "$method" "$docx"
  done
done
echo "DONE"
