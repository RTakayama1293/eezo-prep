#!/bin/bash
# 添付（価格精査一覧）由来のサンプル中間CSVで整えるパイプラインを実行
cd "$(dirname "$0")/../.."
python src/build_eezo_template.py \
  --processed experiments/exp001_price_audit_repro/processed \
  --config config/eezo_policy.json \
  --out experiments/exp001_price_audit_repro/outputs/output_sample.xlsx
