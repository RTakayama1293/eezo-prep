# EEZO「整える」リポジトリ（CCW）

仕入先の見積・送料見積を、EEZOのShopify登録用テンプレートへ整形するための Claude Code on the Web リポジトリ。採算（関西着基準）と要確認一覧（赤旗・空欄）まで一括生成する。

## 何ができるか

- `data/raw/{仕入先}/` に見積を置く → Shopify登録用の出力テンプレート＋採算＋要確認を `outputs/` に生成。
- **推定で埋めない**。見積に無い事実は空欄＋要確認（代理推定は禁止）。
- 消費者負担送料は公開ポリシー（`config/eezo_policy.json` の出典URL）を正とする。

## ディレクトリ構成

```
.
├── CLAUDE.md                 # Claude Code用指示書
├── README.md
├── requirements.txt
├── .gitignore
├── .claude/settings.json     # SessionStartフック（openpyxl自動導入）
├── scripts/setup.sh
├── rules/                    # 絶対ルール（推定禁止・データ不変・規約）
├── skills/                   # ワークフロー・抽出ルール・送料モデル・ドメイン知識
├── config/eezo_policy.json   # 公開送料ポリシーのスナップショット
├── templates/                # 抽出スキーマ（中間CSVの型）
├── data/
│   ├── raw/                  # ← 仕入先ごとの見積（編集禁止）＋出力フォーマット型
│   └── processed/            # ← Claudeが転記した中間CSV
├── src/build_eezo_template.py
├── tests/test_build.py
├── experiments/              # 実験・worked example（exp001_price_audit_repro）
└── outputs/                  # 最終成果物（再生成可能・gitignore）
```

## セットアップ

### Claude Code on the Web
1. このリポジトリを GitHub にプッシュ（プライベート推奨：見積は社外秘）。
2. https://claude.ai/code でリポジトリを選択。
3. SessionStartフックが openpyxl を自動導入。

### ローカル
```bash
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt
```

## 使い方

### 1. 見積を置く
```
data/raw/梅屋/見積_0416.pdf
data/raw/梅屋/送料見積.xlsx        # あれば
data/raw/トワ・ヴェール/見積.pdf    # 送料込なら送料見積なし
```

### 2. 整える（CCWでの指示例）
```
全仕入先を整えて outputs に出力して
```
または個別に:
```
data/raw/梅屋 の見積を data/processed に転記して
```

### 3. ビルド（手動実行する場合）
```bash
python src/build_eezo_template.py \
  --processed data/processed \
  --config config/eezo_policy.json \
  --out outputs/20260605_DATA_EEZO_seibi_output.xlsx
```

### 4. テスト
```bash
python tests/test_build.py        # 簡易
python -m pytest tests/ -q        # pytest導入時
```

## worked example

`experiments/exp001_price_audit_repro/` に、添付（価格精査一覧）由来のサンプル中間CSVと再現手順（log.md）を同梱。`run.sh` で即実行できる。

## メンテナンス

- 公開送料ポリシーが変わったら `config/eezo_policy.json`（確認日・無料ライン・地域別レート）を更新。
- 出力列の増減は `src/build_eezo_template.py` の `COLUMNS` を編集。
- 命名規則: `YYYYMMDD_TYPE_DESCRIPTION_vN`（出力ファイル）。
