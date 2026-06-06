# CLAUDE.md

## プロジェクト概要
- **目的**: 仕入先の見積・送料見積を、EEZOのShopify登録に載せられる出力テンプレートへ「整える」。あわせて関西着基準の採算と要確認一覧を生成する。`20260424_DATA_EEZO価格精査一覧.xlsx` を、手作業ではなく再現可能なパイプラインで作り直す。
- **データ**: `data/raw/{仕入先}/` 配下の見積（PDF/Excel/テキスト）と送料見積。出力は `outputs/`。
- **評価指標**: (1) 推定混入ゼロ（空欄＝事実として未取得）、(2) Shopify取込CSVがそのまま shopify-product-register に渡せること、(3) 採算の赤旗を漏れなく検出すること。

---

## Core Philosophy（基本原則）

### 1. 推定で埋めない（最重要）
見積に書いてある事実だけを転記する。無ければ空欄＋要確認。代理推定（他社送料での仮置き、サイズ不明での送料推定、税込/税抜不明での片側仮定、価格未記載の補完）は**禁止**。詳細は `rules/no-estimation.md`。

### 2. Plan Before Execute（計画先行）
複数仕入先をまとめて処理する前に、対象ディレクトリと送料区分（送料込/産直別途/不明）を一覧化してから着手する。

### 3. Immutability（不変性）
`data/raw/` の見積は**絶対に編集しない**。加工結果は `data/processed/` と `outputs/` にのみ書く。

### 4. 決定論はスクリプトへ
採算計算・送料引き当て・要確認生成・整形は `src/build_eezo_template.py` が決定論的に行う。Claudeの担当は**見積からの転記まで**（属人化と推定の混入を防ぐ）。

### 5. 公開ポリシーが正
消費者負担送料は公開ページ（`config/eezo_policy.json` の出典URL）を正とする。実行前に必ず突き合わせ、相違あれば config を更新（数値の据え置き禁止）。

---

## Critical Rules（絶対ルール）

コミット・出力前の必須チェック:
- [ ] `data/raw/` 配下のファイルを編集していないこと
- [ ] 転記した値はすべて見積に根拠があること（無いものは空欄）
- [ ] 価格に税込/税抜区分が付いていること（不明は「不明」）
- [ ] 仕入送料は送料見積に基づくか、無ければ空欄＋要確認になっていること
- [ ] `config/eezo_policy.json` の送料・無料ラインが出典URLと一致していること

---

## データセット情報

### data/raw/{仕入先}/ （入力）
```
data/raw/
├── _OUTPUT_TEMPLATE.xlsx   # 直下の型（出力フォーマット＋抽出スキーマ）
├── 梅屋/  見積_0416.pdf / 送料見積.xlsx
├── トワ・ヴェール/  見積_20260411.pdf   # 送料込のため送料見積なし
└── ...
```

### data/processed/{仕入先}_products.csv （Claudeが転記）
抽出スキーマは `templates/intermediate_products_schema.csv`。主要列:

| 列 | 説明 | 備考 |
|----|------|------|
| 商品名_見積準拠 | 見積記載の商品名 | 必須 |
| 温度帯 | 常温/冷蔵/冷凍 | 無ければ空欄（推定しない） |
| 仕入単価_値 / _税区分 | 金額と税込/税抜/不明 | 税区分必須 |
| 仕入送料区分 | 送料込/産直別途/不明 | |
| 税率_明記 | 見積に明記がある時のみ 8/10 | |
| 想定販売価格_税込 | 現行Shopify価格 or 目標価格（手入力） | 空なら採算は出ない |

### data/processed/{仕入先}_shipping.csv （送料見積がある時のみ）
抽出スキーマは `templates/intermediate_shipping_schema.csv`。送料体系＝温度帯型 or サイズ型。

---

## Available Commands（利用可能なコマンド）

| キーワード | 内容 | 使い方 |
|-----------|------|--------|
| `転記して` | 見積→中間CSV | 「data/raw/梅屋 の見積を転記して」 |
| `整えて` | 全工程実行 | 「全仕入先を整えて outputs に出力して」 |
| `ポリシー確認` | 送料設定の突合 | 「config の送料を出典URLと確認して」 |
| `テスト` | 採算ロジック検証 | 「tests を実行して」 |

---

## 標準ワークフロー（skills/prep-workflow.md を参照）

1. `config/eezo_policy.json` を出典URLと突き合わせる（必要なら更新）。
2. `data/raw/{仕入先}/` の見積を読み、`data/processed/{仕入先}_products.csv` へ転記（`rules/no-estimation.md`・`skills/extraction-rules.md` に従う）。送料見積があれば `_shipping.csv` も。
3. `python src/build_eezo_template.py --processed data/processed --config config/eezo_policy.json --out outputs/$(date +%Y%m%d)_DATA_EEZO_seibi_output.xlsx`
4. `outputs/` の要確認一覧を提示し、空欄・赤旗の処理方針を確認。
5. `_shopify_import.csv` を shopify-product-register へ渡す。

---

## 技術スタック
- Python 3.x / openpyxl（重い依存なし）

---

## コーディング規約（rules/coding-style.md）
- 型ヒント必須、docstring（Google形式）必須、インデント4スペース、f-string優先。

---

## ドメイン知識（skills/domain-knowledge.md）
- 送料は二層（仕入送料＝コスト／消費者負担送料＝公開ポリシー）。採算は関西着基準で集約。
- 旧・価格精査一覧との違い（送料未受領は仮置きせず空欄、無料ライン15,000円、売上＝商品＋消費者送料／原価＝仕入＋仕入送料の粗利式・決済手数料は販管費で粗利外）は `skills/shipping-model.md` を参照。
