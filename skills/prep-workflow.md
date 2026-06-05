# 整えるワークフロー（prep-workflow）

CCWでこのリポジトリを動かす標準手順。Claudeの担当は **Step 2の転記まで**。Step 3以降の採算・送料引き当て・整形は `src/build_eezo_template.py` が決定論的に行う。

## Step 0: ポリシー突合
- `config/eezo_policy.json` の `_meta.出典` URL を確認し、送料無料ライン・地域別送料が一致するか突き合わせる。
- 相違があれば config を更新（確認日も更新）。数値の据え置き禁止。

## Step 1: 対象の棚卸し
- `data/raw/` 直下の仕入先ディレクトリを一覧化。
- 各仕入先の送料区分（送料込 / 産直別途 / 不明）を見積本文・備考から判定。
- 送料見積ファイルの有無を確認。

## Step 2: 転記（Claude担当）
各仕入先ごとに:
1. 見積を読み、`data/processed/{仕入先}_products.csv` へ転記。
   - スキーマ: `templates/intermediate_products_schema.csv`
   - ルール: `rules/no-estimation.md` ＋ `skills/extraction-rules.md`
   - 事実のみ。無ければ空欄。税込/税抜区分を必ず記録。
2. 送料見積があれば `data/processed/{仕入先}_shipping.csv` へ。
   - スキーマ: `templates/intermediate_shipping_schema.csv`
   - 送料体系（温度帯型 / サイズ型）を記載。サイズ型は自動引当されず要確認になる。
   - 送料見積が無ければファイルを作らない（→ 仕入送料は空欄＋要確認）。

チェックポイント出力:
```
✅ 完了: 梅屋の見積を転記（data/processed/梅屋_products.csv: 4商品）
📄 送料: 梅屋_shipping.csv（温度帯型）
➡️ 次: トワ・ヴェールの転記
```

## Step 3: ビルド
```bash
python src/build_eezo_template.py \
  --processed data/processed \
  --config config/eezo_policy.json \
  --out outputs/$(date +%Y%m%d)_DATA_EEZO_seibi_output.xlsx
```
生成物:
- `outputs/...xlsx`（出力テンプレ / 採算サマリ / 要確認一覧 / 設定）
- `outputs/..._shopify_import.csv`

## Step 4: レビュー
- 要確認一覧を提示し、空欄・赤旗（逆ざや・上代逆転・無料ライン到達）の処理方針を確認。
- 想定販売価格が未入力の行は採算が出ない旨を伝える。

## Step 5: Shopify登録へ接続
`_shopify_import.csv` の各行を shopify-product-register へ。登録前に:
- `title`（EEZO命名規則）/ `productType` / `descriptionHtml`（3ブロック）/ `画像URL` を確定。
- `要確認メモ` が残る行は DRAFT のまま保留。
