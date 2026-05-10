# コード解説

---

## 何を自動化しているか

PowerPointの販売スライドを、ターゲット訴求別に自動生成します。

通常、セールス資料の作成には「企画→デザイン→コピー作成→スライド制作」という複数ステップが必要ですが、このシステムではサービス情報とターゲット設定を入力するだけで、スライド一式が数十秒で出力されます。

---

## 入力から出力までの流れ

```
[Step 1] service_info.yaml を読み込む
  → サービス名・価格・講師情報・特典・Q&Aなど全情報

[Step 2] パターン設定を読み込む（A / B / C）
  → ターゲット・訴求軸・配色パレット

[Step 3] Claude API（claude-opus-4-7）でコンテンツ生成
  → 各スライドのコピー・見出し・ボディをJSON形式で出力
  → 1スライド1メッセージの原則に従い生成

[Step 4] python-pptx でスライドを描画
  → スライドサイズ：13.33 × 7.5 インチ（16:9）
  → 色・フォント・レイアウトはパレット定義から参照
  → 画像アセットが存在する場合は自動挿入

[Step 5] .pptx ファイルとして出力
  → output/ フォルダに保存
  → ファイル名：cc_v6_{パターン}_{タイムスタンプ}.pptx
```

---

## 主要ファイル構成

```
cc-slide-system/
  generate_v6.py        # メイン実行ファイル
  service_info.yaml     # サービス情報（全情報の一元管理）
  .env                  # ANTHROPIC_API_KEY
  assets/
    references/
      profiles/         # 講師写真（差し替え可能）
      actionmap/        # アクションマップ画像（差し替え可能）
  output/               # 生成されたPPTXの保存先
  submission/           # 提出物一式
```

---

## 今回できたこと

- service_info.yamlからの情報読み込みと構造化
- Claude APIを使ったスライドコンテンツのJSON生成
- python-pptxによるスライド自動描画（背景色・テキスト・図形・画像）
- 3ターゲット訴求パターンの設計と生成（A:副業 / B:法人 / C:初心者）
- 画像差し替え可能な設計（ファイルを置くだけで反映）
- 画像なし版でも成立するプレースホルダ構成

---

## 未完成・改善予定のこと

| 項目 | 状況 | 改善案 |
|---|---|---|
| アニメーション | 未実装 | python-pptx の animation API または GIF で対応 |
| Webアプリ化 | 未実装 | FastAPI + フロントエンドでUI化 |
| アイコン・イラスト | 未実装 | SVGアイコンライブラリの組み込み |
| パターン別コピーの精度 | 改善余地あり | プロンプトのチューニングで向上可能 |
| デモURL | 未提供 | Streamlit等でデプロイ予定 |

---

## 実行方法

```bash
# 依存関係のインストール
pip install python-pptx anthropic python-dotenv pyyaml

# .env に API キーを設定
echo "ANTHROPIC_API_KEY=your_key_here" > .env

# スライド生成（3パターン一括）
python3 generate_v6.py
```
