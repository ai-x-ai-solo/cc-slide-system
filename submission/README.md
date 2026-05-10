# Claude Code ブートキャンプ 販売スライド自動生成システム
## 採用評価タスク 提出物

---

## 今回作ったもの

「本質のClaude Code 完全攻略 7dayブートキャンプ」の販売スライドを、
ターゲット別に自動生成するシステムです。

- **入力：** service_info.yaml（サービス情報）＋ パターン設定
- **出力：** ターゲット訴求別 .pptx ファイル（15〜16枚構成）
- **生成：** Python + python-pptx + Claude API（claude-opus-4-7）

---

## 提出物一覧

| ファイル | 内容 |
|---|---|
| `outputs/pattern-a/` | Pattern A スライド（副業・案件獲得訴求） |
| `outputs/pattern-b/` | Pattern B スライド（法人・業務改善訴求） |
| `outputs/pattern-c/` | Pattern C スライド（初心者安心訴求） |
| `docs/design-system-and-workflow.md` | デザインシステム＆ワークフロー解説 |
| `docs/loom-script.md` | デモ動画台本（3分） |
| `docs/self-review.md` | 自己評価 |
| `prompts/master-prompt.md` | マスタープロンプト |
| `prompts/pattern-a-input.md` | Pattern A 入力例 |
| `prompts/pattern-b-input.md` | Pattern B 入力例 |
| `prompts/pattern-c-input.md` | Pattern C 入力例 |
| `assets-guide.md` | 画像アセット運用ガイド |
| `code-summary.md` | コード解説 |
| `CHECKLIST.md` | 提出前チェックリスト |

---

## どこを見れば何が分かるか

- **システムの設計意図・AI活用方法** → `docs/design-system-and-workflow.md`
- **スライドの実物** → `outputs/pattern-*/`
- **プロンプト設計** → `prompts/`
- **コードの説明** → `code-summary.md`
- **画像の差し替え方** → `assets-guide.md`

---

## 3パターンの説明

### Pattern A：副業・案件獲得訴求型
- **ターゲット：** AI副業に興味がある個人・スキルプラス受講生
- **感情：** 「今動けば先行者になれる」
- **主軸：** 月10万円案件 / 7日 / 0.1% / 早い者勝ち
- **デザイン：** 黒×パープル、未来感・高級感、数字を大きく

### Pattern B：法人・業務改善訴求型
- **ターゲット：** 中小企業経営者・管理職・業務改善担当者
- **感情：** 「自社の業務に使えそう」
- **主軸：** 業務効率化 / 自動化 / 人件費削減 / 生産性向上
- **デザイン：** 白×青、信頼感・整理された印象

### Pattern C：初心者安心訴求型
- **ターゲット：** AI初心者・PCが得意でない人・未経験者
- **感情：** 「自分でもできそう」
- **主軸：** 未経験OK / 日本語でOK / 質問し放題 / 一緒に作業
- **デザイン：** 明るい・柔らかい・余白多め・安心感

---

## 画像差し替え設計

本システムは「画像あり版」と「画像なし版」の両方で成立する設計です。

- 画像ファイルが `assets/references/` に存在する場合 → 自動的に画像あり版で生成
- ファイルが存在しない場合 → プレースホルダ（テキスト中心）で成立

詳細は `assets-guide.md` を参照してください。

---

## 前提

- Python 3.10+、python-pptx、anthropic ライブラリが必要です
- Claude API キーを `.env` に設定してください
- スライド生成は `python3 generate_v6.py` で実行します
