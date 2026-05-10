# 画像アセット運用ガイド

---

## 推奨アセット一覧

### 講師プロフィール写真

| ファイル名 | 使用箇所 | 推奨サイズ |
|---|---|---|
| `assets/references/profiles/mikami-profile.png` | みかみ自己紹介スライド（右側） | 正方形 推奨（800×800px以上） |
| `assets/references/profiles/sato-profile.png` | 佐藤自己紹介スライド（右側） | 正方形 推奨（800×800px以上） |

### アクションマップ

| ファイル名 | 使用箇所 | 内容 |
|---|---|---|
| `assets/references/actionmap/actionmap-overview.png` | アクションマップ全体スライド | 全体サムネイル |
| `assets/references/actionmap/actionmap-kensakujidoka.png` | アクションマップスライド 代表例① | 案件探索自動化のサムネイル |
| `assets/references/actionmap/actionmap-slide-seisaku.png` | アクションマップスライド 代表例② | スライド制作のサムネイル |
| `assets/references/actionmap/actionmap-system-kaihatsu.png` | アクションマップスライド 代表例③ | システム開発のサムネイル |

---

## 差し替え方

1. 上記の推奨ファイル名で画像を用意する
2. 対応フォルダに保存する（上書きOK）
3. `python3 generate_v6.py` を実行する
4. ファイルが存在する場合は自動的に画像あり版で生成される

コードを一切触る必要はありません。

---

## 画像が無い場合の代替方法

画像ファイルが存在しない場合、以下の代替表現で自動的に成立します。

| 画像 | 代替表現 |
|---|---|
| mikami-profile.png | 「月商5億 / 受講生6,000名以上」テキストカード |
| sato-profile.png | 「未経験→初月1,000万円案件」ストーリーテキスト |
| actionmap-*.png | 代表3テーマのテキストリスト＋翻訳文 |

---

## 推奨フォーマット

- **形式：** PNG（透明背景対応のため推奨）
- **カラーモード：** RGB
- **最低解像度：** 72dpi（スライド表示用）
- **推奨解像度：** 150dpi以上（印刷・拡大表示に対応）

---

## 注意事項

- ファイル名は完全一致が必要です（大文字・小文字・拡張子を含む）
- 別名のファイルを置いても自動では読み込まれません
- ファイル名を変えた場合は `generate_v6.py` 内の `ASSETS_DIR` 設定を変更してください
