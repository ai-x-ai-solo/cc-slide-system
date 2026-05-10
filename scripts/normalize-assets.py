#!/usr/bin/env python3
"""
画像正規化スクリプト
- アクションマップ: 1200×675px (16:9), 白背景, レターボックス
- プロフィール: 800×360px (横長バナー), 元の比率を活かして拡大
  → 元画像が横長バナーなのでそのまま活かす設計
"""
from PIL import Image, ImageDraw
from pathlib import Path

BASE = Path(__file__).parent.parent
SRC_AM = BASE / "assets/references/actionmap"
SRC_PR = BASE / "assets/references/profiles"
OUT_AM = BASE / "assets/processed/actionmap"
OUT_PR = BASE / "assets/processed/profiles"
OUT_AM.mkdir(parents=True, exist_ok=True)
OUT_PR.mkdir(parents=True, exist_ok=True)


def letterbox(img, target_w, target_h, bg_color):
    """アスペクト比を保ちつつ target サイズに収め、余白は bg_color で埋める"""
    iw, ih = img.size
    scale = min(target_w / iw, target_h / ih)
    new_w = int(iw * scale)
    new_h = int(ih * scale)
    resized = img.resize((new_w, new_h), Image.LANCZOS)
    canvas = Image.new("RGBA", (target_w, target_h), bg_color)
    paste_x = (target_w - new_w) // 2
    paste_y = (target_h - new_h) // 2
    if resized.mode == "RGBA":
        canvas.paste(resized, (paste_x, paste_y), resized)
    else:
        canvas.paste(resized, (paste_x, paste_y))
    return canvas


# ===== アクションマップ: 1200×675 白背景 =====
am_files = {
    "actionmap-kensakujidoka.png":    "actionmap-kensakujidoka-normalized.png",
    "actionmap-slide-seisaku.png":    "actionmap-slide-seisaku-normalized.png",
    "actionmap-system-kaihatsu.png":  "actionmap-system-kaihatsu-normalized.png",
    "actionmap-overview.png":         "actionmap-overview-normalized.png",
}
for src_name, dst_name in am_files.items():
    src = SRC_AM / src_name
    if not src.exists():
        print(f"⚠️  スキップ（ファイルなし）: {src_name}")
        continue
    img = Image.open(src).convert("RGBA")
    out = letterbox(img, 1200, 675, (255, 255, 255, 255))
    out.convert("RGB").save(OUT_AM / dst_name, "PNG", optimize=True)
    print(f"✅ {dst_name}  ({out.size[0]}×{out.size[1]})")


# ===== プロフィール: 600×600px 正方形（黒背景で完全統一）=====
# バナー部分を BANNER_W×BANNER_H に強制統一 → 600×600 黒背景の中央に配置
# → 2人のバナーがピクセル単位で全く同じサイズで表示される
CANVAS_SIZE = 600
BANNER_W, BANNER_H = 600, 270  # バナーを強制的にこのサイズに揃える

pr_files = {
    "mikami-profile.png": "mikami-profile-normalized.png",
    "sato-profile.png":   "sato-profile-normalized.png",
}
for src_name, dst_name in pr_files.items():
    src = SRC_PR / src_name
    if not src.exists():
        print(f"⚠️  スキップ（ファイルなし）: {src_name}")
        continue
    img = Image.open(src).convert("RGBA")
    iw, ih = img.size
    print(f"   {src_name}: 元サイズ {iw}×{ih}px")
    # バナーを強制的に BANNER_W×BANNER_H にリサイズ（比率崩れを最小化するため幅基準）
    banner = img.resize((BANNER_W, BANNER_H), Image.LANCZOS)
    # 600×600 の黒背景キャンバスに中央配置
    canvas = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (13, 13, 13, 255))
    paste_y = (CANVAS_SIZE - BANNER_H) // 2
    canvas.paste(banner, (0, paste_y))
    canvas.convert("RGB").save(OUT_PR / dst_name, "PNG", optimize=True)
    print(f"✅ {dst_name}  ({CANVAS_SIZE}×{CANVAS_SIZE}px / バナー部分 {BANNER_W}×{BANNER_H}px で統一)")


print("\n✨ 正規化完了")
print(f"  アクションマップ → {OUT_AM}")
print(f"  プロフィール     → {OUT_PR}")
