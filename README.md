# 立岩小学校 校内研修サイト（2026-08-24「AIには、人に頼むように。」）

8/24 生成AI研修（Gemini・NotebookLM・60分）の当日ページと、あとから読む資料庫。
旧「教師のスライドづくり」本編は `slides/` に温存（資料編としてリンク）。

- 公開URL: https://a-tozak.github.io/tateiwa-kenshu/
- 本文・図版・比較スライド（SVG）・テンプレート（pptx）・配布資料はすべて外﨑顯博の自作。
  挿絵・表紙・避難所デモ画像は School Stock 素材スタジオ制作のオリジナル。
- 児童の個人情報・学校内部資料は含まない。

## 構成

```
index.html          当日ページ（8/24版：頼み方3つ・画像生成・NotebookLM・触る時間・プロンプト帳）
slides/             資料編＝旧「教師のスライドづくり」全編（大原則5つ・Before/After・テンプレ）
slide/              当日の投影スライド（HTML。矢印キーでページ送り・F11で全画面）
handout/            配布資料（A4×10p）。index.html → Chrome焼き → 立岩小_研修配布資料_2026-08-24.pdf
guide/              あとから読む資料庫（考え方・基本・小技・教科別導入・GAS・NotebookLM・素材集）
templates/          文言差し替え型テンプレ pptx 5本（t5=学級開き見本・build_t5.py）
workspace_build/    テンプレの生成スクリプト（python-pptx）
assets/             挿絵（水彩・自作）＋避難所デモ
qr/                 当日投影用QR（8.24版）
```

## 更新のしかた（自分用メモ）

- **配布資料を直したら**: `handout/index.html` を編集 → Chromeで `--print-to-pdf` → `handout/check/` のPNGで全ページ目視
- **テンプレを直したら**: `workspace_build/build_templates.py`（t1〜t4）または `build_t5.py` → `python3` 実行で `templates/` 上書き
- **反映**: commit → push（GitHub Pagesに数分で反映）

© 2026 外﨑顯博 / LiFE with
