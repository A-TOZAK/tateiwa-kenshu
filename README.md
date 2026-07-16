# 立岩小学校 校内研修「教師のスライドづくり」

校内研修（60分）の当日ページと、あとから読む資料庫。

- 公開URL: https://a-tozak.github.io/tateiwa-kenshu/
- 本文・図版・比較スライド（SVG）・テンプレート（pptx）はすべて外﨑顯博の自作。
  挿絵はSchool Stock素材スタジオ制作のオリジナルイラスト。
- 児童の個人情報・学校内部資料は含まない。

## 構成

```
index.html          当日ページ（60分の見取り図・大原則・Before/After・テンプレ配布・ハンズオン）
guide/              あとから読む資料庫（考え方・基本・小技・教科別導入・GAS・NotebookLM・素材集）
templates/          文言差し替え型テンプレ pptx 4本
workspace_build/    テンプレの生成スクリプト（python-pptx）
assets/             挿絵（水彩・自作）
qr/                 当日投影用QR
```

## 更新のしかた（自分用メモ）

- **日付が決まったら**: `index.html` の `.hero .meta` に `<span>日程 <b>◯月◯日（◯）</b></span>` を1行足す
- **テンプレを直したら**: `workspace_build/build_templates.py` を編集 → `python3 build_templates.py` → `templates/` が上書きされる
- **反映**: commit → push（GitHub Pagesに数分で反映）

© 2026 外﨑顯博 / LiFE with
