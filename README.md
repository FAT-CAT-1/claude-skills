# claude-skills

Claude Code で使う、自作のスキル（作業手順書）3つです。
どれも、実際の制作で同じ失敗を繰り返した経験を、ほかの題材でも使える手順にまとめたものです。

## 中身

| フォルダー | 一言で |
|---|---|
| `feedback-round/` | レビューや他の AI の意見を受け取って成果物に反映するときの手順（項目の3分類・裁定の仰ぎ方・数値の検算・記録・停止条件・失敗パターン30件） |
| `slides/` | PowerPoint（.pptx）をスクリプトで作り、**実際に描いた絵を目で見て**崩れを防ぐ手順と部品（Python） |
| `homepage/` | 日本語のホームページを「情報は正しいが平板」にしないための手順・デザインシステム・実測検査（HTML/CSS/Python） |

## 入れ方

使いたいフォルダーを、フォルダーごと `~/.claude/skills/` の下へコピーします。

```
~/.claude/skills/feedback-round/SKILL.md
~/.claude/skills/slides/SKILL.md
~/.claude/skills/homepage/SKILL.md
```

コピーしたあと Claude Code を起動し直すと、内容に合う依頼のときに自動で使われます。

## 使う前に知っておくこと

- **slides は Windows 向けです。** 既定の書体はメイリオ（Meiryo）で、検査（`verify.py`）は PowerPoint 本体を優先して使います（無ければ LibreOffice）。必要なもの: `python-pptx`（必須）、`pywin32`（PowerPoint で検査する場合）、`pymupdf`（LibreOffice で検査する場合）。
- **homepage のテンプレートは Google Fonts を読み込みます。** 作ったページを公開すると、閲覧した人の情報（IP アドレスなど）が Google に送られます。気になる場合は書体を自分のサーバーに置いて読み込んでください。
- **書体ファイルは同梱していません。** 使っている書体（Zen Old Mincho / Noto Sans JP）は SIL Open Font License（OFL）の書体です。
- `homepage/assets/template.html` の地図ピンのアイコンは [Lucide](https://lucide.dev)（ISC License）の「map-pin」です。
- `homepage/verify.py` は Google Chrome（ヘッドレス）で実際に描画して測ります。
- 手順書の中の「実例」は、元になった制作の固有の内容を外して書き直してあります。サンプルの文面（例: `slides/example_build.py`）は架空です。

## 作り方について

これらのスキルは、Claude（Anthropic の AI）と一緒に書きました。

## 作者・ライセンス

作者: FAT-CAT-1（GitHub）

[MIT License](LICENSE)
