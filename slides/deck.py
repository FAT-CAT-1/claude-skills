# -*- coding: utf-8 -*-
"""deck.py — PowerPoint スライド生成の共通部品（/slides スキル本体）

    from deck import Deck
    d = Deck(theme='lecture', ratio='16:9')
    d.board('大きな*1行*', sub=['補足'], notes=d.note('0:00-0:20', '台詞...'))
    d.save('out.pptx')

設計の根拠はすべて `SKILL.md` にある。
ここには **実際に事故ったので入れた対策** だけをコメントで残す。

⚠この部品を勝手に「軽く」しないこと。下の4つは全部、実機で崩れたのを見てから入れたもの:
  L1 日本語の書体は a:ea / a:cs にも書く（python-pptx は latin しか書かない）
  L2 補足は文ごとに段落を分ける（1段落だと末尾数文字が次行に取り残される）
  L3 棒グラフの軸は全行で共有する（行ごとに組むと同じ値が違う長さで描かれる＝図が嘘をつく）
  L4 文字は必ず実際に描いて見る（数値チェックは折り返しの崩れを捕まえない）
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import difflib
import os
import re


# ══════════════════════════════════════════════════════════
# テーマ
# ══════════════════════════════════════════════════════════
def _c(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


THEMES = {
    # 作品の世界観を出す発表向け（レトロ/ゲーム/クリエイティブ）。暗室前提。
    'neon': dict(
        bg='03020A', fg='E8E8F0', muted='8A8A9C',
        a1='00F5FF', a2='FFD93D', a3='FF6AD5',
        rule='2A2740', track='141225',
        font='Meiryo',
    ),
    # 大学の講義・学術発表向け。**明るい教室でも読める**ことを最優先。
    'lecture': dict(
        bg='FFFFFF', fg='1A1A22', muted='5A5A66',
        a1='0B5FA5', a2='B25000', a3='0F7A5F',
        rule='C8C8D0', track='ECECF2',
        font='Meiryo',
    ),
    # 提案・ビジネスプレゼン向け。落ち着いた紺。
    'business': dict(
        bg='0E1B2E', fg='F2F4F8', muted='94A0B4',
        a1='4FC3F7', a2='FFC64D', a3='7BDCA0',
        rule='27384F', track='1A2942',
        font='Meiryo',
    ),
    # 印刷・コピー・プロジェクタが弱い会場向け。色に頼らない高コントラスト。
    'mono': dict(
        bg='FFFFFF', fg='111111', muted='555555',
        a1='000000', a2='444444', a3='777777',
        rule='BBBBBB', track='E6E6E6',
        font='Meiryo',
    ),
}

RATIOS = {'16:9': (13.333, 7.5), '4:3': (10.0, 7.5)}

# 会場の後ろから読める下限。これを下回る指定は build 時に警告する。
MIN_HEADLINE_PT = 28
MIN_BODY_PT = 14


# ══════════════════════════════════════════════════════════
# 文字装飾のマークアップ
#   *強調*  = アクセント1 ／ **強調** = アクセント2 ／ ~控えめ~ = 薄い色
#   改行は \n。完全に制御したい時は [(文字列, 役割), ...] を直接渡してもよい。
# ══════════════════════════════════════════════════════════
_MARKUP = re.compile(r'\*\*(.+?)\*\*|\*(.+?)\*|~(.+?)~', re.S)


def parse_markup(s):
    """'報告を*あと*ではなく' -> [('報告を','fg'), ('あと','a1'), ('ではなく','fg')]"""
    if not isinstance(s, str):
        return s
    out, pos = [], 0
    for m in _MARKUP.finditer(s):
        if m.start() > pos:
            out.append((s[pos:m.start()], 'fg'))
        if m.group(1) is not None:
            out.append((m.group(1), 'a2'))
        elif m.group(2) is not None:
            out.append((m.group(2), 'a1'))
        else:
            out.append((m.group(3), 'muted'))
        pos = m.end()
    if pos < len(s):
        out.append((s[pos:], 'fg'))
    return out or [('', 'fg')]


def _lines(spec):
    """文字列（\n区切り）／[(文字,役割)]／その入れ子 を「行のリスト」に正規化する"""
    if isinstance(spec, str):
        return [parse_markup(ln) for ln in spec.split('\n')]
    if spec and isinstance(spec[0], (list,)):
        return [parse_markup(x) if isinstance(x, str) else x for x in spec]
    if spec and isinstance(spec[0], tuple):
        return [spec]
    return [parse_markup(x) if isinstance(x, str) else x for x in spec]


# 文字幅のざっくり見積り（pptx はテキストを実測できないので必要）。
# 全角≈1.0em・半角≈0.55em。棒グラフの軸そろえと、はみ出し警告に使う。
def text_width_in(s, size_pt):
    w = 0.0
    for ch in s:
        w += 1.0 if ord(ch) > 0x2E80 else 0.55
    return w * size_pt / 72.0


# ══════════════════════════════════════════════════════════
# ノート（＝台本の唯一の正）
#   1行目 [時刻] ／ ※◯◯版では飛ばす ／ 台詞 ／ ⚠警告
#   この形にしておくと notes_to_script.py が紙用の台本へ機械的に書き出せる。
#   ⚠ノートに書いた台詞が正。別ファイルに原稿を持たない（二重管理は必ず食い違う）。
# ══════════════════════════════════════════════════════════
def note(time=None, say='', warn=None, skip=None):
    parts = []
    if time:
        parts.append('[%s]' % time)
    if skip:
        parts.append('※%s では飛ばす' % '／'.join(skip if isinstance(skip, (list, tuple)) else [skip]))
    if say:
        parts.append(say.strip())
    for w in ([warn] if isinstance(warn, str) else (warn or [])):
        parts.append('⚠ ' + w)
    return '\n'.join(parts)


class Deck:
    def __init__(self, theme='lecture', ratio='16:9', font=None):
        if theme not in THEMES:
            raise ValueError('theme は %s のいずれか' % list(THEMES))
        self.t = dict(THEMES[theme])
        self.theme_name = theme
        if font:
            self.t['font'] = font
        self.font = self.t['font']
        w, h = RATIOS[ratio]
        self.W, self.H = Inches(w), Inches(h)
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = self.W, self.H
        self._blank = self.prs.slide_layouts[6]
        self.warnings = []
        self.skips = {}          # 版名 -> [飛ばす枚番]
        self.note = staticmethod(note).__func__

    # ── 内部 ────────────────────────────────────────────
    def _col(self, role):
        return _c(self.t.get({'fg': 'fg', 'muted': 'muted',
                              'a1': 'a1', 'a2': 'a2', 'a3': 'a3'}.get(role, 'fg')))

    def _font(self, run, size, color, bold=True):
        f = run.font
        f.name = self.font
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = color
        # L1: python-pptx は latin しか書かない。ea/cs を明示しないと
        #     環境によって日本語だけ別書体に化ける（実機で発生）。
        rPr = run._r.get_or_add_rPr()
        for tag in ('a:ea', 'a:cs'):
            rPr.append(rPr.makeelement(qn(tag), {'typeface': self.font}))

    def _slide(self, notes=''):
        s = self.prs.slides.add_slide(self._blank)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.W, self.H)
        bg.fill.solid()
        bg.fill.fore_color.rgb = _c(self.t['bg'])
        bg.line.fill.background()
        bg.shadow.inherit = False
        if notes:
            s.notes_slide.notes_text_frame.text = notes
            for m in re.finditer(r'※(.+?) では飛ばす', notes):
                for v in m.group(1).split('／'):
                    self.skips.setdefault(v, []).append(len(self.prs.slides._sldIdLst))
        return s

    def _text(self, s, top, height, spec, size, align=PP_ALIGN.CENTER,
              left=None, width=None, line_spacing=1.25, bold=True):
        left = Inches(0.6) if left is None else left
        width = (self.W - Inches(1.2)) if width is None else width
        tb = s.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, line in enumerate(_lines(spec)):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            for text, role in line:
                r = p.add_run()
                r.text = text
                self._font(r, size, self._col(role), bold)
            # はみ出しの早期警告（見積り。最終判断は必ず目視）
            plain = ''.join(t for t, _ in line)
            if text_width_in(plain, size) > width / 914400 * 1.02:
                self.warnings.append(
                    '%d枚目: 「%s」が横幅に収まらない見込み（%.1f pt では長い）'
                    % (len(self.prs.slides._sldIdLst), plain[:24], size))
        return tb

    def _rect(self, s, left, top, w, h, fill_hex, line_hex=None):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
        r.fill.solid()
        r.fill.fore_color.rgb = _c(fill_hex)
        if line_hex:
            r.line.color.rgb = _c(line_hex)
            r.line.width = Pt(1)
        else:
            r.line.fill.background()
        r.shadow.inherit = False
        return r

    # ══════════════════════════════════════════════════
    # 型1: 看板（1枚に大きな1行）— 話し手が主役の発表
    # ══════════════════════════════════════════════════
    def board(self, headline, sub=None, notes='', size=48, sub_size=20):
        if size < MIN_HEADLINE_PT:
            self.warnings.append('%d枚目: 見出し %dpt は後ろの席から読めない見込み'
                                 % (len(self.prs.slides._sldIdLst) + 1, size))
        s = self._slide(notes)
        if sub:
            # L2: 補足は必ず「1文＝1行」。1段落に2文入れると幅が足りない時に
            #     文の途中で折り返し、末尾の数文字だけが次行に取り残される（実機で発生）。
            sub = self._split_sentences(sub)
            self._text(s, self.H * 0.29, self.H * 0.32, headline, size)
            self._text(s, self.H * 0.62, self.H * 0.18, sub, sub_size, line_spacing=1.5)
        else:
            self._text(s, self.H * 0.31, self.H * 0.39, headline, size)
        return s

    def _split_sentences(self, sub):
        out = []
        for item in (sub if isinstance(sub, (list, tuple)) else [sub]):
            if isinstance(item, str):
                parts = [x for x in re.split(r'(?<=。)', item) if x.strip()]
                out.extend(parts if len(parts) > 1 else [item])
            else:
                out.append(item)
        return [parse_markup(x) if isinstance(x, str) else x for x in out]

    # ══════════════════════════════════════════════════
    # 型2: 説明（見出し＋中身）— 図表や箇条を見せる回
    # ══════════════════════════════════════════════════
    def explain(self, heading, body=None, caption=None, notes='',
                heading_size=32, body_size=22, bullet='—'):
        s = self._slide(notes)
        self._text(s, self.H * 0.08, self.H * 0.12, heading, heading_size)
        if body:
            items = body if isinstance(body, (list, tuple)) else [body]
            if len(items) > 6:
                self.warnings.append('%d枚目: 箇条が%d個。看板ではなく資料になっている（6個までを推奨）'
                                     % (len(self.prs.slides._sldIdLst), len(items)))
            lines = []
            for it in items:
                if isinstance(it, str):
                    lines.append(parse_markup((bullet + ' ' if bullet else '') + it))
                else:
                    lines.append(([(bullet + ' ', 'muted')] if bullet else []) + list(it))
            self._text(s, self.H * 0.26, self.H * 0.52, lines, body_size,
                       align=PP_ALIGN.LEFT, left=self.W * 0.12,
                       width=self.W * 0.76, line_spacing=1.9, bold=False)
        if caption:
            self._text(s, self.H * 0.86, self.H * 0.09, caption, max(MIN_BODY_PT, body_size - 6))
        return s

    def blank(self, notes=''):
        """デモ中・間を取る時。背景だけ。"""
        return self._slide(notes)

    # ══════════════════════════════════════════════════
    # レシピ1: 横棒グラフ
    #   L3: 目盛りの原点と全長を **全行で共有** する。行ごとに組むと
    #       ラベルの長さの違いがそのまま棒の長さに乗り、同じ値が違う長さで
    #       描かれる（実測 72%が526pxと489px）＝図が嘘をつく。
    # ══════════════════════════════════════════════════
    def bars(self, heading, rows, caption=None, notes='',
             heading_size=26, label_size=18, unit='%', max_value=100):
        s = self._slide(notes)
        self._text(s, self.H * 0.07, self.H * 0.11, heading, heading_size)
        label_w = max(text_width_in(r[0], label_size) for r in rows) + 0.25
        pct_w = text_width_in('000' + unit, label_size) + 0.2
        left = Inches(0.7)
        track_l = left + Inches(label_w)
        track_w = self.W - track_l - Inches(pct_w) - Inches(0.7)
        n = len(rows)
        band = self.H * 0.52 / max(n, 1)
        bar_h = min(band * 0.46, Inches(0.42))
        top = self.H * 0.26
        for i, row in enumerate(rows):
            label, value = row[0], row[1]
            hot = len(row) > 2 and row[2]
            note_ = row[3] if len(row) > 3 else None
            role = 'a2' if hot else 'a1'
            y = top + band * i
            self._text(s, y - Inches(0.06), bar_h + Inches(0.12),
                       [[(label, 'a2' if hot else 'muted')]], label_size,
                       align=PP_ALIGN.RIGHT, left=left, width=Inches(label_w) - Inches(0.2))
            self._rect(s, track_l, y, track_w, bar_h, self.t['track'], self.t['rule'])
            self._rect(s, track_l, y, Emu(int(track_w * min(value, max_value) / max_value)),
                       bar_h, self.t[role])
            self._text(s, y - Inches(0.06), bar_h + Inches(0.12),
                       [[('%g%s' % (value, unit), role)]], label_size,
                       align=PP_ALIGN.LEFT, left=track_l + track_w + Inches(0.15),
                       width=Inches(pct_w))
            if note_:
                self._text(s, y + bar_h, Inches(0.3), [[(note_, 'a2' if hot else 'muted')]],
                           max(11, label_size - 6), align=PP_ALIGN.RIGHT,
                           left=left, width=Inches(label_w) - Inches(0.2))
        if caption:
            self._text(s, self.H * 0.86, self.H * 0.1, caption, label_size)
        return s

    # ══════════════════════════════════════════════════
    # レシピ2: 左右の比較（対比を1枚で見せる）
    #   ⚠どちらの列が何なのかを必ず書く。口で言うだけでは客席が追えない。
    # ══════════════════════════════════════════════════
    def compare(self, heading, left, right, caption=None, notes='',
                heading_size=22, title_size=34, body_size=18):
        s = self._slide(notes)
        self._text(s, self.H * 0.09, self.H * 0.09, heading, heading_size)
        colw = self.W * 0.40
        for side, (tag, title, body) in enumerate(
                [(left[0], left[1], left[2]), (right[0], right[1], right[2])]):
            x = Inches(0.7) if side == 0 else self.W - Inches(0.7) - colw
            role = 'muted' if side == 0 else 'a1'
            self._text(s, self.H * 0.22, self.H * 0.07, [[(tag, role)]],
                       max(MIN_BODY_PT, body_size - 2), left=x, width=colw)
            self._text(s, self.H * 0.31, self.H * 0.16, title, title_size, left=x, width=colw)
            self._text(s, self.H * 0.53, self.H * 0.22, body, body_size,
                       left=x, width=colw, line_spacing=1.6, bold=False)
        rx = (self.W - Emu(9525 * 2)) // 2
        self._rect(s, rx, self.H * 0.24, Emu(9525 * 2), self.H * 0.46, self.t['rule'])
        if caption:
            self._text(s, self.H * 0.83, self.H * 0.1, caption, max(MIN_BODY_PT, body_size))
        return s

    # ══════════════════════════════════════════════════
    # レシピ3: 2列の表（役割・対応・定義など）
    # ══════════════════════════════════════════════════
    def table(self, heading, rows, caption=None, notes='',
              heading_size=32, row_size=20):
        s = self._slide(notes)
        self._text(s, self.H * 0.08, self.H * 0.12, heading, heading_size)
        # ⚠左右の列幅は**実際の文字から測る**。片方を固定比（W*0.42 など）にすると、
        #   中身が短いときに余白ごと中央寄せされて**塊が左に寄って見える**（実機で発生）。
        keyw = max(text_width_in(r[0], row_size) for r in rows) + 0.2
        valw = max(text_width_in(r[1], row_size) for r in rows) + 0.2
        gap = Inches(0.45)
        left = (self.W - (Inches(keyw) + gap + Inches(valw))) // 2
        n = len(rows)
        band = self.H * 0.50 / max(n, 1)
        top = self.H * 0.28
        for i, row in enumerate(rows):
            key, val = row[0], row[1]
            hot = len(row) > 2 and row[2]
            y = top + band * i
            self._text(s, y, band * 0.8, [[(key, 'a2' if hot else 'fg')]], row_size,
                       align=PP_ALIGN.RIGHT, left=left, width=Inches(keyw))
            self._text(s, y, band * 0.8, [[(val, 'a2' if hot else 'muted')]], row_size,
                       align=PP_ALIGN.LEFT, left=left + Inches(keyw) + gap,
                       width=Inches(valw))
        if caption:
            self._text(s, self.H * 0.85, self.H * 0.1, caption, max(MIN_BODY_PT, row_size - 4))
        return s

    # ══════════════════════════════════════════════════
    # レシピ4: 手順・変化（→ でつなぐ）
    # ══════════════════════════════════════════════════
    def steps(self, heading, items, caption=None, notes='',
              heading_size=30, step_size=22):
        s = self._slide(notes)
        self._text(s, self.H * 0.13, self.H * 0.12, heading, heading_size)
        n = len(items)
        if n > 5:
            self.warnings.append('%d枚目: 手順が%d個。1枚は5個までを推奨' % (len(self.prs.slides._sldIdLst), n))
        # ⚠語ごとに幅が違うので、等分の枠に置くと矢印の位置が揃わず、
        #   長い語のときは矢印とぶつかる（実機で「会場の形を変える →」が接触した）。
        #   → 語は実測幅の枠に、矢印は**専用の固定幅の枠**に置き、全体を中央寄せする。
        ws = [Inches(text_width_in(it, step_size) + 0.12) for it in items]
        aw = Inches(text_width_in('→', step_size) + 0.55)
        total = sum(ws, Emu(0)) + aw * (n - 1)
        avail = self.W - Inches(0.8)
        if total > avail:
            self.warnings.append('%d枚目: 手順が横幅に収まらない。語を短くするか個数を減らす'
                                 % (len(self.prs.slides._sldIdLst)))
        x = (self.W - total) // 2
        top = self.H * 0.42
        for i, it in enumerate(items):
            hot = i == n - 1
            self._text(s, top, self.H * 0.14, [[(it, 'a2' if hot else 'fg')]], step_size,
                       left=x, width=ws[i])
            x += ws[i]
            if i < n - 1:
                self._text(s, top, self.H * 0.14, [[('→', 'muted')]], step_size,
                           left=x, width=aw)
                x += aw
        if caption:
            self._text(s, self.H * 0.64, self.H * 0.1, caption, max(MIN_BODY_PT, step_size - 6))
        return s

    # ══════════════════════════════════════════════════
    # レシピ5: 数字を並べる（実績・規模）
    #   ⚠ここに置く数字は、出典から機械的に取れるものだけにする。
    #     手で書いた数字は必ず腐る（実際に腐った）。
    # ══════════════════════════════════════════════════
    def numbers(self, heading, rows, caption=None, notes='',
                heading_size=30, num_size=32, desc_size=20):
        s = self._slide(notes)
        self._text(s, self.H * 0.08, self.H * 0.12, heading, heading_size)
        # ⚠table と同じ理由で、右列の幅も実際の文字から測る（固定比だと左に寄る）
        numw = max(text_width_in(str(r[0]), num_size) for r in rows) + 0.2
        descw = max(text_width_in(r[1], desc_size) for r in rows) + 0.2
        gap = Inches(0.5)
        left = (self.W - (Inches(numw) + gap + Inches(descw))) // 2
        n = len(rows)
        band = self.H * 0.50 / max(n, 1)
        top = self.H * 0.27
        for i, row in enumerate(rows):
            big, desc = str(row[0]), row[1]
            role = row[2] if len(row) > 2 else 'a1'
            y = top + band * i
            self._text(s, y, band * 0.8, [[(big, role)]], num_size,
                       align=PP_ALIGN.RIGHT, left=left, width=Inches(numw))
            self._text(s, y, band * 0.8, [[(desc, 'muted')]], desc_size,
                       align=PP_ALIGN.LEFT, left=left + Inches(numw) + gap,
                       width=Inches(descw))
        if caption:
            self._text(s, self.H * 0.85, self.H * 0.1, caption, max(MIN_BODY_PT, desc_size - 4))
        return s

    # ── 保存 ────────────────────────────────────────────
    def _texts(self, prs=None):
        """全枚の本文とノートを平たい文字列の並びにする（比較用）"""
        prs = prs or self.prs
        out = []
        for i, s in enumerate(prs.slides, 1):
            for sh in s.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip():
                    out.append('%d本文: %s' % (i, sh.text_frame.text.strip()))
            if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
                out.append('%dノート: %s' % (i, s.notes_slide.notes_text_frame.text.strip()))
        return out

    def save(self, path, force=False):
        """⚠**上書き前に、既にあるファイルとの食い違いを必ず調べる。**

        発端＝実際に起きた事故: 完成後、本人が PowerPoint で6箇所を直していた。
        スクリプト側は古い文言のままだったので、**ビルドし直せば手直しが全部消える**状態
        だった（気づいたのは偶然）。「スクリプトで作る」と「当日その場で自分で直せる」は
        正面衝突する。だから機構で止める。

        止まったときの正しい対応は次のどちらか:
          1. 手直しをスクリプト側へ書き戻してから、もう一度実行する（推奨・以後ズレない）
          2. 手直しを捨ててよいと確認できたときだけ save(..., force=True)
        ⚠ 迷ったら 1。**force=True は「消えても構わない」と確認できた時だけ。**
        """
        if os.path.exists(path) and not force:
            try:
                old = self._texts(Presentation(path))
            except Exception:
                old = None
            if old is not None:
                diff = [d for d in difflib.unified_diff(old, self._texts(), lineterm='', n=0)
                        if not d.startswith(('---', '+++', '@@'))]
                if diff:
                    print('■ 中断: %s は、これから書き込む内容と食い違っています。' % path)
                    print('  （誰かが PowerPoint で直した可能性があります。上書きすると消えます）')
                    for d in diff[:40]:
                        print('    ' + d.replace('\n', '⏎'))
                    if len(diff) > 40:
                        print('    …ほか %d 行' % (len(diff) - 40))
                    print('  → 手直しをスクリプトへ書き戻すか、消えてよいなら save(..., force=True)')
                    raise SystemExit(1)
        self.prs.save(path)
        n = len(self.prs.slides._sldIdLst)
        print('生成: %s（%d枚 ／ テーマ=%s）' % (path, n, self.theme_name))
        for v, idx in sorted(self.skips.items()):
            print('  %s: %s 枚目を飛ばす（残り %d 枚）'
                  % (v, ','.join(map(str, idx)), n - len(idx)))
        if self.warnings:
            print('  ⚠ 見積り警告 %d件（最終判断は必ず目視）:' % len(self.warnings))
            for w in self.warnings:
                print('    - ' + w)
        print('  ⚠ まだ完成ではない。verify.py を通して全枚を目で見ること。')
        return path
