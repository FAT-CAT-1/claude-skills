#!/usr/bin/env python3
"""
homepage スキル — 出荷前の実測検査 v2.0

    python verify.py path/to/index.html

やること:
  1. 設計思想（テーゼ）がファイルに固定されているかを検査する
  2. 数値（級数・余白・画像点数・書体分離・濃色帯・横はみ出し）を実測する
  3. 同一パターンの反復（few-shot rut）を検出する
  4. 目視用の画像を自動で書き出す ← 撮るだけでは終わらない。必ず読むこと

⚠ この検査が全部緑でも完成ではない。最後に書き出される画像を必ず自分の目で見ること。
   実測比較の結果、機械検査が全項目PASSした後に見出しの折り返し崩れが見つかった実績がある。
"""

import http.server, io, json, os, re, shutil, socketserver, subprocess, sys, threading

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REQ = {"h1_min": 56, "h2_min": 40, "pad_min": 100, "img_min": 1, "neg_rules_min": 2}
WIDTHS = [390, 768, 1280]

PROBE = """<!doctype html><meta charset="utf-8"><title>probe</title>
<style>html,body{margin:0}iframe{border:0;display:block}</style>
<iframe id="f" src="TARGET" style="width:WIDTHpx;height:2400px"></iframe>
<pre id="out">PENDING</pre>
<script>
function lum(c){var m=/(\\d+),\\s*(\\d+),\\s*(\\d+)/.exec(c||'');if(!m)return 1;
  var v=[m[1],m[2],m[3]].map(function(x){x=x/255;return x<=.03928?x/12.92:Math.pow((x+.055)/1.055,2.4)});
  return .2126*v[0]+.7152*v[1]+.0722*v[2];}
window.addEventListener('load',function(){setTimeout(function(){
  var d=document.getElementById('f').contentDocument, w=document.getElementById('f').contentWindow;
  var vw=d.documentElement.clientWidth, r={vw:vw, scrollW:d.documentElement.scrollWidth};
  // ⚠ 最初の1つではなく最大を測る。小さい帯の見出しを拾うと誤判定する（実際に踏んだ）
  function px(sel){var m=0;d.querySelectorAll(sel).forEach(function(e){
    m=Math.max(m,Math.round(parseFloat(w.getComputedStyle(e).fontSize))||0);});return m;}
  r.h1=px('h1'); r.h2=px('h2');
  var h1e=d.querySelector('h1'), be=d.body;
  r.h1font=h1e?w.getComputedStyle(h1e).fontFamily.split(',')[0].replace(/["']/g,''):'';
  r.bodyfont=be?w.getComputedStyle(be).fontFamily.split(',')[0].replace(/["']/g,''):'';
  var pads=[];
  d.querySelectorAll('section,.section,main>div,article').forEach(function(e){
    var p=parseFloat(w.getComputedStyle(e).paddingTop)||0; if(p>0)pads.push(Math.round(p));});
  r.pad=pads.length?Math.max.apply(null,pads):0;
  r.img=d.querySelectorAll('img,svg image,picture').length;
  r.h2texts=[].slice.call(d.querySelectorAll('h2')).map(function(e){
    return (e.textContent||'').replace(/\\s+/g,' ').trim().slice(0,42);});
  function isDark(e){return e&&e.nodeType===1&&lum(w.getComputedStyle(e).backgroundColor)<0.35;}
  var dark=0;
  d.querySelectorAll('*').forEach(function(e){var b=e.getBoundingClientRect();
    if(b.width>=vw*0.9&&b.height>200&&isDark(e)&&!isDark(e.parentElement))dark++;});
  r.darkband=dark;
  // few-shot rut 検出: 主要ブロックの構造署名を集め、重複の度合いを見る
  var sigs=[];
  d.querySelectorAll('main > section, main > div').forEach(function(e){
    var st=w.getComputedStyle(e);
    var kids=[].slice.call(e.querySelectorAll(':scope > *, :scope > * > *')).slice(0,8)
      .map(function(k){return k.tagName;}).join('-');
    var g=w.getComputedStyle(e.querySelector('[class*=grid],[style*=grid]')||e).gridTemplateColumns||'';
    sigs.push(kids+'|'+(st.backgroundColor||'')+'|'+g.split(' ').length);});
  r.sigs=sigs;
  var over=[];
  d.querySelectorAll('*').forEach(function(e){var b=e.getBoundingClientRect();
    if(b.right>vw+2&&w.getComputedStyle(e).overflowX!=='auto'&&w.getComputedStyle(e.parentElement||d.body).overflowX!=='auto')
      over.push((e.tagName+'.'+(typeof e.className==='string'?e.className:'')).slice(0,38));});
  r.overflow=over.slice(0,5);
  document.getElementById('out').textContent='@@'+JSON.stringify(r)+'@@';
},2600);});
</script>"""


def find_chrome():
    for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
              "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
              "google-chrome", "chromium"]:
        if os.path.exists(p) or shutil.which(p):
            return p
    sys.exit("Chrome が見つかりません。パスを find_chrome() に足してください。")


def serve(directory):
    class Quiet(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *a):
            pass
    handler = lambda *a, **k: Quiet(*a, directory=directory, **k)
    httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd, httpd.server_address[1]


def chrome_run(chrome, args, timeout=140):
    return subprocess.run([chrome, "--headless=new", "--disable-gpu"] + args,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=timeout).stdout


def measure(chrome, port, target, width, workdir):
    name = f"_probe_{width}.html"
    with open(os.path.join(workdir, name), "w", encoding="utf-8") as f:
        f.write(PROBE.replace("TARGET", target).replace("WIDTH", str(width)))
    out = chrome_run(chrome, ["--virtual-time-budget=9000", "--dump-dom",
                              f"http://127.0.0.1:{port}/{name}"])
    os.remove(os.path.join(workdir, name))
    m = re.search(r"@@(\{.*?\})@@", out, re.S)
    return json.loads(m.group(1)) if m else None


# ---------- ① 設計思想（テーゼ）の検査 ----------
NEG_PATTERNS = [
    r"使わない", r"並べない", r"しない", r"置かない", r"書かない", r"引かない",
    r"超えない", r"限る", r"\bno\s+\w", r"\bnever\b", r"\bavoid\b", r"don't", r"do not",
]


def check_thesis(path):
    """HTML/CSS の先頭に Design: テーゼと否定制約があるか。"""
    head = io.open(path, encoding="utf-8", errors="replace").read(6000)
    css = ""
    m = re.search(r'href=["\']([^"\']+\.css)["\']', head)
    if m:
        cp = os.path.join(os.path.dirname(path), m.group(1))
        if os.path.exists(cp):
            css = io.open(cp, encoding="utf-8", errors="replace").read(3000)
    blob = head + "\n" + css
    has_design = bool(re.search(r"Design\s*[:：]", blob))
    name = re.search(r"Design\s*[:：]\s*[「\"']([^」\"'\n]{2,40})", blob)
    # ⚠ 否定制約は改行をまたいで書かれる。行頭の * を除いて連結してから数える
    rules = re.search(r"(Visual rules|視覚の規則)\s*[:：]([\s\S]{0,600}?)"
                      r"(?=Information rule|情報の規則|-{5}|\*/|$)", blob)
    neg = 0
    if rules:
        seg = re.sub(r"\n\s*\*?", " ", rules.group(2))
        for part in re.split(r"[;；。]", seg):
            if part.strip() and any(re.search(p, part, re.I) for p in NEG_PATTERNS):
                neg += 1
    info = bool(re.search(r"(Information rule|情報の規則)\s*[:：]", blob))
    # テンプレートの置換漏れ（ファイル全体を見る）
    whole = io.open(path, encoding="utf-8", errors="replace").read()
    holes = len(re.findall(r"\{\{[^}]{0,40}\}\}", whole))
    return {"design": has_design, "name": name.group(1) if name else "",
            "neg": neg, "info": info, "holes": holes}


# ---------- ② 目視用の画像を書き出す ----------
def shoot(chrome, port, target, outdir):
    os.makedirs(outdir, exist_ok=True)
    made = []
    full = os.path.join(outdir, "review-desktop.png")
    chrome_run(chrome, ["--hide-scrollbars", "--virtual-time-budget=12000",
                        "--window-size=1280,6600", f"--screenshot={full}",
                        f"http://127.0.0.1:{port}/{target}"], timeout=200)
    if os.path.exists(full):
        made.append(full)
    probe = os.path.join(os.path.dirname(os.path.join(outdir, "x")), "..")
    return made


def make_thumb(src, dst, scale=4):
    try:
        from PIL import Image
    except ImportError:
        return None
    im = Image.open(src)
    w, h = im.size
    im.resize((max(1, w // scale), max(1, h // scale))).save(dst)
    return dst


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    path = os.path.abspath(sys.argv[1])
    if not os.path.exists(path):
        sys.exit(f"見つかりません: {path}")
    workdir, target = os.path.dirname(path), os.path.basename(path)
    outdir = os.path.join(workdir, "_review")
    chrome = find_chrome()
    httpd, port = serve(workdir)

    fails, warns, results = [], [], {}
    try:
        for w in WIDTHS:
            r = measure(chrome, port, target, w, workdir)
            if not r:
                sys.exit(f"計測に失敗しました（幅{w}）。Chrome が起動できているか確認してください。")
            results[w] = r
        shots = shoot(chrome, port, target, outdir)
    finally:
        httpd.shutdown()

    d = results[1280]
    th = check_thesis(path)

    def check(ok, label, got, want):
        print(f"  [{'PASS' if ok else 'FAIL'}] {label:<28} {got:>13}   基準 {want}")
        if not ok:
            fails.append(label)

    print("=" * 66)
    print(f"  実測: {target}")
    print("=" * 66)
    print("  ■ 設計思想（reference/DESIGN_THESIS.md）")
    check(th["design"], "テーゼがファイルにある", "あり" if th["design"] else "なし", "先頭コメントに Design:")
    check(bool(th["name"]), "名前のある概念", th["name"] or "なし", "「◯◯」の形")
    check(th["neg"] >= REQ["neg_rules_min"], "否定制約の数", f'{th["neg"]}件', f'≥{REQ["neg_rules_min"]}件')
    check(th["holes"] == 0, "テンプレの置換漏れ", f'{th["holes"]}箇所', "0箇所")
    if not th["info"]:
        warns.append("Information rule が無い（情報の置き方の規則。SHOULD）")

    print("-" * 66)
    print("  ■ 数値")
    check(d["h1"] >= REQ["h1_min"], "h1 実描画 (1280px)", f'{d["h1"]}px', f'≥{REQ["h1_min"]}px')
    check(d["h2"] >= REQ["h2_min"], "h2 実描画 (1280px)", f'{d["h2"]}px', f'≥{REQ["h2_min"]}px')
    check(d["pad"] >= REQ["pad_min"], "セクション上下 padding", f'{d["pad"]}px', f'≥{REQ["pad_min"]}px')
    check(d["img"] >= REQ["img_min"], "画像点数", f'{d["img"]}点', f'≥{REQ["img_min"]}点')
    check(bool(d["h1font"]) and d["h1font"] != d["bodyfont"], "見出しと本文の書体分離",
          f'{d["h1font"]}'[:13], "異なること")
    check(d["darkband"] >= 1, "濃色の全幅帯", f'{d["darkband"]}箇所', "≥1箇所")

    # few-shot rut
    sigs = d.get("sigs") or []
    uniq = len(set(sigs))
    ratio = (uniq / len(sigs)) if sigs else 1
    check(len(sigs) < 3 or ratio >= 0.5, "構造の多様性(反復検出)",
          f'{uniq}/{len(sigs)}種', "半数以上が別構造")

    print("-" * 66)
    for w in WIDTHS:
        r = results[w]
        ok = r["scrollW"] <= r["vw"] + 1
        check(ok, f"横はみ出しなし ({w}px)", f'{r["scrollW"]}px', f'≤{r["vw"]}px')
        if not ok and r["overflow"]:
            print(f"         → 犯人候補: {', '.join(r['overflow'])}")

    if results[390]["h1"] < 40:
        warns.append(f'モバイル h1 が {results[390]["h1"]}px（40px以上が望ましい）')

    print("-" * 66)
    print("  ■ 見出しの拾い読み（機能ラベルが並んでいないか）")
    for t in d["h2texts"]:
        print(f"      ・{t}")
    ng = [t for t in d["h2texts"] if t in
          ("基本情報", "ご利用案内", "アクセス", "支援の内容", "お問い合わせ",
           "サービス", "会社概要", "料金", "特徴", "メニュー")]
    if len(ng) >= 2:
        warns.append(f"見出しが機能ラベル中心です: {' / '.join(ng)}")

    if warns:
        print("-" * 66)
        for w_ in warns:
            print(f"  [WARN] {w_}")

    # ---------- 目視用の画像 ----------
    print("=" * 66)
    if shots:
        thumb = make_thumb(shots[0], os.path.join(outdir, "review-25percent.png"), 4)
        print("  ■ 目視用の画像を書き出しました。⚠ 撮っただけでは検査になっていません。")
        print(f"      全体      : {shots[0]}")
        if thumb:
            print(f"      25%縮小   : {thumb}")
        print("      → この2枚を必ず読み、SKILL.md §5 の4項目を目で確かめること。")
    else:
        print("  [WARN] スクリーンショットを書き出せませんでした。手動で描画して確認すること。")

    print("=" * 66)
    if fails:
        print(f"  ✗ {len(fails)} 件が基準未達: {', '.join(fails)}")
        print("  → SKILL.md §3 の該当要件を読み直して直す。")
        sys.exit(1)
    print("  ✓ 数値はすべて基準内。ただしこれは下限の確認にすぎない。")
    print("  ⚠ 上の画像を目で見て、良し悪しを1行で言語化するまで完成ではない。")


if __name__ == "__main__":
    main()
