"""Assemble exported layer SVGs + timeline plan into web-ready slide data files.

- swaps the 96-dpi rasters PowerPoint embeds in its SVG export for high-res crops of the originals
- namespaces SVG ids per layer so many layers can live in one document
- transcodes slide videos (crop / rotate / downscale) with ffmpeg
- writes static/anim/slides/slideNN.js  ->  PPTAnim.register('slideNN', {...})
"""
import base64, io, json, os, re, subprocess, sys
import numpy as np
from lxml import etree
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SCR = os.environ.get('PPTX2WEB_WORK', os.path.join(HERE, '_work'))
OUT = os.path.join(HERE, '..', '..', 'static', 'anim')
PPT = os.path.join(SCR, 'u', 'ppt')
EMU_PX = 9525            # EMU per px at 96 dpi (slide = 1280 x 720 px)
UPSCALE = 2.5            # raster density relative to 1280-px layout
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

XML = {}  # slide index -> slide XML file name, filled from plan.json

os.makedirs(os.path.join(OUT, 'slides'), exist_ok=True)
os.makedirs(os.path.join(OUT, 'media'), exist_ok=True)


def rels(n):
    t = etree.parse(os.path.join(PPT, 'slides', '_rels', XML[n] + '.rels'))
    return {r.get('Id'): r.get('Target') for r in t.getroot()}


_orig_cache = {}
def original(path):
    if path not in _orig_cache:
        _orig_cache[path] = Image.open(path).convert('RGBA')
    return _orig_cache[path]


def crop_src(im, src):
    """Apply a:srcRect (values in 1/1000 %) — negative values pad with transparency."""
    if src is None:
        return im
    W, H = im.size
    l, t, r, b = [int(src.get(k, 0)) / 100000 for k in 'ltrb']
    box = (round(l * W), round(t * H), round(W - r * W), round(H - b * H))
    if box[0] >= 0 and box[1] >= 0 and box[2] <= W and box[3] <= H:
        return im.crop(box)
    canvas = Image.new('RGBA', (box[2] - box[0], box[3] - box[1]), (0, 0, 0, 0))
    canvas.paste(im, (-box[0], -box[1]))
    return canvas


def slide_candidates(n):
    """High-res candidate images for every blip on the slide (pictures and picture fills)."""
    root = etree.parse(os.path.join(PPT, 'slides', XML[n])).getroot()
    rl = rels(n)
    cands = []
    for blipfill in root.iter(P + 'blipFill', A + 'blipFill'):
        blip = blipfill.find(A + 'blip')
        if blip is None or blip.get(R + 'embed') not in rl:
            continue
        pic = blipfill.getparent()
        if pic.find(f'.//{A}videoFile') is not None:
            continue
        path = os.path.normpath(os.path.join(PPT, 'slides', rl[blip.get(R + 'embed')]))
        if not path.lower().endswith(('.png', '.jpg', '.jpeg')):
            continue
        cands.append(crop_src(original(path), blipfill.find(A + 'srcRect')))
    return cands


def thumb(im, size=(24, 24)):
    im = im.convert('RGBA').resize(size, Image.BILINEAR)
    a = np.asarray(im).astype(np.float32) / 255
    return a[..., :3] * a[..., 3:4]  # premultiplied so transparent areas compare equal


def upgrade_images(svg, cands, stats):
    def repl(m):
        attrs, mime, data = m.group(1), m.group(2), m.group(3)
        w = float(re.search(r'width="([\d.]+)"', attrs).group(1))
        h = float(re.search(r'height="([\d.]+)"', attrs).group(1))
        try:
            low = Image.open(io.BytesIO(base64.b64decode(data)))
        except Exception:
            return m.group(0)
        tl = thumb(low)
        best, err = None, 1e9
        for c in cands:
            if abs(c.width / c.height - low.width / low.height) > 0.08 * (low.width / low.height):
                continue
            e = float(np.mean((thumb(c) - tl) ** 2))
            if e < err:
                best, err = c, e
        if best is None or err > 0.004:
            stats['kept'] += 1
            return m.group(0)
        tw = int(min(best.width, max(low.width, w) * UPSCALE))
        th = max(1, round(tw * best.height / best.width))
        if tw <= low.width * 1.15:
            stats['kept'] += 1
            return m.group(0)
        hi = best.resize((tw, th), Image.LANCZOS)
        buf = io.BytesIO()
        has_alpha = hi.getextrema()[3][0] < 255
        if mime == 'jpeg' or not has_alpha:
            hi.convert('RGB').save(buf, 'JPEG', quality=88, optimize=True)
            mime2 = 'jpeg'
        else:
            hi.save(buf, 'PNG', optimize=True)
            mime2 = 'png'
        # keep whichever is smaller if the "upgrade" isn't actually bigger in pixels
        stats['upgraded'] += 1
        return f'<image{attrs}xlink:href="data:image/{mime2};base64,{base64.b64encode(buf.getvalue()).decode()}"'
    return re.sub(r'<image([^>]*?)xlink:href="data:image/(png|jpeg);base64,([A-Za-z0-9+/=]+)"', repl, svg)


def slide_runs(n):
    """(text, px size) for every text run on the slide (+ chart text sizes) — used to undo the
    integer rounding PowerPoint applies to font-size in its SVG export (14pt = 18.67px -> "19")."""
    root = etree.parse(os.path.join(PPT, 'slides', XML[n])).getroot()
    runs, sizes = [], set()
    for run in root.iter(A + 'r'):
        rp = run.find(A + 'rPr')
        if rp is not None and rp.get('sz'):
            px = int(rp.get('sz')) / 100 * 4 / 3
            runs.append((run.findtext(A + 't') or '', px))
            sizes.add(px)
    for tgt in rels(n).values():
        if 'charts/' in tgt:
            ch = etree.parse(os.path.normpath(os.path.join(PPT, 'slides', tgt))).getroot()
            sizes |= {int(e.get('sz')) / 100 * 4 / 3 for e in ch.iter(A + 'defRPr', A + 'rPr') if e.get('sz')}
    return runs, sizes


def fix_font_sizes(svg, runs, sizes, stats):
    import html

    def repl(m):
        attrs, content = m.group(1), m.group(2)
        fm = re.search(r'font-size="([\d.]+)"', attrs)
        if not fm:
            return m.group(0)
        N = float(fm.group(1))
        txt = html.unescape(content).strip()
        cand = {round(px, 3) for t, px in runs if txt and txt in t and abs(px - N) <= 0.5 + 1e-6}
        if not cand:
            cand = {round(px, 3) for px in sizes if abs(px - N) <= 0.5 + 1e-6}
        if not cand:
            stats['font_kept'] += 1
            return m.group(0)
        px = min(cand, key=lambda c: abs(c - N))
        stats['font_fixed'] += 1
        return '<text' + attrs.replace(fm.group(0), f'font-size="{px:g}"') + '>' + content + '</text>'
    return re.sub(r'<text([^>]*)>([^<]*)</text>', repl, svg)


ANCHOR_RE = re.compile(r'<rect x="(?:0|1279)" y="(?:0|719)" width="1" height="1" fill="#FFFFFF" fill-opacity="0"/>')


def layer_inner(path, prefix):
    raw = open(path, encoding='utf-8').read()
    root = etree.fromstring(raw.encode('utf-8'))
    assert root.get('width') == '1282' and root.get('height') == '722', (path, root.get('width'))
    parts = []
    for ch in root:
        name = etree.QName(ch).localname
        if name == 'g' and (ch.get('transform') or '').startswith('translate('):
            # outer translate maps absolute slide px -> local export box; drop it to stay absolute
            parts.extend(etree.tostring(c, encoding='unicode') for c in ch)
        else:
            parts.append(etree.tostring(ch, encoding='unicode'))
    s = ''.join(parts)
    s = ANCHOR_RE.sub('', s)
    s = re.sub(r'\s+xmlns(:\w+)?="[^"]*"', '', s)
    s = s.replace(',Roboto_MSFontService', '').replace(',Arial_MSFontService', '').replace(',Cambria Math_MSFontService', '')
    s = re.sub(r'\bid="([^"]+)"', lambda m: f'id="{prefix}{m.group(1)}"', s)
    s = re.sub(r'url\(#([^)]+)\)', lambda m: f'url(#{prefix}{m.group(1)})', s)
    s = re.sub(r'href="#([^"]+)"', lambda m: f'href="#{prefix}{m.group(1)}"', s)
    return s


def video_for(n, sid):
    """Transcode the video behind picture `sid` so it can be shown as a plain axis-aligned <video>."""
    root = etree.parse(os.path.join(PPT, 'slides', XML[n])).getroot()
    rl = rels(n)
    for pic in root.iter(P + 'pic'):
        if pic.find(f'{P}nvPicPr/{P}cNvPr').get('id') != sid:
            continue
        vf = pic.find(f'.//{A}videoFile')
        src = os.path.normpath(os.path.join(PPT, 'slides', rl[vf.get(R + 'link')]))
        xfrm = pic.find(f'{P}spPr/{A}xfrm')
        off, ext = xfrm.find(A + 'off'), xfrm.find(A + 'ext')
        x, y = int(off.get('x')) / EMU_PX, int(off.get('y')) / EMU_PX
        w, h = int(ext.get('cx')) / EMU_PX, int(ext.get('cy')) / EMU_PX
        rot = int(xfrm.get('rot', 0)) / 60000 % 360
        filters = []
        sr = pic.find(f'{P}blipFill/{A}srcRect')
        if sr is not None:
            l, t, r, b = [int(sr.get(k, 0)) / 100000 for k in 'ltrb']
            filters.append(f'crop=iw*{1-l-r:.5f}:ih*{1-t-b:.5f}:iw*{l:.5f}:ih*{t:.5f}')
        cx, cy = x + w / 2, y + h / 2
        if rot in (90, 270):
            filters.append('transpose=1' if rot == 90 else 'transpose=2')
            w, h = h, w
        elif rot:
            print('  !! unsupported video rotation', rot)
        x, y = cx - w / 2, cy - h / 2
        tw = int(min(1600, w * UPSCALE)) // 2 * 2
        filters.append(f'scale={tw}:-2')
        name = f'slide{n:02d}_{sid}.mp4'
        dst = os.path.join(OUT, 'media', name)
        if not os.path.exists(dst):
            th = tw * h / w
            kbps = int(tw * th / 1e6 * 2200) + 300
            # this ffmpeg build lacks libx264; Windows MediaFoundation H.264 is universally playable
            subprocess.run(['ffmpeg', '-y', '-v', 'error', '-i', src, '-vf', ','.join(filters), '-an',
                            '-c:v', 'h264_mf', '-b:v', f'{kbps}k', '-pix_fmt', 'nv12',
                            '-movflags', '+faststart', dst], check=True)
        # poster = first frame, for the not-yet-playing state
        poster = name.replace('.mp4', '.jpg')
        pdst = os.path.join(OUT, 'media', poster)
        if not os.path.exists(pdst):
            subprocess.run(['ffmpeg', '-y', '-v', 'error', '-i', dst, '-frames:v', '1', '-q:v', '4', pdst], check=True)
        return {'video': 'media/' + name, 'poster': 'media/' + poster,
                'box': [round(x, 2), round(y, 2), round(w, 2), round(h, 2)]}
    raise KeyError(sid)


def main():
    plan = json.load(open(os.path.join(SCR, 'plan.json')))
    titles = {}
    for n_str, sp in plan.items():
        n = int(n_str)
        XML[n] = sp['xml']
        cands = slide_candidates(n)
        runs, sizes = slide_runs(n)
        stats = {'upgraded': 0, 'kept': 0, 'font_fixed': 0, 'font_kept': 0}
        layers = []
        for k, L in enumerate(sp['layers']):
            ev = L['events']
            for e in ev:
                if e['k'] == 'play':
                    e['v'] = 0
            item = {'ev': ev} if ev else {}
            if L['hidden']:
                item['hidden'] = 1
            if any(e['k'] == 'rot' for e in ev):
                item['c'] = L['center']
            if L['kind'] == 'video':
                item.update(video_for(n, L['ids'][0]))
            else:
                svg = layer_inner(os.path.join(SCR, 'exp', f's{n}_{k}.svg'), f's{n}l{k}_')
                svg = fix_font_sizes(svg, runs, sizes, stats)
                item['svg'] = upgrade_images(svg, cands, stats)
            layers.append(item)
        data = {'w': 1280, 'h': 720, 'duration': sp['duration'], 'layers': layers}
        sid = f'slide{n:02d}'
        js = f'/* generated by tools/pptx2web from slide {n} */\nPPTAnim.register({json.dumps(sid)}, {json.dumps(data, separators=(",", ":"))});\n'
        open(os.path.join(OUT, 'slides', sid + '.js'), 'w', encoding='utf-8').write(js)
        print(f'{sid}: {len(layers)} layers, {len(js)/1024:.0f} KB, images upgraded={stats["upgraded"]} kept={stats["kept"]}, '
              f'font sizes fixed={stats["font_fixed"]} kept={stats["font_kept"]}')


if __name__ == '__main__':
    main()
