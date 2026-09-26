"""Parse slide XML timing trees into flat per-shape event lists and plan export layers.

Output: plan.json = {slideN: {layers:[{indices, ids, events, hidden, kind, center}], duration}}
"""
import json, os, re, sys, math
from lxml import etree

WORK = os.environ.get('PPTX2WEB_WORK', os.path.join(os.path.dirname(os.path.abspath(__file__)), '_work'))
U = os.path.join(WORK, 'u', 'ppt', 'slides')
PNS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P = '{%s}' % PNS
CLICK_GAP = 800  # ms pause used in place of "wait for click"

def ln(e):
    return etree.QName(e).localname if isinstance(e.tag, str) else None

def parse_timeline(root):
    """Return (events_by_target, first_class_by_target, total_ms)."""
    tm = root.find(f'.//{P}timing')
    if tm is None:
        return {}, {}, 0
    begins, ends = {}, {}
    events = []   # (spid, kind, begin, dur, data)
    first_cls = {}

    def ctn_of(node):
        return node.find(f'{P}cTn') if ln(node) in ('par', 'seq') else None

    def resolve_begin(ctn, parent_begin, prev_end, in_seq):
        best = None
        conds = ctn.findall(f'{P}stCondLst/{P}cond')
        if not conds:
            return prev_end if in_seq else parent_begin
        for c in conds:
            d = c.get('delay', '0')
            evt = c.get('evt')
            tn = c.find(f'{P}tn')
            if evt in ('onBegin', 'onEnd') and tn is not None:
                ref = int(tn.get('val'))
                base = begins.get(ref) if evt == 'onBegin' else ends.get(ref)
                if base is None:
                    continue
                t = base + (0 if d == 'indefinite' else int(d))
            elif d == 'indefinite':
                continue
            else:
                t = (prev_end if in_seq else parent_begin) + int(d)
            best = t if best is None else min(best, t)
        if best is None:  # pure click trigger -> auto advance after a short pause
            best = (prev_end if in_seq else parent_begin) + CLICK_GAP
        return best

    def walk_container(node, parent_begin, prev_end, in_seq):
        ctn = ctn_of(node)
        b = resolve_begin(ctn, parent_begin, prev_end, in_seq)
        cid = int(ctn.get('id'))
        begins[cid] = b
        cls = ctn.get('presetClass')
        end = b
        child_seq = ln(node) == 'seq'
        prev = b
        kids = ctn.find(f'{P}childTnLst')
        if kids is not None:
            for ch in kids:
                k = ln(ch)
                if k in ('par', 'seq'):
                    e = walk_container(ch, b, prev, child_seq)
                elif k in ('set', 'animEffect', 'animMotion', 'animRot', 'anim', 'animScale', 'cmd', 'animClr'):
                    e = walk_behavior(ch, b, cls)
                else:
                    continue
                prev = e
                end = max(end, e)
        dur = ctn.get('dur')
        if dur and dur.isdigit():
            end = b + int(dur)
        ends[cid] = end
        return end

    def walk_behavior(el, parent_begin, cls):
        k = ln(el)
        cb = el.find(f'{P}cBhvr')
        ctn = cb.find(f'{P}cTn')
        b = resolve_begin(ctn, parent_begin, parent_begin, False)
        dur = ctn.get('dur', '1')
        dur = int(dur) if dur.isdigit() else 0
        spid = cb.find(f'{P}tgtEl/{P}spTgt').get('spid')
        if cls and spid not in first_cls:
            first_cls[spid] = (b, cls)
        elif cls and b < first_cls[spid][0]:
            first_cls[spid] = (b, cls)
        if k == 'set':
            attr = cb.findtext(f'{P}attrNameLst/{P}attrName')
            val = el.find(f'{P}to/{P}strVal').get('val')
            if attr == 'style.visibility':
                events.append((spid, 'vis', b, 0, val == 'visible'))
        elif k == 'animEffect':
            events.append((spid, 'fade', b, dur, el.get('transition', 'in') == 'in'))
        elif k == 'animMotion':
            events.append((spid, 'move', b, dur, sample_path(el.get('path'))))
        elif k == 'animRot':
            events.append((spid, 'rot', b, dur, int(el.get('by')) / 60000.0))
        elif k == 'cmd':
            events.append((spid, 'play', b, dur, el.get('cmd')))
        else:
            print('  !! unsupported behavior', k, file=sys.stderr)
        cid = int(ctn.get('id'))
        begins[cid] = b
        ends[cid] = b + dur
        return b + dur

    root_par = tm.find(f'{P}tnLst/{P}par')
    total = walk_container(root_par, 0, 0, False)
    # root/mainSeq are 'indefinite'; recompute total as last event end
    total = max([b + d for (_, _, b, d, _) in events] or [0])
    by = {}
    for spid, kind, b, d, data in events:
        by.setdefault(spid, []).append({'k': kind, 't': b, 'd': d, 'v': data})
    for v in by.values():
        v.sort(key=lambda e: e['t'])
    return by, {k: v[1] for k, v in first_cls.items()}, total


def sample_path(path):
    """PowerPoint motion path (fractions of slide W/H) -> polyline sampled by arc length."""
    toks = re.findall(r'[MLCZEmlcze]|-?[\d.]+(?:e-?\d+)?', path)
    pts, cur, i, cmd = [], (0.0, 0.0), 0, None
    while i < len(toks):
        t = toks[i]
        if t.isalpha():
            cmd = t.upper(); i += 1
            if cmd in ('E', 'Z'):
                continue
        nums = lambda n: [float(x) for x in toks[i:i + n]]
        if cmd in ('M', 'L'):
            x, y = nums(2); i += 2
            cur = (x, y); pts.append(cur)
        elif cmd == 'C':
            x1, y1, x2, y2, x, y = nums(6); i += 6
            p0 = cur
            for s in range(1, 13):
                u = s / 12
                bx = (1-u)**3*p0[0] + 3*(1-u)**2*u*x1 + 3*(1-u)*u*u*x2 + u**3*x
                by_ = (1-u)**3*p0[1] + 3*(1-u)**2*u*y1 + 3*(1-u)*u*u*y2 + u**3*y
                pts.append((bx, by_))
            cur = (x, y)
        else:
            i += 1
    # resample to <=120 points uniformly by arc length
    if len(pts) < 2:
        return [[0, 0], [0, 0]]
    acc = [0.0]
    for a, b in zip(pts, pts[1:]):
        acc.append(acc[-1] + math.hypot(b[0]-a[0], b[1]-a[1]))
    L = acc[-1] or 1
    n = min(120, max(2, len(pts)))
    out, j = [], 0
    for s in range(n):
        target = L * s / (n - 1)
        while j < len(acc) - 2 and acc[j+1] < target:
            j += 1
        seg = acc[j+1] - acc[j] or 1
        u = (target - acc[j]) / seg
        a, b = pts[j], pts[j+1]
        out.append([round(a[0] + (b[0]-a[0])*u, 6), round(a[1] + (b[1]-a[1])*u, 6)])
    return out


def top_level_ids(root):
    tree = root.find(f'{P}cSld/{P}spTree')
    ids = []
    for ch in tree:
        k = ln(ch)
        if k in ('nvGrpSpPr', 'grpSpPr', 'extLst', None):
            continue
        if k == 'AlternateContent':
            ch = ch.find('{%s}Choice' % MC)[0]
            k = ln(ch)
        cnv = ch.find('.//{%s}cNvPr' % PNS)
        is_video = ch.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}videoFile') is not None
        ids.append((cnv.get('id'), is_video, cnv.get('name')))
    return ids


def slide_files():
    """Slide XML file names in presentation order (file numbers need not match slide order)."""
    ppt = os.path.dirname(U)
    pres = etree.parse(os.path.join(ppt, 'presentation.xml')).getroot()
    rel = etree.parse(os.path.join(ppt, '_rels', 'presentation.xml.rels')).getroot()
    target = {r.get('Id'): r.get('Target') for r in rel}
    rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    return [os.path.basename(target[s.get(rid)]) for s in pres.find(f'{P}sldIdLst')]


def main():
    shapes = json.load(open(os.path.join(WORK, 'shapes.json'), encoding='utf-8-sig'))
    plan = {}
    for n, xml in enumerate(slide_files(), 1):
        root = etree.parse(os.path.join(U, xml)).getroot()
        events, first_cls, total = parse_timeline(root)
        com = [s for s in shapes if s['slide'] == n]
        tl = top_level_ids(root)
        assert len(tl) == len(com), (n, len(tl), len(com))
        layers = []
        for (sid, is_video, name), c in zip(tl, com):
            assert str(c['id']) == sid, (n, sid, c['id'])
            if not c['visible']:
                continue
            ev = events.get(sid, [])
            hidden = first_cls.get(sid) == 'entr'
            unique = is_video or any(e['k'] in ('move', 'rot') for e in ev)
            sig = json.dumps([ev, hidden])
            kind = 'video' if is_video else 'svg'
            if layers and not unique and not layers[-1]['unique'] and layers[-1]['sig'] == sig:
                layers[-1]['indices'].append(c['idx']); layers[-1]['ids'].append(sid)
                continue
            px = 4 / 3  # points -> px (96 dpi)
            layers.append({
                'indices': [c['idx']], 'ids': [sid], 'names': [name], 'sig': sig, 'unique': unique,
                'events': ev, 'hidden': hidden, 'kind': kind,
                'center': [round((c['left'] + c['width'] / 2) * px, 2), round((c['top'] + c['height'] / 2) * px, 2)],
                'box': [round(c['left'] * px, 2), round(c['top'] * px, 2), round(c['width'] * px, 2), round(c['height'] * px, 2), c['rot']],
            })
        for L in layers:
            del L['sig'], L['unique']
        plan[n] = {'xml': xml, 'layers': layers, 'duration': total}
        print(f'slide{n}: {len(com)} shapes -> {len(layers)} layers, animated targets={len(events)}, duration={total}ms')
    json.dump(plan, open(os.path.join(WORK, 'plan.json'), 'w'), indent=1)


if __name__ == '__main__':
    main()
