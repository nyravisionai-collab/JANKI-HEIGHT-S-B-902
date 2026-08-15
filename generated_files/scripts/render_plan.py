import math
import ezdxf
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle as MplCircle, Arc as MplArc, Polygon, Rectangle
from matplotlib.lines import Line2D

doc = ezdxf.readfile('plan.dxf')
msp = doc.modelspace()

X0, X1, Y0, Y1 = 50350, 53220, -7930, -6520

# ---------- layer style map ----------
def style(e):
    lay = e.dxf.layer.upper()
    t = e.dxftype()
    if lay in ('ELETRICAL',):
        return dict(color='#C00000', lw=1.6, alpha=1.0)
    if lay in ('ELECTRIACL',):
        return dict(color='#0047AB', lw=1.6, alpha=1.0)
    if lay in ('ELE. TEXT',):
        return dict(color='#555555', lw=0.8, alpha=1.0)
    if lay in ('WALL',):
        return dict(color='#222222', lw=1.4, alpha=1.0)
    if lay in ('WINDOW', 'd-w', '01_AR_door', 'door'):
        return dict(color='#3FA7D6', lw=0.9, alpha=0.9)
    if lay in ('TEXT', 'TEX'):
        return dict(color='#444444', lw=0.6, alpha=1.0)
    if t in ('TEXT', 'MTEXT'):
        return dict(color='#444444', lw=0.6, alpha=1.0)
    return dict(color='#B0B0B0', lw=0.5, alpha=0.55)

def in_region(x, y, m=20):
    return X0 - m <= x <= X1 + m and Y0 - m <= y <= Y1 + m

def draw_entity(ax, e, transform=None):
    """transform: (sx, sy, rot_rad, tx, ty) or None"""
    st = style(e)
    t = e.dxftype()
    def T(x, y):
        if transform is None:
            return x, y
        sx, sy, r, tx, ty = transform
        x, y = x * sx, y * sy
        if r:
            x2 = x * math.cos(r) - y * math.sin(r)
            y2 = x * math.sin(r) + y * math.cos(r)
            x, y = x2, y2
        return x + tx, y + ty
    if t == 'LINE':
        s, en = e.dxf.start, e.dxf.end
        p1, p2 = T(s.x, s.y), T(en.x, en.y)
        if in_region(*p1) or in_region(*p2):
            ax.add_line(Line2D([p1[0], p2[0]], [p1[1], p2[1]], **st))
    elif t == 'LWPOLYLINE':
        pts = [(p[0], p[1]) for p in e.get_points()]
        pts = [T(*p) for p in pts]
        if not any(in_region(*p) for p in pts):
            return
        if e.closed and len(pts) > 2:
            ax.add_patch(Polygon(pts, closed=True, fill=False, **st))
        else:
            for a, b in zip(pts, pts[1:]):
                ax.add_line(Line2D([a[0], b[0]], [a[1], b[1]], **st))
    elif t == 'CIRCLE':
        c = e.dxf.center
        x, y = T(c.x, c.y)
        r = e.dxf.radius * (transform[0] if transform else 1)
        if in_region(x, y, r):
            ax.add_patch(MplCircle((x, y), r, fill=False, **st))
    elif t == 'ARC':
        c = e.dxf.center
        x, y = T(c.x, c.y)
        r = e.dxf.radius * (transform[0] if transform else 1)
        a0, a1 = e.dxf.start_angle, e.dxf.end_angle
        if transform and transform[2]:
            a0 += math.degrees(transform[2]); a1 += math.degrees(transform[2])
        if in_region(x, y, r):
            ax.add_patch(MplArc((x, y), 2 * r, 2 * r, theta1=a0, theta2=a1, **st))
    elif t == 'ELLIPSE':
        c = e.dxf.center
        x, y = T(c.x, c.y)
        maj = math.hypot(e.dxf.major_axis[0], e.dxf.major_axis[1])
        ratio = e.dxf.ratio
        r = maj * (transform[0] if transform else 1)
        if in_region(x, y, r):
            ang = math.degrees(math.atan2(e.dxf.major_axis[1], e.dxf.major_axis[0]))
            ax.add_patch(MplArc((x, y), 2 * r, 2 * r * ratio, angle=ang, theta1=0, theta2=360, **st))
    elif t == 'SPLINE':
        try:
            pts = [(p[0], p[1]) for p in e.control_points]
        except Exception:
            return
        pts = [T(*p) for p in pts]
        if not any(in_region(*p) for p in pts):
            return
        for a, b in zip(pts, pts[1:]):
            ax.add_line(Line2D([a[0], b[0]], [a[1], b[1]], **st))
    elif t == 'INSERT':
        p = e.dxf.insert
        sx, sy = e.dxf.xscale, e.dxf.yscale
        r = math.radians(e.dxf.rotation)
        if transform:
            p = T(p.x, p.y)
            # crude: keep nested transform simple
            sx *= transform[0]; sy *= transform[1]
            r += transform[2]
        else:
            p = (p.x, p.y)
        try:
            blk = doc.blocks.get(e.dxf.name)
        except Exception:
            return
        for sub in blk:
            draw_entity(ax, sub, (sx, sy, r, p[0], p[1]))
    elif t in ('TEXT', 'MTEXT'):
        p = e.dxf.insert
        x, y = T(p.x, p.y)
        if not in_region(x, y):
            return
        txt = e.plain_text() if t == 'MTEXT' else e.dxf.text
        h = e.dxf.char_height if e.dxf.hasattr('char_height') else e.dxf.get('height', 2.5)
        txt = txt.replace('\\P', '\n').replace('{', '').replace('}', '')
        # skip heavy formatting strings
        txt = ''.join(ch for ch in txt if ord(ch) >= 32 or ch == '\n')
        if txt.strip():
            ax.text(x, y, txt, fontsize=max(h * 0.55, 3.0), color='#333333', va='bottom', ha='left', clip_on=True)

fig, ax = plt.subplots(figsize=(26, 13.5), dpi=115)
for e in msp:
    if e.dxftype() in ('LEADER', 'DIMENSION', 'HATCH', 'WIPEOUT', 'DIMENSION'):
        continue
    if e.dxf.layer.upper() == 'WIPEOUT':
        continue
    try:
        draw_entity(ax, e)
    except Exception:
        pass

ax.set_xlim(X0, X1)
ax.set_ylim(Y0, Y1)
ax.set_aspect('equal')
ax.set_facecolor('white')
fig.patch.set_facecolor('white')
ax.axis('off')
fig.savefig('/tmp/dwgconv/electrical_plan_overview.png', facecolor='white', bbox_inches='tight', pad_inches=0.15)
plt.close(fig)

from PIL import Image
import numpy as np
im = np.array(Image.open('/tmp/dwgconv/electrical_plan_overview.png').convert('L'))
print("size:", im.shape, "nonwhite:", int((im < 250).sum()), f"({100*(im<250).mean():.2f}%)")
