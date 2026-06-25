#!/usr/bin/env python3
"""
End-to-End AI Chat Application Architecture (Bangkok Bank InnoHub)
Skeleton-and-zoom deck: a persistent 5-layer pill rail appears on every
content slide; each slide zooms into one layer in the detail panel.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
NAVY      = RGBColor(0x0B, 0x1F, 0x3A)
NAVY2     = RGBColor(0x14, 0x2E, 0x53)
BLUE      = RGBColor(0x1E, 0x5A, 0xA8)
BLUE_LT   = RGBColor(0x3C, 0x82, 0xD6)
SKY       = RGBColor(0xE8, 0xF1, 0xFB)
GOLD      = RGBColor(0xC9, 0xA2, 0x27)
GOLD_LT   = RGBColor(0xF2, 0xE6, 0xBE)
TEAL      = RGBColor(0x1A, 0x9C, 0x9C)
GREEN     = RGBColor(0x2E, 0x9E, 0x5B)
GREEN_LT  = RGBColor(0xE3, 0xF3, 0xE9)
RED       = RGBColor(0xC0, 0x39, 0x3D)
RED_LT    = RGBColor(0xF7, 0xE4, 0xE4)
ORANGE    = RGBColor(0xE0, 0x82, 0x2B)
ORANGE_LT = RGBColor(0xFB, 0xEC, 0xD9)
PURPLE    = RGBColor(0x6B, 0x4C, 0xA8)
PURPLE_LT = RGBColor(0xEC, 0xE6, 0xF6)
GREY      = RGBColor(0x5A, 0x66, 0x77)
GREY_LT   = RGBColor(0xF1, 0xF3, 0xF6)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
INK       = RGBColor(0x1A, 0x22, 0x30)
MUTE      = RGBColor(0x8A, 0x93, 0x9F)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
FONT = "Calibri"; MONO = "Consolas"
TOTAL = 15

# the five layers (label, color, glyph, dark-text?)
LAYERS = [
    ("Client",            BLUE_LT, "💬", False),
    ("Agent Orchestrator",TEAL,    "🧠", False),
    ("AI Hub Gateway",    GOLD,    "🏛️", True),
    ("LLM Providers",     PURPLE,  "🤖", False),
    ("Data & RAG",        GREEN,   "🗄️", False),
]

# ---------------------------------------------------------------- helpers
def slide(): return prs.slides.add_slide(BLANK)
def bg(s, color=WHITE):
    s.background.fill.solid(); s.background.fill.fore_color.rgb = color

def blend(c, t, other=WHITE):
    return RGBColor(round(c[0]+(other[0]-c[0])*t),
                    round(c[1]+(other[1]-c[1])*t),
                    round(c[2]+(other[2]-c[2])*t))

def _set_line(shape, color, w=1.0):
    if color is None: shape.line.fill.background()
    else: shape.line.color.rgb = color; shape.line.width = Pt(w)

def box(s, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        shadow=False, radius=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    _set_line(sp, line, lw); sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = el.makeelement(qn('a:outerShdw'),
            {'blurRad':'60000','dist':'30000','dir':'5400000','rotWithShape':'0'})
        clr = el.makeelement(qn('a:srgbClr'), {'val':'1A2230'})
        a = el.makeelement(qn('a:alpha'), {'val':'30000'})
        clr.append(a); sh.append(clr); ef.append(sh); el.append(ef)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try: sp.adjustments[0] = radius
        except Exception: pass
    return sp

def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ('a:buChar','a:buAutoNum'):
        for e in pPr.findall(qn(tag)): pPr.remove(e)
    if pPr.find(qn('a:buNone')) is None:
        pPr.append(pPr.makeelement(qn('a:buNone'), {}))

def _bullet(p, color="1E5AA8"):
    pPr = p._p.get_or_add_pPr(); pPr.set('indent','-182880'); pPr.set('marL','228600')
    pPr.append(pPr.makeelement(qn('a:buFont'), {'typeface':'Arial'}))
    pPr.append(pPr.makeelement(qn('a:buChar'), {'char':'•'}))

def _para(tf, txt, size, color, bold=False, align=PP_ALIGN.LEFT, italic=False,
          font=FONT, sa=2, sb=0, first=False, bullet=False, ls=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(sa); p.space_before = Pt(sb)
    if ls: p.line_spacing = ls
    runs = txt if isinstance(txt, list) else [(txt, color, bold, italic)]
    for seg in runs:
        t,c,b,i = (seg+(color,bold,italic))[:4] if isinstance(seg, tuple) else (seg,color,bold,italic)
        r = p.add_run(); r.text = t
        r.font.size=Pt(size); r.font.bold=b; r.font.italic=i; r.font.name=font; r.font.color.rgb=c
    _bullet(p) if bullet else _no_bullet(p)
    return p

def text(s, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, ln in enumerate(lines):
        _para(tf, ln['t'], ln.get('s',14), ln.get('c',INK), ln.get('b',False),
              ln.get('a',PP_ALIGN.LEFT), ln.get('i',False), ln.get('f',FONT),
              ln.get('sa',2), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False), ls=ln.get('ls'))
    return tb

def boxtext(sp, lines, anchor=MSO_ANCHOR.MIDDLE, ml=0.08, mr=0.08):
    tf = sp.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=Inches(ml); tf.margin_right=Inches(mr); tf.margin_top=Inches(0.03); tf.margin_bottom=Inches(0.03)
    for i, ln in enumerate(lines):
        _para(tf, ln['t'], ln.get('s',12), ln.get('c',INK), ln.get('b',False),
              ln.get('a',PP_ALIGN.CENTER), ln.get('i',False), ln.get('f',FONT),
              ln.get('sa',1), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False), ls=ln.get('ls'))
    return sp

def arrow(s, x1, y1, x2, y2, color=BLUE, w=2.25, dash=None):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    cn.line.color.rgb=color; cn.line.width=Pt(w); cn.shadow.inherit=False
    le = cn.line._get_or_add_ln()
    le.append(le.makeelement(qn('a:headEnd'), {'type':'none'}))
    le.append(le.makeelement(qn('a:tailEnd'), {'type':'triangle','w':'med','len':'med'}))
    if dash: le.append(le.makeelement(qn('a:prstDash'), {'val':dash}))
    return cn

def chip(s, x, y, w, h, label, fill, tcolor=WHITE, size=11, bold=True, icon=None, line=None):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=1.2, radius=0.5)
    boxtext(sp, [{'t':(f"{icon}  " if icon else "")+label,'s':size,'c':tcolor,'b':bold}])
    return sp

def icon_circle(s, x, y, d, glyph, fill, gcolor=WHITE, gsize=20, line=None):
    c = box(s, x, y, d, d, fill=fill, line=line, lw=1.2, shape=MSO_SHAPE.OVAL)
    boxtext(c, [{'t':glyph,'s':gsize,'c':gcolor,'b':True}])
    return c

def footer(s, n):
    box(s, 0, 7.16, 13.333, 0.34, fill=GREY_LT, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 7.16, 8, 0.34, [{'t':"Bangkok Bank InnoHub  |  AI Technology Team",'s':9,'c':GREY}],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 9.8, 7.16, 3.03, 0.34, [{'t':f"Confidential  ·  {n} / {TOTAL}",'s':9,'c':GREY,'a':PP_ALIGN.RIGHT}],
         anchor=MSO_ANCHOR.MIDDLE)

# ---- persistent skeleton rail + header ----
RAIL_X, PILL_W, PILL_H = 0.42, 2.55, 0.78
RAIL_TOP, GAP = 1.45, 0.30
CX = 3.55                      # content area left edge
CW = 13.333 - CX - 0.45        # content width

def pill_y(i): return RAIL_TOP + i*(PILL_H+GAP)

def chrome(s, kicker, title, n, active):
    """active: int index, or 'all' to highlight every layer."""
    bg(s)
    # top header
    box(s, 0, 0, 13.333, 1.0, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0, 1.0, 13.333, 0.06, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 0.12, 11.4, 0.34, [{'t':kicker,'s':12,'c':GOLD,'b':True}])
    text(s, 0.5, 0.40, 11.6, 0.6, [{'t':title,'s':23,'c':WHITE,'b':True}])
    bd = box(s, 12.5, 0.26, 0.55, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(bd, [{'t':str(n),'s':16,'c':NAVY,'b':True}])
    # rail label
    text(s, RAIL_X, 1.12, PILL_W, 0.3, [{'t':"ARCHITECTURE MAP",'s':9.5,'c':MUTE,'b':True}])
    # pills + flow arrows
    for i,(lbl,clr,gly,dark) in enumerate(LAYERS):
        y = pill_y(i)
        on = (active=='all') or (active==i)
        fill = clr if on else blend(clr, 0.66)
        tcol = (NAVY if dark else WHITE) if on else MUTE
        w = PILL_W + (0.0 if not on or active=='all' else 0.0)
        sp = box(s, RAIL_X, y, w, PILL_H, fill=fill, line=(GOLD if (on and active!='all') else None),
                 lw=2.5, shadow=on, radius=0.5)
        boxtext(sp, [{'t':f"{gly}  {lbl}",'s':(12 if on else 10.5),'c':tcol,'b':True}], ml=0.12, mr=0.12)
        # down request arrow between pills
        if i < len(LAYERS)-1:
            ax = RAIL_X + PILL_W/2
            arrow(s, ax-0.18, y+PILL_H+0.02, ax-0.18, y+PILL_H+GAP-0.02,
                  color=(blend(NAVY,0.55) if active=='all' else RGBColor(0xB9,0xC2,0xCE)), w=1.5)
            arrow(s, ax+0.18, y+PILL_H+GAP-0.02, ax+0.18, y+PILL_H+0.02,
                  color=(blend(GREEN,0.4) if active=='all' else RGBColor(0xCBD6 if False else 0xC9,0xDE,0xCF)), w=1.5)
        # zoom marker on active single layer
        if on and active!='all':
            zx = RAIL_X+PILL_W
            arrow(s, zx-0.02, y+PILL_H/2, CX-0.22, y+PILL_H/2 if False else y+PILL_H/2, color=GOLD, w=2.25)
            mg = icon_circle(s, zx-0.30, y-0.18, 0.42, "🔍", GOLD, NAVY, 13)
    # request/response legend at rail bottom
    text(s, RAIL_X, pill_y(4)+PILL_H+0.05, PILL_W, 0.3,
         [{'t':[("▼ request   ",blend(NAVY,0.3),True,False),("▲ response",blend(GREEN,0.1),True,False)],'s':8.5,'a':PP_ALIGN.CENTER}])
    footer(s, n)

def panel(s, color, title, sub, glyph, x=CX, y=1.25, w=CW, h=5.75):
    """Zoom detail panel with a colored header; returns inner top y."""
    card = box(s, x, y, w, h, fill=WHITE, line=blend(color,0.4), lw=1.4, shadow=True, radius=0.03)
    box(s, x, y, w, 0.72, fill=color, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    icon_circle(s, x+0.18, y+0.13, 0.46, glyph, WHITE, color, 16, line=None)
    dark = (color==GOLD or color==GOLD_LT)
    tcol = NAVY if dark else WHITE
    text(s, x+0.78, y+0.06, w-2.0, 0.4, [{'t':title,'s':16,'c':tcol,'b':True,'sa':0}])
    text(s, x+0.78, y+0.40, w-2.0, 0.3, [{'t':sub,'s':10.5,'c':(NAVY if dark else SKY)}])
    chip(s, x+w-1.55, y+0.18, 1.35, 0.36, "ZOOM ‹", blend(WHITE,0.0) if False else WHITE, color, size=10, line=None)
    return y+0.9

# ================================================================ SLIDE 1 — Title
s = slide(); bg(s, NAVY)
box(s, 0,0,13.333,7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 8.3,-1.5,7,10.5, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 2.0, 0.12, 2.8, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 0.8, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 1.95, 7.4, 2.6, [
    {'t':"End-to-End AI Chat",'s':44,'c':WHITE,'b':True,'sa':2},
    {'t':"Application Architecture",'s':44,'c':GOLD,'b':True,'sa':10},
    {'t':"A five-layer journey from user input to intelligent",'s':18,'c':SKY,'sa':1},
    {'t':"response — with governance, PII protection & security.",'s':18,'c':SKY},
])
# hero pill stack (the loved design)
hx, hy, hw, hh = 9.4, 1.55, 3.1, 0.82
text(s, hx-0.05, 1.15, hw, 0.3, [{'t':"THE SKELETON",'s':10,'c':GOLD,'b':True,'a':PP_ALIGN.CENTER}])
for i,(lbl,clr,gly,dark) in enumerate(LAYERS):
    y = hy + i*(hh+0.22)
    sp = box(s, hx, y, hw, hh, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(sp, [{'t':f"{gly}  {lbl}",'s':13,'c':(NAVY if dark else WHITE),'b':True}])
    if i < 4:
        arrow(s, hx+hw/2, y+hh+0.01, hx+hw/2, y+hh+0.21, color=GOLD, w=1.75)
text(s, 0.95, 6.45, 8, 0.5, [{'t':"Executive Briefing  ·  AI Technology Team  ·  2026",'s':13,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ================================================================ SLIDE 2 — Skeleton / Agenda
s = slide()
box(s, 0,0,13.333,1.0, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 0,1.0,13.333,0.06, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
bg_card = bg  # noop
s.background.fill.solid(); s.background.fill.fore_color.rgb = WHITE
box(s, 0,0,13.333,1.0, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 0,1.0,13.333,0.06, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
text(s, 0.5, 0.12, 11, 0.34, [{'t':"ARCHITECTURE MAP",'s':12,'c':GOLD,'b':True}])
text(s, 0.5, 0.40, 11.6, 0.6, [{'t':"The Skeleton — Five Cooperating Layers",'s':23,'c':WHITE,'b':True}])
bd = box(s, 12.5, 0.26, 0.55, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.OVAL); boxtext(bd,[{'t':"2",'s':16,'c':NAVY,'b':True}])
footer(s, 2)
text(s, 0.5, 1.2, 12.3, 0.35, [{'t':"This stack is the backbone of every slide that follows. Each section zooms into one layer.",'s':12.5,'c':GREY,'i':True}])
# big vertical stack on left, descriptors on right
sx, sy, sw, sh2 = 0.95, 1.85, 3.4, 0.86
DESCR = [
    "Secure web/mobile chat UI — captures the user's natural-language request.",
    "Plans the task, calls tools, generates SQL, applies input/output guardrails & PII handling.",
    "Central control plane: policy, PII masking, quotas and provider routing + logging.",
    "External & private LLMs — generate the reasoning and response (placeholders only).",
    "Transactional DB, vector knowledge base, document store and the audit trail.",
]
ZOOMS = ["Slide 3","Slides 4–6","Slide 7","Slides 8–9","Slide 10"]
for i,(lbl,clr,gly,dark) in enumerate(LAYERS):
    y = sy + i*(sh2+0.18)
    sp = box(s, sx, y, sw, sh2, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(sp, [{'t':f"{gly}  {lbl}",'s':15,'c':(NAVY if dark else WHITE),'b':True}])
    if i<4: arrow(s, sx+sw/2-0.2, y+sh2+0.005, sx+sw/2-0.2, y+sh2+0.18, color=blend(NAVY,0.4), w=1.6)
    if i<4: arrow(s, sx+sw/2+0.2, y+sh2+0.18, sx+sw/2+0.2, y+sh2+0.005, color=blend(GREEN,0.2), w=1.6)
    # descriptor card
    dc = box(s, 4.7, y, 6.45, sh2, fill=GREY_LT, line=blend(clr,0.45), lw=1.2, radius=0.08)
    box(s, 4.7, y, 0.12, sh2, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 4.95, y+0.07, 6.0, sh2-0.1, [{'t':DESCR[i],'s':11,'c':INK,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
    # zoom tag
    zt = box(s, 11.3, y+0.16, 1.55, 0.54, fill=NAVY, line=GOLD, lw=1.0, radius=0.4)
    boxtext(zt, [{'t':[("🔍 ",GOLD,True,False),(ZOOMS[i],WHITE,True,False)],'s':10}])
# request/response legend
text(s, 0.95, sy+5*(sh2+0.18)-0.05, 4, 0.3,
     [{'t':[("▼ request   ",blend(NAVY,0.3),True,False),("▲ response",GREEN,True,False)],'s':10}])

# ================================================================ SLIDE 3 — L1 Client (zoom)
s = slide(); chrome(s, "ZOOM · LAYER 1", "Client Layer", 3, 0)
iy = panel(s, BLUE_LT, "1 · Client", "The secure entry & exit point for the user", "💬")
# left: role + responsibilities
text(s, CX+0.25, iy, 4.5, 0.4, [{'t':"Role",'s':13,'c':NAVY,'b':True}])
for t in ["Render a secure chat experience","Authenticate the user (SSO / OAuth2)","Manage session & conversation state","Stream the response back in real time"]:
    pass
text(s, CX+0.25, iy+0.36, 4.55, 2.2, [
    {'t':"Render a secure chat experience",'s':11.5,'c':INK,'bu':True,'sa':5},
    {'t':"Authenticate the user (SSO / OAuth2)",'s':11.5,'c':INK,'bu':True,'sa':5},
    {'t':"Manage session & conversation state",'s':11.5,'c':INK,'bu':True,'sa':5},
    {'t':"Stream the response back in real time",'s':11.5,'c':INK,'bu':True,'sa':5},
])
text(s, CX+0.25, iy+2.5, 4.5, 0.35, [{'t':"Handover",'s':13,'c':NAVY,'b':True}])
hc = box(s, CX+0.25, iy+2.86, 4.55, 0.95, fill=SKY, line=BLUE, lw=1.0, radius=0.08)
boxtext(hc, [{'t':"The raw query is passed to the Agent Orchestrator, where PII handling begins.",'s':11,'c':INK}], ml=0.18, mr=0.18)
# right: chat mockup
mx = CX+5.1
ph = box(s, mx, iy, 3.9, 3.9, fill=GREY_LT, line=RGBColor(0xCF,0xD6,0xE0), lw=1.2, radius=0.05)
box(s, mx, iy, 3.9, 0.6, fill=NAVY, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
icon_circle(s, mx+0.16, iy+0.12, 0.36, "🤖", GOLD, NAVY, 13)
text(s, mx+0.6, iy+0.1, 3.0, 0.5, [{'t':"InnoHub Assistant",'s':11,'c':WHITE,'b':True,'sa':0},{'t':"● Secure channel",'s':8,'c':GREEN_LT}])
ub = box(s, mx+0.5, iy+0.85, 3.2, 1.55, fill=BLUE, line=None, radius=0.18)
boxtext(ub, [{'t':"Show me transaction history for customer Somchai Jantapan with account number 1234567890 from last month",'s':11,'c':WHITE,'a':PP_ALIGN.LEFT}], ml=0.16, mr=0.16)
text(s, mx+0.5, iy+2.42, 3.2, 0.25, [{'t':"You · just now",'s':8.5,'c':GREY,'a':PP_ALIGN.RIGHT}])
ib = box(s, mx+0.25, iy+2.9, 3.0, 0.5, fill=WHITE, line=RGBColor(0xCF,0xD6,0xE0), lw=1.0, radius=0.5)
boxtext(ib, [{'t':"Type a message…",'s':10,'c':GREY,'a':PP_ALIGN.LEFT}], ml=0.14)
icon_circle(s, mx+3.3, iy+2.92, 0.46, "➤", BLUE, WHITE, 12)

# ================================================================ SLIDE 4 — L2 Orchestrator (zoom overview)
s = slide(); chrome(s, "ZOOM · LAYER 2", "Agent Orchestrator", 4, 1)
iy = panel(s, TEAL, "2 · Agent Orchestrator", "Plans the task, calls tools, governs the flow", "🧠")
text(s, CX+0.25, iy, CW-0.5, 0.35, [{'t':"The orchestrator turns one request into an executable, governed plan:",'s':11.5,'c':GREY,'i':True}])
dec = [
    ("Task Type","DB retrieval + LLM-assisted SQL", BLUE, SKY),
    ("Tools","SQL tool · RAG retriever · formatter", PURPLE, PURPLE_LT),
    ("Data Sources","Transaction DB · knowledge base", GREEN, GREEN_LT),
    ("Constraints","Read-only · PII placeholders kept", RED, RED_LT),
    ("Plan","generate → validate → run → answer", GOLD, GOLD_LT),
    ("Governance","route via AI Hub · log each step", NAVY, GREY_LT),
]
bw, bh = (CW-0.5-2*0.2)/3, 1.35
for i,(t,d,ac,bgc) in enumerate(dec):
    col=i%3; row=i//3
    x = CX+0.25+col*(bw+0.2); y = iy+0.5+row*(bh+0.22)
    c = box(s, x, y, bw, bh, fill=bgc, line=ac, lw=1.2, shadow=True, radius=0.08)
    box(s, x, y, bw, 0.42, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,y,bw,0.42,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':t,'s':12,'c':(NAVY if ac==GOLD else WHITE),'b':True}])
    text(s, x+0.16, y+0.5, bw-0.3, 0.75, [{'t':d,'s':11,'c':INK,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
text(s, CX+0.25, iy+0.5+2*(bh+0.22)+0.02, CW-0.5, 0.4,
     [{'t':[("Next zoom: ",TEAL,True,False),("input guardrails & PII redaction (slide 5), then dynamic SQL tool generation (slide 6).",INK,False,False)],'s':11}])

# ================================================================ SLIDE 5 — L2 Input Guardrails & PII
s = slide(); chrome(s, "ZOOM · LAYER 2", "Input Guardrails & PII Redaction", 5, 1)
iy = panel(s, TEAL, "2 · Agent Orchestrator — Guardrails", "Scan, then redact before anything leaves", "🛡️")
# 4 checks
checks=[("🔐","PII",RED,RED_LT),("⚠️","Injection",ORANGE,ORANGE_LT),("🚫","Toxicity",PURPLE,PURPLE_LT),("📋","Policy",BLUE,SKY)]
cw4=(CW-0.5-3*0.15)/4
for i,(ic,t,ac,bgc) in enumerate(checks):
    x=CX+0.25+i*(cw4+0.15)
    c=box(s,x,iy,cw4,0.62,fill=bgc,line=ac,lw=1.1,radius=0.12)
    boxtext(c,[{'t':[(ic+" ",ac,True,False),(t,NAVY,True,False)],'s':11.5}])
# redaction flow
fy=iy+0.78
orig=box(s, CX+0.25, fy, 4.1, 1.25, fill=WHITE, line=RED, lw=1.3, shadow=True, radius=0.06)
text(s, CX+0.4, fy+0.06, 3.8,0.25, [{'t':"ORIGINAL",'s':9.5,'c':RED,'b':True}])
text(s, CX+0.4, fy+0.34, 3.85,0.85, [{'t':[("…customer ",INK,False,False),("Somchai Jantapan",RED,True,False),(", account ",INK,False,False),("1234567890",RED,True,False),(", last month",INK,False,False)],'s':10.5,'ls':1.0}])
sc=box(s, CX+4.55, fy+0.28, 1.2, 0.7, fill=NAVY, line=GOLD, lw=1.3, shadow=True, radius=0.12)
boxtext(sc,[{'t':"🛡️ SCAN",'s':10,'c':GOLD,'b':True}])
arrow(s, CX+4.35, fy+0.62, CX+4.55, fy+0.62, color=GREY, w=2.25)
arrow(s, CX+5.75, fy+0.62, CX+5.95, fy+0.62, color=GREY, w=2.25)
red=box(s, CX+5.97, fy, CW-6.5, 1.25, fill=WHITE, line=GREEN, lw=1.3, shadow=True, radius=0.06)
text(s, CX+6.12, fy+0.06, 3.0,0.25, [{'t':"REDACTED ✓",'s':9.5,'c':GREEN,'b':True}])
text(s, CX+6.12, fy+0.34, CW-6.7,0.85, [{'t':[("…customer ",INK,False,False),("CUSTOMER_NAME_001",GREEN,True,False),(", account ",INK,False,False),("ACCOUNT_ID_001",GREEN,True,False),(", last month",INK,False,False)],'s':10.5,'ls':1.0}])
# mapping table
ty=fy+1.45
text(s, CX+0.25, ty, 6, 0.3, [{'t':"PII Mapping Table (held securely for re-insertion)",'s':12,'c':NAVY,'b':True}])
rows=[("Original","PII Type","Placeholder",True),("Somchai Jantapan","Customer Name","CUSTOMER_NAME_001",False),
      ("1234567890","Account ID","ACCOUNT_ID_001",False),("\"last month\"","Temporal context","(no redaction)",False)]
cws=[3.3,2.7,3.4]; tyy=ty+0.36
for r,(a,b,cc,hdr) in enumerate(rows):
    xx=CX+0.25
    for ci,val in enumerate((a,b,cc)):
        rbg = NAVY if hdr else (SKY if r in (1,2) else WHITE)
        cell=box(s,xx,tyy,cws[ci],0.42,fill=rbg,line=RGBColor(0xD5,0xDC,0xE6),lw=0.75,shape=MSO_SHAPE.RECTANGLE)
        col = GOLD if hdr else (RED if (ci==2 and r in(1,2)) else (GREY if r==3 else INK))
        boxtext(cell,[{'t':val,'s':10.5,'c':col,'b':(hdr or (ci==2 and r in(1,2))),'a':PP_ALIGN.LEFT}],ml=0.12)
        xx+=cws[ci]
    tyy+=0.42

# ================================================================ SLIDE 6 — L2 SQL tool generation
s = slide(); chrome(s, "ZOOM · LAYER 2", "Dynamic Tool Generation — SQL", 6, 1)
iy = panel(s, TEAL, "2 · Agent Orchestrator — Tooling", "The agent writes its own parameterized query", "⚙️")
code=box(s, CX+0.25, iy, 5.3, 3.0, fill=RGBColor(0x10,0x1A,0x2B), line=NAVY2, lw=1.0, shadow=True, radius=0.04)
box(s, CX+0.25, iy, 5.3, 0.4, fill=NAVY2, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
for k,clr in enumerate([RED,GOLD,GREEN]):
    box(s, CX+0.45+k*0.2, iy+0.13, 0.13,0.13, fill=clr, line=None, shape=MSO_SHAPE.OVAL)
text(s, CX+1.2, iy+0.06, 3.5,0.3, [{'t':"generated_query.sql",'s':10,'c':SKY,'a':PP_ALIGN.LEFT}])
sqlb=text(s, CX+0.45, iy+0.5, 5.0, 2.4, [])
sql=[ [("SELECT ",BLUE_LT,True),("txn_date, amount, channel",WHITE,False)],
      [("FROM ",BLUE_LT,True),("transactions t",WHITE,False)],
      [("JOIN ",BLUE_LT,True),("customers c ",WHITE,False),("ON ",BLUE_LT,True),("c.id=t.cust_id",WHITE,False)],
      [("WHERE ",BLUE_LT,True),("c.name = ",WHITE,False),(":CUSTOMER_NAME_001",GOLD,True)],
      [("  AND ",BLUE_LT,True),("t.account = ",WHITE,False),(":ACCOUNT_ID_001",GOLD,True)],
      [("  AND ",BLUE_LT,True),("t.txn_date >= ",WHITE,False),("date_trunc(",GREEN_LT,False)],
      [("       'month', now()-interval '1 month')",GREEN_LT,False)],
      [("ORDER BY ",BLUE_LT,True),("t.txn_date ",WHITE,False),("DESC;",BLUE_LT,True)] ]
for j,segs in enumerate(sql):
    p=sqlb.text_frame.paragraphs[0] if j==0 else sqlb.text_frame.add_paragraph()
    p.space_after=Pt(1); _no_bullet(p)
    for tt,cc,bb in segs:
        r=p.add_run(); r.text=tt; r.font.size=Pt(11); r.font.name=MONO; r.font.color.rgb=cc; r.font.bold=bb
# safeguards
sx=CX+5.8
text(s, sx, iy, CW-5.8, 0.35, [{'t':"Built-in safeguards",'s':13,'c':NAVY,'b':True}])
guards=[("🔒","Placeholders bound, never real values",RED),("👁️","Read-only — no writes / DDL",BLUE),
        ("🧪","Validated vs schema & row-level access",PURPLE),("⏱️","Row caps, timeouts & cost ceilings",GREEN)]
yy=iy+0.42
for ic,t,clr in guards:
    icon_circle(s, sx, yy, 0.4, ic, clr, WHITE, 13)
    text(s, sx+0.55, yy+0.02, CW-6.4, 0.55, [{'t':t,'s':11,'c':INK,'ls':1.0}])
    yy+=0.66
text(s, CX+0.25, iy+3.2, CW-0.5, 0.5,
     [{'t':[("Result: ",GREEN,True,False),("rows return still keyed to placeholders — real data is restored only at the secured output stage.",INK,False,False)],'s':11}])

# ================================================================ SLIDE 7 — L3 AI Hub Gateway
s = slide(); chrome(s, "ZOOM · LAYER 3", "AI Hub Gateway", 7, 2)
iy = panel(s, GOLD, "3 · AI Hub Gateway", "One governed control plane for every model call", "🏛️")
mods=[("🛂","Policy Engine","access control · routing rules"),
      ("🎭","PII Masking","verify redaction · block leakage"),
      ("📦","Quota & Rate","budgets · throttling · cost"),
      ("🔀","Provider Router","select · balance · fail-over"),
      ("🧾","Logging & Tracing","every prompt, token & decision")]
bw=(CW-0.5-2*0.2)/3; bh=1.12
for i,(ic,t,d) in enumerate(mods):
    col=i%3; row=i//3
    x=CX+0.25+col*(bw+0.2); y=iy+row*(bh+0.2)
    c=box(s,x,y,bw,bh,fill=GOLD_LT,line=GOLD,lw=1.1,shadow=True,radius=0.08)
    icon_circle(s, x+0.14, y+0.16, 0.5, ic, NAVY, GOLD, 15)
    text(s, x+0.72, y+0.13, bw-0.85, 0.35, [{'t':t,'s':12,'c':NAVY,'b':True,'sa':0}])
    text(s, x+0.72, y+0.5, bw-0.85, 0.5, [{'t':d,'s':9.5,'c':GREY}])
# mini flow at bottom
fy=iy+2*(bh+0.2)+0.05
for i,(lbl,clr) in enumerate([("Agent",TEAL),("AI Hub Gateway",GOLD),("LLM Provider",PURPLE)]):
    w=2.5 if i==1 else 2.0
    x=CX+0.25+ (0 if i==0 else (2.0+0.55 if i==1 else 2.0+0.55+2.5+0.55))
    c=box(s,x,fy,w,0.55,fill=clr,line=None,radius=0.4)
    boxtext(c,[{'t':lbl,'s':11,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<2: arrow(s, x+w, fy+0.18, x+w+0.55, fy+0.18, color=PURPLE if i==1 else TEAL, w=2.25)
    if i<2: arrow(s, x+w+0.55, fy+0.38, x+w, fy+0.38, color=GREEN, w=1.75)
text(s, CX+8.0, fy+0.02, CW-8.2, 0.55, [{'t':"One gateway = one audit point & one policy surface.",'s':10.5,'c':NAVY,'b':True,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 8 — L4 LLM Providers
s = slide(); chrome(s, "ZOOM · LAYER 4", "LLM Providers & Routing", 8, 3)
iy = panel(s, PURPLE, "4 · LLM Providers", "Pluggable models behind the gateway", "🤖")
text(s, CX+0.25, iy, 4.6, 0.35, [{'t':"Provider pool",'s':13,'c':NAVY,'b':True}])
provs=[("OpenAI GPT","frontier reasoning"),("Anthropic Claude","long-context & safety"),
       ("Google Gemini","multimodal"),("Private / On-prem","sensitive workloads")]
yy=iy+0.42
for t,d in provs:
    c=box(s, CX+0.25, yy, 4.6, 0.6, fill=PURPLE_LT, line=PURPLE, lw=1.0, radius=0.1)
    text(s, CX+0.45, yy+0.04, 4.3, 0.3, [{'t':t,'s':11.5,'c':PURPLE,'b':True,'sa':0}])
    text(s, CX+0.45, yy+0.32, 4.3, 0.25, [{'t':d,'s':9,'c':GREY}])
    yy+=0.7
# routing features
rx=CX+5.2
text(s, rx, iy, CW-5.2, 0.35, [{'t':"Routing & resilience",'s':13,'c':NAVY,'b':True}])
feats=[("🔀","Best-fit model per task type",BLUE),("⚖️","Load balancing across providers",TEAL),
       ("🛟","Automatic fail-over on outage",GREEN),("🎭","Only placeholders are sent out",RED)]
yy=iy+0.42
for ic,t,clr in feats:
    icon_circle(s, rx, yy, 0.42, ic, clr, WHITE, 13)
    text(s, rx+0.58, yy+0.04, CW-5.9, 0.5, [{'t':t,'s':11.5,'c':INK,'ls':1.0}])
    yy+=0.72
note=box(s, CX+0.25, iy+3.1, CW-0.5, 0.65, fill=SKY, line=BLUE, lw=1.0, radius=0.08)
boxtext(note,[{'t':[("Output side (slide 9): ",BLUE,True,False),("the model replies with placeholders, then output guardrails + PII re-insertion produce the final answer.",INK,False,False)],'s':11}],ml=0.2,mr=0.2)

# ================================================================ SLIDE 9 — L4 Output Guardrails & re-insertion
s = slide(); chrome(s, "ZOOM · LAYER 4", "Output Guardrails & PII Re-insertion", 9, 3)
iy = panel(s, PURPLE, "4 · LLM Providers — Output", "Restore real data only inside the trust boundary", "🔓")
stg=[("LLM Response","…for CUSTOMER_NAME_001, account ACCOUNT_ID_001: 14 txns",PURPLE,PURPLE_LT),
     ("Output Guardrails","check leaked PII, hallucination, toxicity, disclosure",RED,RED_LT),
     ("PII Re-insertion","restore real values from the secure mapping table",GREEN,GREEN_LT),
     ("Final Answer","…for Somchai Jantapan, account 1234567890: 14 txns",BLUE,SKY)]
bw=(CW-0.5-3*0.18)/4
for i,(t,d,ac,bgc) in enumerate(stg):
    x=CX+0.25+i*(bw+0.18)
    c=box(s,x,iy,bw,1.95,fill=bgc,line=ac,lw=1.2,shadow=True,radius=0.06)
    box(s,x,iy,bw,0.46,fill=ac,line=None,shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,iy,bw,0.46,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),[{'t':f"{i+1}. {t}",'s':10.5,'c':WHITE,'b':True}])
    text(s, x+0.12, iy+0.52, bw-0.24, 1.3, [{'t':d,'s':10,'c':INK,'ls':1.0}])
    if i<3: arrow(s, x+bw, iy+0.95, x+bw+0.18, iy+0.95, color=GREY, w=2.0)
text(s, CX+0.25, iy+2.15, CW-0.5, 0.3, [{'t':"Output guardrail checks",'s':12.5,'c':NAVY,'b':True}])
oc=[("🔍","No PII leakage",RED),("🎯","Grounded & accurate",BLUE),("🧹","Safe content",PURPLE),("📜","Disclosure rules",GREEN)]
cw4=(CW-0.5-3*0.15)/4
for i,(ic,t,clr) in enumerate(oc):
    x=CX+0.25+i*(cw4+0.15)
    c=box(s,x,iy+2.5,cw4,0.8,fill=WHITE,line=clr,lw=1.1,shadow=True,radius=0.08)
    icon_circle(s, x+0.12, iy+2.62, 0.46, ic, clr, WHITE, 13)
    text(s, x+0.66, iy+2.62, cw4-0.75, 0.6, [{'t':t,'s':10.5,'c':NAVY,'b':True,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 10 — L5 Data & RAG
s = slide(); chrome(s, "ZOOM · LAYER 5", "Data & Knowledge Layer", 10, 4)
iy = panel(s, GREEN, "5 · Data & RAG", "Grounding the answer in real, governed data", "🗄️")
stores=[("🗄️","Transactional DB","read-only SQL · row-level scope",BLUE,SKY),
        ("📚","Vector Knowledge Base","semantic RAG retrieval",PURPLE,PURPLE_LT),
        ("📄","Document Store","policies, products, statements",ORANGE,ORANGE_LT),
        ("📈","Audit Trail","immutable log of every access",GREEN,GREEN_LT)]
bw=(CW-0.5-0.2)/2; bh=1.0
for i,(ic,t,d,ac,bgc) in enumerate(stores):
    col=i%2; row=i//2
    x=CX+0.25+col*(bw+0.2); y=iy+row*(bh+0.2)
    c=box(s,x,y,bw,bh,fill=bgc,line=ac,lw=1.1,shadow=True,radius=0.08)
    icon_circle(s, x+0.16, y+0.2, 0.6, ic, ac, WHITE, 17)
    text(s, x+0.95, y+0.16, bw-1.1, 0.35, [{'t':t,'s':13,'c':NAVY,'b':True,'sa':0}])
    text(s, x+0.95, y+0.55, bw-1.1, 0.35, [{'t':d,'s':10,'c':GREY}])
# RAG flow strip
fy=iy+2*(bh+0.2)+0.0
for i,(lbl,clr) in enumerate([("Redacted query",TEAL),("Retrieve (SQL+RAG)",GREEN),("Ground response",BLUE)]):
    w=(CW-0.5-2*0.5)/3; x=CX+0.25+i*(w+0.5)
    c=box(s,x,fy,w,0.55,fill=clr,line=None,radius=0.4)
    boxtext(c,[{'t':lbl,'s':10.5,'c':WHITE,'b':True}])
    if i<2: arrow(s, x+w, fy+0.28, x+w+0.5, fy+0.28, color=GREY, w=2.25)

# ================================================================ SLIDE 11 — End-to-End Sequence (all)
s = slide(); chrome(s, "CROSS-CUTTING · ALL LAYERS", "End-to-End Request Lifecycle", 11, 'all')
iy = panel(s, NAVY2, "Sequence View", "How one query traverses all five layers", "🔄")
actors=[("User",BLUE_LT),("Orchestrator",TEAL),("AI Hub",GOLD),("LLM",PURPLE),("Data/RAG",GREEN)]
ax0=CX+0.7; span=CW-1.6; xs=[ax0+span*i/4 for i in range(5)]
top=iy+0.05
for (nm,clr),x in zip(actors,xs):
    h=box(s, x-0.85, top, 1.7, 0.45, fill=clr, line=None, radius=0.3)
    boxtext(h,[{'t':nm,'s':10,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    ll=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x),Inches(top+0.45),Inches(x),Inches(iy+3.55))
    ll.line.color.rgb=RGBColor(0xC4,0xCC,0xD6); ll.line.width=Pt(1.0); ll.shadow.inherit=False
    led=ll.line._get_or_add_ln(); led.append(led.makeelement(qn('a:prstDash'),{'val':'dash'}))
msgs=[(0,1,"1 · query (raw)",BLUE),(1,1,"2 · guardrails + PII redaction",RED),
      (1,4,"3 · retrieve (SQL / RAG)",GREEN),(4,1,"4 · placeholdered rows",GREEN),
      (1,2,"5 · masked prompt → gateway",TEAL),(2,3,"6 · policy ✓, route",GOLD),
      (3,2,"7 · model response",PURPLE),(2,1,"8 · logged + masked",GOLD),
      (1,1,"9 · output guardrails + re-insertion",RED),(1,0,"10 · final answer (real data)",BLUE)]
yy=top+0.62
for a,b,lbl,clr in msgs:
    x1,x2=xs[a],xs[b]
    if a==b:
        box(s, x1, yy-0.02, 0.9, 0.26, fill=None, line=clr, lw=1.4, radius=0.4)
        text(s, x1+1.0, yy-0.16, 6.0, 0.3, [{'t':lbl,'s':9.5,'c':INK,'b':True}])
    else:
        arrow(s, x1, yy+0.1, x2, yy+0.1, color=clr, w=1.9)
        text(s, min(x1,x2)+0.05, yy-0.16, abs(x2-x1)+0.1, 0.3, [{'t':lbl,'s':9.5,'c':INK,'b':True,'a':PP_ALIGN.CENTER}])
    yy+=0.29
text(s, CX+0.25, iy+3.7, CW-0.5, 0.3, [{'t':[("🛡 guardrails at 2 & 9   ·   ",RED,True,False),("🏛 gateway governs 5–8   ·   ",GOLD,True,False),("🎭 real PII never crosses the dashed boundary",GREEN,True,False)],'s':9.5}])

# ================================================================ SLIDE 12 — Observability (all)
s = slide(); chrome(s, "CROSS-CUTTING · ALL LAYERS", "Logging, Monitoring & Audit", 12, 'all')
iy = panel(s, NAVY2, "Observability", "Every layer emits telemetry into one pipeline", "📊")
coll=box(s, CX+0.25, iy, CW-0.5, 0.55, fill=GOLD, line=None, shadow=True, radius=0.2)
boxtext(coll,[{'t':"📥  Telemetry Collector — structured logs · traces · metrics · events",'s':12,'c':NAVY,'b':True}])
pil=[("📈","Metrics","latency, tokens, cost, error & throttle rates",BLUE,SKY),
     ("🧵","Traces","end-to-end span per request, correlation IDs",PURPLE,PURPLE_LT),
     ("🧾","Audit Logs","immutable prompts, decisions & PII actions",GREEN,GREEN_LT)]
bw=(CW-0.5-2*0.2)/3
for i,(ic,t,d,ac,bgc) in enumerate(pil):
    x=CX+0.25+i*(bw+0.2)
    c=box(s,x,iy+0.75,bw,1.5,fill=bgc,line=ac,lw=1.2,shadow=True,radius=0.06)
    icon_circle(s, x+0.16, iy+0.92, 0.55, ic, ac, WHITE, 16)
    text(s, x+0.8, iy+0.92, bw-0.9, 0.4, [{'t':t,'s':13,'c':NAVY,'b':True}])
    text(s, x+0.18, iy+1.55, bw-0.35, 0.65, [{'t':d,'s':10.5,'c':GREY,'ls':1.0}])
out=box(s, CX+0.25, iy+2.45, CW-0.5, 1.05, fill=SKY, line=BLUE, lw=1.0, radius=0.06)
text(s, CX+0.45, iy+2.52, 4, 0.3, [{'t':"What this enables",'s':12,'c':BLUE,'b':True}])
items=[("🔎","Full traceability per interaction"),("⚖️","Audit readiness (BOT, PDPA)"),
       ("💰","Cost & quota accountability"),("🚨","Real-time anomaly detection")]
for i,(ic,t) in enumerate(items):
    x=CX+0.5+(i%2)*((CW-1.0)/2); y=iy+2.85+(i//2)*0.32
    text(s, x, y, (CW-1.0)/2, 0.3, [{'t':[(ic+"  ",INK,False,False),(t,INK,False,False)],'s':10.5}])

# ================================================================ SLIDE 13 — Security (all)
s = slide(); chrome(s, "CROSS-CUTTING · ALL LAYERS", "Defense-in-Depth", 13, 'all')
iy = panel(s, NAVY2, "Security & Compliance", "Layered controls — no single point of failure", "🔐")
ctrls=[("🔐","Data Protection","redaction · tokenization · encryption · secure vault",RED),
       ("🛡️","Guardrails","input/output filtering · injection defense · screening",ORANGE),
       ("🏛️","Governance","central policy · allow-lists · quotas · approvals",GOLD),
       ("👤","Access Control","RBAC · row-level scope · least-privilege tools",BLUE),
       ("📜","Compliance","PDPA & BOT · audit trails · residency & retention",PURPLE),
       ("🔭","Monitoring","continuous logging · anomaly detection · response",GREEN)]
bw=(CW-0.5-0.2)/2; bh=1.05
for i,(ic,t,d,clr) in enumerate(ctrls):
    col=i%2; row=i//2
    x=CX+0.25+col*(bw+0.2); y=iy+row*(bh+0.15)
    c=box(s,x,y,bw,bh,fill=WHITE,line=blend(clr,0.4),lw=1.0,shadow=True,radius=0.06)
    box(s,x,y,0.12,bh,fill=clr,line=None,shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, x+0.28, y+0.28, 0.5, ic, clr, WHITE, 15)
    text(s, x+0.92, y+0.13, bw-1.05, 0.35, [{'t':t,'s':12.5,'c':NAVY,'b':True}])
    text(s, x+0.92, y+0.5, bw-1.05, 0.45, [{'t':d,'s':9.5,'c':GREY,'ls':1.0}])

# ================================================================ SLIDE 14 — Key Takeaways
s = slide(); chrome(s, "WRAP-UP", "Key Takeaways", 14, 'all')
iy = panel(s, NAVY2, "What to Remember", "Five ideas, one for each layer", "✅")
tk=[("①","Layered by design","Five cooperating layers turn one query into a governed response.",BLUE_LT),
    ("②","PII stays inside the boundary","Redact in, re-insert out — models only ever see placeholders.",TEAL),
    ("③","Guardrails at every stage","Input, processing & output checks make compliance continuous.",GOLD),
    ("④","One gateway governs all","The AI Hub centralizes policy, masking, routing, quota & logging.",PURPLE),
    ("⑤","Observable end-to-end","Metrics, traces & audit logs deliver trust & regulatory readiness.",GREEN)]
yy=iy
for ic,t,d,clr in tk:
    c=box(s, CX+0.25, yy, CW-0.5, 0.66, fill=WHITE, line=blend(clr,0.4), lw=1.0, shadow=True, radius=0.1)
    box(s, CX+0.25, yy, 0.12, 0.66, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, CX+0.42, yy+0.13, 0.4, ic, clr, (NAVY if clr==GOLD else WHITE), 15)
    text(s, CX+1.0, yy+0.05, 3.6, 0.6, [{'t':t,'s':12.5,'c':NAVY,'b':True}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, CX+4.7, yy+0.05, CW-5.0, 0.6, [{'t':d,'s':10.5,'c':GREY,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
    yy+=0.74

# ================================================================ SLIDE 15 — Closing
s = slide(); bg(s, NAVY)
box(s, 0,0,13.333,7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, -2,4.4,18,3.4, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 2.3, 0.12, 2.2, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 1.2, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 2.3, 8.0, 2.0, [
    {'t':"Building Trustworthy AI,",'s':38,'c':WHITE,'b':True,'sa':2},
    {'t':"Layer by Layer.",'s':38,'c':GOLD,'b':True,'sa':12},
    {'t':"Intelligence, security and compliance — engineered together,",'s':16,'c':SKY,'sa':1},
    {'t':"from the first keystroke to the final, audited response.",'s':16,'c':SKY},
])
# horizontal recap of the skeleton
rw=2.3
for i,(lbl,clr,gly,dark) in enumerate(LAYERS):
    x=1.15+i*(rw+0.05)
    cc=box(s, x, 5.25, rw, 0.6, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(cc,[{'t':f"{gly} {lbl}",'s':10.5,'c':(NAVY if dark else WHITE),'b':True}])
    if i<4: arrow(s, x+rw, 5.55, x+rw+0.05, 5.55, color=GOLD, w=1.5)
text(s, 1.15, 6.4, 11, 0.5, [{'t':"Thank you  ·  AI Technology Team  ·  Questions & discussion welcome",'s':13,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ---------------------------------------------------------------- save
out = "/home/user/Langchain-langgraph-course/AI_Chat_Architecture_BangkokBank.pptx"
prs.save(out)
print("SAVED", out, "slides:", len(prs.slides._sldIdLst))
