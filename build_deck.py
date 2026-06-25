#!/usr/bin/env python3
"""
AI Chat Application Architecture — Simplified (Bangkok Bank InnoHub)
Topology:  CLIENT <-> AI AGENTS <-> AI HUB <-> LLM PROVIDER  (+ DATABASE off the agents)
PII redaction / re-insertion happens at the AI HUB.
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
TEAL_LT   = RGBColor(0xDE, 0xF1, 0xF1)
GREEN     = RGBColor(0x2E, 0x9E, 0x5B)
GREEN_LT  = RGBColor(0xE3, 0xF3, 0xE9)
RED       = RGBColor(0xC0, 0x39, 0x3D)
RED_LT    = RGBColor(0xF7, 0xE4, 0xE4)
ORANGE    = RGBColor(0xE0, 0x82, 0x2B)
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
FONT = "Calibri"
TOTAL = 5

# ---------------------------------------------------------------- helpers
def slide(): return prs.slides.add_slide(BLANK)
def bg(s, color=WHITE):
    s.background.fill.solid(); s.background.fill.fore_color.rgb = color

def blend(c, t, other=WHITE):
    return RGBColor(round(c[0]+(other[0]-c[0])*t), round(c[1]+(other[1]-c[1])*t), round(c[2]+(other[2]-c[2])*t))

def _set_line(shape, color, w=1.0):
    if color is None: shape.line.fill.background()
    else: shape.line.color.rgb = color; shape.line.width = Pt(w)

def box(s, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, shadow=False, radius=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    _set_line(sp, line, lw); sp.shadow.inherit = False
    if shadow:
        spPr = sp._element.spPr
        ef = spPr.find(qn('a:effectLst'))          # inherit=False already created one
        if ef is None:
            ef = spPr.makeelement(qn('a:effectLst'), {}); spPr.append(ef)
        sh = ef.makeelement(qn('a:outerShdw'), {'blurRad':'60000','dist':'30000','dir':'5400000','rotWithShape':'0'})
        clr = ef.makeelement(qn('a:srgbClr'), {'val':'1A2230'}); a = ef.makeelement(qn('a:alpha'), {'val':'30000'})
        clr.append(a); sh.append(clr); ef.append(sh)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try: sp.adjustments[0] = radius
        except Exception: pass
    return sp

def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ('a:buChar','a:buAutoNum'):
        for e in pPr.findall(qn(tag)): pPr.remove(e)
    if pPr.find(qn('a:buNone')) is None: pPr.append(pPr.makeelement(qn('a:buNone'), {}))

def _bullet(p):
    pPr = p._p.get_or_add_pPr(); pPr.set('indent','-182880'); pPr.set('marL','228600')
    pPr.append(pPr.makeelement(qn('a:buFont'), {'typeface':'Arial'}))
    pPr.append(pPr.makeelement(qn('a:buChar'), {'char':'•'}))

def _para(tf, txt, size, color, bold=False, align=PP_ALIGN.LEFT, italic=False, font=FONT, sa=2, sb=0, first=False, bullet=False, ls=None):
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
        _para(tf, ln['t'], ln.get('s',14), ln.get('c',INK), ln.get('b',False), ln.get('a',PP_ALIGN.LEFT),
              ln.get('i',False), ln.get('f',FONT), ln.get('sa',2), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False), ls=ln.get('ls'))
    return tb

def boxtext(sp, lines, anchor=MSO_ANCHOR.MIDDLE, ml=0.08, mr=0.08):
    tf = sp.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=Inches(ml); tf.margin_right=Inches(mr); tf.margin_top=Inches(0.03); tf.margin_bottom=Inches(0.03)
    for i, ln in enumerate(lines):
        _para(tf, ln['t'], ln.get('s',12), ln.get('c',INK), ln.get('b',False), ln.get('a',PP_ALIGN.CENTER),
              ln.get('i',False), ln.get('f',FONT), ln.get('sa',1), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False), ls=ln.get('ls'))
    return sp

def arrow(s, x1, y1, x2, y2, color=BLUE, w=2.25, dash=None, bi=False):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    cn.line.color.rgb=color; cn.line.width=Pt(w); cn.shadow.inherit=False
    le = cn.line._get_or_add_ln()
    le.append(le.makeelement(qn('a:headEnd'), {'type':('triangle' if bi else 'none'),'w':'med','len':'med'}))
    le.append(le.makeelement(qn('a:tailEnd'), {'type':'triangle','w':'med','len':'med'}))
    if dash: le.append(le.makeelement(qn('a:prstDash'), {'val':dash}))
    return cn

def chip(s, x, y, w, h, label, fill, tcolor=WHITE, size=11, bold=True, icon=None, line=None, lw=1.2):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=lw, radius=0.5)
    boxtext(sp, [{'t':(f"{icon}  " if icon else "")+label,'s':size,'c':tcolor,'b':bold}])
    return sp

def icon_circle(s, x, y, d, glyph, fill, gcolor=WHITE, gsize=20, line=None):
    c = box(s, x, y, d, d, fill=fill, line=line, lw=1.2, shape=MSO_SHAPE.OVAL)
    boxtext(c, [{'t':glyph,'s':gsize,'c':gcolor,'b':True}])
    return c

def numbadge(s, x, y, d, n, fill=GOLD, tcolor=NAVY):
    c = box(s, x, y, d, d, fill=fill, line=None, shape=MSO_SHAPE.OVAL, shadow=True)
    boxtext(c, [{'t':str(n),'s':13,'c':tcolor,'b':True}])
    return c

def header(s, kicker, title, n):
    bg(s)
    box(s, 0, 0, 13.333, 1.0, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0, 1.0, 13.333, 0.06, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 0.12, 11.4, 0.34, [{'t':kicker,'s':12,'c':GOLD,'b':True}])
    text(s, 0.5, 0.40, 11.6, 0.6, [{'t':title,'s':24,'c':WHITE,'b':True}])
    bd = box(s, 12.5, 0.26, 0.55, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(bd, [{'t':str(n),'s':16,'c':NAVY,'b':True}])
    # footer
    box(s, 0, 7.16, 13.333, 0.34, fill=GREY_LT, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 7.16, 8, 0.34, [{'t':"Bangkok Bank InnoHub  |  AI Technology Team",'s':9,'c':GREY}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 9.8, 7.16, 3.03, 0.34, [{'t':f"Confidential  ·  {n} / {TOTAL}",'s':9,'c':GREY,'a':PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)

# ---------- shared: the topology diagram ----------
def draw_topology(s, top=2.35, scale_small=False):
    """Draws CLIENT <-> AI AGENTS <-> AI HUB <-> LLM PROVIDER + DATABASE.
       Returns dict of anchor coords for annotation."""
    # main boxes
    cy = top
    cl = box(s, 0.55, cy+0.35, 2.15, 1.3, fill=BLUE, line=None, shadow=True, radius=0.12)
    boxtext(cl, [{'t':"💻",'s':24,'c':WHITE,'sa':2},{'t':"CLIENT",'s':14,'c':WHITE,'b':True},{'t':"web / mobile chat",'s':9,'c':SKY}])
    # AI AGENTS container (taller, holds sub-agents)
    ag = box(s, 3.35, cy, 3.55, 2.4, fill=TEAL_LT, line=TEAL, lw=1.6, shadow=True, radius=0.06)
    boxtext(box(s,3.35,cy,3.55,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE), [{'t':"🤖  AI AGENTS",'s':14,'c':TEAL,'b':True}])
    subs = [("🧭","Router / Orchestrator"),("🔎","Data Retrieval Agent"),("📊","Analysis Agent")]
    for i,(ic,t) in enumerate(subs):
        sc = box(s, 3.55, cy+0.55+i*0.58, 3.15, 0.5, fill=WHITE, line=TEAL, lw=1.0, radius=0.2)
        boxtext(sc, [{'t':[(ic+"  ",TEAL,True,False),(t,NAVY,True,False)],'s':10.5}])
    # AI HUB
    hub = box(s, 7.55, cy+0.2, 2.55, 2.0, fill=GOLD, line=None, shadow=True, radius=0.08)
    boxtext(box(s,7.55,cy+0.28,2.55,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE), [{'t':"🏛️  AI HUB",'s':14,'c':NAVY,'b':True}])
    for i,(ic,t) in enumerate([("🔢","Token counting"),("🎭","PII redact / re-insert"),("🛡️","Guardrails")]):
        hc = box(s, 7.72, cy+0.78+i*0.42, 2.2, 0.36, fill=blend(WHITE,0.12,GOLD_LT), line=None, radius=0.3)
        boxtext(hc, [{'t':f"{ic}  {t}",'s':9.5,'c':NAVY,'b':True}])
    # LLM PROVIDER
    lm = box(s, 10.75, cy+0.35, 2.05, 1.3, fill=PURPLE, line=None, shadow=True, radius=0.12)
    boxtext(lm, [{'t':"🧠",'s':24,'c':WHITE,'sa':2},{'t':"LLM PROVIDER",'s':12.5,'c':WHITE,'b':True},{'t':"generates SQL & text",'s':9,'c':PURPLE_LT}])
    # DATABASE under agents
    db = box(s, 4.0, cy+2.95, 2.25, 1.05, fill=GREEN, line=None, shadow=True, shape=MSO_SHAPE.CAN)
    boxtext(db, [{'t':"🗄️  DATABASE",'s':12,'c':WHITE,'b':True}])
    # bidirectional arrows
    midy = cy+1.0
    arrow(s, 2.72, midy, 3.33, midy, color=NAVY, w=2.5, bi=True)            # client <-> agents
    arrow(s, 6.92, midy, 7.53, midy, color=NAVY, w=2.5, bi=True)            # agents <-> hub
    arrow(s, 10.12, midy, 10.73, midy, color=NAVY, w=2.5, bi=True)         # hub <-> llm
    arrow(s, 5.1, cy+2.42, 5.1, cy+2.93, color=GREEN, w=2.5, bi=True)      # agents <-> db
    return {'cy':cy,'midy':midy}

# ================================================================ SLIDE 1 — Title
s = slide(); bg(s, NAVY)
box(s, 0,0,13.333,7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 8.4,-1.5,7,10.5, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 2.0, 0.12, 2.5, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 0.85, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 2.0, 8.0, 2.2, [
    {'t':"AI Chat Application",'s':44,'c':WHITE,'b':True,'sa':2},
    {'t':"Architecture — Simplified",'s':44,'c':GOLD,'b':True,'sa':12},
    {'t':"How a question flows across the agents, the AI Hub and",'s':18,'c':SKY,'sa':1},
    {'t':"the LLM — with PII protected at the Hub.",'s':18,'c':SKY},
])
# horizontal chain preview
chain=[("💻 CLIENT",BLUE),("🤖 AI AGENTS",TEAL),("🏛️ AI HUB",GOLD),("🧠 LLM",PURPLE)]
x=1.15
for i,(lbl,clr) in enumerate(chain):
    w=2.35
    c=box(s, x, 5.0, w, 0.6, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(c, [{'t':lbl,'s':12,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<3: arrow(s, x+w, 5.3, x+w+0.18, 5.3, color=GOLD, w=2.0, bi=True)
    x+=w+0.18
db=box(s, 3.5+0.05, 5.85, 2.35, 0.55, fill=GREEN, line=None, shadow=True, shape=MSO_SHAPE.CAN)
boxtext(db, [{'t':"🗄️ DATABASE",'s':11,'c':WHITE,'b':True}])
arrow(s, 4.67, 5.62, 4.67, 5.83, color=GREEN, w=2.0, bi=True)
text(s, 0.95, 6.65, 9, 0.5, [{'t':"Executive Briefing  ·  AI Technology Team  ·  2026",'s':12.5,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ================================================================ SLIDE 2 — Architecture
s = slide(); header(s, "ARCHITECTURE", "The Building Blocks", 2)
text(s, 0.55, 1.2, 12.3, 0.35, [{'t':"Four components in a simple chain — the Database sits next to the agents.",'s':12.5,'c':GREY,'i':True}])
anch = draw_topology(s, top=2.05)
# annotations: guardrails & PII location
icon_circle(s, 2.80, 2.50, 0.46, "🛡️", RED, WHITE, 13)
lc = box(s, 0.9, 4.55, 2.95, 0.85, fill=RED_LT, line=RED, lw=1.0, radius=0.1)
boxtext(lc, [{'t':"🛡️ Input & output guardrails run here — no PII redaction at this step",'s':9.5,'c':INK}], ml=0.12, mr=0.12)
rc = box(s, 7.3, 4.55, 3.1, 0.85, fill=GOLD_LT, line=GOLD, lw=1.0, radius=0.1)
boxtext(rc, [{'t':"🎭 PII is redacted before the LLM and re-inserted after — only at the AI Hub",'s':9.5,'c':INK}], ml=0.12, mr=0.12)
# legend strip
ly=6.55
for i,(lbl,clr) in enumerate([("Request / Response (bi-directional)",NAVY),("PII boundary at AI Hub",GOLD),("Guardrails",RED),("Data access",GREEN)]):
    x=0.6+i*3.1
    box(s, x, ly+0.06, 0.3, 0.16, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, x+0.4, ly-0.04, 2.7, 0.4, [{'t':lbl,'s':9.5,'c':INK}])

# ================================================================ SLIDE 3 — End-to-end flow
s = slide(); header(s, "END-TO-END FLOW", "How One Request Is Answered", 3)
text(s, 0.55, 1.2, 12.3, 0.32, [{'t':"Eleven steps from the user's question to the answer on screen.",'s':12.5,'c':GREY,'i':True}])
LANE = {'CLIENT':BLUE,'AGENTS':TEAL,'AI HUB':GOLD,'LLM':PURPLE,'DB':GREEN,'GUARD':RED}
steps = [
 (1,"GUARD","User submits the request; the first input guardrail runs (no PII redaction yet)."),
 (2,"AGENTS","Router agent reads the prompt and routes it to the Data Retrieval agent first."),
 (3,"AI HUB","Data Retrieval agent sends the query request to the AI Hub."),
 (4,"AI HUB","AI Hub counts input tokens, runs guardrails, and redacts PII."),
 (5,"LLM","LLM provider generates the SQL from the redacted prompt."),
 (6,"AI HUB","AI Hub re-inserts the PII into the generated SQL."),
 (7,"AGENTS","AI Hub sends the SQL back to the Retrieval agent."),
 (8,"DB","Retrieval agent runs the SQL against the Database to fetch the data."),
 (9,"AGENTS","Analysis agent analyses the data and writes the final natural-language response."),
 (10,"GUARD","Output guardrail checks the response before it leaves."),
 (11,"CLIENT","Client shows the result to the user."),
]
def step_row(x, y, w, n, lane, desc):
    numbadge(s, x, y+0.06, 0.5, n, fill=LANE[lane], tcolor=(NAVY if lane in ('AI HUB',) else WHITE))
    lab = {'GUARD':'GUARDRAIL'}.get(lane, lane)
    lc = box(s, x+0.62, y+0.1, 1.45, 0.42, fill=blend(LANE[lane],0.78), line=LANE[lane], lw=1.0, radius=0.4)
    boxtext(lc, [{'t':lab,'s':9,'c':(RED if lane=='GUARD' else NAVY),'b':True}])
    text(s, x+2.18, y, w-2.2, 0.66, [{'t':desc,'s':10.8,'c':INK,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
colw=6.0; rh=0.84
for i,(n,lane,desc) in enumerate(steps):
    col = 0 if i<6 else 1
    row = i if i<6 else i-6
    x = 0.55 + col*(colw+0.25)
    y = 1.65 + row*rh
    step_row(x, y, colw, n, lane, desc)

# ================================================================ SLIDE 4 — Guardrails & PII
s = slide(); header(s, "SECURITY", "Where Guardrails & PII Protection Happen", 4)
text(s, 0.55, 1.2, 12.3, 0.35, [{'t':"Three control points keep the flow safe — PII is handled only at the AI Hub.",'s':12.5,'c':GREY,'i':True}])
cards = [
 ("🛡️","1 · Input Guardrail","at CLIENT → AGENTS",
  ["Content & policy moderation","Prompt-injection / jailbreak checks","Scope & intent validation","No PII redaction at this step"], RED, RED_LT),
 ("🏛️","2 · AI Hub (PII boundary)","every LLM call",
  ["Counts input / output tokens","Redacts PII before the LLM","LLM sees placeholders only","Re-inserts PII into the result"], GOLD, GOLD_LT),
 ("✅","3 · Output Guardrail","at AGENTS → CLIENT",
  ["Checks for leaked PII","Toxicity & safety screening","Grounded / no hallucination","Required disclosures attached"], GREEN, GREEN_LT),
]
cw=3.95; cx=0.6
for i,(ic,t,sub,items,ac,bgc) in enumerate(cards):
    x=cx+i*(cw+0.2); y=1.75
    c=box(s, x, y, cw, 4.0, fill=WHITE, line=ac, lw=1.4, shadow=True, radius=0.05)
    box(s, x, y, cw, 0.95, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    icon_circle(s, x+0.25, y+0.22, 0.55, ic, WHITE, ac, 17)
    text(s, x+0.92, y+0.16, cw-1.05, 0.4, [{'t':t,'s':13.5,'c':(NAVY if ac==GOLD else WHITE),'b':True,'sa':0}])
    text(s, x+0.92, y+0.55, cw-1.05, 0.3, [{'t':sub,'s':10,'c':(NAVY if ac==GOLD else SKY),'i':True}])
    yy=y+1.2
    for it in items:
        icon_circle(s, x+0.28, yy, 0.28, "•", bgc, ac, 12)
        text(s, x+0.68, yy-0.06, cw-0.9, 0.5, [{'t':it,'s':11,'c':INK,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
        yy+=0.62
# bottom banner
bn=box(s, 0.6, 6.0, 12.15, 0.72, fill=NAVY, line=None, radius=0.1)
boxtext(bn, [{'t':[("Key point:  ",GOLD,True,False),("real customer PII is exposed to the LLM only as redacted placeholders — the AI Hub re-inserts the real values, so sensitive data never leaves the trust boundary.",WHITE,False,False)],'s':12}], ml=0.3, mr=0.3)

# ================================================================ SLIDE 5 — Closing
s = slide(); bg(s, NAVY)
box(s, 0,0,13.333,7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, -2,4.5,18,3.3, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 2.2, 0.12, 2.0, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 1.1, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 2.2, 9.5, 1.8, [
    {'t':"Simple Chain, Strong Controls.",'s':36,'c':WHITE,'b':True,'sa':10},
    {'t':"Client → Agents → AI Hub → LLM, with the Database beside the agents",'s':16,'c':SKY,'sa':1},
    {'t':"and PII protected at the Hub on every model call.",'s':16,'c':SKY},
])
# recap chain
chain=[("💻 CLIENT",BLUE),("🤖 AI AGENTS",TEAL),("🏛️ AI HUB",GOLD),("🧠 LLM",PURPLE)]
x=1.15
for i,(lbl,clr) in enumerate(chain):
    w=2.5
    c=box(s, x, 5.05, w, 0.6, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(c, [{'t':lbl,'s':12,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<3: arrow(s, x+w, 5.35, x+w+0.2, 5.35, color=GOLD, w=2.0, bi=True)
    x+=w+0.2
db=box(s, 3.6, 5.95, 2.5, 0.55, fill=GREEN, line=None, shadow=True, shape=MSO_SHAPE.CAN)
boxtext(db, [{'t':"🗄️ DATABASE",'s':11,'c':WHITE,'b':True}])
arrow(s, 4.85, 5.7, 4.85, 5.93, color=GREEN, w=2.0, bi=True)
text(s, 1.15, 6.7, 11, 0.5, [{'t':"Thank you  ·  AI Technology Team  ·  Questions welcome",'s':13,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ---------------------------------------------------------------- save
out = "/home/user/Langchain-langgraph-course/AI_Chat_Architecture_BangkokBank.pptx"
prs.save(out)
print("SAVED", out, "slides:", len(prs.slides._sldIdLst))
