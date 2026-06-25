#!/usr/bin/env python3
"""
Generates: End-to-End AI Chat Application Architecture (Bangkok Bank InnoHub)
A rich, diagram-driven executive briefing deck.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
NAVY      = RGBColor(0x0B, 0x1F, 0x3A)   # deep banking navy
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

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

FONT = "Calibri"
FONT_H = "Calibri"

# ---------------------------------------------------------------- helpers
def slide():
    return prs.slides.add_slide(BLANK)

def bg(s, color=WHITE):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color

def _set_line(shape, color, w=1.0):
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = color
        shape.line.width = Pt(w)

def box(s, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        shadow=False, radius=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    _set_line(sp, line, lw)
    sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = el.makeelement(qn('a:outerShdw'),
                            {'blurRad':'60000','dist':'30000','dir':'5400000','rotWithShape':'0'})
        clr = el.makeelement(qn('a:srgbClr'), {'val':'1A2230'})
        alpha = el.makeelement(qn('a:alpha'), {'val':'32000'})
        clr.append(alpha); sh.append(clr); ef.append(sh); el.append(ef)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp

def _para(tf, text, size, color, bold=False, align=PP_ALIGN.LEFT, italic=False,
          font=FONT, space_after=2, space_before=0, first=False, bullet=False,
          level=0, line_spacing=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    p.level = level
    if line_spacing:
        p.line_spacing = line_spacing
    runs = text if isinstance(text, list) else [(text, color, bold, italic)]
    for seg in runs:
        t, c, b, i = (seg + (color, bold, italic))[:4] if isinstance(seg, tuple) else (seg, color, bold, italic)
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.bold = b; r.font.italic = i
        r.font.name = font; r.font.color.rgb = c
    if bullet:
        _bullet_char(p)
    else:
        _no_bullet(p)
    return p

def _no_bullet(p):
    pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
    for tag in ('a:buChar','a:buAutoNum'):
        for e in pPr.findall(qn(tag)):
            pPr.remove(e)
    if pPr.find(qn('a:buNone')) is None:
        pPr.append(pPr.makeelement(qn('a:buNone'), {}))

def _bullet_char(p, char="•", color="1E5AA8"):
    pPr = p.get_or_add_pPr()
    pPr.set('indent', '-182880'); pPr.set('marL', '228600')
    bf = pPr.makeelement(qn('a:buFont'), {'typeface':'Arial'})
    bc = pPr.makeelement(qn('a:buChar'), {'char':char})
    pPr.append(bf); pPr.append(bc)

def text(s, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, wrap=True, shrink=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    for i, ln in enumerate(lines):
        _para(tf, ln['t'], ln.get('s',14), ln.get('c',INK), ln.get('b',False),
              ln.get('a',PP_ALIGN.LEFT), ln.get('i',False), ln.get('f',FONT),
              ln.get('sa',2), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False),
              level=ln.get('lvl',0), line_spacing=ln.get('ls'))
    return tb

def boxtext(sp, lines, anchor=MSO_ANCHOR.MIDDLE, ml=0.08, mr=0.08, mt=0.04, mb=0.04):
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left=Inches(ml); tf.margin_right=Inches(mr)
    tf.margin_top=Inches(mt); tf.margin_bottom=Inches(mb)
    for i, ln in enumerate(lines):
        _para(tf, ln['t'], ln.get('s',12), ln.get('c',INK), ln.get('b',False),
              ln.get('a',PP_ALIGN.CENTER), ln.get('i',False), ln.get('f',FONT),
              ln.get('sa',1), ln.get('sb',0), first=(i==0), bullet=ln.get('bu',False),
              line_spacing=ln.get('ls'))
    return sp

def arrow(s, x1, y1, x2, y2, color=BLUE, w=2.25, dash=None):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    cn.line.color.rgb = color; cn.line.width = Pt(w)
    cn.shadow.inherit = False
    le = cn.line._get_or_add_ln()
    he = le.makeelement(qn('a:headEnd'), {'type':'none'})
    te = le.makeelement(qn('a:tailEnd'), {'type':'triangle','w':'med','len':'med'})
    le.append(he); le.append(te)
    if dash:
        d = le.makeelement(qn('a:prstDash'), {'val':dash}); le.append(d)
    return cn

def chip(s, x, y, w, h, label, fill, tcolor=WHITE, size=11, bold=True, icon=None):
    sp = box(s, x, y, w, h, fill=fill, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    boxtext(sp, [{'t': (f"{icon}  " if icon else "")+label, 's':size, 'c':tcolor, 'b':bold}])
    return sp

def header(s, kicker, title, n):
    box(s, 0, 0, 13.333, 1.18, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0, 1.18, 13.333, 0.07, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.55, 0.12, 11.5, 0.4, [{'t':kicker,'s':12,'c':GOLD,'b':True,'f':FONT}])
    text(s, 0.55, 0.42, 11.6, 0.7, [{'t':title,'s':25,'c':WHITE,'b':True,'f':FONT_H}])
    # slide number badge
    bd = box(s, 12.5, 0.34, 0.55, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(bd, [{'t':str(n),'s':16,'c':NAVY,'b':True}])

def footer(s, n):
    box(s, 0, 7.18, 13.333, 0.32, fill=GREY_LT, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 7.18, 8, 0.32, [{'t':"Bangkok Bank InnoHub  |  AI Technology Team",
         's':9,'c':GREY,'b':False}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 10.0, 7.18, 2.83, 0.32, [{'t':f"Confidential  ·  {n} / 16",
         's':9,'c':GREY,'a':PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)

def icon_circle(s, x, y, d, glyph, fill, gcolor=WHITE, gsize=20):
    c = box(s, x, y, d, d, fill=fill, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(c, [{'t':glyph,'s':gsize,'c':gcolor,'b':True}])
    return c

# ================================================================ SLIDE 1 — Title
s = slide(); bg(s, NAVY)
# layered background geometry
box(s, 0, 0, 13.333, 7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 8.6, -1.5, 7, 10.5, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 10.3, -1.5, 5, 10.5, fill=BLUE, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
# gold rule
box(s, 0.9, 2.05, 0.12, 2.7, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
# brand row
chip(s, 0.95, 0.85, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 2.0, 8.2, 2.4, [
    {'t':"End-to-End AI Chat",'s':46,'c':WHITE,'b':True,'sa':2,'f':FONT_H},
    {'t':"Application Architecture",'s':46,'c':GOLD,'b':True,'sa':10,'f':FONT_H},
    {'t':"From User Input to Intelligent Response —",'s':19,'c':SKY,'b':False,'sa':1},
    {'t':"with Governance, PII Protection & Security",'s':19,'c':SKY,'b':False},
])
# right-side mini stack visual
for i,(lbl,clr) in enumerate([("Client",BLUE_LT),("Agent Orchestrator",TEAL),
        ("AI Hub Gateway",GOLD),("LLM Providers",PURPLE),("Data & RAG",GREEN)]):
    yy = 1.9 + i*0.62
    cc = box(s, 10.55, yy, 2.25, 0.5, fill=clr, line=None, radius=0.5)
    boxtext(cc, [{'t':lbl,'s':11,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<4:
        arrow(s, 11.67, yy+0.5, 11.67, yy+0.62, color=WHITE, w=1.5)
text(s, 0.95, 6.4, 9, 0.5, [{'t':"Executive Briefing  ·  AI Technology Team  ·  2026",
     's':13,'c':GREY,'b':False}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ================================================================ SLIDE 2 — Executive Summary
s = slide(); bg(s); header(s, "EXECUTIVE SUMMARY", "Why Architecture Matters for Enterprise AI", 2); footer(s,2)
# key message banner
mb = box(s, 0.55, 1.45, 12.25, 0.92, fill=SKY, line=BLUE, lw=1.0, radius=0.12)
boxtext(mb, [
    {'t':[("Key message:  ", BLUE, True, False),
          ("Enterprise AI applications require a sophisticated, multi-layered architecture that "
           "balances intelligence, security, compliance, and performance.", INK, False, False)],'s':14.5}],
    anchor=MSO_ANCHOR.MIDDLE, ml=0.25, mr=0.25)
cards = [
    ("🧠","Intelligent Agents","User input flows through agents that plan and orchestrate complex, multi-step workflows.", BLUE, SKY),
    ("⚙️","Dynamic Tooling","Agentic systems generate tools on-the-fly — such as SQL queries — under strict PII protection.", TEAL, GREEN_LT),
    ("🛡️","Multi-Stage Guardrails","Input, processing and output guardrails enforce compliance at every stage of the pipeline.", RED, RED_LT),
    ("🏛️","Central AI Hub","A gateway enforces policy, PII masking, quota and provider routing in one control point.", GOLD, GOLD_LT),
    ("📊","Full Observability","End-to-end logging, tracing and monitoring maintain complete audit trails.", PURPLE, PURPLE_LT),
    ("✅","Trust by Design","Security and governance are built in — not bolted on after deployment.", GREEN, GREEN_LT),
]
cx, cy, cw, ch = 0.55, 2.62, 3.94, 1.95
for i,(ic,t,d,ac,bgc) in enumerate(cards):
    col = i%3; row = i//3
    x = cx + col*(cw+0.18); y = cy + row*(ch+0.16)
    c = box(s, x, y, cw, ch, fill=WHITE, line=RGBColor(0xDD,0xE3,0xEC), lw=1.0, shadow=True, radius=0.06)
    box(s, x, y, 0.12, ch, fill=ac, line=None, shape=MSO_SHAPE.RECTANGLE)
    ico = box(s, x+0.28, y+0.24, 0.66, 0.66, fill=bgc, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(ico, [{'t':ic,'s':22,'c':ac}])
    text(s, x+1.06, y+0.22, cw-1.2, 0.5, [{'t':t,'s':14.5,'c':NAVY,'b':True}])
    text(s, x+0.3, y+0.96, cw-0.55, ch-1.0, [{'t':d,'s':11.5,'c':GREY,'ls':1.02}])

# ================================================================ SLIDE 3 — System Architecture Overview
s = slide(); bg(s); header(s, "SYSTEM ARCHITECTURE", "Five-Layer Reference Architecture", 3); footer(s,3)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"Requests flow left → right; responses flow right → left. Guardrails sit at every boundary.",
     's':12.5,'c':GREY,'i':True}])
# 5 vertical layers
layers = [
    ("1  CLIENT", "Web Application", "💬", BLUE, "User chat UI\nSession & auth\nResponse render"),
    ("2  ORCHESTRATION", "Agent Planner", "🧠", TEAL, "Task planning\nTool calling\nState / memory"),
    ("3  AI HUB GATEWAY", "Central Control", "🏛️", GOLD, "Policy engine\nPII masking\nQuota & routing"),
    ("4  PROVIDERS", "External LLMs", "🤖", PURPLE, "OpenAI · Claude\nGemini · Local\nFail-over"),
]
lx, ly, lw, lh = 0.55, 1.95, 2.78, 3.05
for i,(tag,name,ic,clr,body) in enumerate(layers):
    x = lx + i*(lw+0.30)
    card = box(s, x, ly, lw, lh, fill=WHITE, line=clr, lw=1.6, shadow=True, radius=0.05)
    box(s, x, ly, lw, 0.5, fill=clr, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,ly,lw,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':tag,'s':11.5,'c':WHITE,'b':True}])
    ico = box(s, x+lw/2-0.42, ly+0.7, 0.84, 0.84, fill=clr, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(ico, [{'t':ic,'s':28,'c':WHITE}])
    text(s, x+0.1, ly+1.62, lw-0.2, 0.4, [{'t':name,'s':14,'c':NAVY,'b':True,'a':PP_ALIGN.CENTER}])
    tb = text(s, x+0.18, ly+2.06, lw-0.36, 0.95, [], anchor=MSO_ANCHOR.TOP)
    tf = tb.text_frame
    for j,line in enumerate(body.split("\n")):
        _para(tf, line, 11, GREY, align=PP_ALIGN.CENTER, first=(j==0), space_after=2)
    # connecting arrows with guardrail marks
    if i<3:
        ax = x+lw
        arrow(s, ax+0.02, ly+lh/2-0.12, ax+0.28, ly+lh/2-0.12, color=BLUE, w=2.5)
        arrow(s, ax+0.28, ly+lh/2+0.12, ax+0.02, ly+lh/2+0.12, color=GREEN, w=2.5)
        gm = box(s, ax+0.04, ly+lh/2+0.34, 0.22, 0.22, fill=RED, line=None, shape=MSO_SHAPE.OVAL)
        boxtext(gm, [{'t':"🛡",'s':9,'c':WHITE}])
# legend for arrows
text(s, 8.95, 4.62, 4.1, 0.4, [
    {'t':[("▶ request  ",BLUE,True,False),("◀ response  ",GREEN,True,False),("🛡 guardrail",RED,True,False)],'s':10}])
# bottom data layer band
db = box(s, 0.55, 5.28, 12.25, 1.15, fill=NAVY, line=None, shadow=True, radius=0.05)
boxtext(box(s,0.7,5.36,3.0,0.4,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':"5   DATA & KNOWLEDGE LAYER",'s':12.5,'c':GOLD,'b':True,'a':PP_ALIGN.LEFT}])
for j,(ic,lbl) in enumerate([("🗄️","Transactional Database"),("📚","RAG Knowledge Base"),
        ("📄","Document Store"),("📈","Logs & Audit Trail")]):
    bx = 0.85 + j*3.0
    dchip = box(s, bx, 5.78, 2.78, 0.5, fill=NAVY2, line=GOLD, lw=1.0, radius=0.4)
    boxtext(dchip, [{'t':f"{ic}  {lbl}",'s':11.5,'c':WHITE,'b':True}])
# arrow orchestration->data
arrow(s, 4.6, 5.0, 4.6, 5.26, color=GOLD, w=2.5)
arrow(s, 5.6, 5.26, 5.6, 5.0, color=GOLD, w=2.5)

# ================================================================ SLIDE 4 — The User Query
s = slide(); bg(s); header(s, "STEP 1 · USER QUERY", "The User Initiates a Request", 4); footer(s,4)
# Left: chat mockup
ph = box(s, 0.7, 1.6, 5.0, 5.1, fill=GREY_LT, line=RGBColor(0xCF,0xD6,0xE0), lw=1.2, shadow=True, radius=0.04)
box(s, 0.7, 1.6, 5.0, 0.7, fill=NAVY, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
icon_circle(s, 0.92, 1.74, 0.42, "🤖", GOLD, NAVY, 16)
text(s, 1.45, 1.66, 4.0, 0.6, [
    {'t':"InnoHub Assistant",'s':13,'c':WHITE,'b':True,'sa':0},
    {'t':"● Online · Secure channel",'s':9,'c':GREEN_LT}])
# assistant greeting bubble
ab = box(s, 0.95, 2.55, 3.2, 0.62, fill=WHITE, line=RGBColor(0xDD,0xE3,0xEC), lw=1.0, radius=0.25)
boxtext(ab, [{'t':"Hi! How can I help you today?",'s':10.5,'c':INK,'a':PP_ALIGN.LEFT}], ml=0.14, mr=0.14)
# user bubble
ub = box(s, 1.4, 3.45, 4.05, 1.5, fill=BLUE, line=None, radius=0.18)
boxtext(ub, [{'t':"Show me transaction history for customer Somchai Jantapan with account number 1234567890 from last month",
     's':12,'c':WHITE,'a':PP_ALIGN.LEFT,'b':False}], anchor=MSO_ANCHOR.MIDDLE, ml=0.18, mr=0.18)
text(s, 1.4, 5.0, 4.05, 0.3, [{'t':"You · just now",'s':9,'c':GREY,'a':PP_ALIGN.RIGHT}])
# input bar
ibar = box(s, 0.95, 5.95, 4.5, 0.55, fill=WHITE, line=RGBColor(0xCF,0xD6,0xE0), lw=1.0, radius=0.5)
boxtext(ibar, [{'t':"Type a message…",'s':10.5,'c':GREY,'a':PP_ALIGN.LEFT}], ml=0.16)
icon_circle(s, 4.95, 5.99, 0.46, "➤", BLUE, WHITE, 13)

# Right: what happens
text(s, 6.1, 1.55, 6.6, 0.5, [{'t':"What happens behind the scenes",'s':16,'c':NAVY,'b':True}])
steps4 = [
    ("①","User submits query","Natural-language request entered in the secure web application.", BLUE),
    ("②","Received by Orchestrator","The Agent Orchestrator captures the message and opens a traced session.", TEAL),
    ("③","Intent recognised","Agent detects this needs database retrieval + tool calling, not a simple chat reply.", PURPLE),
    ("④","Next: PII handling","PII identification and redaction begin before any external call is made.", RED),
]
yy = 2.15
for ic,t,d,clr in steps4:
    icon_circle(s, 6.1, yy+0.05, 0.5, ic, clr, WHITE, 16)
    text(s, 6.78, yy, 6.0, 0.45, [{'t':t,'s':14,'c':NAVY,'b':True,'sa':0}])
    text(s, 6.78, yy+0.42, 5.95, 0.6, [{'t':d,'s':11.5,'c':GREY,'ls':1.02}])
    yy += 1.16
# callout
co = box(s, 6.1, 6.75, 6.7, 0.0+0.0, fill=None, line=None)

# ================================================================ SLIDE 5 — Input Guardrails & PII
s = slide(); bg(s); header(s, "STEP 2 · INPUT GUARDRAILS", "Input Validation & PII Redaction", 5); footer(s,5)
# four scan checks
text(s, 0.55, 1.32, 12, 0.35, [{'t':"Input guardrails scan every query across four dimensions before it proceeds:",'s':12.5,'c':GREY,'i':True}])
checks = [("🔐","PII Detection","names · accounts · emails",RED,RED_LT),
          ("⚠️","Prompt Injection","jailbreak / override attempts",ORANGE,ORANGE_LT),
          ("🚫","Toxic Language","harmful / abusive content",PURPLE,PURPLE_LT),
          ("📋","Policy Violations","out-of-scope requests",BLUE,SKY)]
for i,(ic,t,d,ac,bgc) in enumerate(checks):
    x = 0.55 + i*3.12
    c = box(s, x, 1.72, 2.95, 0.95, fill=bgc, line=ac, lw=1.2, radius=0.1)
    boxtext(c, [{'t':[(ic+"  ",ac,True,False),(t,NAVY,True,False)],'s':12.5},
                {'t':d,'s':9.5,'c':GREY,'sb':1}], anchor=MSO_ANCHOR.MIDDLE)
# redaction flow
fy = 2.95
orig = box(s, 0.55, fy, 3.85, 1.5, fill=WHITE, line=RED, lw=1.4, shadow=True, radius=0.05)
text(s, 0.7, fy+0.08, 3.6, 0.3, [{'t':"ORIGINAL QUERY",'s':10,'c':RED,'b':True}])
boxtext(box(s,0.68,fy+0.4,3.6,1.0,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':[("Show me transaction history for customer ",INK,False,False),
               ("Somchai Jantapan",RED,True,False),(" with account number ",INK,False,False),
               ("1234567890",RED,True,False),(" from last month",INK,False,False)],'s':11,'a':PP_ALIGN.LEFT}],
        anchor=MSO_ANCHOR.TOP)
# scanner
sc = box(s, 4.62, fy+0.25, 1.55, 1.0, fill=NAVY, line=GOLD, lw=1.4, shadow=True, radius=0.12)
boxtext(sc, [{'t':"🛡️",'s':20,'c':WHITE,'sa':0},{'t':"PII SCANNER",'s':10,'c':GOLD,'b':True}])
arrow(s, 4.4, fy+0.75, 4.6, fy+0.75, color=GREY, w=2.5)
arrow(s, 6.17, fy+0.75, 6.4, fy+0.75, color=GREY, w=2.5)
# redacted
red = box(s, 6.42, fy, 6.38, 1.5, fill=WHITE, line=GREEN, lw=1.4, shadow=True, radius=0.05)
text(s, 6.57, fy+0.08, 6.0, 0.3, [{'t':"REDACTED QUERY  ✓ safe to process",'s':10,'c':GREEN,'b':True}])
boxtext(box(s,6.57,fy+0.4,6.1,1.0,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':[("Show me transaction history for customer ",INK,False,False),
               ("CUSTOMER_NAME_001",GREEN,True,False),(" with account number ",INK,False,False),
               ("ACCOUNT_ID_001",GREEN,True,False),(" from last month",INK,False,False)],'s':11,'a':PP_ALIGN.LEFT}],
        anchor=MSO_ANCHOR.TOP)
# mapping table
ty = 4.72
text(s, 0.55, ty, 6, 0.35, [{'t':"PII Mapping Table (held securely for re-insertion)",'s':12.5,'c':NAVY,'b':True}])
rows = [("Original Data","PII Type","Placeholder",True,NAVY,GOLD),
        ("Somchai Jantapan","Customer Name","CUSTOMER_NAME_001",False,WHITE,INK),
        ("1234567890","Account ID","ACCOUNT_ID_001",False,SKY,INK),
        ("\"last month\"","Temporal context","(no redaction needed)",False,WHITE,GREY)]
cw3 = [4.3, 3.6, 4.4]; tx = 0.55; tyy = ty+0.42
for r,(a,b,cc,hdr,rbg,fg) in enumerate(rows):
    xx = tx
    for ci,val in enumerate((a,b,cc)):
        cell = box(s, xx, tyy, cw3[ci], 0.46, fill=(NAVY if hdr else rbg),
                   line=RGBColor(0xD5,0xDC,0xE6), lw=0.75, shape=MSO_SHAPE.RECTANGLE)
        col = GOLD if hdr else (RED if (ci==2 and r in (1,2)) else fg)
        boxtext(cell, [{'t':val,'s':11,'c':col,'b':(hdr or (ci==2 and r in(1,2))),'a':PP_ALIGN.LEFT}], ml=0.14)
        xx += cw3[ci]
    tyy += 0.46

# ================================================================ SLIDE 6 — Agent Planner
s = slide(); bg(s); header(s, "STEP 3 · AGENT PLANNER", "The Orchestrator Analyzes the Task", 6); footer(s,6)
text(s, 0.55, 1.34, 12.2, 0.4, [{'t':"The Agent Planner receives the redacted query and decomposes it into an executable plan.",'s':12.5,'c':GREY,'i':True}])
# central brain
brain = box(s, 5.35, 1.9, 2.6, 1.0, fill=TEAL, line=None, shadow=True, radius=0.2)
boxtext(brain, [{'t':"🧠  AGENT PLANNER",'s':13,'c':WHITE,'b':True},{'t':"reason · plan · route",'s':10,'c':GREEN_LT}])
# decision attributes
dec = [
    ("Task Type","Database retrieval +\nLLM-assisted SQL generation", BLUE, SKY, 0.55, 3.2),
    ("Required Tools","SQL query tool · RAG retriever ·\nResponse formatter", PURPLE, PURPLE_LT, 4.79, 3.2),
    ("Data Sources","Transaction DB ·\nKnowledge base", GREEN, GREEN_LT, 9.03, 3.2),
    ("Constraints","Read-only · row-level scope ·\nPII placeholders preserved", RED, RED_LT, 0.55, 4.55),
    ("Execution Plan","1) Generate SQL  2) Validate\n3) Execute  4) Synthesize answer", GOLD, GOLD_LT, 4.79, 4.55),
    ("Governance","Route via AI Hub ·\nlog every step", NAVY, GREY_LT, 9.03, 4.55),
]
for t,d,ac,bgc,x,y in dec:
    c = box(s, x, y, 3.75, 1.15, fill=bgc, line=ac, lw=1.2, shadow=True, radius=0.08)
    box(s, x, y, 3.75, 0.42, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,y,3.75,0.42,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':t,'s':12,'c':(NAVY if ac==GOLD else WHITE),'b':True}])
    tb = text(s, x+0.18, y+0.48, 3.4, 0.62, [], anchor=MSO_ANCHOR.TOP)
    for j,line in enumerate(d.split("\n")):
        _para(tb.text_frame, line, 10.5, INK, align=PP_ALIGN.LEFT, first=(j==0), space_after=1)
    # connector from brain
    arrow(s, 6.65, 2.9, x+1.87, y-0.02, color=RGBColor(0xB9,0xC2,0xCE), w=1.25, dash='dash')

# ================================================================ SLIDE 7 — Tool Generation (SQL)
s = slide(); bg(s); header(s, "STEP 4 · AGENTIC TOOLING", "Dynamic Tool Generation — SQL Query", 7); footer(s,7)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"The agent writes its own tool: a parameterized SQL query built from the redacted, placeholdered request.",'s':12.5,'c':GREY,'i':True}])
# left: generated SQL code block
code = box(s, 0.55, 1.85, 6.55, 3.05, fill=RGBColor(0x10,0x1A,0x2B), line=NAVY2, lw=1.0, shadow=True, radius=0.04)
box(s, 0.55, 1.85, 6.55, 0.42, fill=NAVY2, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
for dx,clr in [(0.78,RED),(0.98,GOLD),(1.18,GREEN)]:
    box(s, dx, 1.99, 0.14, 0.14, fill=clr, line=None, shape=MSO_SHAPE.OVAL)
boxtext(box(s,1.5,1.85,5,0.42,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':"generated_query.sql",'s':10.5,'c':SKY,'b':False,'a':PP_ALIGN.LEFT}])
sqlb = text(s, 0.8, 2.42, 6.1, 2.4, [], anchor=MSO_ANCHOR.TOP)
MONO = "Consolas"
sql_lines = [
    [("SELECT ",BLUE_LT,True),("txn_date, amount, channel, balance",WHITE,False)],
    [("FROM ",BLUE_LT,True),("transactions t",WHITE,False)],
    [("JOIN ",BLUE_LT,True),("customers c ",WHITE,False),("ON ",BLUE_LT,True),("c.id = t.cust_id",WHITE,False)],
    [("WHERE ",BLUE_LT,True),("c.name = ",WHITE,False),(":CUSTOMER_NAME_001",GOLD,True)],
    [("  AND ",BLUE_LT,True),("t.account = ",WHITE,False),(":ACCOUNT_ID_001",GOLD,True)],
    [("  AND ",BLUE_LT,True),("t.txn_date >= ",WHITE,False),("date_trunc('month',",GREEN_LT,False)],
    [("        current_date - interval '1 month')",GREEN_LT,False)],
    [("ORDER BY ",BLUE_LT,True),("t.txn_date ",WHITE,False),("DESC",BLUE_LT,True),(";",WHITE,False)],
]
for j,segs in enumerate(sql_lines):
    p = sqlb.text_frame.paragraphs[0] if j==0 else sqlb.text_frame.add_paragraph()
    p.space_after = Pt(1); _no_bullet(p)
    for tt,cc,bb in segs:
        r=p.add_run(); r.text=tt; r.font.size=Pt(11.5); r.font.name=MONO; r.font.color.rgb=cc; r.font.bold=bb
# right: safeguards
text(s, 7.35, 1.85, 5.5, 0.4, [{'t':"Built-in safeguards on every generated tool",'s':14,'c':NAVY,'b':True}])
guards = [
    ("🔒","Bind PII placeholders","Parameters stay as :PLACEHOLDER tokens — real values are never embedded in the query string.", RED),
    ("👁️","Read-only enforcement","Only SELECT is permitted; INSERT / UPDATE / DELETE / DDL are rejected by the policy layer.", BLUE),
    ("🧪","Static validation","Query is parsed & validated against schema and row-level access rules before execution.", PURPLE),
    ("⏱️","Limits & timeouts","Row caps, query timeout and cost ceilings protect the production database.", GREEN),
]
yy=2.4
for ic,t,d,clr in guards:
    icon_circle(s, 7.35, yy, 0.44, ic, clr, WHITE, 14)
    text(s, 7.95, yy-0.04, 4.9, 0.35, [{'t':t,'s':12.5,'c':NAVY,'b':True,'sa':0}])
    text(s, 7.95, yy+0.32, 4.85, 0.6, [{'t':d,'s':10.5,'c':GREY,'ls':1.0}])
    yy+=1.06
# bottom flow strip
by=5.25
for i,(lbl,clr) in enumerate([("Redacted intent",TEAL),("Generate SQL",BLUE),("Validate & bind",PURPLE),
        ("Execute (read-only)",GREEN),("Result set",GOLD)]):
    x=0.55+i*2.5
    c=box(s, x, by, 2.3, 0.62, fill=clr, line=None, radius=0.4)
    boxtext(c, [{'t':lbl,'s':11,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<4: arrow(s, x+2.3, by+0.31, x+2.5, by+0.31, color=GREY, w=2.5)
text(s, 0.55, 6.05, 12.2, 0.5, [{'t':[("Result:  ",GREEN,True,False),
     ("rows are returned still keyed to placeholders — real customer data is re-inserted only at the final, secured output stage.",INK,False,False)],'s':12}])

# ================================================================ SLIDE 8 — AI Hub Gateway
s = slide(); bg(s); header(s, "STEP 5 · AI HUB GATEWAY", "Central Policy & Control Plane", 8); footer(s,8)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"Every call to an external model passes through one governed gateway — the single point of control.",'s':12.5,'c':GREY,'i':True}])
# left input
li = box(s, 0.55, 2.4, 2.2, 2.0, fill=TEAL, line=None, shadow=True, radius=0.08)
boxtext(li, [{'t':"🧠",'s':24,'c':WHITE,'sa':2},{'t':"Agent",'s':13,'c':WHITE,'b':True},
             {'t':"redacted request",'s':10,'c':GREEN_LT}])
# central hub
hub = box(s, 3.35, 1.85, 5.6, 4.6, fill=GOLD_LT, line=GOLD, lw=1.6, shadow=True, radius=0.04)
box(s, 3.35, 1.85, 5.6, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
boxtext(box(s,3.35,1.85,5.6,0.55,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':"🏛️  AI HUB GATEWAY",'s':15,'c':NAVY,'b':True}])
hub_mod = [
    ("🛂","Policy Engine","Access control, allow/deny lists, request classification & routing rules"),
    ("🎭","PII Masking Layer","Verifies redaction, applies extra masking, blocks leakage to providers"),
    ("📦","Quota & Rate Mgmt","Per-team budgets, token quotas, throttling and cost attribution"),
    ("🔀","Provider Router","Selects best model, load-balances and fails over across providers"),
    ("🧾","Logging & Tracing","Captures every prompt, response, token count and policy decision"),
]
my = 2.55
for ic,t,d in hub_mod:
    m = box(s, 3.6, my, 5.1, 0.72, fill=WHITE, line=RGBColor(0xE0,0xD3,0xA8), lw=1.0, radius=0.1)
    icon_circle(s, 3.72, my+0.13, 0.46, ic, NAVY, GOLD, 14)
    text(s, 4.32, my+0.06, 4.3, 0.32, [{'t':t,'s':11.5,'c':NAVY,'b':True,'sa':0}])
    text(s, 4.32, my+0.36, 4.3, 0.32, [{'t':d,'s':9,'c':GREY}])
    my += 0.78
# right providers
prov = box(s, 9.55, 2.4, 3.25, 2.0, fill=PURPLE, line=None, shadow=True, radius=0.08)
boxtext(box(s,9.55,2.45,3.25,0.4,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':"🤖  LLM PROVIDERS",'s':12.5,'c':WHITE,'b':True}])
for j,p in enumerate(["OpenAI GPT","Anthropic Claude","Google Gemini","Private / On-prem"]):
    pc = box(s, 9.75, 2.92+j*0.36, 2.85, 0.3, fill=PURPLE_LT, line=None, radius=0.5)
    boxtext(pc, [{'t':p,'s':9.5,'c':PURPLE,'b':True}])
# arrows
arrow(s, 2.75, 3.2, 3.33, 3.2, color=TEAL, w=2.75)
arrow(s, 8.97, 3.2, 9.53, 3.2, color=PURPLE, w=2.75)
arrow(s, 9.53, 3.6, 8.97, 3.6, color=GREEN, w=2.25)
arrow(s, 3.33, 3.6, 2.75, 3.6, color=GREEN, w=2.25)
# value strip bottom
vb = box(s, 0.55, 6.55, 12.25, 0.0, fill=None, line=None)
text(s, 0.55, 4.6, 2.2, 1.8, [{'t':"⬅ response\nflows back\nmasked & logged",'s':10,'c':GREEN,'b':True,'a':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
text(s, 9.55, 4.6, 3.25, 1.8, [{'t':"One gateway =\none audit point,\none policy surface",'s':10.5,'c':PURPLE,'b':True,'a':PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 9 — Output Guardrails & Re-insertion
s = slide(); bg(s); header(s, "STEP 6 · OUTPUT GUARDRAILS", "Response Assembly & PII Re-insertion", 9); footer(s,9)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"The model answers using placeholders. Real data is restored only here — inside the trust boundary.",'s':12.5,'c':GREY,'i':True}])
# stage row
fy=2.0
stg = [
    ("LLM Response","…history for CUSTOMER_NAME_001, account ACCOUNT_ID_001: 14 transactions…", PURPLE, PURPLE_LT),
    ("Output Guardrails","Check for leaked PII, hallucination, toxicity, policy & disclosure compliance", RED, RED_LT),
    ("PII Re-insertion","Restore real values from the secure mapping table held by the orchestrator", GREEN, GREEN_LT),
    ("Final Answer","…history for Somchai Jantapan, account 1234567890: 14 transactions…", BLUE, SKY),
]
bw=2.92
for i,(t,d,ac,bgc) in enumerate(stg):
    x=0.55+i*(bw+0.13)
    c=box(s, x, fy, bw, 2.0, fill=bgc, line=ac, lw=1.3, shadow=True, radius=0.06)
    box(s, x, fy, bw, 0.5, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,fy,bw,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':f"{i+1}. {t}",'s':12,'c':(WHITE),'b':True}])
    boxtext(box(s,x+0.12,fy+0.55,bw-0.24,1.4,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':d,'s':10.5,'c':INK,'a':PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.TOP)
    if i<3: arrow(s, x+bw, fy+1.0, x+bw+0.13, fy+1.0, color=GREY, w=2.5)
# checks under output guardrail
text(s, 0.55, 4.35, 12.2, 0.35, [{'t':"Output guardrail checks",'s':13,'c':NAVY,'b':True}])
oc = [("🔍","No PII leakage","placeholders all resolved / none exposed externally",RED),
      ("🎯","Grounded & accurate","answer matches retrieved data; no fabrication",BLUE),
      ("🧹","Safe content","no toxic, biased or non-compliant language",PURPLE),
      ("📜","Disclosure rules","required disclaimers & data-use notices attached",GREEN)]
for i,(ic,t,d,clr) in enumerate(oc):
    x=0.55+i*3.12
    c=box(s, x, 4.75, 2.95, 1.4, fill=WHITE, line=clr, lw=1.2, shadow=True, radius=0.08)
    icon_circle(s, x+0.18, 4.92, 0.5, ic, clr, WHITE, 15)
    text(s, x+0.78, 4.86, 2.1, 0.5, [{'t':t,'s':11.5,'c':NAVY,'b':True,'sa':0}])
    text(s, x+0.18, 5.5, 2.6, 0.6, [{'t':d,'s':10,'c':GREY,'ls':1.0}])

# ================================================================ SLIDE 10 — End-to-End Sequence
s = slide(); bg(s); header(s, "END-TO-END FLOW", "Complete Request Lifecycle — Sequence View", 10); footer(s,10)
actors = [("User",BLUE,1.4),("Orchestrator",TEAL,3.85),("AI Hub",GOLD,6.3),("LLM",PURPLE,8.75),("Data/RAG",GREEN,11.2)]
top=1.75
for name,clr,x in actors:
    h=box(s, x-1.0, top, 2.0, 0.5, fill=clr, line=None, radius=0.3)
    boxtext(h, [{'t':name,'s':11.5,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    # lifeline
    ll=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x),Inches(top+0.5),Inches(x),Inches(6.95))
    ll.line.color.rgb=RGBColor(0xC4,0xCC,0xD6); ll.line.width=Pt(1.0)
    ll.shadow.inherit=False
    led=ll.line._get_or_add_ln(); led.append(led.makeelement(qn('a:prstDash'),{'val':'dash'}))
msgs = [
    (1.4,3.85,"1 · query (raw)",BLUE),
    (3.85,3.85,"2 · input guardrails + PII redaction",RED),
    (3.85,11.2,"3 · retrieve (generated SQL / RAG)",GREEN),
    (11.2,3.85,"4 · placeholdered result set",GREEN),
    (3.85,6.3,"5 · masked prompt → gateway",TEAL),
    (6.3,8.75,"6 · policy ✓, route to provider",GOLD),
    (8.75,6.3,"7 · model response (placeholders)",PURPLE),
    (6.3,3.85,"8 · logged + masked response",GOLD),
    (3.85,3.85,"9 · output guardrails + PII re-insertion",RED),
    (3.85,1.4,"10 · final answer (real data)",BLUE),
]
yy=2.45
for x1,x2,lbl,clr in msgs:
    if x1==x2:
        # self-call
        box(s, x1, yy-0.02, 1.2, 0.32, fill=None, line=clr, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
        text(s, x1+0.15, yy-0.18, 5.2, 0.3, [{'t':lbl,'s':10,'c':INK,'b':True,'a':PP_ALIGN.LEFT}])
    else:
        arrow(s, x1, yy+0.12, x2, yy+0.12, color=clr, w=2.0)
        midx=min(x1,x2);
        text(s, min(x1,x2)+0.1, yy-0.2, abs(x2-x1)+0.2, 0.3,
             [{'t':lbl,'s':10,'c':INK,'b':True,'a':PP_ALIGN.CENTER if abs(x2-x1)>1.5 else PP_ALIGN.LEFT}])
    yy+=0.46
text(s, 0.55, 6.95, 12, 0.3, [{'t':[("🛡 Guardrails fire at steps 2 & 9   ·   ",RED,True,False),
     ("🏛 Gateway governs steps 5–8   ·   ",GOLD,True,False),("🎭 Real PII exists only outside the dashed trust boundary",GREEN,True,False)],'s':10}])

# ================================================================ SLIDE 11 — Observability
s = slide(); bg(s); header(s, "OBSERVABILITY", "Logging, Monitoring & Audit Trails", 11); footer(s,11)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"Every layer emits structured telemetry into a unified observability pipeline.",'s':12.5,'c':GREY,'i':True}])
# pipeline
src = ["Client","Orchestrator","AI Hub","Providers","Data Layer"]
sy=1.95
for i,nm in enumerate(src):
    x=0.55+i*1.55
    c=box(s, x, sy, 1.42, 0.55, fill=NAVY, line=None, radius=0.3)
    boxtext(c, [{'t':nm,'s':9.5,'c':WHITE,'b':True}])
    arrow(s, x+0.71, sy+0.55, 4.0+0.0, sy+0.95 if False else sy+0.9, color=RGBColor(0xC4,0xCC,0xD6), w=1.0)
collector = box(s, 0.55, 2.95, 7.4, 0.6, fill=GOLD, line=None, shadow=True, radius=0.2)
boxtext(collector, [{'t':"📥  Telemetry Collector  —  structured logs · traces · metrics · events",'s':12,'c':NAVY,'b':True}])
# three pillars
pil = [("📈","Metrics","Latency, token usage, cost per request, error & throttle rates, model mix",BLUE,SKY),
       ("🧵","Traces","End-to-end span per request across all five layers with correlation IDs",PURPLE,PURPLE_LT),
       ("🧾","Audit Logs","Immutable record of prompts, policy decisions, PII actions & approvals",GREEN,GREEN_LT)]
for i,(ic,t,d,ac,bgc) in enumerate(pil):
    x=0.55+i*4.18
    c=box(s, x, 3.85, 3.95, 1.55, fill=bgc, line=ac, lw=1.2, shadow=True, radius=0.06)
    icon_circle(s, x+0.2, 4.02, 0.55, ic, ac, WHITE, 16)
    text(s, x+0.85, 3.98, 3.0, 0.4, [{'t':t,'s':14,'c':NAVY,'b':True}])
    text(s, x+0.22, 4.62, 3.55, 0.7, [{'t':d,'s':10.5,'c':GREY,'ls':1.02}])
    arrow(s, 4.25, 3.55, x+1.97, 3.83, color=GOLD, w=1.5, dash='dash')
# outcomes strip
out=box(s, 0.55, 5.6, 12.25, 1.15, fill=SKY, line=BLUE, lw=1.0, radius=0.06)
boxtext(box(s,0.7,5.66,4,0.4,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
        [{'t':"What this enables",'s':12.5,'c':BLUE,'b':True,'a':PP_ALIGN.LEFT}])
for i,(ic,t) in enumerate([("🔎","Full traceability for every customer interaction"),
        ("⚖️","Regulatory & audit readiness (BOT, PDPA)"),
        ("💰","Cost & quota accountability per team"),
        ("🚨","Real-time anomaly & abuse detection")]):
    x=0.85+ (i%2)*6.1; y=6.05+(i//2)*0.42
    text(s, x, y, 6.0, 0.35, [{'t':[(ic+"  ",INK,False,False),(t,INK,False,False)],'s':11}])

# ================================================================ SLIDE 12 — Security & Compliance
s = slide(); bg(s); header(s, "SECURITY & COMPLIANCE", "Defense-in-Depth Across the Stack", 12); footer(s,12)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"Controls are layered so that no single failure exposes customer data.",'s':12.5,'c':GREY,'i':True}])
layers12 = [
    ("🔐","Data Protection","PII redaction & tokenization · encryption in transit & at rest · secure mapping vault", RED),
    ("🛡️","Guardrails","Input & output filtering · prompt-injection defense · content & toxicity screening", ORANGE),
    ("🏛️","Governance Gateway","Centralized policy · provider allow-lists · quota enforcement · approval workflows", GOLD),
    ("👤","Access Control","Role-based access · row-level data scoping · least-privilege tool permissions", BLUE),
    ("📜","Compliance","PDPA & BOT alignment · audit trails · data-residency & retention policies", PURPLE),
    ("🔭","Monitoring","Continuous logging · anomaly detection · incident response & alerting", GREEN),
]
for i,(ic,t,d,clr) in enumerate(layers12):
    col=i%2; row=i//2
    x=0.55+col*6.25; y=1.85+row*1.6
    c=box(s, x, y, 6.05, 1.42, fill=WHITE, line=RGBColor(0xDD,0xE3,0xEC), lw=1.0, shadow=True, radius=0.06)
    box(s, x, y, 0.14, 1.42, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, x+0.32, y+0.42, 0.6, ic, clr, WHITE, 18)
    text(s, x+1.1, y+0.18, 4.8, 0.4, [{'t':t,'s':14.5,'c':NAVY,'b':True}])
    text(s, x+1.1, y+0.66, 4.8, 0.7, [{'t':d,'s':10.8,'c':GREY,'ls':1.02}])

# ================================================================ SLIDE 13 — Tech building blocks
s = slide(); bg(s); header(s, "IMPLEMENTATION", "Reference Building Blocks", 13); footer(s,13)
text(s, 0.55, 1.32, 12.2, 0.4, [{'t':"Indicative technology choices per layer — provider-agnostic and swappable.",'s':12.5,'c':GREY,'i':True}])
cols = [
    ("Client", BLUE, SKY, ["Web / mobile chat UI","Auth (OAuth2 / SSO)","Session management","Streaming responses"]),
    ("Orchestration", TEAL, GREEN_LT, ["LangChain / LangGraph","Agent + tool calling","State & memory store","Retry & fallback"]),
    ("AI Hub Gateway", GOLD, GOLD_LT, ["API gateway / proxy","Policy engine","PII masking service","Quota & key vault"]),
    ("Providers", PURPLE, PURPLE_LT, ["OpenAI · Claude","Gemini · open models","On-prem / private","Embedding models"]),
    ("Data & RAG", GREEN, GREEN_LT, ["SQL / warehouse","Vector database","Document store","Audit log sink"]),
]
cw=2.42; cx=0.55
for i,(t,ac,bgc,items) in enumerate(cols):
    x=cx+i*(cw+0.06)
    c=box(s, x, 1.95, cw, 4.55, fill=bgc, line=ac, lw=1.2, shadow=True, radius=0.04)
    box(s, x, 1.95, cw, 0.62, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,1.95,cw,0.62,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':t,'s':12.5,'c':(NAVY if ac==GOLD else WHITE),'b':True}])
    yy=2.75
    for it in items:
        ic=box(s, x+0.15, yy, cw-0.3, 0.82, fill=WHITE, line=None, shadow=False, radius=0.1)
        boxtext(ic, [{'t':it,'s':10.5,'c':INK,'b':True}])
        yy+=0.92

# ================================================================ SLIDE 14 — Risks & Mitigations
s = slide(); bg(s); header(s, "RISK MANAGEMENT", "Key Risks & Mitigations", 14); footer(s,14)
risks = [
    ("PII / data leakage","Redaction + tokenization, output-side leak checks, secure mapping vault, gateway masking", RED),
    ("Prompt injection","Input guardrails, instruction isolation, tool allow-lists, read-only data access", ORANGE),
    ("Hallucination","RAG grounding, output validation against source data, confidence & disclosure rules", PURPLE),
    ("Provider outage","Multi-provider routing, automatic fail-over, graceful degradation via the AI Hub", BLUE),
    ("Cost / abuse","Per-team quotas, rate limiting, token budgets, anomaly detection & alerting", GREEN),
    ("Compliance gaps","Immutable audit trails, PDPA/BOT controls, data residency & retention policies", TEAL),
]
# header row
hr=2.0
box(s, 0.55, hr, 4.2, 0.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 4.75, hr, 8.05, 0.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
boxtext(box(s,0.55,hr,4.2,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),[{'t':"RISK",'s':12,'c':GOLD,'b':True,'a':PP_ALIGN.LEFT}],ml=0.2)
boxtext(box(s,4.75,hr,8.05,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),[{'t':"MITIGATION",'s':12,'c':GOLD,'b':True,'a':PP_ALIGN.LEFT}],ml=0.2)
yy=hr+0.5
for i,(r,m,clr) in enumerate(risks):
    rbg = WHITE if i%2==0 else GREY_LT
    box(s, 0.55, yy, 4.2, 0.72, fill=rbg, line=RGBColor(0xE2,0xE6,0xEC), lw=0.75, shape=MSO_SHAPE.RECTANGLE)
    box(s, 4.75, yy, 8.05, 0.72, fill=rbg, line=RGBColor(0xE2,0xE6,0xEC), lw=0.75, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0.55, yy, 0.1, 0.72, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    boxtext(box(s,0.7,yy,4.0,0.72,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':r,'s':11.5,'c':NAVY,'b':True,'a':PP_ALIGN.LEFT}],ml=0.12)
    boxtext(box(s,4.9,yy,7.8,0.72,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':m,'s':11,'c':INK,'a':PP_ALIGN.LEFT}],ml=0.12)
    yy+=0.72

# ================================================================ SLIDE 15 — Key Takeaways
s = slide(); bg(s); header(s, "KEY TAKEAWAYS", "What to Remember", 15); footer(s,15)
tk = [
    ("①","Layered by design","Five cooperating layers turn a single user query into a governed, intelligent response.", BLUE),
    ("②","PII never leaves the boundary","Redaction in, re-insertion out — external models only ever see placeholders.", RED),
    ("③","Guardrails at every stage","Input, processing and output controls make compliance continuous, not occasional.", ORANGE),
    ("④","One gateway to govern them all","The AI Hub centralizes policy, masking, routing, quotas and logging.", GOLD),
    ("⑤","Observable end-to-end","Metrics, traces and audit logs deliver trust, accountability and regulatory readiness.", GREEN),
]
yy=1.7
for ic,t,d,clr in tk:
    c=box(s, 0.7, yy, 12.0, 0.96, fill=WHITE, line=RGBColor(0xDD,0xE3,0xEC), lw=1.0, shadow=True, radius=0.08)
    box(s, 0.7, yy, 0.14, 0.96, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, 0.98, yy+0.2, 0.56, ic, clr, WHITE, 20)
    text(s, 1.75, yy+0.13, 4.7, 0.7, [{'t':t,'s':15,'c':NAVY,'b':True}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 6.45, yy+0.13, 6.0, 0.7, [{'t':d,'s':11.5,'c':GREY,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
    yy+=1.06

# ================================================================ SLIDE 16 — Closing
s = slide(); bg(s, NAVY)
box(s, 0, 0, 13.333, 7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, -2, 4.6, 18, 3, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 2.5, 0.12, 2.2, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 1.4, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 2.5, 11, 2.0, [
    {'t':"Building Trustworthy AI,",'s':40,'c':WHITE,'b':True,'sa':2,'f':FONT_H},
    {'t':"Layer by Layer.",'s':40,'c':GOLD,'b':True,'sa':12,'f':FONT_H},
    {'t':"Intelligence, security and compliance — engineered together,",'s':17,'c':SKY,'sa':1},
    {'t':"from the first keystroke to the final, audited response.",'s':17,'c':SKY},
])
# mini recap chips
for i,(lbl) in enumerate(["Client","Orchestration","AI Hub","Providers","Data & RAG"]):
    x=1.15+i*2.35
    cc=box(s, x, 5.35, 2.2, 0.55, fill=NAVY2, line=GOLD, lw=1.0, radius=0.4)
    boxtext(cc, [{'t':lbl,'s':11,'c':WHITE,'b':True}])
    if i<4: arrow(s, x+2.2, 5.62, x+2.35, 5.62, color=GOLD, w=1.5)
text(s, 1.15, 6.5, 11, 0.5, [{'t':"Thank you  ·  AI Technology Team  ·  Questions & discussion welcome",'s':13,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ---------------------------------------------------------------- save
out = "/home/user/Langchain-langgraph-course/AI_Chat_Architecture_BangkokBank.pptx"
prs.save(out)
print("SAVED", out, "slides:", len(prs.slides._sldIdLst))
