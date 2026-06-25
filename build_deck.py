#!/usr/bin/env python3
"""
AI Chat Application Architecture — Full Walkthrough (Bangkok Bank InnoHub)
Topology:  CLIENT <-> AI AGENTS <-> AI HUB <-> LLM PROVIDER  (+ DATABASE)
A persistent mini-chain skeleton appears on every step slide; the active
component is highlighted as we zoom into each of the 11 steps.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
NAVY      = RGBColor(0x0B, 0x1F, 0x3A); NAVY2 = RGBColor(0x14, 0x2E, 0x53)
BLUE      = RGBColor(0x1E, 0x5A, 0xA8); BLUE_LT = RGBColor(0x3C, 0x82, 0xD6)
SKY       = RGBColor(0xE8, 0xF1, 0xFB)
GOLD      = RGBColor(0xC9, 0xA2, 0x27); GOLD_LT = RGBColor(0xF2, 0xE6, 0xBE)
TEAL      = RGBColor(0x1A, 0x9C, 0x9C); TEAL_LT = RGBColor(0xDE, 0xF1, 0xF1)
GREEN     = RGBColor(0x2E, 0x9E, 0x5B); GREEN_LT = RGBColor(0xE3, 0xF3, 0xE9)
RED       = RGBColor(0xC0, 0x39, 0x3D); RED_LT = RGBColor(0xF7, 0xE4, 0xE4)
ORANGE    = RGBColor(0xE0, 0x82, 0x2B)
PURPLE    = RGBColor(0x6B, 0x4C, 0xA8); PURPLE_LT = RGBColor(0xEC, 0xE6, 0xF6)
GREY      = RGBColor(0x5A, 0x66, 0x77); GREY_LT = RGBColor(0xF1, 0xF3, 0xF6)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF); INK = RGBColor(0x1A, 0x22, 0x30)
MUTE      = RGBColor(0x8A, 0x93, 0x9F)

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]; FONT = "Calibri"; MONO = "Consolas"
TOTAL = 17

# ---------------------------------------------------------------- helpers
def slide(): return prs.slides.add_slide(BLANK)
def bg(s, color=WHITE): s.background.fill.solid(); s.background.fill.fore_color.rgb = color
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
        ef = spPr.find(qn('a:effectLst'))
        if ef is None: ef = spPr.makeelement(qn('a:effectLst'), {}); spPr.append(ef)
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
    if dash: le.append(le.makeelement(qn('a:prstDash'), {'val':dash}))
    le.append(le.makeelement(qn('a:headEnd'), {'type':('triangle' if bi else 'none'),'w':'med','len':'med'}))
    le.append(le.makeelement(qn('a:tailEnd'), {'type':'triangle','w':'med','len':'med'}))
    return cn

def chip(s, x, y, w, h, label, fill, tcolor=WHITE, size=11, bold=True, icon=None, line=None, lw=1.2):
    sp = box(s, x, y, w, h, fill=fill, line=line, lw=lw, radius=0.5)
    boxtext(sp, [{'t':(f"{icon}  " if icon else "")+label,'s':size,'c':tcolor,'b':bold}])
    return sp

def icon_circle(s, x, y, d, glyph, fill, gcolor=WHITE, gsize=20, line=None):
    c = box(s, x, y, d, d, fill=fill, line=line, lw=1.2, shape=MSO_SHAPE.OVAL)
    boxtext(c, [{'t':glyph,'s':gsize,'c':gcolor,'b':True}]); return c

def header(s, kicker, title, page, badge=None):
    bg(s)
    box(s, 0, 0, 13.333, 1.0, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0, 1.0, 13.333, 0.06, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 0.12, 11.4, 0.34, [{'t':kicker,'s':12,'c':GOLD,'b':True}])
    text(s, 0.5, 0.40, 11.6, 0.6, [{'t':title,'s':23,'c':WHITE,'b':True}])
    bd = box(s, 12.5, 0.26, 0.55, 0.55, fill=GOLD, line=None, shape=MSO_SHAPE.OVAL)
    boxtext(bd, [{'t':str(badge if badge is not None else page),'s':15,'c':NAVY,'b':True}])
    box(s, 0, 7.16, 13.333, 0.34, fill=GREY_LT, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.5, 7.16, 8, 0.34, [{'t':"Bangkok Bank InnoHub  |  AI Technology Team",'s':9,'c':GREY}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 9.8, 7.16, 3.03, 0.34, [{'t':f"Confidential  ·  {page} / {TOTAL}",'s':9,'c':GREY,'a':PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)

# ---------- persistent mini-chain skeleton ----------
CHAIN = {'CLIENT':('💻 CLIENT',BLUE,0.7,2.4),'AGENTS':('🤖 AI AGENTS',TEAL,3.5,2.9),
         'HUB':('🏛️ AI HUB',GOLD,6.8,2.6),'LLM':('🧠 LLM',PURPLE,9.8,2.85)}
def minichain(s, active=set(), guard=False, y=1.3):
    nh=0.6
    for a,b,x1,x2 in [('CLIENT','AGENTS',3.1,3.5),('AGENTS','HUB',6.4,6.8),('HUB','LLM',9.4,9.8)]:
        on = a in active and b in active
        arrow(s, x1, y+nh/2, x2, y+nh/2, color=(GOLD if on else RGBColor(0xC4,0xCC,0xD6)), w=(3.0 if on else 1.5), bi=True)
    for k,(lbl,clr,x,w) in CHAIN.items():
        on = k in active
        sp=box(s,x,y,w,nh,fill=(clr if on else blend(clr,0.7)),line=(GOLD if on else None),lw=2.5,shadow=on,radius=0.5)
        boxtext(sp,[{'t':lbl,'s':12,'c':((NAVY if clr==GOLD else WHITE) if on else MUTE),'b':True}])
    on='DB' in active; dbx,dbw=4.05,1.8
    arrow(s, dbx+dbw/2, y+nh, dbx+dbw/2, y+0.92, color=(GOLD if on else RGBColor(0xC4,0xCC,0xD6)), w=(3.0 if on else 1.5), bi=True)
    db=box(s,dbx,y+0.92,dbw,0.5,fill=(GREEN if on else blend(GREEN,0.7)),line=(GOLD if on else None),lw=2.5,shadow=on,shape=MSO_SHAPE.CAN)
    boxtext(db,[{'t':"🗄️ DATABASE",'s':9.5,'c':(WHITE if on else MUTE),'b':True}])
    if guard:
        icon_circle(s, 3.04, y-0.26, 0.42, "🛡️", RED, WHITE, 12)
        text(s, 2.5, y+0.62, 1.5, 0.3, [{'t':"guardrail",'s':8,'c':RED,'b':True,'a':PP_ALIGN.CENTER}])

# ---------- example-card frame ----------
EX_X, EX_Y, EX_W, EX_H = 7.45, 3.0, 5.35, 3.55
def card_frame(s, title, accent, x=EX_X, y=EX_Y, w=EX_W, h=EX_H):
    box(s, x, y, w, h, fill=WHITE, line=blend(accent,0.4), lw=1.3, shadow=True, radius=0.04)
    box(s, x, y, w, 0.5, fill=accent, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    boxtext(box(s,x,y,w,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE),
            [{'t':title,'s':12,'c':(NAVY if accent==GOLD else WHITE),'b':True,'a':PP_ALIGN.LEFT}], anchor=MSO_ANCHOR.MIDDLE, ml=0.2)
    return x+0.22, y+0.62, w-0.44

def codeblock(s, x, y, w, lines, h=2.3):
    box(s, x, y, w, h, fill=RGBColor(0x10,0x1A,0x2B), line=NAVY2, lw=1.0, radius=0.04)
    tb = text(s, x+0.18, y+0.14, w-0.3, h-0.2, [])
    for j,segs in enumerate(lines):
        p = tb.text_frame.paragraphs[0] if j==0 else tb.text_frame.add_paragraph()
        p.space_after = Pt(1); _no_bullet(p)
        for tt,cc,bb in segs:
            r=p.add_run(); r.text=tt; r.font.size=Pt(10.5); r.font.name=MONO; r.font.color.rgb=cc; r.font.bold=bb
    return tb

SQL_PH = [[("SELECT ",BLUE_LT,True),("txn_date, amount, channel",WHITE,False)],
          [("FROM ",BLUE_LT,True),("transactions",WHITE,False)],
          [("WHERE ",BLUE_LT,True),("name = ",WHITE,False),(":NAME_001",GOLD,True)],
          [("  AND ",BLUE_LT,True),("acct = ",WHITE,False),(":ACCT_001",GOLD,True)],
          [("  AND ",BLUE_LT,True),("month = last",WHITE,False)]]
SQL_REAL = [[("SELECT ",BLUE_LT,True),("txn_date, amount, channel",WHITE,False)],
            [("FROM ",BLUE_LT,True),("transactions",WHITE,False)],
            [("WHERE ",BLUE_LT,True),("name = ",WHITE,False),("'Somchai J.'",GREEN_LT,True)],
            [("  AND ",BLUE_LT,True),("acct = ",WHITE,False),("'1234567890'",GREEN_LT,True)],
            [("  AND ",BLUE_LT,True),("month = last",WHITE,False)]]

# ================================================================ SLIDE 1 — Title
s = slide(); bg(s, NAVY)
box(s, 0,0,13.333,7.5, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE)
box(s, 8.4,-1.5,7,10.5, fill=NAVY2, line=None, shape=MSO_SHAPE.PARALLELOGRAM)
box(s, 0.9, 1.95, 0.12, 2.5, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)
chip(s, 0.95, 0.8, 3.2, 0.5, "BANGKOK BANK · InnoHub", GOLD, NAVY, size=12, icon="◆")
text(s, 1.15, 1.95, 8.0, 2.2, [
    {'t':"AI Chat Application",'s':44,'c':WHITE,'b':True,'sa':2},
    {'t':"Architecture Walkthrough",'s':44,'c':GOLD,'b':True,'sa':12},
    {'t':"A step-by-step journey from the user's question to the",'s':18,'c':SKY,'sa':1},
    {'t':"answer — with PII protected at the AI Hub.",'s':18,'c':SKY},
])
chain=[("💻 CLIENT",BLUE),("🤖 AI AGENTS",TEAL),("🏛️ AI HUB",GOLD),("🧠 LLM",PURPLE)]
x=1.15
for i,(lbl,clr) in enumerate(chain):
    w=2.35; c=box(s, x, 4.95, w, 0.6, fill=clr, line=None, shadow=True, radius=0.5)
    boxtext(c, [{'t':lbl,'s':12,'c':(NAVY if clr==GOLD else WHITE),'b':True}])
    if i<3: arrow(s, x+w, 5.25, x+w+0.18, 5.25, color=GOLD, w=2.0, bi=True)
    x+=w+0.18
db=box(s, 3.55, 5.8, 2.35, 0.55, fill=GREEN, line=None, shadow=True, shape=MSO_SHAPE.CAN)
boxtext(db, [{'t':"🗄️ DATABASE",'s':11,'c':WHITE,'b':True}])
arrow(s, 4.72, 5.57, 4.72, 5.78, color=GREEN, w=2.0, bi=True)
text(s, 0.95, 6.6, 9, 0.5, [{'t':"Executive Briefing  ·  AI Technology Team  ·  2026",'s':12.5,'c':GREY}])
box(s, 0, 7.18, 13.333, 0.32, fill=GOLD, line=None, shape=MSO_SHAPE.RECTANGLE)

# ================================================================ SLIDE 2 — Agenda
s = slide(); header(s, "AGENDA", "What We'll Cover", 2)
items=[("🧩","The Architecture","Four components — Client, AI Agents, AI Hub, LLM — plus the Database.",BLUE,SKY),
       ("①","The 11-Step Journey","Follow one real request through the whole system, step by step.",TEAL,TEAL_LT),
       ("🔐","Guardrails & PII","Where checks run and how PII is protected at the AI Hub.",GOLD,GOLD_LT),
       ("✅","Key Takeaways","The few ideas worth remembering.",GREEN,GREEN_LT)]
for i,(ic,t,d,ac,bgc) in enumerate(items):
    y=1.55+i*1.32
    c=box(s, 0.7, y, 11.95, 1.15, fill=WHITE, line=blend(ac,0.4), lw=1.0, shadow=True, radius=0.06)
    box(s, 0.7, y, 0.14, 1.15, fill=ac, line=None, shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, 1.0, y+0.28, 0.6, ic, ac, (NAVY if ac==GOLD else WHITE), 18)
    text(s, 1.85, y+0.18, 4.0, 0.5, [{'t':t,'s':16,'c':NAVY,'b':True}])
    text(s, 1.85, y+0.62, 10.5, 0.45, [{'t':d,'s':12,'c':GREY}])

# ================================================================ SLIDE 3 — Architecture
s = slide(); header(s, "ARCHITECTURE", "The Building Blocks", 3)
text(s, 0.55, 1.18, 12.3, 0.35, [{'t':"Four components in a simple chain — the Database sits next to the agents.",'s':12.5,'c':GREY,'i':True}])
cy=2.0
cl = box(s, 0.55, cy+0.35, 2.15, 1.3, fill=BLUE, line=None, shadow=True, radius=0.12)
boxtext(cl, [{'t':"💻",'s':24,'c':WHITE,'sa':2},{'t':"CLIENT",'s':14,'c':WHITE,'b':True},{'t':"web / mobile chat",'s':9,'c':SKY}])
box(s, 3.35, cy, 3.55, 2.4, fill=TEAL_LT, line=TEAL, lw=1.6, shadow=True, radius=0.06)
boxtext(box(s,3.35,cy,3.55,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE), [{'t':"🤖  AI AGENTS",'s':14,'c':TEAL,'b':True}])
for i,(ic,t) in enumerate([("🧭","Router / Orchestrator"),("🔎","Data Retrieval Agent"),("📊","Analysis Agent")]):
    sc = box(s, 3.55, cy+0.55+i*0.58, 3.15, 0.5, fill=WHITE, line=TEAL, lw=1.0, radius=0.2)
    boxtext(sc, [{'t':[(ic+"  ",TEAL,True,False),(t,NAVY,True,False)],'s':10.5}])
box(s, 7.55, cy+0.2, 2.55, 2.0, fill=GOLD, line=None, shadow=True, radius=0.08)
boxtext(box(s,7.55,cy+0.28,2.55,0.5,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE), [{'t':"🏛️  AI HUB",'s':14,'c':NAVY,'b':True}])
for i,(ic,t) in enumerate([("🔢","Token counting"),("🎭","PII redact / re-insert"),("🛡️","Guardrails")]):
    hc = box(s, 7.72, cy+0.78+i*0.42, 2.2, 0.36, fill=GOLD_LT, line=None, radius=0.3)
    boxtext(hc, [{'t':f"{ic}  {t}",'s':9.5,'c':NAVY,'b':True}])
lm = box(s, 10.75, cy+0.35, 2.05, 1.3, fill=PURPLE, line=None, shadow=True, radius=0.12)
boxtext(lm, [{'t':"🧠",'s':24,'c':WHITE,'sa':2},{'t':"LLM PROVIDER",'s':12.5,'c':WHITE,'b':True},{'t':"generates SQL & text",'s':9,'c':PURPLE_LT}])
db = box(s, 4.0, cy+2.95, 2.25, 1.05, fill=GREEN, line=None, shadow=True, shape=MSO_SHAPE.CAN)
boxtext(db, [{'t':"🗄️  DATABASE",'s':12,'c':WHITE,'b':True}])
midy=cy+1.0
arrow(s, 2.72, midy, 3.33, midy, color=NAVY, w=2.5, bi=True)
arrow(s, 6.92, midy, 7.53, midy, color=NAVY, w=2.5, bi=True)
arrow(s, 10.12, midy, 10.73, midy, color=NAVY, w=2.5, bi=True)
arrow(s, 5.1, cy+2.42, 5.1, cy+2.93, color=GREEN, w=2.5, bi=True)
lc = box(s, 0.9, 6.5, 2.95, 0.55, fill=RED_LT, line=RED, lw=1.0, radius=0.1)
boxtext(lc, [{'t':"🛡️ Guardrails at the client edge — no PII redaction here",'s':9,'c':INK}], ml=0.1, mr=0.1)
rc = box(s, 7.3, 6.5, 3.1, 0.55, fill=GOLD_LT, line=GOLD, lw=1.0, radius=0.1)
boxtext(rc, [{'t':"🎭 PII redacted & re-inserted only at the AI Hub",'s':9,'c':INK}], ml=0.1, mr=0.1)

# ================================================================ STEP SLIDES (4–14)
STEPS = [
 (1,{'CLIENT'},True,BLUE,"User Sends a Request","Client submits the prompt; the first input guardrail runs",
    ["User types a natural-language question in the chat app.","Input guardrail checks content, policy & prompt-injection.","No PII redaction at this step — that happens at the AI Hub.","Clean requests are passed to the AI Agents."]),
 (2,{'AGENTS'},False,TEAL,"Router Agent Plans the Work","Router understands intent and routes to the Data Retrieval Agent",
    ["Router / Orchestrator interprets the user's intent.","Decides the request needs data retrieval first.","Hands the task to the Data Retrieval Agent.","Will orchestrate the remaining steps & state."]),
 (3,{'AGENTS','HUB'},False,TEAL,"Retrieval Agent Calls the AI Hub","Data Retrieval Agent sends the query request to the AI Hub",
    ["The agent needs a SQL query generated for it.","It forwards the request to the central AI Hub.","Every model call goes through the Hub — one control point.","The Hub will govern and protect this call."]),
 (4,{'HUB'},False,GOLD,"AI Hub Secures the Input","Token counting, guardrails, and PII redaction",
    ["Counts input tokens for quota & cost tracking.","Runs gateway guardrails and policy checks.","Redacts PII into placeholders (name, account…).","Only the redacted prompt is sent onward to the LLM."]),
 (5,{'HUB','LLM'},False,PURPLE,"LLM Generates the SQL","The provider model writes the query from the redacted prompt",
    ["The LLM receives placeholders only — never real PII.","Generates parameterized, read-only SQL.","Returns the generated SQL to the AI Hub.","Output tokens are counted as well."]),
 (6,{'HUB'},False,GOLD,"AI Hub Re-inserts PII","Real values are restored into the generated SQL",
    ["Hub maps placeholders back to the real values.","Uses the secure mapping kept for this request.","SQL becomes executable against real data.","PII never left the trust boundary."]),
 (7,{'HUB','AGENTS'},False,GOLD,"SQL Returns to the Agent","AI Hub sends the completed SQL back to the Retrieval Agent",
    ["The ready-to-run SQL is returned to the agent.","Retrieval Agent receives it from the Hub.","The exchange is logged and traced.","Next: execute it against the database."]),
 (8,{'AGENTS','DB'},False,GREEN,"Retrieval Agent Queries the Database","Run the SQL to fetch the customer's data",
    ["Read-only query with row-level access scope.","Returns the matching transaction rows.","No LLM is involved in data access itself.","Results are handed to the Analysis Agent."]),
 (9,{'AGENTS','HUB','LLM'},False,TEAL,"Analysis Agent Writes the Answer","Analyse the data and compose a natural-language response",
    ["Analysis Agent reviews the returned rows.","Generates the answer via AI Hub → LLM (PII-safe).","Summarises the result in clear language.","Prepares the response for delivery."]),
 (10,{'AGENTS'},True,RED,"Output Guardrail","Check the response before it leaves the system",
    ["Scan the answer for any leaked PII.","Toxicity and safety screening.","Confirm it is grounded in the retrieved data.","Attach any required disclosures."]),
 (11,{'CLIENT'},False,BLUE,"Client Shows the Result","The user sees the final answer",
    ["Response is streamed back to the chat UI.","Presented in clear natural language.","The full interaction is logged for audit.","Session is ready for the next question."]),
]
def draw_example(s, n):
    if n==1:
        ix,iy,iw = card_frame(s, "💬  Chat", BLUE)
        bb=box(s, ix, iy, iw, 1.4, fill=BLUE, line=None, radius=0.18)
        boxtext(bb,[{'t':"Show me transaction history for Somchai Jantapan, account 1234567890, last month",'s':10.5,'c':WHITE,'a':PP_ALIGN.LEFT}],ml=0.16,mr=0.16)
        nb=box(s, ix, iy+1.55, iw, 0.95, fill=RED_LT, line=RED, lw=1.0, radius=0.1)
        boxtext(nb,[{'t':[("Guardrail: ",RED,True,False),("policy + prompt-injection checks. PII is NOT redacted here.",INK,False,False)],'s':10}],ml=0.14,mr=0.14)
    elif n==2:
        ix,iy,iw = card_frame(s, "🧭  Routing decision", TEAL)
        c1=box(s, ix, iy, iw, 0.6, fill=TEAL_LT, line=TEAL, lw=1.0, radius=0.1)
        boxtext(c1,[{'t':[("Intent: ",TEAL,True,False),("retrieve transactions",INK,False,False)],'s':11}])
        arrow(s, ix+iw/2, iy+0.65, ix+iw/2, iy+0.95, color=TEAL, w=2.0)
        c2=box(s, ix, iy+1.0, iw, 0.6, fill=TEAL, line=None, radius=0.1)
        boxtext(c2,[{'t':"→  Route to Data Retrieval Agent",'s':11,'c':WHITE,'b':True}])
        c3=box(s, ix, iy+1.85, iw, 0.55, fill=GREY_LT, line=None, radius=0.1)
        boxtext(c3,[{'t':"then → Analysis Agent (later)",'s':10,'c':GREY}])
    elif n==3:
        ix,iy,iw = card_frame(s, "📨  Request to AI Hub", TEAL)
        b=box(s, ix, iy, iw, 1.9, fill=GREY_LT, line=blend(TEAL,0.4), lw=1.0, radius=0.06)
        text(s, ix+0.16, iy+0.12, iw-0.3, 1.7, [
            {'t':[("POST ",TEAL,True,False),("/ai-hub/generate-sql",INK,False,False)],'s':10.5,'f':MONO,'sa':4},
            {'t':"task: text-to-SQL",'s':10,'c':INK,'f':MONO,'sa':2},
            {'t':"prompt: \"transactions for",'s':10,'c':INK,'f':MONO,'sa':0},
            {'t':"  Somchai J., acct 1234567890…\"",'s':10,'c':INK,'f':MONO}])
    elif n==4:
        ix,iy,iw = card_frame(s, "🎭  PII redaction + tokens", GOLD)
        o=box(s, ix, iy, iw, 0.78, fill=WHITE, line=RED, lw=1.1, radius=0.08)
        boxtext(o,[{'t':[("IN  ",RED,True,False),("Somchai J. · acct 1234567890",INK,False,False)],'s':10}],ml=0.12)
        arrow(s, ix+iw/2, iy+0.82, ix+iw/2, iy+1.05, color=GREY, w=2.0)
        r=box(s, ix, iy+1.1, iw, 0.78, fill=WHITE, line=GREEN, lw=1.1, radius=0.08)
        boxtext(r,[{'t':[("OUT  ",GREEN,True,False),(":NAME_001 · acct :ACCT_001",INK,False,False)],'s':10}],ml=0.12)
        tb=box(s, ix, iy+2.05, iw, 0.5, fill=GOLD_LT, line=GOLD, lw=1.0, radius=0.3)
        boxtext(tb,[{'t':"🔢 input tokens counted: 42",'s':10.5,'c':NAVY,'b':True}])
    elif n==5:
        ix,iy,iw = card_frame(s, "🧠  Generated SQL (placeholders)", PURPLE)
        codeblock(s, ix, iy, iw, SQL_PH, h=1.95)
        text(s, ix, iy+2.0, iw, 0.4, [{'t':"LLM sees placeholders only — no real PII.",'s':9.5,'c':PURPLE,'i':True}])
    elif n==6:
        ix,iy,iw = card_frame(s, "🎭  SQL after PII re-insert", GOLD)
        codeblock(s, ix, iy, iw, SQL_REAL, h=1.95)
        text(s, ix, iy+2.0, iw, 0.4, [{'t':"Placeholders → real values, ready to run.",'s':9.5,'c':GREEN,'i':True}])
    elif n==7:
        ix,iy,iw = card_frame(s, "↩️  SQL handed back", GOLD)
        codeblock(s, ix, iy, iw, SQL_REAL, h=1.95)
        rr=box(s, ix, iy+2.05, iw, 0.5, fill=TEAL, line=None, radius=0.3)
        boxtext(rr,[{'t':"→ delivered to Data Retrieval Agent",'s':10.5,'c':WHITE,'b':True}])
    elif n==8:
        ix,iy,iw = card_frame(s, "🗄️  Query result", GREEN)
        rows=[("Date","Amount","Channel",True),("05 Jun","− 1,200","ATM",False),
              ("12 Jun","+ 8,500","Transfer",False),("21 Jun","− 350","QR",False)]
        cw=[iw*0.34,iw*0.34,iw*0.32]; ry=iy
        for r,(a,b,c,hdr) in enumerate(rows):
            xx=ix
            for ci,val in enumerate((a,b,c)):
                cell=box(s,xx,ry,cw[ci],0.5,fill=(GREEN if hdr else (GREEN_LT if r%2 else WHITE)),line=RGBColor(0xD5,0xDC,0xE6),lw=0.75,shape=MSO_SHAPE.RECTANGLE)
                boxtext(cell,[{'t':val,'s':10,'c':(WHITE if hdr else INK),'b':hdr,'a':PP_ALIGN.LEFT}],ml=0.1)
                xx+=cw[ci]
            ry+=0.5
        text(s, ix, ry+0.05, iw, 0.4, [{'t':"14 transactions returned (sample).",'s':9.5,'c':GREEN,'i':True}])
    elif n==9:
        ix,iy,iw = card_frame(s, "📊  Draft answer", TEAL)
        bb=box(s, ix, iy, iw, 2.3, fill=TEAL_LT, line=TEAL, lw=1.0, radius=0.12)
        boxtext(bb,[{'t':"\"Last month Somchai had 14 transactions totalling ฿7,950 net — mostly QR payments, with one ฿8,500 transfer in on 12 Jun.\"",'s':11,'c':INK,'a':PP_ALIGN.LEFT}],ml=0.18,mr=0.18)
    elif n==10:
        ix,iy,iw = card_frame(s, "🛡️  Output checks", RED)
        for i,(ic,t) in enumerate([("✓","No leaked PII"),("✓","Safe & non-toxic"),("✓","Grounded in the data"),("✓","Disclosure attached")]):
            yy=iy+i*0.62
            icon_circle(s, ix, yy, 0.4, ic, GREEN, WHITE, 14)
            text(s, ix+0.55, yy+0.02, iw-0.7, 0.5, [{'t':t,'s':11.5,'c':INK}], anchor=MSO_ANCHOR.MIDDLE)
    elif n==11:
        ix,iy,iw = card_frame(s, "💬  Final answer", BLUE)
        bb=box(s, ix, iy, iw, 2.3, fill=WHITE, line=blend(BLUE,0.3), lw=1.2, radius=0.12)
        boxtext(bb,[{'t':"Last month, Somchai Jantapan (acct 1234567890) made 14 transactions, ฿7,950 net. Want a breakdown by channel?",'s':11,'c':INK,'a':PP_ALIGN.LEFT}],ml=0.18,mr=0.18)

for (n,active,guard,accent,title,action,bullets) in STEPS:
    s = slide(); page = 3+n
    header(s, f"STEP {n} OF 11", title, page, badge=n)
    minichain(s, active=active, guard=guard, y=1.3)
    a=box(s,0.6,2.98,6.55,0.66,fill=accent,line=None,shadow=True,radius=0.12)
    boxtext(a,[{'t':action,'s':12.5,'c':(NAVY if accent==GOLD else WHITE),'b':True,'a':PP_ALIGN.LEFT}],ml=0.2,mr=0.2)
    yy=3.88
    for b in bullets:
        icon_circle(s,0.66,yy+0.02,0.26,"•",blend(accent,0.7),accent,12)
        text(s,1.06,yy-0.08,6.0,0.6,[{'t':b,'s':11.5,'c':INK,'ls':1.0}],anchor=MSO_ANCHOR.MIDDLE)
        yy+=0.66
    draw_example(s, n)

# ================================================================ SLIDE 15 — Security & PII
s = slide(); header(s, "SECURITY", "Where Guardrails & PII Protection Happen", 15)
text(s, 0.55, 1.18, 12.3, 0.35, [{'t':"Three control points keep the flow safe — PII is handled only at the AI Hub.",'s':12.5,'c':GREY,'i':True}])
cards = [
 ("🛡️","1 · Input Guardrail","at CLIENT → AGENTS",["Content & policy moderation","Prompt-injection / jailbreak checks","Scope & intent validation","No PII redaction at this step"], RED, RED_LT),
 ("🏛️","2 · AI Hub (PII boundary)","every LLM call",["Counts input / output tokens","Redacts PII before the LLM","LLM sees placeholders only","Re-inserts PII into the result"], GOLD, GOLD_LT),
 ("✅","3 · Output Guardrail","at AGENTS → CLIENT",["Checks for leaked PII","Toxicity & safety screening","Grounded / no hallucination","Required disclosures attached"], GREEN, GREEN_LT),
]
cw=3.95
for i,(ic,t,sub,items,ac,bgc) in enumerate(cards):
    x=0.6+i*(cw+0.2); y=1.75
    box(s, x, y, cw, 3.95, fill=WHITE, line=ac, lw=1.4, shadow=True, radius=0.05)
    box(s, x, y, cw, 0.95, fill=ac, line=None, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    icon_circle(s, x+0.25, y+0.22, 0.55, ic, WHITE, ac, 17)
    text(s, x+0.92, y+0.16, cw-1.05, 0.4, [{'t':t,'s':13.5,'c':(NAVY if ac==GOLD else WHITE),'b':True,'sa':0}])
    text(s, x+0.92, y+0.55, cw-1.05, 0.3, [{'t':sub,'s':10,'c':(NAVY if ac==GOLD else SKY),'i':True}])
    yy=y+1.2
    for it in items:
        icon_circle(s, x+0.28, yy, 0.28, "•", bgc, ac, 12)
        text(s, x+0.68, yy-0.06, cw-0.9, 0.5, [{'t':it,'s':11,'c':INK,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)
        yy+=0.6
bn=box(s, 0.6, 5.95, 12.15, 0.78, fill=NAVY, line=None, radius=0.1)
boxtext(bn, [{'t':[("Key point:  ",GOLD,True,False),("real customer PII reaches the LLM only as redacted placeholders — the AI Hub re-inserts the real values, so sensitive data never leaves the trust boundary.",WHITE,False,False)],'s':12}], ml=0.3, mr=0.3)

# ================================================================ SLIDE 16 — Takeaways
s = slide(); header(s, "WRAP-UP", "Key Takeaways", 16)
tk=[("①","One simple chain","Client → Agents → AI Hub → LLM, with the Database beside the agents.",BLUE),
    ("②","Agents do the orchestration","A Router, a Retrieval agent and an Analysis agent split the work.",TEAL),
    ("③","The AI Hub is the control point","Token counting, guardrails and PII redaction / re-insertion live here.",GOLD),
    ("④","PII never leaves the boundary","The LLM only ever sees placeholders; real values are restored by the Hub.",PURPLE),
    ("⑤","Guardrails top and tail the flow","An input check (no PII) and an output check protect every request.",GREEN)]
for i,(ic,t,d,clr) in enumerate(tk):
    y=1.6+i*1.04
    box(s, 0.7, y, 11.95, 0.92, fill=WHITE, line=blend(clr,0.4), lw=1.0, shadow=True, radius=0.08)
    box(s, 0.7, y, 0.14, 0.92, fill=clr, line=None, shape=MSO_SHAPE.RECTANGLE)
    icon_circle(s, 0.98, y+0.21, 0.5, ic, clr, (NAVY if clr==GOLD else WHITE), 17)
    text(s, 1.75, y+0.1, 4.1, 0.75, [{'t':t,'s':14.5,'c':NAVY,'b':True}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 5.95, y+0.1, 6.5, 0.75, [{'t':d,'s':11.5,'c':GREY,'ls':1.0}], anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 17 — Closing
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
chain=[("💻 CLIENT",BLUE),("🤖 AI AGENTS",TEAL),("🏛️ AI HUB",GOLD),("🧠 LLM",PURPLE)]
x=1.15
for i,(lbl,clr) in enumerate(chain):
    w=2.5; c=box(s, x, 5.05, w, 0.6, fill=clr, line=None, shadow=True, radius=0.5)
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
