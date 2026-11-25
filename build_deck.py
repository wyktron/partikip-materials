#!/usr/bin/env python3
"""Build the partikip.md pitch deck for Deeptech Bridge Accelerator."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

def strip_style(sh):
    el = sh._element
    style = el.find(qn('p:style'))
    if style is not None:
        el.remove(style)

INK    = RGBColor.from_string("211A16")
MUTED  = RGBColor.from_string("6B625C")
ACCENT = RGBColor.from_string("E4572E")
PAPER  = RGBColor.from_string("FFF8F3")
SOFT   = RGBColor.from_string("F8EDE5")
WHITE  = RGBColor.from_string("FFFFFF")
RULE   = RGBColor.from_string("EADCD2")
GREEN  = RGBColor.from_string("15803D")

FONT = "Noto Sans"
SW, SH = 13.333, 7.5
ML = 0.9
CW = SW - 2 * ML

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

def slide(bg_color=PAPER):
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background(); bg.shadow.inherit = False
    strip_style(bg)
    return s

def text(s, x, y, w, h, runs, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT,
         space_after=6, line=1.25, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = runs if isinstance(runs, list) and runs and isinstance(runs[0], list) else [runs]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line
        items = para if isinstance(para, list) else [(para, {})]
        for t, fmt in items:
            r = p.add_run(); r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(fmt.get("size", size))
            f.bold = fmt.get("bold", bold)
            f.color.rgb = fmt.get("color", color)
    return box

def card(s, x, y, w, h, fill=WHITE, line=RULE, radius=0.055):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(0.9)
    sh.shadow.inherit = False
    strip_style(sh)
    try: sh.adjustments[0] = radius
    except Exception: pass
    return sh

def bar(s, x, y, w, h=0.05, color=ACCENT):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    sh.line.fill.background(); sh.shadow.inherit = False
    strip_style(sh)
    return sh

def kicker(s, label, color=ACCENT):
    text(s, ML, 0.5, CW, 0.3, [[(label.upper(), {"size": 11.5, "bold": True, "color": color})]])

def title(s, t, y=0.9, size=38, w=None, color=INK):
    text(s, ML, y, w or CW, 1.3, t, size=size, bold=True, color=color, line=1.06)

def notes(s, body):
    s.notes_slide.notes_text_frame.text = body

# ---------------------------------------------------------------- 1. cover
s = slide()
bar(s, ML, 0.0, SW, 0.12, ACCENT)
text(s, ML, 0.8, CW, 0.3, [[("PARTIKIP.MD", {"size": 12, "bold": True, "color": ACCENT})]])
text(s, ML, 1.55, 8.6, 2.6, "Singurătatea,\nînvinsă invitație\ncu invitație.", size=46, bold=True, line=1.03)
text(s, ML, 4.55, 7.6, 1.3,
     "Platformă digitală pentru întâlniri reale în Republica Moldova — cu intenția declarată de la început și participare prin consimțământ.",
     size=16, color=MUTED, line=1.3)
text(s, ML, 6.55, 7.0, 0.4, [[("Deeptech Bridge Accelerator", {"size": 13, "color": MUTED}),
                               ("   ·   Chișinău   2026", {"size": 13, "color": MUTED})]])
text(s, SW - ML - 2.6, 6.55, 2.6, 0.4, [[("partikip.md", {"size": 13, "bold": True, "color": INK})]], align=PP_ALIGN.RIGHT)

# right stats
cx = 9.55
for i, (big, small) in enumerate([("825", "commit-uri"), ("67", "migrații"), ("TRL 7", "pilot live")]):
    y = 1.55 + i * 1.45
    card(s, cx, y, 2.9, 1.15, fill=WHITE)
    text(s, cx + 0.3, y + 0.2, 2.4, 0.55, big, size=28, bold=True, color=ACCENT)
    text(s, cx + 0.3, y + 0.78, 2.4, 0.3, small, size=12, color=MUTED)
notes(s, "partikip.md este o platformă pentru întâlniri reale, nu o aplicație de dating. Ideea centrală: intenția e declarată înainte, iar participarea cere acceptul gazdei. Fără chat.")

# ---------------------------------------------------------------- 2. problema
s = slide()
kicker(s, "Problema")
title(s, "Tot mai mulți tineri ajung singuri.")
rows = [
    ("Plecarea", "Emigrația golește orașele mici, iar comunitățile se subțiază."),
    ("Ruptura", "Cercurile sociale se destramă după facultate și la primul job."),
    ("Momentul", "„Vreau să ies, dar nu am cu cine” apare săptămânal, nu o dată pe an."),
    ("Alternatives", "Grupuri haotice, aplicații de dating cu chat, evenimente fără nicio verificare."),
]
for i, (head, body) in enumerate(rows):
    y = 2.15 + i * 1.12
    bar(s, ML, y + 0.08, 0.32, 0.05, ACCENT)
    text(s, ML, y, 2.15, 0.35, [[(head.upper(), {"size": 10.5, "bold": True, "color": ACCENT})]])
    text(s, ML + 2.2, y - 0.04, 9.2, 0.85, body, size=15.5, color=MUTED, line=1.3)
text(s, ML, 6.75, CW, 0.4, [[("Moldova: declin demografic + singurătate urbană — o problemă de sănătate publică și de capital social.", {"size": 12.5, "color": MUTED})]])
notes(s, "Problema nu e lipsa de oameni, ci lipsa unui cadru sigur și clar în care să te întâlnești. Aplicațiile existente optimizează conversația, nu întâlnirea.")

# ---------------------------------------------------------------- 3. soluția
s = slide()
kicker(s, "Soluția")
title(s, "Intenție întâi. Consimțământ la participare.")
text(s, ML, 2.0, 11.4, 0.7, "O invitație spune clar ce urmează; gazda decide cine vine; chatul nu există.", size=16, color=MUTED)
cards = [
    ("01", "Publici o invitație", "Concert, cafea, plimbare, sport. Intenția e explicită: romantic, prietenie, networking, practică de limbă."),
    ("02", "Ceri să participi", "Fără conversații și fără ghosting. O singură fotografie, un profil clar."),
    ("03", "Gazda acceptă", "Gazda e verificată, are istoric, iar moderatorii pot interveni."),
]
for i, (n, head, body) in enumerate(cards):
    x = ML + i * 3.95
    card(s, x, 3.0, 3.6, 2.7, fill=WHITE)
    text(s, x + 0.35, 3.25, 0.8, 0.4, n, size=13, bold=True, color=ACCENT)
    text(s, x + 0.35, 3.7, 2.9, 0.5, head, size=17, bold=True)
    text(s, x + 0.35, 4.3, 2.9, 1.2, body, size=12.5, color=MUTED, line=1.28)
text(s, ML, 6.3, CW, 0.4, [[("fără chat   ·   fără feed infinit   ·   cu verificarea gazdei", {"size": 13, "bold": True, "color": ACCENT})]])
notes(s, "Trei reguli care schimbă comportamentul: intenția e publică, participarea cere accept, conversația nu e produsul. Produsul este întâlnirea.")

# ---------------------------------------------------------------- 4. produsul
s = slide()
kicker(s, "Produsul")
title(s, "De la invitație la întâlnire, în patru pași.")
steps = [
    ("01", "Descoperă", "Feed, calendar, „Sunt norocos” — invitații aprobate, filtrate pe oraș și intenție."),
    ("02", "Cere", "Trimiți cererea; gazda vede profilul și decide. Fără mesaje."),
    ("03", "Întâlnește-te", "Locație și oră clare; gazda poate publica update-uri scurte despre schimbări."),
    ("04", "Revino", "Recurență, notificări e-mail/push, leaderboard și istoric."),
]
for i, (n, head, body) in enumerate(steps):
    x = ML + i * 2.95
    text(s, x, 2.15, 0.7, 0.4, n, size=12, bold=True, color=ACCENT)
    text(s, x, 2.55, 2.6, 0.5, head, size=18, bold=True)
    text(s, x, 3.15, 2.55, 1.9, body, size=12.5, color=MUTED, line=1.3)
    bar(s, x, 5.2, 2.55, 0.04, RULE)
text(s, ML, 5.65, CW, 0.9, "Verificarea gazdelor, moderarea, jurnalul de audit și notificările sunt integrate în produs — nu lipite peste.",
     size=14, color=INK, line=1.3)
notes(s, "Fluxul e scurt intenționat. Cu cât sunt mai puține etape de conversație, cu atât mai puține oportunități de hărțuire și de abandon.")

# ---------------------------------------------------------------- 5. tehnologie
s = slide()
kicker(s, "Tehnologie")
title(s, "Încrederea stă în backend.")
items = [
    ("RLS pe fiecare tabel", "Accesul la date e impus de PostgreSQL, nu de client."),
    ("67 de migrații", "Schemată versionată, RLS, trigger-e de limitare și audit."),
    ("Edge functions", "E-mail tranzacțional și notificări push, cu verificare JWT."),
    ("Moderare + audit", "Coadă de moderare, jurnal de acțiuni, suspendare."),
    ("Verificarea gazdelor", "Programare video, limită de încercări, reguli stricte."),
    ("MCP pentru agenți AI", "Asistenți externi pot căuta invitații publice, cu OAuth."),
]
for i, (head, body) in enumerate(items):
    col, row = i % 3, i // 3
    x = ML + col * 3.95
    y = 2.1 + row * 1.95
    card(s, x, y, 3.6, 1.65, fill=WHITE)
    text(s, x + 0.3, y + 0.22, 3.0, 0.35, head, size=14, bold=True)
    text(s, x + 0.3, y + 0.65, 3.0, 0.85, body, size=11.5, color=MUTED, line=1.25)
text(s, ML, 6.35, CW, 0.4, [[("React PWA  ·  PostgreSQL (Supabase)  ·  21 de ecrane  ·  825 commit-uri", {"size": 13, "bold": True, "color": ACCENT})]])
notes(s, "Tehnologia nu e spectaculoasă, dar e disciplina care lipsește produselor sociale locale: datele sunt protejate la nivel de bază de date, iar regulile sunt impuse server-side.")

# ---------------------------------------------------------------- 6. piața
s = slide()
kicker(s, "Piața")
title(s, "Mică, dar densă — și repetabilă.")
nums = [("2,5M", "populația Republicii Moldova"), ("~450k", "tineri 18–40, urban"), ("3", "orașe pilot")]
for i, (big, small) in enumerate(nums):
    x = ML + i * 3.95
    card(s, x, 2.1, 3.6, 1.5, fill=WHITE)
    text(s, x + 0.3, 2.28, 3.0, 0.55, big, size=30, bold=True, color=ACCENT)
    text(s, x + 0.3, 2.92, 3.0, 0.4, small, size=12, color=MUTED)
text(s, ML, 4.05, 11.4, 1.2, "Modelul se copiază în piețe mici cu aceeași rană: Balcanii, țările baltice și diaspora din UE. Nu concurăm cu rețelele globale pe feed — concurăm pe densitate locală.",
     size=15.5, color=INK, line=1.32)
text(s, ML, 5.6, CW, 0.9, "Țintă la 12 luni: 800+ utilizatori înregistrați în 3 orașe, cu cel puțin 120 de invitații create și jumătate cu 3+ confirmați.",
     size=13.5, color=MUTED, line=1.3)
text(s, ML, 6.6, CW, 0.3, [[("Estimări de piață; cifrele de țintă sunt asumate, nu realizate.", {"size": 10.5, "color": MUTED})]])
notes(s, "Nu avem nevoie de o piață mare cât de la început. Avem nevoie de un oraș în care invitațiile se umplu. Densitatea e metrica.")

# ---------------------------------------------------------------- 7. business model
s = slide()
kicker(s, "Model de business")
title(s, "Gratuit pentru oameni. Plătit de cine are interes.")
bm = [
    ("Evenimente sponsorizate", "Venue-uri și branduri locale plătesc pentru vizibilitate contextuală, marcată clar."),
    ("Parteneriate cu venue-uri", "Abonament lunar pentru cafenele și centre culturale care găzduiesc evenimente."),
    ("Premium pentru gazde", "Vizibilitate și instrumente pentru organizatorii recurenți."),
]
for i, (head, body) in enumerate(bm):
    x = ML + i * 3.95
    card(s, x, 2.1, 3.6, 2.3, fill=WHITE)
    text(s, x + 0.3, 2.35, 3.0, 0.7, head, size=15, bold=True, line=1.15)
    text(s, x + 0.3, 3.15, 3.0, 1.0, body, size=12, color=MUTED, line=1.3)
text(s, ML, 4.85, 11.4, 0.8, "Nu vindem date personale și nu introducem reclame în feed-ul de invitații. Monetizarea nu a fost pornită — este următorul pas după densitate.",
     size=14.5, color=INK, line=1.3)
notes(s, "Ordinea contează: mai întâi densitate și încredere, apoi bani de la partea care are interes comercial. Utilizatorul nu plătește.")

# ---------------------------------------------------------------- 8. stadiu
s = slide()
kicker(s, "Stadiul actual")
title(s, "Pilot live, TRL 7.")
stats = [
    ("~30", "vizitatori unici / 30 zile"),
    ("0", "venituri, 0 clienți plătitori"),
    ("#1", "hackathon Moldova Digital Summit"),
    ("825", "commit-uri, 67 migrații"),
]
for i, (big, small) in enumerate(stats):
    x = ML + (i % 2) * 5.85
    y = 2.1 + (i // 2) * 1.75
    card(s, x, y, 5.5, 1.45, fill=WHITE)
    text(s, x + 0.35, y + 0.2, 1.8, 0.6, big, size=28, bold=True, color=ACCENT)
    text(s, x + 2.1, y + 0.38, 3.1, 0.7, small, size=13, color=MUTED, line=1.2)
text(s, ML, 5.85, 11.4, 0.9, "Pilot public în Chișinău. Relansarea tehnică pe infrastructură cu backup automat este în curs — lecția învățată după prima lună de inactivitate.",
     size=13.5, color=MUTED, line=1.3)
notes(s, "Suntem la TRL 7: produs demonstrat în mediu real, cu utilizatori reali, dar fără încă o piață pornită. Cifrele mici sunt reale și le spunem ca atare.")

# ---------------------------------------------------------------- 9. roadmap
s = slide()
kicker(s, "Roadmap")
title(s, "Până la Demo Day: densitate, nu promisiuni.", size=34)
phases = [
    ("Q4 2026", "Backend stabil + backup automat. Relansare. 50 de invitații reale în Chișinău."),
    ("Q1 2027", "Bălți și Iași. 10 parteneriate cu venue-uri. Primele evenimente sponsorizate."),
    ("Mai 2027", "800+ utilizatori, retenție 30 zile >25%, 10.000 EUR venituri. Pachet de metrici pentru pre-seed."),
]
for i, (when, what) in enumerate(phases):
    y = 2.4 + i * 1.45
    card(s, ML, y, 11.4, 1.15, fill=WHITE)
    text(s, ML + 0.35, y + 0.28, 1.8, 0.4, when, size=14, bold=True, color=ACCENT)
    text(s, ML + 2.3, y + 0.26, 8.7, 0.75, what, size=13.5, color=INK, line=1.25)
text(s, ML, 6.75, CW, 0.3, [[("Fiecare fază are o singură metrică de trecere: invitații care se umplu.", {"size": 12.5, "color": MUTED})]])
notes(s, "Roadmap-ul e construit invers din Demo Day. Nu adăugăm funcții până nu avem densitate într-un oraș.")

# ---------------------------------------------------------------- 10. competiție
s = slide()
kicker(s, "Competiție")
title(s, "Toți rezolvă altceva.")
rows = [
    ("Tinder / Bumble BFF", "Dating și chat; intenția rămâne ambiguă."),
    ("Meetup", "Evenimente generice, fără verificare locală."),
    ("Facebook Events", "Haos, fără consimțământ și fără urmărire a participării."),
    ("Telegram / WhatsApp", "Grupuri private, fără încredere și fără istoric."),
    ("Timeleft", "Cine plătite, nu comunitate locală."),
]
for i, (head, body) in enumerate(rows):
    y = 1.85 + i * 0.82
    text(s, ML, y, 3.4, 0.35, head, size=14, bold=True)
    text(s, ML + 3.5, y, 7.4, 0.5, body, size=13, color=MUTED)
text(s, ML, 6.15, 11.4, 0.8, "Diferența partikip: intenție explicită + consimțământ + verificare + densitate locală. Nu o funcție — un protocol.",
     size=14, bold=True, color=ACCENT)
notes(s, "Concurența nu e o altă aplicație de dating. Concurența e obiceiul: grupul de WhatsApp și evenimentul de pe Facebook.")

# ---------------------------------------------------------------- 11. echipa
s = slide()
kicker(s, "Echipa")
title(s, "Un fondator care construiește.")
card(s, ML, 2.0, 6.6, 3.0, fill=WHITE)
text(s, ML + 0.4, 2.25, 5.8, 0.5, "Victor Girbu — founder", size=20, bold=True)
text(s, ML + 0.4, 2.85, 5.8, 1.8, "Produs și engineering. 825 de commit-uri, 21 de ecrane, 67 de migrații, de la zero la un produs live. Background tehnic și în comunitate.",
     size=13, color=MUTED, line=1.35)
text(s, ML, 5.35, 6.0, 0.35, [[("Ce lipsește și cum acoperim", {"size": 12, "bold": True, "color": ACCENT})]])
text(s, ML, 5.75, 6.2, 1.3, "Juridic/GDPR-DSA (DPO fractional), fundraising (advisor), growth și moderare la scară (primul hire după rundă).",
     size=12.5, color=MUTED, line=1.3)
card(s, 8.0, 2.0, 4.3, 3.0, fill=SOFT, line=None)
text(s, 8.3, 2.25, 3.7, 0.4, "De ce noi", size=13, bold=True, color=ACCENT)
text(s, 8.3, 2.75, 3.7, 2.0, "Cunoaștem piața din interior, avem produsul construit și învățăm din fiecare lansare. Nu vindem o idee — vindem execuție.",
     size=12.5, color=INK, line=1.35)
notes(s, "Echipa e mică, dar produsul există. Acceleratorul poate acoperi exact golurile pe care le recunoaștem: juridic, fundraising, growth.")

# ---------------------------------------------------------------- 12. cererea
s = slide()
kicker(s, "Cererea")
title(s, "120.000 EUR pre-seed.")
text(s, ML, 2.0, 11.0, 0.6, "18 luni de runway pentru a transforma un produs viu într-o piață locală cu metrici de pre-seed.", size=15, color=MUTED)
alloc = [("Produs & engineering", 35), ("Go-to-market", 30), ("Juridic & conformitate", 15), ("Echipă", 15), ("Infrastructură & rezervă", 5)]
for i, (label, pct) in enumerate(alloc):
    y = 3.0 + i * 0.62
    text(s, ML, y, 3.2, 0.35, label, size=12.5, color=INK)
    bar(s, ML + 3.4, y + 0.11, 5.2 * pct / 35, 0.22, ACCENT)
    text(s, ML + 3.4 + 5.3 * pct / 35, y, 1.0, 0.35, f"{pct}%", size=12.5, bold=True, color=ACCENT)
card(s, ML, 6.2, 11.4, 0.85, fill=SOFT, line=None)
text(s, ML + 0.35, 6.38, 10.6, 0.4, [[("hello@partikip.com", {"size": 15, "bold": True, "color": INK}),
                                       ("      partikip.md", {"size": 15, "color": MUTED})]])
notes(s, "Cererea e aliniată cu formularul: 120k pre-seed, 18 luni, prag minim 60k pentru un pilot de 6 luni în România.")

prs.save("partikip-pitch-deck.pptx")
print("saved partikip-pitch-deck.pptx", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
