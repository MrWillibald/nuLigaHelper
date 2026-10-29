"""Build both German handbook editions using the web application's visual design.

Optional documentation dependencies: reportlab and Pillow. Run:
    python docs/benutzer/build_handbook.py
    python docs/benutzer/build_handbook.py --edition members
"""

from argparse import ArgumentParser
from dataclasses import dataclass
from html import escape
from pathlib import Path
import re

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, CondPageBreak, Flowable, Frame, KeepTogether, PageBreak,
    PageTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / 'docs' / 'benutzer'
W, H = A4
PAGE_WIDTH = W - 44*mm
FONT_DIR = Path('/usr/share/fonts/truetype/liberation')
for name, filename in [
    ('Web', 'LiberationSans-Regular.ttf'), ('WebBold', 'LiberationSans-Bold.ttf'),
    ('WebItalic', 'LiberationSans-Italic.ttf'), ('WebBoldItalic', 'LiberationSans-BoldItalic.ttf'),
]:
    pdfmetrics.registerFont(TTFont(name, str(FONT_DIR / filename)))
pdfmetrics.registerFontFamily('Web', normal='Web', bold='WebBold',
                            italic='WebItalic', boldItalic='WebBoldItalic')

# Keep the printed handbooks aligned with the actual app, without loading any
# configuration, database, network resource, or application runtime.
css = (ROOT / 'nuligahelper/web/static/style.css').read_text(encoding='utf-8')
tokens = dict(re.findall(r'--([\w-]+):\s*(#[0-9a-fA-F]{6})', css))
def color(name):
    return colors.HexColor(tokens[name])

NAVY, DARK, BG, TEXT, MUTED, LINE, ORANGE = [
    color(name) for name in ['navy', 'navy-dark', 'bg', 'text', 'muted', 'line', 'kids']
]
RAINBOW = [color(name) for name in ['men','male-ab','male-cd','women','female-ab','female-cd','kids']]
PALE_ORANGE = colors.HexColor('#FFF5E6')


def style(name, **kwargs):
    return ParagraphStyle(name, fontName='Web', fontSize=10.4, leading=14.5,
                          textColor=TEXT, spaceAfter=8, allowWidows=0,
                          allowOrphans=0, **kwargs)

STYLES = {
    'body': style('body'),
    'subheading': ParagraphStyle(
        'subheading', fontName='WebBold', fontSize=12.2, leading=16.5,
        textColor=NAVY, spaceBefore=13, spaceAfter=7, keepWithNext=True),
    'chapter': ParagraphStyle('chapter', fontName='WebBold', fontSize=18.5,
        leading=23.5, textColor=NAVY),
    'caption': ParagraphStyle('caption', fontName='Web', fontSize=8.9,
        leading=12.2, textColor=MUTED),
    'cell': ParagraphStyle('cell', fontName='Web', fontSize=9.3, leading=13.1,
        textColor=TEXT),
    'cellhead': ParagraphStyle('cellhead', fontName='WebBold', fontSize=9.3,
        leading=13.1, textColor=colors.white),
    'toc': ParagraphStyle('toc', fontName='Web', fontSize=9.5, leading=14,
        textColor=NAVY, spaceBefore=0),
    'step': ParagraphStyle('step', fontName='Web', fontSize=10.4, leading=14.5,
        textColor=TEXT, allowWidows=0, allowOrphans=0),
    'checklist': ParagraphStyle('checklist', fontName='Web', fontSize=12.3,
        leading=19, textColor=TEXT),
    'cover-title': ParagraphStyle('cover-title', fontName='WebBold', fontSize=35,
        leading=39, textColor=NAVY),
    'cover-subtitle': ParagraphStyle('cover-subtitle', fontName='WebBold', fontSize=19,
        leading=24, textColor=NAVY),
    'cover-body': ParagraphStyle('cover-body', fontName='Web', fontSize=12,
        leading=18, textColor=TEXT),
}


@dataclass(frozen=True)
class Edition:
    key: str
    source: str
    output: str
    label: str
    title: str
    description: str

EDITIONS = [
    Edition('all', 'BENUTZERHANDBUCH.md', 'nuLigaHelper_Benutzerhandbuch.pdf',
            'GESAMTAUSGABE', 'Benutzerhandbuch',
            'Die vollständige Anleitung für Mitglieder, Mannschaftsverantwortliche und Admins.'),
    Edition('members', 'MITGLIEDERHANDBUCH.md', 'nuLigaHelper_Mitgliederhandbuch.pdf',
            'MITGLIEDERAUSGABE', 'Mitgliederhandbuch',
            'Deine Anleitung für Anmeldung, eigene Helferdienste und den nächsten Heimspieltag.'),
]


def inline(text):
    text = escape(text)
    text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'`(.*?)`', r'\1', text)
    return re.sub(r'\[([^]]+)\]\(([^)]+)\)', r'\1', text)


def paragraph(text, kind='body'):
    return Paragraph(inline(text), STYLES[kind])


def bookmark(title):
    return re.sub(r'\W+', '-', title)


class ChapterHeading(Flowable):
    """The app's dark section icon paired with a clear section heading."""
    keepWithNext = True

    def __init__(self, number, title):
        super().__init__()
        self.number = number
        self.title = f'{number} {title}'
        self.p = paragraph(title, 'chapter')

    def wrap(self, width, height):
        self.width = width
        self.pw, self.ph = self.p.wrap(width-46, height)
        self.height = max(32,self.ph)+22
        return self.width,self.height

    def draw(self):
        c=self.canv
        c.setFillColor(NAVY)
        c.roundRect(0,self.height-34,32,32,9,fill=1,stroke=0)
        c.setFillColor(ORANGE)
        c.setFont('WebBold',13)
        c.drawCentredString(16,self.height-23,str(self.number))
        self.p.drawOn(c,46,self.height-self.ph-2)


class Step(Flowable):
    def __init__(self, number, text):
        super().__init__()
        self.number=number
        self.p=paragraph(text,'step')

    def wrap(self,width,height):
        self.width=width
        _,self.ph=self.p.wrap(width-30,height)
        self.height=max(self.ph,19)+8
        return self.width,self.height

    def draw(self):
        c=self.canv
        c.setFillColor(ORANGE)
        c.circle(9,self.height-10,9,fill=1,stroke=0)
        c.setFillColor(NAVY)
        c.setFont('WebBold',9)
        c.drawCentredString(9,self.height-13,str(self.number))
        self.p.drawOn(c,30,self.height-self.ph)


class Notice(Flowable):
    """Give a small number of important rules the app's warm notice treatment."""
    def __init__(self,text):
        super().__init__()
        self.p=paragraph(text)

    def wrap(self,width,height):
        self.width=width
        _,self.ph=self.p.wrap(width-28,height)
        self.height=self.ph+24
        return width,self.height+8

    def draw(self):
        c=self.canv
        c.setFillColor(PALE_ORANGE)
        c.roundRect(0,0,self.width,self.height,9,fill=1,stroke=0)
        c.setFillColor(ORANGE)
        c.roundRect(0,0,4,self.height,2,fill=1,stroke=0)
        self.p.drawOn(c,14,12)


class ChecklistStep(Step):
    """A printable tick box, with enough space for the reader to mark the page."""
    def __init__(self, number, text):
        super().__init__(number, text)
        self.p = paragraph(text, 'checklist')

    def wrap(self, width, height):
        width, height = super().wrap(width, height)
        self.height += 13
        return width, self.height

    def draw(self):
        c = self.canv
        c.setStrokeColor(NAVY)
        c.setLineWidth(1.2)
        c.roundRect(0, self.height-19, 17, 17, 3, fill=0, stroke=1)
        self.p.drawOn(c, 30, self.height-self.ph)


class Figure(Flowable):
    def __init__(self,path,caption):
        super().__init__()
        self.path=path
        self.caption=paragraph(caption,'caption')
        with PILImage.open(path) as im:
            self.iw,self.ih=im.size

    def wrap(self,width,height):
        self.width=width
        maxw=width-14
        if self.iw<800:
            maxw=300 if '09-profil' not in self.path.name else 225
            if '06-bestaetigen' in self.path.name:
                maxw=355
        scale=min(maxw/self.iw,420/self.ih)
        self.dw,self.dh=self.iw*scale,self.ih*scale
        _,self.ch=self.caption.wrap(width-4,height)
        self.height=self.dh+self.ch+31
        return width,self.height

    def draw(self):
        c=self.canv
        y=self.ch+18
        x=(self.width-self.dw)/2
        c.setStrokeColor(LINE)
        c.setFillColor(colors.white)
        c.roundRect(x-5,y-5,self.dw+10,self.dh+10,9,fill=1,stroke=1)
        c.drawImage(str(self.path),x,y,width=self.dw,height=self.dh,mask='auto')
        self.caption.drawOn(c,2,3)


class Cover(Flowable):
    def __init__(self,edition,chapters,figures):
        super().__init__()
        self.edition=edition
        self.chapters=chapters
        self.figures=figures

    def wrap(self,width,height):
        self.width=width
        self.height=height-1
        return width,self.height

    def draw(self):
        c=self.canv
        top=self.height-17
        c.setFillColor(ORANGE)
        c.roundRect(0,top-22,155,23,7,fill=1,stroke=0)
        c.setFillColor(NAVY)
        c.setFont('WebBold',9.7)
        c.drawCentredString(77.5,top-14,self.edition.label)
        c.setFont('WebBold',21)
        c.drawString(0,top-67,'nuLigaHelper')
        p=paragraph(self.edition.title,'cover-title')
        _,ph=p.wrap(self.width,200)
        p.drawOn(c,0,top-79-ph)
        y=top-79-ph-19
        p=paragraph('Heimspieltage gemeinsam organisieren','cover-subtitle')
        _,ph=p.wrap(self.width,100)
        p.drawOn(c,0,y-ph)
        y-=ph+20
        p=paragraph(self.edition.description,'cover-body')
        _,ph=p.wrap(self.width,100)
        p.drawOn(c,0,y-ph)
        y-=ph+28
        for i,(label,target) in enumerate([
            ('Anmelden','4 Anmelden und abmelden'),
            ('Dienst übernehmen','9 Einen freien Dienst verbindlich übernehmen'),
            ('Ersatz suchen','10 Ersatz suchen und einen Dienst übernehmen'),
        ]):
            x=i*(self.width+10)/3
            cw=(self.width-20)/3
            c.setFillColor(NAVY)
            c.roundRect(x,y-29,cw,29,8,fill=1,stroke=0)
            c.setFillColor(colors.white)
            c.setFont('WebBold',10)
            c.drawCentredString(x+cw/2,y-19,label)
            c.linkRect(label,bookmark(target),(x,y-29,x+cw,y),relative=1,thickness=0)
        y-=62
        iw,ih=1069,300
        dw=self.width
        dh=dw*ih/iw
        c.setStrokeColor(LINE)
        c.setFillColor(colors.white)
        c.roundRect(0,y-dh,self.width,dh,12,fill=1,stroke=1)
        c.drawImage(str(DOCS/'screenshots/00-spielkarte.png'),7,y-dh+7,
                    width=dw-14,height=dh-14,mask='auto')
        y-=dh+28
        p=paragraph('**Eine Zusage ist verbindlich.** Du bleibst verantwortlich, bis eine geeignete Person den Dienst im System übernimmt. Am Vortag um **09:00 Uhr** werden besetzte Dienste fest.','cover-body')
        _,ph=p.wrap(self.width-30,150)
        c.setFillColor(PALE_ORANGE)
        c.roundRect(0,y-ph-25,self.width,ph+25,10,fill=1,stroke=0)
        p.drawOn(c,15,y-ph-12)
        c.setFillColor(MUTED)
        c.setFont('WebBold',10)
        c.drawString(0,20,f'{self.chapters} Kapitel   ·   {self.figures} erläuterte Abbildungen')
        c.setFont('Web',9)
        c.drawString(0,4,'Version 0.30  |  Stand 25. Oktober 2026')


class Manual(BaseDocTemplate):
    def __init__(self,edition):
        self.edition=edition
        self.target=DOCS/'pdf'/edition.output
        super().__init__(str(self.target),pagesize=A4,leftMargin=22*mm,
            rightMargin=22*mm,topMargin=27*mm,bottomMargin=21*mm,
            title=f'nuLigaHelper {edition.title}',author='TuS Raubling Handball',
            subject=edition.description,lang='de-DE')
        frame=Frame(self.leftMargin,self.bottomMargin,self.width,self.height,
            leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0)
        self.addPageTemplates(PageTemplate(id='manual',frames=[frame],onPage=self.page_chrome))

    def page_chrome(self,c,doc):
        c.saveState()
        c.setFillColor(BG)
        c.rect(0,0,W,H,fill=1,stroke=0)
        c.setFillColor(NAVY)
        c.rect(0,H-42,W,42,fill=1,stroke=0)
        c.setFont('WebBold',12)
        c.setFillColor(colors.white)
        c.drawString(18*mm,H-26,'TuS Raubling Handball')
        c.setFont('WebBold',8.7)
        c.drawRightString(W-18*mm,H-25,self.edition.label)
        for i,col in enumerate(RAINBOW):
            c.setFillColor(col)
            c.rect(i*W/7,H-47,W/7+0.1,5,fill=1,stroke=0)
        c.setFillColor(colors.HexColor('#E8EDF4'))
        c.roundRect(12*mm,14*mm-2,W-24*mm,H-37*mm,14,fill=1,stroke=0)
        c.setFillColor(colors.white)
        c.setStrokeColor(LINE)
        c.setLineWidth(.55)
        c.roundRect(12*mm,14*mm,W-24*mm,H-37*mm,14,fill=1,stroke=1)
        c.setFillColor(MUTED)
        c.setFont('Web',8)
        c.drawString(18*mm,8*mm,f'nuLigaHelper  ·  {self.edition.title}')
        c.setFillColor(ORANGE)
        c.roundRect(W-27*mm,5*mm,11*mm,7*mm,4,fill=1,stroke=0)
        c.setFillColor(NAVY)
        c.setFont('WebBold',9)
        c.drawCentredString(W-21.5*mm,7.3*mm,f'{doc.page:02}')
        c.restoreState()

    def afterFlowable(self,flowable):
        if isinstance(flowable,ChapterHeading):
            key=bookmark(flowable.title)
            self.canv.bookmarkPage(key)
            self.canv.addOutlineEntry(flowable.title,key,level=0)
            self.notify('TOCEntry',(0,flowable.title,self.page,key))


def make_table(rows):
    if len(rows[0])==2:
        widths=[PAGE_WIDTH*.29,PAGE_WIDTH*.71]
    elif 'Jugend' in rows[0][1]:
        widths=[PAGE_WIDTH*.25,PAGE_WIDTH*.30,PAGE_WIDTH*.45]
    else:
        widths=[PAGE_WIDTH*.20,PAGE_WIDTH*.18,PAGE_WIDTH*.62]
    cells=[[paragraph(cell,'cellhead' if i==0 else 'cell') for cell in row]
           for i,row in enumerate(rows)]
    t=Table(cells,colWidths=widths,repeatRows=1,hAlign='LEFT',cornerRadii=[8,8,8,8])
    commands=[('BACKGROUND',(0,0),(-1,0),NAVY),
        ('GRID',(0,0),(-1,-1),.5,LINE),('VALIGN',(0,0),(-1,-1),'TOP'),
        ('LEFTPADDING',(0,0),(-1,-1),9),('RIGHTPADDING',(0,0),(-1,-1),9),
        ('TOPPADDING',(0,0),(-1,-1),9),('BOTTOMPADDING',(0,0),(-1,-1),9)]
    for i in range(1,len(rows)):
        if i%2==0:
            commands.append(('BACKGROUND',(0,i),(-1,i),BG))
    t.setStyle(TableStyle(commands))
    return t


def parse_body(lines):
    story=[]
    i=0
    checklist=False
    while i<len(lines):
        line=lines[i].strip()
        if not line:
            i+=1
            continue
        if line == '### Vor deinem Einsatz':
            checklist = True
            story.extend([PageBreak(), paragraph('Deine Checkliste für den Spieltag', 'chapter'),
                          Spacer(1, 14), paragraph('Prüfe diese Punkte vor deinem Einsatz. '
                          'Die Kästchen kannst du auf einer ausgedruckten Seite abhaken.'),
                          Spacer(1, 18)])
        elif line.startswith('### '):
            story.append(paragraph(line[4:],'subheading'))
        elif line.startswith('|'):
            rows=[]
            while i<len(lines) and lines[i].strip().startswith('|'):
                cells=[c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-+:?',c) for c in cells):
                    rows.append(cells)
                i+=1
            t=make_table(rows)
            if t.wrap(PAGE_WIDTH,H)[1]<285:
                story.append(KeepTogether([t]))
            else:
                story.append(t)
            story.append(Spacer(1,10))
            continue
        elif line.startswith('!['):
            match=re.match(r'!\[([^]]*)\]\(([^)]+)\)',line)
            j=i+1
            while j<len(lines) and not lines[j].strip():
                j+=1
            caption=match.group(1)
            if j<len(lines) and lines[j].strip().startswith('*Abbildung'):
                caption=lines[j].strip().strip('*')
                i=j
            story.append(Figure(DOCS/match.group(2),caption))
            story.append(Spacer(1,10))
        elif re.match(r'^\d+\. ',line):
            steps=[]
            while i < len(lines) and re.match(r'^\d+\. ', lines[i].strip()):
                match=re.match(r'^(\d+)\. (.*)',lines[i].strip())
                cls = ChecklistStep if checklist else Step
                steps.append(cls(match.group(1),match.group(2)))
                i += 1
            total = sum(step.wrap(PAGE_WIDTH, H)[1] for step in steps)
            if total < 290:
                if (story and isinstance(story[-1], Paragraph)
                        and story[-1].style.name == 'subheading'):
                    steps.insert(0, story.pop())
                story.append(KeepTogether(steps))
            else:
                story.extend(steps)
            continue
        elif line.startswith('- '):
            st=ParagraphStyle('bullet',parent=STYLES['body'],leftIndent=14,
                              bulletIndent=0,bulletColor=ORANGE)
            story.append(Paragraph(inline(line[2:]),st,bulletText='•'))
        else:
            texts=[line]
            while i+1<len(lines) and lines[i+1].strip() and not re.match(
                r'^(#|\||!\[|\d+\. |- )',lines[i+1].strip()):
                i+=1
                texts.append(lines[i].strip())
            text=' '.join(texts)
            notice=any(text.startswith(prefix) for prefix in [
                '**Die wichtigste Regel:', '**Beispiel für die Frist:',
                '**Nach der Frist:', '**Kurzfristiger Ausfall nach der Sperre:',
                '**Auch eine Kuchenzusage ist verbindlich.',
                '**Mindestens einen erreichbaren Kontakt behalten:',
                '**Ausblenden ändert nur die Anzeige.',
                '**Erinnerungen unterstützen deine Planung, ersetzen sie aber nicht.',
            ])
            story.append(Notice(text) if notice else paragraph(text))
        i+=1
    return story


def build(edition):
    source=(DOCS/edition.source).read_text(encoding='utf-8')
    pieces=re.split(r'^## (\d+) (.+)$',source,flags=re.M)
    chapters=[(int(pieces[i]),pieces[i+1],pieces[i+2]) for i in range(1,len(pieces),3)]
    figures=len(re.findall(r'^!\[',source,re.M))
    story=[Cover(edition,len(chapters),figures),PageBreak()]
    story.append(paragraph('Inhalt','chapter'))
    story.append(Spacer(1,18))
    toc=TableOfContents()
    toc.levelStyles=[STYLES['toc']]
    toc.dotsMinLevel=0
    story.extend([toc,PageBreak()])
    intro=pieces[0].split('## Inhalt')[0]
    intro_lines=[line for line in intro.splitlines() if not line.startswith('#')]
    story.extend([paragraph('Bevor du beginnst','chapter'),Spacer(1,18)])
    story.extend(parse_body(intro_lines))
    for number,title,body in chapters:
        story.extend([PageBreak() if number == 1 else CondPageBreak(180), Spacer(1, 10)])
        body_flowables = parse_body(body.splitlines())
        # Keep a short introduction with the figure or table it explains.
        for index, flowable in enumerate(body_flowables[:-1]):
            following = body_flowables[index + 1]
            if isinstance(flowable, Paragraph) and isinstance(following, (Figure, Table, KeepTogether)):
                flowable.keepWithNext = True
        # Flatten a short opening group so ReportLab cannot separate a chapter
        # title from the first table, figure, or complete sequence of steps.
        opening = [ChapterHeading(number, title)]
        taken = 0
        for flowable in body_flowables:
            if isinstance(flowable, (PageBreak, Spacer)):
                break
            opening.extend(flowable._content if isinstance(flowable, KeepTogether) else [flowable])
            taken += 1
            if not isinstance(flowable, Paragraph):
                break
        height = sum(item.wrap(PAGE_WIDTH, H)[1] + item.getSpaceBefore()
                     + item.getSpaceAfter() for item in opening)
        if taken and height < 630:
            story.append(KeepTogether(opening))
            body_flowables = body_flowables[taken:]
        else:
            story.append(ChapterHeading(number, title))
        story.extend(body_flowables)
    target=DOCS/'pdf'/edition.output
    target.parent.mkdir(parents=True,exist_ok=True)
    Manual(edition).multiBuild(story)
    print(target)


if __name__=='__main__':
    parser=ArgumentParser(description=__doc__)
    parser.add_argument('--edition',choices=['both','all','members'],default='both')
    args=parser.parse_args()
    for edition in EDITIONS:
        if args.edition in ['both',edition.key]:
            build(edition)
