# BILIM AI SMART LESSON PRO — Google Colab: бір ұяшық
# Gemini + суреттері кестенің ішінде Word ҚМЖ + PowerPoint + HTML сабақ + 6 Premium карточка
import os, json, re, html, uuid, tempfile, subprocess, textwrap, base64, hashlib
from pathlib import Path
from datetime import date
from xml.sax.saxutils import escape as xe
from google import genai
from google.genai import types
import gradio as gr
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import cairosvg

# Render: Environment > GEMINI_API_KEY (never hardcode a key)
API_KEY = os.environ.get('GEMINI_API_KEY', '').strip()
MODEL = os.environ.get('GEMINI_MODEL', 'gemini-3.5-flash-lite')
client = genai.Client(api_key=API_KEY, http_options=types.HttpOptions(timeout=120000)) if API_KEY else None
ROOT = Path(tempfile.gettempdir()) / 'bilim_ai_pro'
ROOT.mkdir(parents=True, exist_ok=True)

def E(x): return html.escape(str(x or ''),quote=True)
def S(x):
    if isinstance(x,list): return '\n'.join(map(str,x))
    if isinstance(x,dict): return '\n'.join(f'{k}: {v}' for k,v in x.items())
    return str(x or '')
def uniq(ext): return str(ROOT/(uuid.uuid4().hex[:12]+ext))
def normalize(data):
    if not isinstance(data,dict): raise ValueError('Gemini JSON объект қайтармады')
    for name in ('intro','middle','end'):
        data.setdefault(name,{})
        for field in ('teacher','student','assessment','resources'):
            data[name][field]=S(data[name].get(field,''))
    data['cards']=data.get('cards',[])
    data['quiz']=data.get('quiz',[])
    if len(data['cards'])!=6 or len(data['quiz'])!=10:
        raise ValueError('Gemini дәл 6 карточка және 10 тест сұрағын қайтаруы тиіс.')
    for i,c in enumerate(data['cards']):
        if not isinstance(c,dict): raise ValueError('Карточка құрылымы қате')
        c.setdefault('visual_labels',[])
        if not isinstance(c['visual_labels'],list): c['visual_labels']=[]
        c['visual_labels']=[S(v) for v in c['visual_labels'][:4]]
    for q in data['quiz']:
        if len(q.get('options',[]))!=4 or not isinstance(q.get('answer'),int) or not 0<=q['answer']<=3:
            raise ValueError('Тест сұрағының құрылымы қате')
    return data

# Әдістер сабақ тақырыбына қарай таңдалады; сабақтан сабаққа әртүрлі.
BRAINSTORM = ["Миға шабуыл", "Суретті сөйлет", "Ойлан–Жұптас–Бөліс", "Болжау", "Қате пікірді тап", "Ассоциация", "Проблемалық жағдаят", "Құпия сұрақ"]
CRITICAL = ["Алты ойлау қалпағы", "Фишбоун", "Венн диаграммасы", "INSERT", "Джигсо", "Кейс-стади", "ПОПС формуласы", "SWOT талдау"]
ASSESS = ["Смарт сағат", "Бағдаршам", "Екі жұлдыз, бір тілек", "Жетістік баспалдағы", "Бас бармақ", "Өзін-өзі бағалау парағы", "Үш шапалақ", "Білім монеталары"]
REFLECT = ["3–2–1 Exit Ticket", "Бес саусақ", "Білім чемоданы", "SMS хабарлама", "Табыс ағашы", "Бүгін мен білдім", "Рефлексия нысанасы", "Болжам – Жаңа ақпарат"]
GAME_SITES = ["Quizlet — https://quizlet.com", "Kahoot — https://kahoot.com", "Wordwall — https://wordwall.net", "LearningApps — https://learningapps.org", "Wayground — https://wayground.com"]

def select_methods(subject, grade, topic, objective):
    key = f"{subject}|{grade}|{topic}|{objective}".encode('utf-8')
    h = hashlib.sha256(key).digest()
    return dict(brainstorm=BRAINSTORM[h[0]%len(BRAINSTORM)],
                critical=CRITICAL[h[1]%len(CRITICAL)],
                assessment=ASSESS[h[2]%len(ASSESS)],
                reflection=REFLECT[h[3]%len(REFLECT)],
                game=GAME_SITES[h[4]%len(GAME_SITES)])

def generate(subject,grade,topic,objective,duration,values):
    if client is None: raise RuntimeError("GEMINI_API_KEY құпия кілті Render Environment Variables бөлімінде орнатылмаған")
    schema={
      'type':'OBJECT','properties':{
        'lesson_goal':{'type':'STRING'},'values':{'type':'STRING'},'ai_usage':{'type':'STRING'},
        'intro':{'type':'OBJECT','properties':{k:{'type':'STRING'} for k in ['teacher','student','assessment','resources']},'required':['teacher','student','assessment','resources']},
        'middle':{'type':'OBJECT','properties':{k:{'type':'STRING'} for k in ['teacher','student','assessment','resources']},'required':['teacher','student','assessment','resources']},
        'end':{'type':'OBJECT','properties':{k:{'type':'STRING'} for k in ['teacher','student','assessment','resources']},'required':['teacher','student','assessment','resources']},
        'explanation':{'type':'STRING'},
        'concepts':{'type':'ARRAY','items':{'type':'OBJECT','properties':{'term':{'type':'STRING'},'meaning':{'type':'STRING'},'example':{'type':'STRING'}},'required':['term','meaning','example']}},
        'steps':{'type':'ARRAY','items':{'type':'STRING'}},
        'diagram':{'type':'OBJECT','properties':{'labels':{'type':'ARRAY','items':{'type':'STRING'}},'caption':{'type':'STRING'}},'required':['labels','caption']},
        'cards':{'type':'ARRAY','items':{'type':'OBJECT','properties':{
            'title':{'type':'STRING'},'level':{'type':'STRING'},'kind':{'type':'STRING'},'context':{'type':'STRING'},'question':{'type':'STRING'},
            'hint':{'type':'STRING'},'answer':{'type':'STRING'},'visual_labels':{'type':'ARRAY','items':{'type':'STRING'}}},
            'required':['title','level','kind','context','question','hint','answer','visual_labels']}},
        'quiz':{'type':'ARRAY','items':{'type':'OBJECT','properties':{'q':{'type':'STRING'},'options':{'type':'ARRAY','items':{'type':'STRING'}},'answer':{'type':'INTEGER'},'explanation':{'type':'STRING'}},'required':['q','options','answer','explanation']}},
        'reflection':{'type':'STRING'},'homework':{'type':'STRING'}
      },'required':['lesson_goal','values','ai_usage','intro','middle','end','explanation','concepts','steps','diagram','cards','quiz','reflection','homework']}
    selected=select_methods(subject,grade,topic,objective)
    prompt=f'''Қазақстандық тәжірибелі педагог-әдіскер ретінде бір тақырыпқа 4 өзара үйлесімді материал әзірле.
Пән: {subject}; Сынып: {grade}; Тақырып: {topic}; Оқу мақсаты (кодын өзгертпе): {objective}; Ұзақтығы: {duration} минут; Құндылық: {values}.
ТЕК осы тақырып бойынша жұмыс істе. БАСҚА ТАҚЫРЫПТЫҢ суретін, түсіндірмесін, есептерін қоспа.
Тақырыптық түсіндірме (кемінде 160 сөз), кемінде 4 ұғым, 4 қадамдық мысал, мұғалім және оқушы әрекеттері нақты жазылсын.
Сабақ кезеңдері: басы, ортасы, соңы; нақты бағалау критерийлері мен дескрипторлар көрсетілсін.
ДӘЛ 6 суретті карточка: 1-2 жеңіл, 3-4 орташа, 5-6 күрделі. Түрлері: бос орын, сәйкестендіру, суретті талдау, есеп, функционалдық сауаттылық, шығармашылық. Әр карточкадағы visual_labels сол карточканың тақырыбына нақты сәйкес келсін.
ДӘЛ 10 тест, әрқайсысында 4 жауап және 0-3 дұрыс индекс. Сызбаға 3-4 қысқа нақты термин бер.
ОСЫ САБАҚҚА ӘДІСТЕР:
Қызығушылықты ояту: {selected['brainstorm']} — 2 нақты тақырыптық сұрақ және өткізу жолы.
Сыни ойлау: {selected['critical']} — тақырыпқа сай нақты талдау тапсырмасы және күтілетін жауап.
Қалыптастырушы бағалау: {selected['assessment']} — нақты 3 дескриптор мен балл қою ережесі.
Рефлексия: {selected['reflection']} — 3 нақты сөйлем/сұрақ және өткізу жолы.
Ойын ресурсы: {selected['game']} — тек басты сайт URL-ін көрсет, нақты ойынның жасалмаған сілтемесін ОЙДАН ШЫҒАРМА.
Осы әдістерді intro/middle/end teacher, student, assessment, resources өрістеріне нақты орналастыр.
Жаңа сабақта оқушының сыни ойлауын және функционалдық сауаттылығын дамыт.
Нейрондық желі IP желісі емес. Әр пәнге өзіне сәйкес диаграмма жаса. Қазақ тілінде жауап бер.''' 
    last=None
    for attempt in range(2):
        try:
            resp=client.models.generate_content(model=MODEL,contents=prompt,
              config=types.GenerateContentConfig(response_mime_type='application/json',response_schema=schema,temperature=0.25,max_output_tokens=11000))
            result=normalize(json.loads(resp.text))
            result['methods']=selected
            for phase, label, key in [('intro','Қызығушылықты ояту','brainstorm'),('middle','Сыни ойлау','critical'),('end','Рефлексия','reflection')]:
                if selected[key].lower() not in result[phase]['teacher'].lower():
                    result[phase]['teacher'] += '\n' + label + ': ' + selected[key]
            result['middle']['resources'] += '\nОйын платформасы: ' + selected['game']
            result['middle']['assessment'] += '\nБағалау әдісі: ' + selected['assessment']
            return result
        except Exception as e:
            last=e
    raise RuntimeError(f'Gemini жауап құрылымын жасау мүмкін болмады: {str(last)[:260]}')

def svg_diagram(labels,caption=''):
    labels=[S(x).strip() for x in (labels or []) if S(x).strip()][:4]
    if not labels: labels=['Негізгі ұғым','Талдау','Нәтиже']
    w=1000;n=len(labels); gap=14; boxw=(w-64-gap*(n-1))/n
    nodes=[]
    for i,label in enumerate(labels):
        x=32+i*(boxw+gap); color=['#fce7f3','#dbeafe','#dcfce7','#fef3c7'][i%4]
        words=textwrap.wrap(label,19,break_long_words=False)[:3]
        lines=''.join(f'<text x="{x+boxw/2:.0f}" y="{116+j*25}" font-size="19" text-anchor="middle" font-weight="700" fill="#17356c">{xe(line)}</text>' for j,line in enumerate(words))
        nodes.append(f'<rect x="{x:.0f}" y="62" width="{boxw:.0f}" height="136" rx="20" fill="{color}" stroke="#9bb4d9" stroke-width="2"/>{lines}')
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 255"><rect width="1000" height="255" rx="25" fill="#f1f6ff"/>{"".join(nodes)}<text x="500" y="237" text-anchor="middle" font-size="17" fill="#4b6387">{xe(S(caption)[:95])}</text></svg>'
def diagram_png(labels,caption=''):
    p=uniq('.png'); cairosvg.svg2png(bytestring=svg_diagram(labels,caption).encode(),write_to=p,output_width=1100);return p

def make_docx(meta,d):
    # 11-сынып үлгісі: тапсырмалар, суреттер, бағалау және ресурстар
    # сабақ барысының БІР кестесінің ІШІНДЕ орналасады.
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.table import WD_TABLE_ALIGNMENT

    doc=Document()
    sec=doc.sections[0]
    sec.page_width=Cm(21);sec.page_height=Cm(29.7)
    sec.top_margin=Cm(1.45);sec.bottom_margin=Cm(1.45)
    sec.left_margin=Cm(1.25);sec.right_margin=Cm(1.25)
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman';normal.font.size=Pt(10)
    normal.paragraph_format.space_after=Pt(3)

    def write(cell,text,bold=False):
        cell.text=''
        p=cell.paragraphs[0]
        r=p.add_run(S(text));r.bold=bold;r.font.name='Times New Roman';r.font.size=Pt(9)
        return p

    def add(cell,title,body='',bold=True):
        p=cell.add_paragraph()
        r=p.add_run(S(title));r.bold=bold;r.font.name='Times New Roman';r.font.size=Pt(9)
        if body:
            p=cell.add_paragraph(S(body))
            for r in p.runs:r.font.name='Times New Roman';r.font.size=Pt(9)

    def shade(cell,color):
        tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd')
        shd.set(qn('w:fill'),color);tcPr.append(shd)

    def border_table(table):
        tblPr=table._tbl.tblPr
        borders=OxmlElement('w:tblBorders')
        for edge in ('top','left','bottom','right','insideH','insideV'):
            el=OxmlElement('w:'+edge)
            el.set(qn('w:val'),'single');el.set(qn('w:sz'),'6')
            el.set(qn('w:color'),'4B85B4');borders.append(el)
        tblPr.append(borders)

    def add_pic(cell,labels,caption='',width=6.5):
        labels=labels if isinstance(labels,list) else []
        if not labels:return
        path=diagram_png(labels,caption)
        p=cell.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path,width=Cm(width))

    heading=doc.add_paragraph()
    heading.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=heading.add_run('Қысқа мерзімді (сабақ) жоспары')
    run.bold=True;run.font.size=Pt(13)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(S(meta['topic']));r.bold=True;r.font.size=Pt(11)

    info=doc.add_table(rows=0,cols=2);info.autofit=False
    info.columns[0].width=Cm(4.3);info.columns[1].width=Cm(14.2)
    fields=[('Бекітемін','Директордың оқу ісі жөніндегі орынбасары'),
            ('Ұзақ мерзімді жоспар бөлімі',meta['unit']),
            ('Мектеп',meta['school']),('Күні',meta['day']),
            ('Мұғалімнің аты-жөні',meta['teacher']),
            ('Сынып',meta['grade']),('Қатысқан / қатыспаған оқушылар саны','_____ / _____'),
            ('Сабақ тақырыбы',meta['topic']),('Оқу бағдарламасына сәйкес оқу мақсаты',meta['objective']),
            ('Сабақ мақсаты',d.get('lesson_goal','')),
            ('Құндылықтарды дарыту',d.get('values',''))]
    for label,value in fields:
        cells=info.add_row().cells
        write(cells[0],label,True);write(cells[1],value)
        shade(cells[0],'EAF3FC')
    border_table(info)

    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Сабақ барысы');r.bold=True;r.font.size=Pt(11)
    tab=doc.add_table(rows=1,cols=5);tab.autofit=False
    widths=[2.2,8.1,3.1,2.65,2.45]
    for col,w in zip(tab.columns,widths):col.width=Cm(w)
    tab.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,title in enumerate(['Сабақ кезеңі / уақыт','Педагогтің әрекеті','Оқушының әрекеті','Бағалау','Ресурстар']):
        write(tab.cell(0,i),title,True);shade(tab.cell(0,i),'DCECFB')
    dur=int(meta['duration']);start=max(5,round(dur*.12));finish=max(5,round(dur*.12))
    phases=[('Сабақтың басы',start,'intro'),('Сабақтың ортасы',dur-start-finish,'middle'),('Сабақтың соңы',finish,'end')]
    for title,mins,key in phases:
        block=d.get(key,{})
        row=tab.add_row().cells
        write(row[0],f'{title}\n{mins} минут',True)
        write(row[1],block.get('teacher',''))
        write(row[2],block.get('student',''))
        write(row[3],block.get('assessment',''))
        write(row[4],block.get('resources',''))
        if key=='intro':
            add(row[1],'Қызығушылықты ояту',d.get('intro',{}).get('teacher',''))
        if key=='middle':
            add(row[1],'Жаңа тақырыпты меңгеру',d.get('explanation',''))
            diag=d.get('diagram',{})
            add(row[1],'Тақырыптық сызба')
            add_pic(row[1],diag.get('labels',[]),diag.get('caption',''),7.2)
            for i,card in enumerate(d.get('cards',[]),1):
                add(row[1],f'{i}-тапсырма. {card.get("title","")} ({card.get("level","")})',
                    S(card.get('context',''))+'\n'+S(card.get('question','')))
                add_pic(row[1],card.get('visual_labels',[]),'',6.7)
                add(row[1],'Оқушы жауабы: __________________________',bold=False)
                add(row[2],f'{i}-тапсырма',card.get('kind','')+' бойынша тапсырманы орындайды.')
                add(row[3],f'{i}-тапсырма дескрипторы',
                    '1) Негізгі ұғымды дұрыс қолданады (1 балл); 2) шешімін дәлелдейді (1 балл); 3) қорытынды жасайды (1 балл).')
                add(row[4],f'{i}-карточка','Суретті жұмыс парағы')
        if key=='end':
            add(row[1],'Рефлексия',d.get('reflection',''))
            add(row[1],'Үй тапсырмасы',d.get('homework',''))
        for c in row:
            c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    border_table(tab)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name='Times New Roman'
                        if not r.font.size:r.font.size=Pt(9)
    path=uniq('.docx');doc.save(path);return path

STYLE='''*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;background:#edf3ff;color:#19335c;margin:0}.top{padding:30px;background:linear-gradient(115deg,#2465ef,#8c39e8);color:white}.top h1{color:white;margin:0}.wrap{max-width:1180px;margin:25px auto;padding:0 18px}.panel{background:white;padding:25px;border-radius:20px;margin:16px 0;box-shadow:0 8px 25px #122c5210}.panel h2{color:#173f95}.diagram svg{width:100%}.choice{display:block;background:#f1f6ff;padding:12px;margin:8px 0;border-radius:10px;cursor:pointer}button{background:#245ee8;color:white;border:0;border-radius:11px;padding:13px 23px;cursor:pointer;font-weight:700}#score{padding:14px;font-weight:700}'''
CARD_STYLE='''*{box-sizing:border-box}body{margin:0;background:#eaf1ff;font-family:Segoe UI,Arial,sans-serif;color:#17335d}.toolbar{padding:22px;max-width:1200px;margin:auto}.toolbar button{background:#255ce5;color:white;border:0;padding:12px 20px;border-radius:10px;cursor:pointer}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:24px;max-width:1260px;margin:auto;padding:16px}.card{width:100%;max-width:420px;aspect-ratio:9/16;background:#142743;border:4px solid #d5a955;border-radius:20px;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 14px 28px #0d28522b}.card header,.card footer{color:#f5ce7e;padding:13px 17px;font-size:12px;font-weight:700}.sheet{flex:1;background:white;margin:0 9px;padding:18px;border-radius:12px;overflow:auto}.sheet h2{color:#173f95;font-size:22px}.tag{display:inline-block;background:#e8f0ff;border-radius:20px;color:#2459ac;padding:6px 11px;font-size:12px}.illustration svg{width:100%;margin:10px 0}.task{background:#eff5ff;padding:13px;border-left:5px solid #7c3aed;border-radius:8px;font-size:14px;line-height:1.55}textarea{width:100%;height:72px;border:1px solid #aabbd4;border-radius:9px;padding:9px}details{padding:10px;margin-top:10px;background:#f8faff;border-radius:9px;font-size:13px}summary{cursor:pointer;color:#2459ac;font-weight:700}@media print{@page{size:A4;margin:8mm}.toolbar{display:none}.grid{display:block}.card{width:110mm;height:195mm;min-height:0;margin:0 auto 10mm;break-after:page;box-shadow:none}.sheet{overflow:hidden}}'''

def make_html(meta,d):
    cards=[]
    for i,c in enumerate(d['cards'],1):
        cards.append(f'''<article class="card"><header>✦ BILIM AI · PREMIUM CARD <span style="float:right">{i:02d}/06</span></header><div class="sheet"><small>{E(meta['grade'])}-СЫНЫП · {E(meta['subject'])}</small><h2>{E(c.get('title'))}</h2><span class="tag">{E(c.get('level'))} · {E(c.get('kind'))}</span><p>{E(c.get('context'))}</p><div class="illustration">{svg_diagram(c.get('visual_labels',[]))}</div><div class="task"><b>ТАПСЫРМА</b><p>{E(c.get('question'))}</p></div><p><b>Жауабыңыз:</b></p><textarea></textarea><details><summary>Кеңес пен дұрыс жауап</summary><p>{E(c.get('hint'))}</p><p><b>Жауап:</b> {E(c.get('answer'))}</p></details></div><footer>ОЙЛАН · ОРЫНДА · ТЕКСЕР</footer></article>''')
    cards_file=uniq('.html');Path(cards_file).write_text(f'<!doctype html><html lang="kk"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>6 Premium карточка</title><style>{CARD_STYLE}</style><div class="toolbar"><h1>6 Premium карточка · {E(meta["topic"])}</h1><button onclick="print()">Басып шығару</button></div><div class="grid">{"".join(cards)}</div></html>',encoding='utf-8')
    concepts=''.join(f'<div class="panel"><b>{E(c.get("term"))}</b><p>{E(c.get("meaning"))}</p><small>{E(c.get("example"))}</small></div>' for c in d.get('concepts',[]) if isinstance(c,dict))
    steps=''.join(f'<li>{E(x)}</li>' for x in d.get('steps',[]))
    qs=[]
    for i,q in enumerate(d['quiz']):
        options=''.join(f'<label class="choice"><input type="radio" name="q{i}" value="{j}"> {E(opt)}</label>' for j,opt in enumerate(q['options']))
        qs.append(f'<div class="panel"><b>{i+1}. {E(q["q"])}</b>{options}</div>')
    answer_json=json.dumps([{'answer':q['answer'],'explanation':q['explanation']} for q in d['quiz']],ensure_ascii=False).replace('<','\\u003c')
    page=f'''<!doctype html><html lang="kk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{E(meta['topic'])}</title><style>{STYLE}</style></head><body><div class="top"><h1>✦ BILIM AI · {E(meta['topic'])}</h1><p>{E(meta['subject'])} · {E(meta['grade'])}-сынып · {E(meta['objective'])}</p></div><main class="wrap"><section class="panel"><h2>Тақырыптық түсіндірме</h2><p style="white-space:pre-line">{E(d['explanation'])}</p><div class="diagram">{svg_diagram(d.get('diagram',{}).get('labels',[]),d.get('diagram',{}).get('caption',''))}</div></section><section><h2>Негізгі ұғымдар</h2>{concepts}</section><section class="panel"><h2>Қадамдық мысалдар</h2><ol>{steps}</ol></section><section class="panel"><h2>6 суретті тапсырма</h2><p>6 Premium карточка жеке HTML файлында беріледі, оны интернетсіз ашуға және басып шығаруға болады.</p></section><section><h2>10 сұрақтық тест</h2>{''.join(qs)}<button onclick="checkQuiz()">Тестті тексеру</button><div id="score"></div></section><section class="panel"><h2>Рефлексия</h2><p>{E(d.get('reflection'))}</p></section></main><script>const key={answer_json};function checkQuiz(){{let n=0,notes=[];key.forEach((a,i)=>{{const x=document.querySelector('input[name="q'+i+'"]:checked');if(x&&Number(x.value)===a.answer)n++;else notes.push((i+1)+'. '+a.explanation)}});document.getElementById('score').textContent='Нәтиже: '+n+'/10 ('+n*10+'%). '+notes.join(' | ')}}</script></body></html>'''
    lesson_file=uniq('.html');Path(lesson_file).write_text(page,encoding='utf-8')
    return lesson_file,cards_file

def make_pptx(meta,d):
    """16 editable slides, premium educational style."""
    from pptx import Presentation
    from pptx.util import Inches, Pt as PPTPt
    from pptx.dml.color import RGBColor as PPTColor
    from pptx.enum.text import PP_ALIGN
    prs=Presentation()
    prs.slide_width=Inches(13.333)
    prs.slide_height=Inches(7.5)
    def rect(sl,x,y,w,h,color):
        from pptx.enum.shapes import MSO_SHAPE
        sh=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
        sh.fill.solid();sh.fill.fore_color.rgb=PPTColor.from_string(color)
        sh.line.fill.background()
        return sh
    def tx(sl,content,x,y,w,h,size=19,color='173B86',bold=False):
        sh=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
        tf=sh.text_frame;tf.word_wrap=True
        tf.margin_left=Inches(.08);tf.margin_right=Inches(.08)
        for i,line in enumerate(str(content or '').split('\n')):
            para=tf.paragraphs[0] if i==0 else tf.add_paragraph()
            para.text=line;para.font.name='Arial';para.font.size=PPTPt(size)
            para.font.bold=bold;para.font.color.rgb=PPTColor.from_string(color)
        return sh
    def slide(title,idx):
        sl=prs.slides.add_slide(prs.slide_layouts[6])
        sl.background.fill.solid();sl.background.fill.fore_color.rgb=PPTColor(245,248,255)
        rect(sl,0,0,13.333,.15,'2563EB')
        tx(sl,title,.65,.42,12,.8,28,'173B86',True)
        tx(sl,f'BILIM AI · SMART LESSON PRO                                   {idx:02d}',.7,7.1,11.8,.24,9,'627698')
        return sl
    def panel(sl,content,x,y,w,h,color='EAF2FF',size=18):
        rect(sl,x,y,w,h,color)
        tx(sl,content,x+.16,y+.12,w-.32,h-.24,size,'243C68')
    sl=prs.slides.add_slide(prs.slide_layouts[6]);sl.background.fill.solid()
    sl.background.fill.fore_color.rgb=PPTColor(23,59,134)
    tx(sl,'BILIM AI — SMART LESSON PRO',.8,.85,11,.7,28,'FFFFFF',True)
    tx(sl,meta['topic'],.8,2.0,11.7,2.3,37,'FFFFFF',True)
    tx(sl,f"{meta['subject']} · {meta['grade']}-сынып",.8,5.35,11,.5,22,'DBEAFE')
    sl=slide('Сабақ мақсаты',2);panel(sl,d['lesson_goal'],.8,1.6,11.7,2.0,size=23);panel(sl,meta['objective'],.8,4.0,11.7,1.4,'F2EAFE')
    sl=slide('Қызығушылықты ояту',3);panel(sl,d['intro']['teacher'],.8,1.6,11.7,2.7);panel(sl,d['intro']['student'],.8,4.6,11.7,1.5)
    sl=slide('Тақырыптық түсіндірме',4);panel(sl,d['explanation'],.8,1.6,11.7,4.9,'FFFFFF',17)
    sl=slide('Тақырыптық сызба',5)
    labels=d.get('diagram',{}).get('labels',[])[:4]
    if labels:
        width=11.5/len(labels)
        for i,label in enumerate(labels):panel(sl,label,.8+i*width,2.35,width-.15,2.0,['FCE7F3','DBEAFE','DCFCE7','FEF3C7'][i%4],16)
    tx(sl,d.get('diagram',{}).get('caption',''),.9,5.2,11,.8,17)
    sl=slide('Негізгі ұғымдар',6)
    for i,c in enumerate(d['concepts'][:4]):panel(sl,c['term']+' — '+c['meaning'],.8+(i%2)*6,1.55+(i//2)*2.35,5.7,2.05,['EAF3FF','F1EBFF'][i%2],15)
    sl=slide('Қадамдық мысалдар',7)
    for i,item in enumerate(d['steps'][:5]):panel(sl,f'{i+1}. {item}',.8,1.48+i*1.03,11.7,.91,'EBF2FF' if i%2==0 else 'FFFFFF',15)
    for i,c in enumerate(d['cards'][:6]):
        sl=slide(f"Тапсырма {i+1}: {c['title']}",8+i)
        panel(sl,c['context'],.8,1.55,11.7,1.3,size=17)
        panel(sl,c['question'],.8,3.1,11.7,2.5,'FFFFFF',19)
        tx(sl,c['level']+' · '+c['kind'],.9,6.0,11,.4,15,'7847D6',True)
    sl=slide('Ойын және цифрлық ресурстар',14)
    panel(sl,d.get('methods',{}).get('game','Quizlet / Kahoot / Wordwall'),.8,1.6,11.7,1.5,'EAF2FF',20)
    panel(sl,'Тапсырмаларға арналған сұрақтар мен жауаптар дайын. Нақты ойын платформасында мұғалім өзі жариялайды.',.8,3.5,11.7,1.6,'F3EAFF',17)
    sl=slide('Сыни ойлау және бағалау',15)
    panel(sl,d.get('methods',{}).get('critical','Сыни ойлау'),.8,1.6,11.7,1.3,'EAF2FF',22)
    panel(sl,d.get('middle',{}).get('assessment',''),.8,3.3,11.7,2.5,'FFFFFF',17)
    sl=slide('Қорытынды және рефлексия',16)
    panel(sl,d['reflection'],.8,1.6,11.7,2.0,size=20)
    panel(sl,'Үй тапсырмасы: '+d['homework'],.8,4.1,11.7,1.8,'F2EAFE')
    out=uniq('.pptx');prs.save(out);return out

def run(school,teacher,day,unit,subject,grade,topic,objective,duration,values):
    missing=[name for name,value in [('Мектеп',school),('Мұғалім',teacher),('Пән',subject),('Сынып',grade),('Тақырып',topic),('Оқу мақсаты',objective)] if not S(value).strip()]
    if missing: raise gr.Error('Міндетті өрістер: '+', '.join(missing))
    meta=dict(school=school,teacher=teacher,day=day,unit=unit,subject=subject,grade=grade,topic=topic,objective=objective,duration=int(duration))
    try:
        d=generate(subject,grade,topic,objective,int(duration),values)
        word=make_docx(meta,d);lesson,cards=make_html(meta,d);pptx=make_pptx(meta,d)
        preview=f'''<div class="pro-preview"><h2>🎓 {E(topic)}</h2><p>🎯 <b>Сабақ мақсаты:</b> {E(d['lesson_goal'])}</p><p>💗 <b>Құндылықтар:</b> {E(d['values'])}</p><p>📊 <b>Тапсырмалар:</b> 6 суретті карточка (жеке, жұптық, функционалдық)</p><p>☷ <b>Тест:</b> 10 сұрақ, автоматты тексеру</p></div>'''
        return preview,word,pptx,lesson,cards,'<div class="status-success">✅ &nbsp; Файлдар сәтті дайындалды!</div>'
    except Exception as e:
        raise gr.Error('Генерация қатесі: '+str(e)[:450])

CSS = r"""
:root {color-scheme:light!important;--body-background-fill:#f1f6ff;--block-background-fill:#fff;--body-text-color:#123466;--block-label-text-color:#0a3d94;--input-background-fill:#fff;}
html,body,gradio-app,.gradio-container {background:#f1f6ff!important;color:#16345e!important;color-scheme:light!important;font-family:'Segoe UI',Arial,sans-serif!important;}
.gradio-container{max-width:1680px!important;margin:auto!important;padding:10px 22px 30px!important;}
#bilim-hero{background:linear-gradient(105deg,#1169f9,#3855eb 56%,#a13aee)!important;border-radius:18px!important;padding:20px 22px!important;margin:0 0 10px!important;}
#bilim-hero,#bilim-hero *{color:#fff!important;}
#bilim-hero .hero-grid{display:flex;justify-content:space-between;align-items:center;gap:18px;flex-wrap:wrap;}
#bilim-hero h1{font-size:32px!important;font-weight:850!important;margin:0!important;line-height:1.2!important;}
#bilim-hero p{margin:7px 0 0!important;font-size:14px!important;}
#bilim-hero .hero-icons{display:flex;gap:9px;}
#bilim-hero .hero-icon{background:#ffffff25;border:1px solid #ffffff38;border-radius:12px;min-width:102px;text-align:center;padding:8px;font-size:12px;}
#bilim-hero .hero-icon b{display:block;font-size:22px;}
#form-panel,#result-panel{background:#fff!important;border:1px solid #e7edf7!important;border-radius:17px!important;padding:17px!important;box-shadow:0 5px 20px #183d6a0b!important;}
#form-panel .block,#form-panel .form,#form-panel .wrap,#form-panel .container{background:#fff!important;color:#143a78!important;box-shadow:none!important;}
#form-panel label,#form-panel label span,#form-panel .label-wrap,#form-panel .label-wrap span{color:#083e98!important;background:transparent!important;font-weight:750!important;font-size:14px!important;}
#form-panel input,#form-panel textarea,#form-panel select,#form-panel [role=combobox]{background:#fff!important;color:#1b355a!important;border:1px solid #d8e2ee!important;border-radius:9px!important;}
#form-panel button.primary{background:linear-gradient(100deg,#246cf9,#9639f2)!important;color:#fff!important;border:0!important;border-radius:12px!important;min-height:54px!important;font-size:16px!important;font-weight:800!important;}
#result-panel .status-ready,#result-panel .status-success{display:flex;align-items:center;min-height:48px;border-radius:12px;padding:12px 18px!important;font-weight:800!important;font-size:15px!important;line-height:1.4!important;box-shadow:none!important;}
#result-panel .status-ready{background:#e9f3ff!important;color:#164f9d!important;}
#result-panel .status-success{background:#e5f8ee!important;color:#087443!important;}
#result-panel .status-ready *,#result-panel .status-success *{color:inherit!important;background:transparent!important;}
#result-panel .pro-preview{display:block!important;min-height:180px!important;}
#result-panel .pro-preview{background:#eaf3ff!important;border:0!important;border-radius:14px!important;padding:20px!important;}
#result-panel .pro-preview,#result-panel .pro-preview *{color:#1b355f!important;opacity:1!important;}
#result-panel .pro-preview h2{color:#1251bd!important;font-size:23px!important;}
#result-panel .download-row{border-radius:13px!important;border:1px solid #d0eadc!important;background:#e9faf3!important;padding:11px 13px!important;margin:3px 0!important;align-items:center!important;min-height:82px!important;}
#result-panel .download-row.ppt{background:#fff4e8!important;border-color:#f7dfc1!important;}
#result-panel .download-row.web{background:#e8f2ff!important;border-color:#c9e0ff!important;}
#result-panel .download-row.cards{background:#f3eaff!important;border-color:#e3d3fc!important;}
#result-panel .download-row .download-copy{flex:1!important;min-width:0!important;line-height:1.55!important;}
#result-panel .download-row .download-button{flex:0 0 auto!important;}
#result-panel .download-row .download-button button{padding:10px 18px!important;}
#result-panel .download-row .download-copy strong{display:inline-block!important;}
#result-panel .download-row .download-copy small{display:inline-block!important;margin-top:3px!important;}
#result-panel .download-row{gap:12px!important;}
#result-panel .download-row .block,#result-panel .download-row .form,#result-panel .download-row .wrap{background:transparent!important;border:0!important;box-shadow:none!important;}
#result-panel .download-row .download-copy,#result-panel .download-row .download-copy *{background:transparent!important;color:#1d355b!important;opacity:1!important;}
#result-panel .download-row .download-copy strong{font-size:17px!important;color:#075cce!important;}
#result-panel .download-row.ppt .download-copy strong{color:#e56b08!important;}
#result-panel .download-row.cards .download-copy strong{color:#7c2bdf!important;}
#result-panel .download-row .download-copy small{font-size:12px!important;color:#607391!important;}
#result-panel .download-row .download-button{min-width:132px!important;}
#result-panel .download-row .download-button button,#result-panel .download-row .download-button a{background:#08a76b!important;border:0!important;border-radius:10px!important;color:white!important;font-weight:800!important;min-height:46px!important;box-shadow:none!important;}
#result-panel .download-row.ppt .download-button button,#result-panel .download-row.ppt .download-button a{background:#ff710b!important;}
#result-panel .download-row.web .download-button button,#result-panel .download-row.web .download-button a{background:#1767f5!important;}
#result-panel .download-row.cards .download-button button,#result-panel .download-row.cards .download-button a{background:#8d30eb!important;}
#result-panel .help-panel{background:#fff8e5!important;color:#223b64!important;padding:15px 18px!important;border-radius:13px!important;line-height:1.65!important;}
#result-panel .help-panel *{color:#223b64!important;}
@media(max-width:900px){#bilim-hero h1{font-size:23px!important;}#bilim-hero .hero-icons{width:100%;}#bilim-hero .hero-icon{flex:1;min-width:0;}#form-panel,#result-panel{padding:12px!important;}}
"""

with gr.Blocks(theme=gr.themes.Soft(primary_hue='blue'),css=CSS,title='BILIM AI — SMART LESSON PRO') as demo:
    gr.HTML("""<div id="bilim-hero"><div class="hero-grid"><div><h1>🎓 BILIM AI — SMART LESSON PRO</h1><p>Gemini AI · Суретті Word ҚМЖ · PowerPoint · Интерактивті HTML · 6 Premium карточка</p></div><div class="hero-icons"><div class="hero-icon"><b>📄</b>Word ҚМЖ</div><div class="hero-icon"><b>📊</b>PowerPoint</div><div class="hero-icon"><b>🌐</b>HTML сабақ</div><div class="hero-icon"><b>🎨</b>6 карточка</div></div></div></div>""")
    with gr.Row(equal_height=False):
        with gr.Column(scale=5,min_width=320,elem_id='form-panel'):
            school=gr.Textbox(label='🏫 Мектеп',value='№103 жалпы орта білім беретін мектеп')
            teacher=gr.Textbox(label='👤 Мұғалімнің аты-жөні')
            day=gr.Textbox(label='📅 Күні',value=date.today().strftime('%d.%m.%Y'))
            unit=gr.Textbox(label='📚 Ұзақ мерзімді жоспар бөлімі')
            subject=gr.Textbox(label='💻 Пән',value='Информатика')
            grade=gr.Dropdown([str(i) for i in range(1,12)],value='7',label='🎓 Сынып')
            topic=gr.Textbox(label='📝 Сабақ тақырыбы',lines=2)
            objective=gr.Textbox(label='🎯 Оқу мақсаты (код және мәтін)',lines=3)
            duration=gr.Dropdown([40,45,80,90],value=45,label='⏱ Сабақ ұзақтығы')
            values=gr.Textbox(label='💗 Құндылықтар',value='Жауапкершілік, ынтымақтастық')
            btn=gr.Button('✨ ҚМЖ + презентация + HTML + 6 карточка құру',variant='primary',size='lg')
        with gr.Column(scale=9,min_width=390,elem_id='result-panel'):
            status=gr.HTML('<div class="status-ready">✦ &nbsp; Сабақ мәліметтерін енгізіп, генерация батырмасын басыңыз.</div>')
            preview=gr.HTML('<div class="pro-preview"><h2>🎓 Сабақ нәтижесі</h2><p>Генерациядан кейін сабақтың мақсаты мен материалдары осы жерде көрсетіледі.</p></div>')
            with gr.Row(elem_classes=['download-row','word']):
                gr.HTML('<div class="download-copy">📘 <strong>Word ҚМЖ (.docx)</strong><br><small>Суреттермен безендірілген қысқа мерзімді жоспар</small></div>')
                word=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','ppt']):
                gr.HTML('<div class="download-copy">📙 <strong>PowerPoint презентациясы (.pptx)</strong><br><small>16 слайд, инфографика және тапсырмалар</small></div>')
                ppt=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','web']):
                gr.HTML('<div class="download-copy">🌐 <strong>Интерактивті HTML сабағы</strong><br><small>Түсіндірме, тапсырмалар және 10 сұрақтық тест</small></div>')
                lesson=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','cards']):
                gr.HTML('<div class="download-copy">🎨 <strong>6 Premium карточка (HTML)</strong><br><small>9:16 форматындағы суретті тапсырмалар</small></div>')
                cards=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            gr.HTML('<div class="help-panel"><b>💡 Қолдану нұсқаулығы:</b><br>1. Word ҚМЖ — суретті сабақ жоспары және дескрипторлар.<br>2. PowerPoint — сабақ презентациясы.<br>3. HTML сабақ — түсіндірме, тапсырмалар және автоматты тест.<br>4. Premium карточкалар — 9:16 форматындағы 6 тапсырма.</div>')
    btn.click(fn=run,inputs=[school,teacher,day,unit,subject,grade,topic,objective,duration,values],outputs=[preview,word,ppt,lesson,cards,status],show_progress='full')
demo.queue(default_concurrency_limit=1)
demo.launch(server_name='0.0.0.0',server_port=7860,share=False,show_error=True)
