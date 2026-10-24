# BILIM AI SMART LESSON PRO — Render / GitHub · Interactive Lab
# Gemini + суреттері кестенің ішінде Word ҚМЖ + PowerPoint + HTML сабақ + 6 Premium карточка
import os, json, re, html, uuid, tempfile, subprocess, textwrap, base64, hashlib, secrets, time
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

# Canva OAuth 2.0 Authorization Code + PKCE. No tokens are exposed to the browser.
# Session is kept in memory (single Render instance); restart requires reconnecting.
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
import httpx

CANVA_BASE_URL = os.environ.get('CANVA_BASE_URL', 'https://bilim-ai-smart-lesson-pro.onrender.com').rstrip('/')
CANVA_REDIRECT_URI = CANVA_BASE_URL + '/canva/callback'
CANVA_CLIENT_ID = os.environ.get('CANVA_CLIENT_ID', '').strip()
CANVA_CLIENT_SECRET = os.environ.get('CANVA_CLIENT_SECRET', '').strip()
CANVA_SCOPES = 'design:meta:read design:content:read design:content:write asset:read asset:write'
CANVA_SIGNING_KEY = os.environ.get('CANVA_SESSION_SECRET', '').strip() or CANVA_CLIENT_SECRET
canva_sessions = {}  # temporary in-memory tokens; never write to logs or files
canva_app = FastAPI()

def canva_message(message, ok=False):
    color = '#087443' if ok else '#9a3412'
    return HTMLResponse(f'<html lang="kk"><meta charset="utf-8"><title>Canva</title><body style="font-family:Segoe UI,Arial;padding:48px;background:#f1f6ff"><main style="max-width:650px;margin:auto;background:white;border-radius:18px;padding:32px"><h2 style="color:{color}">{html.escape(message)}</h2><p><a href="/">BILIM AI сайтына қайту</a></p></main></body></html>')

@canva_app.get('/canva/connect')
async def canva_connect():
    if not CANVA_CLIENT_ID or not CANVA_CLIENT_SECRET or not CANVA_SIGNING_KEY:
        return canva_message('Canva кілттері Render Environment ішінде орнатылмаған.')
    from urllib.parse import urlencode
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip('=')
    state = secrets.token_urlsafe(24)
    signer = URLSafeTimedSerializer(CANVA_SIGNING_KEY, salt='bilim-canva-oauth-v1')
    cookie = signer.dumps({'state':state,'verifier':verifier})
    params = dict(response_type='code', client_id=CANVA_CLIENT_ID, redirect_uri=CANVA_REDIRECT_URI,
                  scope=CANVA_SCOPES, state=state, code_challenge=challenge, code_challenge_method='S256')
    response = RedirectResponse('https://www.canva.com/api/oauth/authorize?'+urlencode(params),status_code=302)
    response.set_cookie('bilim_canva_pkce',cookie,httponly=True,secure=True,samesite='lax',max_age=600,path='/canva')
    return response

@canva_app.get('/canva/callback')
async def canva_callback(request:Request):
    if request.query_params.get('error'):
        return canva_message('Canva рұқсат бермеді: '+request.query_params.get('error',''))
    code = request.query_params.get('code','')
    state = request.query_params.get('state','')
    signed = request.cookies.get('bilim_canva_pkce','')
    try:
        data=URLSafeTimedSerializer(CANVA_SIGNING_KEY,salt='bilim-canva-oauth-v1').loads(signed,max_age=600)
    except (BadSignature,SignatureExpired):
        return canva_message('Авторизация мерзімі аяқталды. Canva-ға қайта қосылыңыз.')
    if not code or not state or not secrets.compare_digest(state,data.get('state','')):
        return canva_message('Canva авторизациясының қауіпсіздік тексеруі сәтсіз.')
    try:
        async with httpx.AsyncClient(timeout=25) as http:
            result=await http.post('https://api.canva.com/rest/v1/oauth/token',
                auth=(CANVA_CLIENT_ID,CANVA_CLIENT_SECRET),
                data={'grant_type':'authorization_code','code':code,'redirect_uri':CANVA_REDIRECT_URI,
                      'code_verifier':data['verifier']})
            result.raise_for_status()
            tokens=result.json()
        if not tokens.get('access_token'):
            raise ValueError('Canva access token қайтармады')
        sid=secrets.token_urlsafe(32)
        canva_sessions[sid]={'access_token':tokens['access_token'],
            'refresh_token':tokens.get('refresh_token'),
            'expires_at':time.time()+int(tokens.get('expires_in',3600))}
        response=canva_message('Canva аккаунты сәтті қосылды! OAuth байланысы жұмыс істейді.',True)
        response.set_cookie('bilim_canva_session',sid,httponly=True,secure=True,samesite='lax',max_age=3600,path='/canva')
    except (httpx.HTTPError,ValueError):
        response=canva_message('Canva токенін алу мүмкін болмады. Redirect URL мен рұқсаттарды тексеріңіз.')
    response.delete_cookie('bilim_canva_pkce',path='/canva')
    return response

@canva_app.get('/canva/status')
async def canva_status(request:Request):
    sid=request.cookies.get('bilim_canva_session','')
    session=canva_sessions.get(sid)
    active=bool(session and session['expires_at']>time.time())
    return canva_message('Canva қосылған.' if active else 'Canva қосылмаған немесе сеанс аяқталған.',active)

@canva_app.get('/canva/disconnect')
async def canva_disconnect(request:Request):
    sid=request.cookies.get('bilim_canva_session','')
    canva_sessions.pop(sid,None)
    response=RedirectResponse('/',status_code=302)
    response.delete_cookie('bilim_canva_session',path='/canva')
    return response


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
    """Offline, topic-aware illustrated lesson and six printable premium cards."""
    cards=[]
    for i,c in enumerate(d['cards'],1):
        cards.append(f'''<article class="card"><header>✦ BILIM AI · PREMIUM CARD <span style="float:right">{i:02d}/06</span></header><div class="sheet"><small>{E(meta['grade'])}-СЫНЫП · {E(meta['subject'])}</small><h2>{E(c.get('title'))}</h2><span class="tag">{E(c.get('level'))} · {E(c.get('kind'))}</span><p>{E(c.get('context'))}</p><div class="illustration">{svg_diagram(c.get('visual_labels',[]))}</div><div class="task"><b>ТАПСЫРМА</b><p>{E(c.get('question'))}</p></div><p><b>Жауабыңыз:</b></p><textarea></textarea><details><summary>Кеңес пен дұрыс жауап</summary><p>{E(c.get('hint'))}</p><p><b>Жауап:</b> {E(c.get('answer'))}</p></details></div><footer>ОЙЛАН · ОРЫНДА · ТЕКСЕР</footer></article>''')
    cards_file=uniq('.html')
    Path(cards_file).write_text(f'''<!doctype html><html lang="kk"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>6 Premium карточка</title><style>{CARD_STYLE}</style><div class="toolbar"><h1>6 Premium карточка · {E(meta["topic"])}</h1><button onclick="print()">Басып шығару</button></div><div class="grid">{"".join(cards)}</div></html>''',encoding='utf-8')

    concepts=[c for c in d.get('concepts',[]) if isinstance(c,dict) and c.get('term')]
    quiz=[q for q in d.get('quiz',[]) if isinstance(q,dict)]
    tasks=[c for c in d.get('cards',[]) if isinstance(c,dict)]
    steps=d.get('steps',[])
    topic=str(meta.get('topic','')).lower()
    if any(x in topic for x in ('нейрон','перцептрон','машиналық оқыту')):
        sim='neuron'
    elif any(x in topic for x in ('санау жүй','екілік','ақпаратты өлшеу','бит','байт')):
        sim='binary'
    elif any(x in topic for x in ('ip-мекен','ip адрес','маска','желі адрес')):
        sim='network'
    else:
        sim='sequence'
    payload={
        'topic':meta['topic'],'subject':meta['subject'],'grade':meta['grade'],
        'objective':meta['objective'],'explanation':d.get('explanation',''),
        'concepts':concepts,'steps':steps,'diagram':d.get('diagram',{}),
        'cards':tasks,'quiz':quiz,'reflection':d.get('reflection',''),
        'sim':sim
    }
    data_json=json.dumps(payload,ensure_ascii=False).replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026')
    page=r'''<!doctype html><html lang="kk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>BILIM AI INTERACTIVE LAB</title>
<style>
:root{--ink:#17335e;--blue:#335de5;--violet:#7849d9;--light:#f1f5ff;--green:#087b61}*{box-sizing:border-box}body{margin:0;background:#f0f4ff;color:var(--ink);font-family:Segoe UI,Arial,sans-serif;line-height:1.55}header{background:linear-gradient(115deg,#164dbe,#6739cd);color:white;padding:36px max(20px,calc((100vw - 1140px)/2));position:relative;overflow:hidden}header:after{content:'✦';font-size:200px;opacity:.1;position:absolute;right:5%;top:-95px}header h1{font-size:clamp(29px,4vw,47px);line-height:1.15;max-width:860px;margin:12px 0}header p{max-width:900px}.pill{display:inline-block;padding:5px 12px;background:#ffffff30;border-radius:50px;font-weight:700}.wrap{max-width:1140px;margin:auto;padding:20px}.tabs{display:flex;gap:9px;flex-wrap:wrap;position:sticky;top:0;background:#f0f4fff0;backdrop-filter:blur(8px);z-index:3;padding:12px 0}.tabs a{text-decoration:none;background:white;color:#24468c;padding:9px 14px;border-radius:13px;font-weight:700;box-shadow:0 2px 8px #23407012}.panel{background:white;border:1px solid #dce5f5;border-radius:22px;padding:clamp(18px,3vw,28px);margin:20px 0;box-shadow:0 10px 25px #173e7510}h2{font-size:clamp(23px,3vw,30px);margin:0 0 18px;color:#183f91}h3{color:#254bb2;margin:6px 0}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(235px,1fr));gap:14px}.tile{background:linear-gradient(145deg,#f1f6ff,#fff);border:1px solid #d7e3fa;border-radius:17px;padding:18px}.tile b{color:#284ec5}.tile p{margin:9px 0}.visual{background:#f3f7ff;border-radius:18px;padding:15px;overflow:auto}.visual svg{width:100%;min-width:380px;max-height:320px}.btn,button{background:#315fe2;color:white;border:0;border-radius:12px;padding:11px 17px;font-weight:750;cursor:pointer;font-size:15px}button:hover{filter:brightness(.94)}button:focus-visible,input:focus-visible{outline:3px solid #e3a839}.secondary{background:#ede8ff;color:#5037a0}.good{color:#067453;font-weight:700}.bad{color:#b42339;font-weight:700}.score{background:#e6f8ee;color:#07674d;border-radius:12px;padding:12px 16px;font-weight:800}.question{border:1px solid #dce5f5;border-radius:15px;padding:17px;margin:12px 0;background:#f9fbff}.option{display:block;background:white;border:1px solid #dce5f5;padding:10px;margin:7px 0;border-radius:10px;cursor:pointer}.option:has(input:checked){background:#e9f0ff;border-color:#5275e6}input[type=text],input[type=number]{width:100%;max-width:520px;padding:12px;border:1px solid #afc0dd;border-radius:10px;font-size:16px}.matchrow{display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:center;margin:10px 0}.matchrow select{width:100%;padding:10px;border-radius:9px;border:1px solid #adc0dc;background:white;color:#17335e}.sortable{list-style:none;padding:0}.sortable li{padding:12px;background:#edf3ff;border:1px solid #cbd9f3;border-radius:12px;margin:8px 0;display:flex;justify-content:space-between;gap:10px;align-items:center}.sort-controls{display:flex;gap:5px}.sort-controls button{padding:5px 10px}.task-title{display:flex;align-items:center;gap:10px}.num{background:#365ddd;color:white;width:34px;height:34px;border-radius:10px;display:grid;place-items:center;flex-shrink:0}.taskvisual{background:#f3f7ff;padding:12px;border-radius:14px;margin:10px 0}.taskvisual svg{width:100%;max-height:210px}details{background:#f6f3ff;border-radius:12px;padding:12px;margin-top:10px}summary{cursor:pointer;font-weight:750;color:#5535ac}.sim-box{background:#edf6ff;padding:18px;border-radius:18px}.sim-box input[type=range]{width:100%}.sim-value{font-size:28px;font-weight:850;color:#2352bd}.progress{height:12px;background:#dbe5fa;border-radius:20px;overflow:hidden}.progress span{display:block;height:100%;background:linear-gradient(90deg,#2e7be8,#8b4edc);width:0;transition:width .25s}.muted{color:#607292}.feedback{margin-top:10px}.step{display:none;background:#eaf5ff;border-left:5px solid #3975dc;padding:12px;border-radius:9px;margin-top:10px}.step.show{display:block}@media(max-width:650px){.matchrow{grid-template-columns:1fr}.wrap{padding:10px}.panel{margin:12px 0}.tabs{position:static}}
</style></head><body>
<header><span class="pill">✦ BILIM AI · INTERACTIVE LAB</span><h1 id="heroTitle"></h1><p id="heroSub"></p><div class="pill">Зертте · Ойна · Тексер · Жетістікке жет</div></header>
<main class="wrap"><nav class="tabs"><a href="#theory">Түсіндірме</a><a href="#lab">Зертхана</a><a href="#games">Ойындар</a><a href="#missions">Миссиялар</a><a href="#quiz">Тест</a></nav>
<section class="panel" id="theory"><h2>📖 Тақырыпты зерттейік</h2><div class="visual" id="mainDiagram"></div><p id="explanation"></p><h3>Негізгі ұғымдар</h3><div class="grid" id="concepts"></div><h3>Қадамдық мысал</h3><div id="stepArea"></div><button class="secondary" id="nextStep">Келесі қадам →</button></section>
<section class="panel" id="lab"><h2>🔬 Интерактивті зертхана</h2><p class="muted">Параметрлерді өзгертіп, нәтижесін бақылаңыз.</p><div id="simulator" class="sim-box"></div></section>
<section class="panel" id="games"><h2>🎮 Ойын алаңы</h2><div class="grid"><div class="tile"><h3>1. Сәйкестендіру</h3><p>Ұғымға дұрыс анықтаманы таңдаңыз.</p></div><div class="tile"><h3>2. Бос орынды толтыр</h3><p>Тақырыптық терминді жазыңыз.</p></div><div class="tile"><h3>3. Ретімен орналастыр</h3><p>Қадамдарды жоғары-төмен жылжытыңыз.</p></div></div><div class="question" id="matchGame"></div><div class="question" id="fillGame"></div><div class="question" id="sortGame"></div></section>
<section class="panel" id="missions"><h2>🚀 6 деңгейлік миссия</h2><p class="muted">Миссияны орындап, жауабыңызды тексеріңіз. Дұрыс жауапты бірден ашпай, алдымен ойланыңыз.</p><div id="missionsArea"></div></section>
<section class="panel" id="quiz"><h2>🏆 10 сұрақтық финал</h2><div class="progress"><span id="progress"></span></div><p class="muted" id="quizProgress">0 / 10</p><div id="quizArea"></div><button id="checkQuiz">Тестті тексеру</button><div id="quizResult" class="feedback" aria-live="polite"></div></section>
<section class="panel"><h2>🌟 Рефлексия</h2><p id="reflection"></p><p class="muted">Жетістіктеріңізді талқылап, келесі сабаққа мақсат қойыңыз.</p></section></main>
<script id="lesson-data" type="application/json">__DATA__</script>
<script>
'use strict';
const D=JSON.parse(document.getElementById('lesson-data').textContent);
const $=id=>document.getElementById(id);
function node(tag,txt,cls){const el=document.createElement(tag);if(txt!==undefined)el.textContent=String(txt??'');if(cls)el.className=cls;return el}
function add(parent,child){parent.appendChild(child);return child}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function diagram(labels,caption){labels=(Array.isArray(labels)?labels:[]).slice(0,4);if(!labels.length)labels=['Бастау','Талдау','Нәтиже'];const w=850,n=labels.length,bw=(w-40-(n-1)*22)/n;let s='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 850 225" role="img" aria-label="Тақырыптық сызба"><rect width="850" height="225" rx="20" fill="#edf4ff"/>';labels.forEach((lab,i)=>{const x=20+i*(bw+22);s+='<rect x="'+x+'" y="35" width="'+bw+'" height="130" rx="17" fill="'+['#dbeafe','#ede9fe','#d1fae5','#fef3c7'][i%4]+'" stroke="#9db4dc"/><text x="'+(x+bw/2)+'" y="85" text-anchor="middle" font-size="18" fill="#183c79" font-weight="700">'+esc(String(lab).slice(0,17))+'</text><text x="'+(x+bw/2)+'" y="112" text-anchor="middle" font-size="15" fill="#183c79">'+esc(String(lab).slice(17,33))+'</text>';if(i<n-1)s+='<text x="'+(x+bw+10)+'" y="105" text-anchor="middle" font-size="24" fill="#446ac2">→</text>'});return s+'<text x="425" y="205" text-anchor="middle" font-size="15" fill="#456080">'+esc(String(caption??'').slice(0,90))+'</text></svg>'}
$('heroTitle').textContent=D.topic;$('heroSub').textContent=D.subject+' · '+D.grade+'-сынып · '+D.objective;$('explanation').textContent=D.explanation;$('mainDiagram').innerHTML=diagram(D.diagram?.labels,D.diagram?.caption);$('reflection').textContent=D.reflection;
(D.concepts||[]).forEach((c,i)=>{const t=add($('concepts'),node('div',undefined,'tile'));add(t,node('h3',(i+1)+'. '+c.term));add(t,node('p',c.meaning));add(t,node('small',c.example,'muted'))});
let step=0;const steps=D.steps||[];steps.forEach((s,i)=>{const t=add($('stepArea'),node('div',(i+1)+'. '+s,'step'));if(i===0)t.classList.add('show')});$('nextStep').onclick=()=>{step=Math.min(steps.length-1,step+1);[...$('stepArea').children].forEach((x,i)=>x.classList.toggle('show',i<=step));if(step===steps.length-1)$('nextStep').textContent='Барлық қадам көрсетілді'};
function simulator(){const box=$('simulator');let html='';if(D.sim==='neuron')html='<h3>Нейрон салмақтарын өзгерт</h3><label>Бірінші салмақ <output id="aVal">0.5</output></label><input id="a" type="range" min="-2" max="2" step="0.1" value="0.5"><label>Екінші салмақ <output id="bVal">0.5</output></label><input id="b" type="range" min="-2" max="2" step="0.1" value="0.5"><p>Сигмоид нәтижесі:</p><div class="sim-value" id="simOut"></div>';
else if(D.sim==='binary')html='<h3>Санның екілік көрінісін зертте</h3><label>Ондық сан: <output id="aVal">25</output></label><input id="a" type="range" min="0" max="255" step="1" value="25"><p>Екілік жүйе:</p><div class="sim-value" id="simOut"></div><p id="simExtra"></p>';
else if(D.sim==='network')html='<h3>IP және ішкі желі маскасы</h3><label>Соңғы октет: <output id="aVal">130</output></label><input id="a" type="range" min="0" max="255" step="1" value="130"><label>Префикс</label><select id="b"><option value="24">/24</option><option value="25">/25</option><option value="26">/26</option><option value="27">/27</option><option value="28">/28</option></select><p>Мысал IP: 192.168.1.<span id="ipLast">130</span></p><div class="sim-value" id="simOut"></div>';
else html='<h3>Тақырыптық қадамдар зертханасы</h3><p>Қадамдарды ауыстырып, алгоритм ретін зерттеңіз.</p><label>Қадам нөмірі: <output id="aVal">1</output></label><input id="a" type="range" min="1" max="'+Math.max(1,steps.length)+'" step="1" value="1"><div class="sim-value" style="font-size:20px" id="simOut"></div>';
box.innerHTML=html;const calc=()=>{const a=Number($('a').value);$('aVal').textContent=a;if(D.sim==='neuron'){const b=Number($('b').value);$('bVal').textContent=b;$('simOut').textContent=(1/(1+Math.exp(-(a+b)))).toFixed(3)}else if(D.sim==='binary'){$('simOut').textContent=a.toString(2).padStart(8,'0');$('simExtra').textContent='8 биттік жазылуы · '+a+'₁₀'}else if(D.sim==='network'){const pre=Number($('b').value),size=2**(32-pre),net=Math.floor(a/size)*size;$('ipLast').textContent=a;$('simOut').textContent='Желі: 192.168.1.'+net+'/'+pre}else $('simOut').textContent=steps[a-1]||'Қадам жоқ'};$('a').oninput=calc;if($('b'))$('b').oninput=calc;calc()}simulator();
const defs=(D.concepts||[]).slice(0,5);const mg=$('matchGame');add(mg,node('h3','1. Сәйкестендіру'));const picks=[];defs.forEach((c,i)=>{const r=add(mg,node('div',undefined,'matchrow'));add(r,node('b',c.term));const sel=add(r,node('select'));add(sel,node('option','Анықтаманы таңдаңыз')).value='';defs.forEach((v,j)=>{const o=add(sel,node('option',v.meaning));o.value=j});picks.push(sel)});const mb=add(mg,node('button','Тексеру'));const mf=add(mg,node('div',undefined,'feedback'));mb.onclick=()=>{let n=0;picks.forEach((p,i)=>{if(p.value===String(i))n++});mf.textContent=n+' / '+defs.length+' дұрыс сәйкестік';mf.className='feedback '+(n===defs.length?'good':'bad')};
const fg=$('fillGame');add(fg,node('h3','2. Бос орынды толтыр'));const fc=defs[0];if(fc){add(fg,node('p','Анықтама: '+fc.meaning));const inp=add(fg,node('input'));inp.type='text';inp.placeholder='Терминді жазыңыз';const fb=add(fg,node('button','Жауапты тексеру'));const ff=add(fg,node('div',undefined,'feedback'));fb.onclick=()=>{const clean=x=>String(x).toLocaleLowerCase('kk').trim().replace(/\s+/g,' ');const ok=clean(inp.value)===clean(fc.term);ff.textContent=ok?'Дұрыс! +1 ұпай':'Тағы ойланып көріңіз. Кеңес: '+String(fc.term).slice(0,2)+'…';ff.className='feedback '+(ok?'good':'bad')}}
const sg=$('sortGame');add(sg,node('h3','3. Қадамдарды ретімен орналастыр'));const original=steps.slice(0,5);let order=original.map((_,i)=>i).reverse();const ul=add(sg,node('ul',undefined,'sortable'));function renderSort(){ul.replaceChildren();order.forEach((ix,pos)=>{const li=add(ul,node('li'));add(li,node('span',original[ix]));const cs=add(li,node('div',undefined,'sort-controls'));for(const [symbol,dir] of [['↑',-1],['↓',1]]){const btn=add(cs,node('button',symbol,'secondary'));btn.disabled=pos+dir<0||pos+dir>=order.length;btn.onclick=()=>{[order[pos],order[pos+dir]]=[order[pos+dir],order[pos]];renderSort()}}})}renderSort();const sb=add(sg,node('button','Ретін тексеру'));const sf=add(sg,node('div',undefined,'feedback'));sb.onclick=()=>{const n=order.filter((v,i)=>v===i).length;sf.textContent=n+' / '+order.length+' қадам дұрыс орында';sf.className='feedback '+(n===order.length?'good':'bad')};
(D.cards||[]).forEach((c,i)=>{const box=add($('missionsArea'),node('div',undefined,'question'));const title=add(box,node('div',undefined,'task-title'));add(title,node('span',i+1,'num'));add(title,node('h3',c.title+' · '+c.level));add(box,node('p',c.context));const v=add(box,node('div',undefined,'taskvisual'));v.innerHTML=diagram(c.visual_labels,'Тапсырманың тақырыптық сызбасы');add(box,node('p',c.question));const ans=add(box,node('input'));ans.type='text';ans.placeholder='Өз жауабыңызды енгізіңіз';const btn=add(box,node('button','Жауапты салыстыру'));const feed=add(box,node('div',undefined,'feedback'));btn.onclick=()=>{const a=ans.value.trim().toLocaleLowerCase('kk'),b=String(c.answer||'').trim().toLocaleLowerCase('kk');if(!a){feed.textContent='Алдымен жауап енгізіңіз';feed.className='feedback bad'}else{feed.textContent=a===b?'Дұрыс жауап!':'Жауабыңызды үлгімен салыстырыңыз. Ашық сұрақтарды мұғалім бағалайды.';feed.className='feedback '+(a===b?'good':'muted')}};const det=add(box,node('details'));add(det,node('summary','Кеңес және үлгі жауап'));add(det,node('p','Кеңес: '+c.hint));add(det,node('p','Үлгі жауап: '+c.answer))});
(D.quiz||[]).forEach((q,i)=>{const box=add($('quizArea'),node('div',undefined,'question'));add(box,node('h3',(i+1)+'. '+q.q));(q.options||[]).forEach((opt,j)=>{const lab=add(box,node('label',undefined,'option'));const inp=add(lab,node('input'));inp.type='radio';inp.name='q'+i;inp.value=j;lab.append(' '+opt)});const feed=add(box,node('div',undefined,'feedback'));feed.id='qfeed'+i});const update=()=>{const n=(D.quiz||[]).filter((q,i)=>document.querySelector('input[name="q'+i+'"]:checked')).length;$('quizProgress').textContent=n+' / '+D.quiz.length;$('progress').style.width=(100*n/Math.max(1,D.quiz.length))+'%'};$('quizArea').addEventListener('change',update);$('checkQuiz').onclick=()=>{let n=0;(D.quiz||[]).forEach((q,i)=>{const x=document.querySelector('input[name="q'+i+'"]:checked'),ok=x&&Number(x.value)===q.answer;if(ok)n++;const f=$('qfeed'+i);f.textContent=ok?'✓ Дұрыс':'Түсіндірме: '+q.explanation;f.className='feedback '+(ok?'good':'bad')});$('quizResult').textContent='Нәтиже: '+n+' / '+D.quiz.length+' · '+Math.round(100*n/Math.max(1,D.quiz.length))+'%';$('quizResult').className='score'};
</script></body></html>'''
    page=page.replace('__DATA__',data_json)
    lesson_file=uniq('.html')
    Path(lesson_file).write_text(page,encoding='utf-8')
    return lesson_file,cards_file

# AI SMART DESIGN: тақырыпқа сәйкес стиль, инфографика және әртүрлі композиция.
# Бұл Canva/Gamma API емес: презентация өңделетін PowerPoint ретінде жергілікті құрылады.
DESIGN_THEMES = {
    'ai': {'name':'AI / Футуристік', 'dark':'101C40','main':'635BDF','accent':'11C5C6','pale':'E9E9FF','soft':'E8FBFA'},
    'coding': {'name':'Программалау / Ойын', 'dark':'26345A','main':'F07827','accent':'8B5CF6','pale':'FFF1DF','soft':'F0E9FF'},
    'network': {'name':'Желі / Технология', 'dark':'102F49','main':'0A88AE','accent':'20B99B','pale':'E3F7FD','soft':'E2FAF1'},
    'science': {'name':'Жаратылыстану / Табиғат', 'dark':'184538','main':'279E72','accent':'E6AC3C','pale':'E8F7EB','soft':'FFF4D8'},
    'math': {'name':'Математика / Дәлдік', 'dark':'182C55','main':'3463D5','accent':'ECA94B','pale':'E8EFFF','soft':'FFF4DF'},
    'creative': {'name':'Шығармашылық / Әдебиет', 'dark':'51294D','main':'B6538C','accent':'E6A95E','pale':'FBEAF4','soft':'FFF4E7'},
    'general': {'name':'Заманауи / Білім', 'dark':'1B315A','main':'3868DB','accent':'8E4DDA','pale':'E9F1FF','soft':'F1E9FF'}
}

def choose_design(meta, selected='auto'):
    if selected in DESIGN_THEMES and selected != 'auto':
        return DESIGN_THEMES[selected]
    topic = (str(meta.get('topic',''))+' '+str(meta.get('subject',''))).lower()
    # Neural network before generic network: prevent IP graphics for neural topics.
    if any(x in topic for x in ('нейрон','жасанды интеллект','machine learning','машиналық оқыту','перцептрон','генератив','deep learning')):
        key='ai'
    elif any(x in topic for x in ('scratch','python','код жазу','алгоритм','программалау','kodu','for циклі','while циклі','функция құру','тармақталу')):
        key='coding'
    elif any(x in topic for x in ('ip-мекен','dns','домен','компьютерлік желі','желілік','tcp','хаттама','интернет','желінің')):
        key='network'
    elif any(x in topic for x in ('өсімдік','биология','химия','экология','жаратылыстану','география','табиғат','физика')):
        key='science'
    elif any(x in topic for x in ('математика','алгебра','геометрия','санау жүйесі','ақпаратты өлшеу','формула','теңдеу')):
        key='math'
    elif any(x in topic for x in ('әдебиет','тарих','өнер','тіл','шығарма','музыка')):
        key='creative'
    else:
        key='general'
    return DESIGN_THEMES[key]

def make_pptx(meta,d,design_style='auto',include_visuals=True):
    """16 editable slides, with topic-aware theme and varied vector-based layouts."""
    from pptx import Presentation
    from pptx.util import Inches, Pt as PPTPt
    from pptx.dml.color import RGBColor as PPTColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import MSO_ANCHOR
    from pptx.enum.dml import MSO_THEME_COLOR
    import textwrap
    prs=Presentation()
    prs.slide_width=Inches(13.333)
    prs.slide_height=Inches(7.5)
    t=choose_design(meta,design_style)
    def col(h):return PPTColor.from_string(h)
    def shape(sl,x,y,w,h,fill,kind=MSO_SHAPE.ROUNDED_RECTANGLE,radius=None):
        sh=sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid();sh.fill.fore_color.rgb=col(fill)
        sh.line.fill.background()
        return sh
    def txt(sl,content,x,y,w,h,size=20,color=None,bold=False,align=None):
        sh=sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf=sh.text_frame;tf.clear();tf.word_wrap=True
        tf.margin_left=Inches(.06);tf.margin_right=Inches(.06)
        tf.margin_top=Inches(.02);tf.margin_bottom=Inches(.02)
        tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=tf.paragraphs[0]
        p.text=str(content or '')
        p.font.name='Arial';p.font.size=PPTPt(size)
        p.font.bold=bold;p.font.color.rgb=col(color or t['dark'])
        if align is not None:p.alignment=align
        return sh
    def shorten(v,limit=260):
        value=str(v or '').strip()
        return value if len(value)<=limit else value[:limit].rsplit(' ',1)[0]+'…'
    def base(title,num,variant=0):
        sl=prs.slides.add_slide(prs.slide_layouts[6])
        sl.background.fill.solid();sl.background.fill.fore_color.rgb=col('F8FAFF')
        shape(sl,0,0,.17,7.5,t['main'],MSO_SHAPE.RECTANGLE)
        shape(sl,.55,.38,.12,.65,t['accent'],MSO_SHAPE.RECTANGLE)
        txt(sl,shorten(title,90),.82,.38,11.8,.74,27,t['dark'],True)
        shape(sl,.8,7.05,11.7,.015,'E0E8F2',MSO_SHAPE.RECTANGLE)
        txt(sl,'BILIM AI  •  SMART DESIGN  /  '+t['name'],.85,7.11,10,.23,9,'7084A0')
        txt(sl,f'{num:02d} / 16',11.6,7.1,.85,.23,9,t['main'],True)
        if variant==1:
            shape(sl,12.65,.15,.45,.45,t['pale'],MSO_SHAPE.OVAL)
        return sl
    def panel(sl,content,x,y,w,h,bg=None,size=18,title=None):
        shape(sl,x,y,w,h,bg or t['pale'])
        if title:
            txt(sl,title,x+.18,y+.14,w-.36,.42,15,t['main'],True)
            txt(sl,shorten(content,max(110,int(w*h*25))),x+.18,y+.67,w-.36,h-.82,size,t['dark'])
        else:
            txt(sl,shorten(content,max(90,int(w*h*26))),x+.18,y+.15,w-.36,h-.3,size,t['dark'])
    def infographic(sl,labels,y=2.2):
        labels=[str(x) for x in labels if str(x).strip()][:4]
        if not labels:labels=['Тақырып','Талдау','Қорытынды']
        count=len(labels); gap=.24; w=(11.7-gap*(count-1))/count
        for i,lab in enumerate(labels):
            x=.82+i*(w+gap)
            shape(sl,x,y,w,2.05,[t['pale'],t['soft'],'F0F4FC','FCEFE9'][i%4])
            shape(sl,x+.16,y+.16,.53,.53,[t['main'],t['accent'],t['main'],t['accent']][i%4],MSO_SHAPE.OVAL)
            txt(sl,str(i+1),x+.16,y+.16,.53,.53,17,'FFFFFF',True)
            txt(sl,shorten(lab,62),x+.16,y+.85,w-.32,.92,18,t['dark'],True)
            if i<count-1:
                txt(sl,'→',x+w-.02,y+.74,.25,.45,19,t['main'],True)
    def callout(sl,text,y=5.35):
        shape(sl,.82,y,11.65,.9,'FFFFFF')
        shape(sl,.82,y,.075,.9,t['accent'],MSO_SHAPE.RECTANGLE)
        txt(sl,shorten(text,150),1.06,y+.1,11.1,.7,16,t['dark'])
    # 01 Cover: dark editorial style
    sl=prs.slides.add_slide(prs.slide_layouts[6]);sl.background.fill.solid()
    sl.background.fill.fore_color.rgb=col(t['dark'])
    shape(sl,9.55,.15,3.65,3.65,t['main'],MSO_SHAPE.OVAL)
    shape(sl,10.6,5.0,2.6,2.35,t['accent'],MSO_SHAPE.OVAL)
    shape(sl,.82,1.0,2.7,.08,t['accent'],MSO_SHAPE.RECTANGLE)
    txt(sl,'BILIM AI  /  SMART LESSON PRO',.82,.48,10,.38,17,'FFFFFF',True)
    txt(sl,shorten(meta['topic'],110),.82,1.5,10.3,2.65,39,'FFFFFF',True)
    txt(sl,f"{meta['subject']}  ·  {meta['grade']}-сынып",.82,5.35,10,.6,22,'DCEBFF')
    txt(sl,t['name'],.82,6.35,8,.35,14,'FFFFFF')
    # 02 objective: two-column contrast
    sl=base('Сабақ мақсаты',2)
    panel(sl,meta['objective'],.82,1.48,5.6,4.65,t['pale'],19,'Оқу мақсаты')
    panel(sl,d.get('lesson_goal',''),6.68,1.48,5.82,4.65,t['soft'],19,'Сабақ нәтижесі')
    # 03 opening / brainstorm
    sl=base('Ой қозғау · сабаққа кіріспе',3,1)
    panel(sl,d.get('intro',{}).get('teacher',''),.82,1.5,7.35,4.85,t['pale'],19,'Мұғалімнің сұрағы')
    panel(sl,d.get('intro',{}).get('student',''),8.43,1.5,4.07,4.85,t['soft'],17,'Оқушы әрекеті')
    # 04 explanation: limited summary not wall of text
    sl=base('Жаңа тақырыпты меңгеру',4)
    explanation=d.get('explanation','')
    pieces=[x.strip() for x in str(explanation).replace(chr(10),' ').split('.') if x.strip()]
    summary=['. '.join(pieces[i::3][:2])+'.' for i in range(3)]
    for i,part in enumerate(summary):panel(sl,part,.83,1.52+i*1.63,11.65,1.4,[t['pale'],t['soft'],'FFFFFF'][i],17,f'НЕГІЗГІ ОЙ {i+1}')
    # 05 topic diagram
    sl=base('Тақырыптық инфографика',5)
    if include_visuals:infographic(sl,d.get('diagram',{}).get('labels',[]))
    callout(sl,d.get('diagram',{}).get('caption',''),5.32)
    # 06 concepts: four quadrants
    sl=base('Негізгі ұғымдар',6)
    for i,c in enumerate(d.get('concepts',[])[:4]):
        panel(sl,c.get('meaning',''),.83+(i%2)*5.95,1.48+(i//2)*2.58,5.7,2.3,[t['pale'],t['soft']][i%2],16,c.get('term','Ұғым'))
    # 07 sequence
    sl=base('Қадамдық мысал · қалай орындалады?',7)
    steps=d.get('steps',[])[:5]
    for i,item in enumerate(steps):
        y=1.42+i*1.04
        shape(sl,.85,y,.6,.62,t['main'] if i%2==0 else t['accent'],MSO_SHAPE.OVAL)
        txt(sl,str(i+1),.85,y,.6,.62,18,'FFFFFF',True)
        panel(sl,item,1.64,y,10.8,.78,'FFFFFF' if i%2 else t['pale'],15)
    # 08..13 six activity slides, three layout variations
    for i,c in enumerate(d.get('cards',[])[:6]):
        sl=base(f"{i+1}-тапсырма · {c.get('title','')}",8+i,i%2)
        if i%3==0:
            panel(sl,c.get('context',''),.82,1.5,4.15,4.65,t['soft'],17,'ЖАҒДАЯТ')
            panel(sl,c.get('question',''),5.2,1.5,7.3,4.65,t['pale'],20,'ТАПСЫРМА')
        elif i%3==1:
            panel(sl,c.get('question',''),.82,1.5,11.65,2.3,t['pale'],21,'НЕ ІСТЕУ КЕРЕК?')
            if include_visuals:infographic(sl,c.get('visual_labels',[])[:3],4.05)
            else:panel(sl,c.get('context',''),.82,4.08,11.65,2.1,t['soft'],17)
        else:
            if include_visuals:infographic(sl,c.get('visual_labels',[])[:3],1.45)
            panel(sl,c.get('question',''),.82,4.0,11.65,2.05,t['pale'],19,'ТАЛДАУ / ШЕШІМ')
        txt(sl,str(c.get('level',''))+'  •  '+str(c.get('kind','')),.93,6.52,10,.34,14,t['main'],True)
    # 14 game
    sl=base('Ойын және цифрлық ресурстар',14)
    methods=d.get('methods',{})
    panel(sl,methods.get('game','Quizlet / Kahoot / Wordwall'),.82,1.55,11.65,1.5,t['pale'],22,'ОЙЫН ПЛАТФОРМАСЫ')
    panel(sl,'Мұғалім тақырыпқа сай сұрақтарды таңдалған платформаға енгізеді. Дайын ойын сілтемесі автоматты жасалмайды.',.82,3.45,11.65,2.1,t['soft'],18,'ҚОЛДАНУ')
    # 15 assessment
    sl=base('Сыни ойлау және бағалау',15)
    panel(sl,methods.get('critical','Тақырыптық талдау'),.82,1.5,5.6,4.5,t['pale'],20,'СЫНИ ОЙЛАУ ӘДІСІ')
    panel(sl,d.get('middle',{}).get('assessment',''),6.67,1.5,5.83,4.5,t['soft'],17,'ДЕСКРИПТОРЛАР')
    # 16 reflection
    sl=base('Қорытынды · рефлексия',16)
    panel(sl,d.get('reflection',''),.82,1.5,7.1,4.75,t['pale'],19,'БҮГІНГІ САБАҚ')
    panel(sl,d.get('homework',''),8.16,1.5,4.34,4.75,t['soft'],18,'ҮЙ ТАПСЫРМАСЫ')
    assert len(prs.slides)==16, f'Слайд саны: {len(prs.slides)}'
    out=uniq('.pptx');prs.save(out);return out


# Canva Education: API-free export. No Canva credentials are collected.
def canva_design_pack(meta, data, pptx_path, design_style='auto'):
    import zipfile
    from xml.sax.saxutils import escape as xml_escape
    theme=choose_design(meta,design_style)
    main='#'+theme['main']; dark='#'+theme['dark']; accent='#'+theme['accent']
    def lines(text,limit=32,maxlines=5):
        import textwrap
        value=' '.join(str(text or '').split())
        arr=textwrap.wrap(value,width=limit,break_long_words=False,break_on_hyphens=False)
        if len(arr)>maxlines:arr=arr[:maxlines];arr[-1]=arr[-1][:max(1,limit-2)]+'…'
        return arr
    def svg_text(text,x,y,width=34,size=43,fill='#14284e',maxlines=5,step=None):
        step=step or size*1.3
        return ''.join(f'<text x="{x}" y="{y+i*step:.0f}" font-family="DejaVu Sans,Arial,sans-serif" font-size="{size}" font-weight="700" fill="{fill}">{xml_escape(t)}</text>' for i,t in enumerate(lines(text,width,maxlines)))
    def save_svg_png(svg,name):
        target=Path(ROOT)/('canva_'+uuid.uuid4().hex[:8]+'_'+name)
        cairosvg.svg2png(bytestring=svg.encode('utf-8'),write_to=str(target))
        return target
    outputs=[]
    labels=data.get('diagram',{}).get('labels',[])[:4]
    labels=[str(x) for x in labels if str(x).strip()] or ['Түсіну','Талдау','Қолдану']
    # Landscape infographic, 1600 x 900.
    cells=[]
    cell_w=(1460-20*(len(labels)-1))/len(labels)
    for i,label in enumerate(labels):
        x=70+i*(cell_w+20)
        cells.append(f'<rect x="{x:.0f}" y="345" width="{cell_w:.0f}" height="310" rx="26" fill="white" stroke="#DCE7F4" stroke-width="3"/>')
        cells.append(f'<circle cx="{x+60:.0f}" cy="410" r="33" fill="{main if i%2==0 else accent}"/>')
        cells.append(svg_text(str(i+1),x+47,426,2,36,'#FFFFFF',1))
        cells.append(svg_text(label,x+30,502,max(10,int(cell_w/22)),32,dark,4,47))
    infographic=f'''<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900"><rect width="1600" height="900" fill="#F2F7FF"/><rect width="1600" height="220" fill="{dark}"/><circle cx="1450" cy="70" r="170" fill="{main}" opacity=".45"/>{svg_text(meta['topic'],70,105,37,49,'#FFFFFF',2,64)}<text x="70" y="285" font-size="29" font-family="DejaVu Sans" fill="{main}">ТАҚЫРЫПТЫҚ ИНФОГРАФИКА</text>{''.join(cells)}{svg_text(data.get('diagram',{}).get('caption',''),75,755,70,25,dark,3,35)}</svg>'''
    outputs.append(('infographic_16x9.png',save_svg_png(infographic,'infographic.png')))
    poster=f'''<svg xmlns="http://www.w3.org/2000/svg" width="1080" height="1528"><rect width="1080" height="1528" fill="#F5F8FF"/><rect width="1080" height="480" fill="{dark}"/><circle cx="960" cy="80" r="210" fill="{accent}" opacity=".4"/><text x="74" y="100" font-family="DejaVu Sans" font-size="33" fill="#FFFFFF">BILIM AI • SMART LESSON</text>{svg_text(meta['topic'],74,205,24,55,'#FFFFFF',4,72)}<rect x="65" y="530" width="950" height="850" rx="36" fill="#FFFFFF" stroke="#D9E5F7" stroke-width="3"/><text x="112" y="605" font-size="35" font-family="DejaVu Sans" font-weight="700" fill="{main}">НЕГІЗГІ ҰҒЫМДАР</text>{''.join(f'<rect x="112" y="{665+i*147}" width="850" height="118" rx="20" fill="{theme["pale"]}"/>'+svg_text(label,146,731+i*147,29,33,dark,2,42) for i,label in enumerate(labels[:4]))}<text x="72" y="1470" font-size="25" font-family="DejaVu Sans" fill="{dark}">{xml_escape(str(meta['subject']))} • {xml_escape(str(meta['grade']))}-сынып</text></svg>'''
    outputs.append(('poster_A4_portrait.png',save_svg_png(poster,'poster.png')))
    for i,card in enumerate(data.get('cards',[])[:6],1):
        card_labels=card.get('visual_labels',[])[:3]
        visual=''.join(f'<rect x="76" y="{570+j*124}" width="568" height="96" rx="18" fill="{theme["pale"]}"/>'+svg_text(label,105,628+j*124,27,26,dark,2,33) for j,label in enumerate(card_labels))
        card_svg=f'''<svg xmlns="http://www.w3.org/2000/svg" width="720" height="1280"><rect width="720" height="1280" fill="{dark}"/><rect x="20" y="20" width="680" height="1240" rx="32" fill="#FFFFFF"/><rect x="20" y="20" width="680" height="210" rx="32" fill="{main}"/><text x="68" y="75" font-size="23" fill="#FFFFFF" font-family="DejaVu Sans">BILIM AI · ТАПСЫРМА {i}/6</text>{svg_text(card.get('title',''),68,135,25,34,'#FFFFFF',2,43)}<text x="70" y="280" font-size="25" fill="{main}" font-family="DejaVu Sans">{xml_escape(str(card.get('level','')))}</text>{svg_text(card.get('context',''),70,340,34,25,dark,5,37)}{visual}<rect x="68" y="1005" width="584" height="175" rx="20" fill="#EDF3FF"/><text x="92" y="1047" font-size="25" font-family="DejaVu Sans" fill="{main}">ТАПСЫРМА</text>{svg_text(card.get('question',''),92,1093,37,23,dark,3,31)}</svg>'''
        outputs.append((f'card_{i:02d}_9x16.png',save_svg_png(card_svg,f'card_{i}.png')))
    zip_path=uniq('.zip')
    guide=('CANVA EDUCATION IMPORT PACK\n\n'
           '1. Canva -> Create a design -> Import file / Uploads.\n'
           '2. presentation.pptx: import as a presentation and edit text and shapes.\n'
           '3. infographic_16x9.png: upload as a background/visual asset.\n'
           '4. poster_A4_portrait.png: upload into an A4 design.\n'
           '5. card_01..06_9x16.png: upload as six image assets.\n'
           'PNG images are flat images, not individually editable Canva elements.\n'
           'This is an API-free export; it does not create Canva designs automatically.\n'
           'OAuth and Autofill API are NOT enabled in this version.\n')
    with zipfile.ZipFile(zip_path,'w',compression=zipfile.ZIP_DEFLATED) as z:
        z.write(pptx_path,'presentation.pptx')
        for name,path in outputs:z.write(path,name)
        z.writestr('CANVA_IMPORT_README.txt',guide)
    return zip_path

def run(school,teacher,day,unit,subject,grade,topic,objective,duration,values,design_style,include_visuals):
    missing=[name for name,value in [('Мектеп',school),('Мұғалім',teacher),('Пән',subject),('Сынып',grade),('Тақырып',topic),('Оқу мақсаты',objective)] if not S(value).strip()]
    if missing: raise gr.Error('Міндетті өрістер: '+', '.join(missing))
    meta=dict(school=school,teacher=teacher,day=day,unit=unit,subject=subject,grade=grade,topic=topic,objective=objective,duration=int(duration))
    try:
        d=generate(subject,grade,topic,objective,int(duration),values)
        word=make_docx(meta,d);lesson,cards=make_html(meta,d);pptx=make_pptx(meta,d,design_style,include_visuals)
        canva_pack=canva_design_pack(meta,d,pptx,design_style)
        preview=f'''<div class="pro-preview"><h2>🎓 {E(topic)}</h2><p>🎯 <b>Сабақ мақсаты:</b> {E(d['lesson_goal'])}</p><p>💗 <b>Құндылықтар:</b> {E(d['values'])}</p><p>📊 <b>Тапсырмалар:</b> 6 суретті карточка (жеке, жұптық, функционалдық)</p><p>☷ <b>Тест:</b> 10 сұрақ, автоматты тексеру</p></div>'''
        return preview,word,pptx,lesson,cards,canva_pack,'<div class="status-success">✅ &nbsp; Файлдар сәтті дайындалды!</div>'
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


# Restore the original light blue/white interface even when the browser prefers dark mode.
LIGHT_RESTORE_CSS = r"""
:root, html, body, gradio-app, .gradio-container, .gradio-container.dark {
 color-scheme: light !important; --body-background-fill:#f1f6ff !important;
 --background-fill-primary:#ffffff !important; --background-fill-secondary:#f1f6ff !important;
 --block-background-fill:#ffffff !important; --block-label-background-fill:#ffffff !important;
 --body-text-color:#17345e !important; --block-label-text-color:#164f9d !important;
 --input-background-fill:#ffffff !important; --input-text-color:#17345e !important;
 --button-primary-background-fill:#2563eb !important; --button-primary-text-color:#ffffff !important;
 --button-secondary-background-fill:#e9f1ff !important; --button-secondary-text-color:#17418a !important;
 background:#f1f6ff !important; color:#17345e !important;
}
@media (prefers-color-scheme: dark) {
 :root,html,body,gradio-app,.gradio-container {color-scheme:light !important;background:#f1f6ff !important;color:#17345e !important;}
}
gradio-app .gradio-container, gradio-app main, gradio-app .main, gradio-app .contain {background:#f1f6ff !important;color:#17345e !important;}
#form-panel,#result-panel {background:#ffffff !important;color:#17345e !important;border:1px solid #dce7f7 !important;border-radius:20px !important;}
#form-panel input,#form-panel textarea,#form-panel select,#form-panel [role=combobox] {background:#ffffff !important;color:#17345e !important;border:1px solid #ccd9ec !important;}
#form-panel label,#form-panel label span,#result-panel label {color:#17418a !important;}
#bilim-hero {background:linear-gradient(110deg,#1265f3,#4a55ed 55%,#9739ec) !important;padding:25px !important;}
#bilim-hero h1 {color:white !important;font-size:clamp(26px,3vw,38px) !important;}
#result-panel .download-row {background:#e9faf3 !important;border-radius:14px !important;}
#result-panel .download-row.ppt {background:#fff4e8 !important;}
#result-panel .download-row.web {background:#e8f2ff !important;}
#result-panel .download-row.cards {background:#f3eaff !important;}
#result-panel .download-row .download-button button {background:#08a76b !important;color:white !important;}
#result-panel .download-row.ppt .download-button button {background:#f47716 !important;}
#result-panel .download-row.web .download-button button {background:#1767f5 !important;}
#result-panel .download-row.cards .download-button button {background:#8538df !important;}
#result-panel .pro-preview {background:#eaf3ff !important;color:#17345e !important;}
#result-panel .help-panel {background:#fff8e5 !important;color:#17345e !important;}
"""
CSS = CSS + "\n" + LIGHT_RESTORE_CSS

with gr.Blocks(theme=gr.themes.Soft(primary_hue='blue'),css=CSS,title='BILIM AI — SMART LESSON PRO') as demo:
    gr.HTML("""<div id="bilim-hero"><div class="hero-grid"><div><h1>🎓 BILIM AI — SMART LESSON PRO</h1><p>Gemini AI · Суретті Word ҚМЖ · AI Smart Design · HTML сабақ · 6 Premium карточка</p></div><div class="hero-icons"><div class="hero-icon"><b>📄</b>Word ҚМЖ</div><div class="hero-icon"><b>📊</b>PowerPoint</div><div class="hero-icon"><b>🌐</b>HTML сабақ</div><div class="hero-icon"><b>🎨</b>6 карточка</div></div></div></div>""")
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
            gr.Markdown('### 🎨 AI Smart Design — презентация')
            design_style=gr.Dropdown(choices=[('ЖИ тақырыпқа сай автоматты таңдасын','auto'),('AI / Футуристік','ai'),('Программалау / Ойын','coding'),('Желі / Технология','network'),('Жаратылыстану / Табиғат','science'),('Математика / Дәлдік','math'),('Шығармашылық','creative'),('Заманауи','general')],value='auto',label='Презентация дизайны')
            gr.Markdown('**Слайд саны: 16 · Формат: 16:9**')
            include_visuals=gr.Checkbox(label='Тақырыптық инфографика мен векторлық иллюстрациялар',value=True)
            btn=gr.Button('✨ ҚМЖ + презентация + HTML + 6 карточка құру',variant='primary',size='lg')
        with gr.Column(scale=9,min_width=390,elem_id='result-panel'):
            status=gr.HTML('<div class="status-ready">✦ &nbsp; Сабақ мәліметтерін енгізіп, генерация батырмасын басыңыз.</div>')
            preview=gr.HTML('<div class="pro-preview"><h2>🎓 Сабақ нәтижесі</h2><p>Генерациядан кейін сабақтың мақсаты мен материалдары осы жерде көрсетіледі.</p></div>')
            with gr.Row(elem_classes=['download-row','word']):
                gr.HTML('<div class="download-copy">📘 <strong>Word ҚМЖ (.docx)</strong><br><small>Суреттермен безендірілген қысқа мерзімді жоспар</small></div>')
                word=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','ppt']):
                gr.HTML('<div class="download-copy">📙 <strong>PowerPoint презентациясы (.pptx)</strong><br><small>16 слайд, AI Smart Design және инфографика</small></div>')
                ppt=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','web']):
                gr.HTML('<div class="download-copy">🌐 <strong>Интерактивті HTML сабағы</strong><br><small>Түсіндірме, тапсырмалар және 10 сұрақтық тест</small></div>')
                lesson=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','cards']):
                gr.HTML('<div class="download-copy">🎨 <strong>6 Premium карточка (HTML)</strong><br><small>9:16 форматындағы суретті тапсырмалар</small></div>')
                cards=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            with gr.Row(elem_classes=['download-row','web']):
                gr.HTML('<div class="download-copy">🎨 <strong>Canva Education импорт пакеті (.zip)</strong><br><small>PPTX + инфографика PNG + постер PNG + 6 суретті карточка PNG</small></div>')
                canva_download=gr.DownloadButton('⬇ Жүктеу',elem_classes=['download-button'],variant='primary')
            gr.HTML('<div class="help-panel"><b>🎨 Canva Education OAuth:</b> <a href="/canva/connect" target="_blank">🔗 Canva-ға қосылу</a> · <a href="/canva/status" target="_blank">Байланысты тексеру</a> · <a href="/canva/disconnect">Ажырату</a><br>Canva ZIP импорты сақталған. Автоматты дизайн/Autofill келесі кезеңде іске қосылады.</div>')
            gr.HTML('<div class="help-panel"><b>💡 Қолдану нұсқаулығы:</b><br>1. Word ҚМЖ — суретті сабақ жоспары және дескрипторлар.<br>2. PowerPoint — сабақ презентациясы.<br>3. HTML сабақ — түсіндірме, тапсырмалар және автоматты тест.<br>4. Premium карточкалар — 9:16 форматындағы 6 тапсырма.</div>')
    btn.click(fn=run,inputs=[school,teacher,day,unit,subject,grade,topic,objective,duration,values,design_style,include_visuals],outputs=[preview,word,ppt,lesson,cards,canva_download,status],show_progress='full')
demo.queue(default_concurrency_limit=1)
# Mount Gradio under the same FastAPI app so Canva /canva/callback works on Render.
app = gr.mount_gradio_app(canva_app, demo, path='/')

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app,host='0.0.0.0',port=int(os.environ.get('PORT','7860')))
