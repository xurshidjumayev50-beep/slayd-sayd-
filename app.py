import json
import os
import re
import threading
import time
from io import BytesIO

import google.generativeai as genai
from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

app = Flask(__name__)

# API kalitni sozlash
api_key = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=api_key)

# To'g'ri model nomini kiritamiz
MODEL = "gemini-1.5-flash"



THEMES = {
    "ko'k": {"bg": "1F3A5F", "title": "FFFFFF", "text": "E8EEF5"},
    "yashil": {"bg": "1E5B4F", "title": "FFFFFF", "text": "E6F2EF"},
    "qora": {"bg": "1A1A1A", "title": "FFD166", "text": "EEEEEE"},
    "oq": {"bg": "FFFFFF", "title": "1F3A5F", "text": "333333"},
}

LANGUAGES = {"uz": "o'zbek tilida (lotin yozuvi)", "ru": "rus tilida", "en": "ingliz tilida"}


# ---------- Foydalanish limiti (kalit limitini himoya qilish) ----------
RATE_PER_HOUR = int(os.environ.get("RATE_LIMIT_PER_HOUR", "8"))   # bitta foydalanuvchiga soatiga
DAILY_LIMIT = int(os.environ.get("DAILY_LIMIT", "150"))           # butun sayt bo'yicha kuniga
_hits, _lock = {}, threading.Lock()
_day = {"date": None, "n": 0}


def client_ip():
    # Render proxy IP ni X-Forwarded-For oxiriga qo'shadi; undan oldingisi - foydalanuvchi.
    parts = [p.strip() for p in request.headers.get("X-Forwarded-For", "").split(",") if p.strip()]
    if len(parts) >= 2:
        return parts[-2]
    return parts[-1] if parts else (request.remote_addr or "?")


def check_rate_limit(ip):
    now = time.time()
    today = time.strftime("%Y-%m-%d", time.gmtime(now))
    with _lock:
        if _day["date"] != today:
            _day["date"], _day["n"] = today, 0
        if _day["n"] >= DAILY_LIMIT:
            return "Bugungi umumiy limit tugadi. Ertaga qayta urinib ko'ring."
        recent = [t for t in _hits.get(ip, []) if now - t < 3600]
        if len(recent) >= RATE_PER_HOUR:
            return f"Soatiga {RATE_PER_HOUR} tagacha taqdimot yaratish mumkin. Keyinroq urinib ko'ring."
        recent.append(now)
        _hits[ip] = recent
        _day["n"] += 1
        if len(_hits) > 5000:
            for k in [k for k, v in _hits.items() if not v or now - v[-1] >= 3600]:
                del _hits[k]
    return None


# ---------- Slayd foni ----------
SLIDE_W, SLIDE_H = 13.333, 7.5
BACKGROUNDS = ("sodda", "geometrik", "chiziqli")


def mix(c1, c2, t):
    """Ikki rangni aralashtiradi (t=0 -> c1, t=1 -> c2)."""
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def _shape(slide, kind, left, top, w, h, color):
    sp = slide.shapes.add_shape(kind, Inches(left), Inches(top), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = rgb(color)
    sp.line.fill.background()
    sp.shadow.inherit = False
    style = sp._element.find(qn("p:style"))
    if style is not None:
        sp._element.remove(style)  # tema soyasi/chizig'i ta'sir qilmasin
    tree = slide.shapes._spTree
    tree.remove(sp._element)
    tree.insert(2, sp._element)  # matn ostiga (eng orqaga)
    return sp


def add_background(slide, style, theme):
    bg, tx, ti = theme["bg"], theme["text"], theme["title"]
    if style == "geometrik":
        _shape(slide, MSO_SHAPE.OVAL, SLIDE_W - 4.0, SLIDE_H - 3.6, 6.5, 6.5, mix(bg, tx, 0.07))
        _shape(slide, MSO_SHAPE.OVAL, SLIDE_W - 2.6, -1.4, 3.2, 3.2, mix(bg, tx, 0.12))
    elif style == "chiziqli":
        _shape(slide, MSO_SHAPE.RECTANGLE, 0, SLIDE_H - 0.12, SLIDE_W, 0.12, mix(bg, tx, 0.12))
        _shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 0.28, SLIDE_H, ti)


def rgb(hex_color):
    return RGBColor.from_string(hex_color)


def paint_background(slide, hex_color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(hex_color)


# ---------- AI: mavzudan slaydlar matnini yaratish (Gemini) ----------
def generate_content(topic, count, lang_key):
    """Gemini'dan {"title": ..., "slides": [{"title":..., "bullets":[...]}]} oladi."""
    if not os.environ.get("GEMINI_API_KEY"):
        raise RuntimeError(
            "GEMINI_API_KEY topilmadi. Terminalda: export GEMINI_API_KEY='kalitingiz'"
        )
    from google import genai  # pip install google-genai

    language = LANGUAGES.get(lang_key, LANGUAGES["uz"])
    prompt = f"""Quyidagi mavzu bo'yicha taqdimot (slaydlar) tayyorla.

Mavzu: {topic}
Kontent slaydlari soni: {count} ta (sarlavha slaydini hisobga olmagan holda)
Til: {language}

Talablar:
- Har bir slaydda sarlavha va 3-5 ta qisqa, aniq band bo'lsin (har biri bir jumla, 15 so'zdan oshmasin).
- Slaydlar mantiqiy ketma-ketlikda bo'lsin: kirish, asosiy qismlar, xulosa.
- Faqat tekshirib bo'ladigan, aniq ma'lumot yoz; ishonchsiz raqam va faktlar o'ylab topma.

Javobni FAQAT JSON ko'rinishida qaytar, boshqa hech narsa yozma:
{{"title": "taqdimot nomi", "slides": [{{"title": "slayd sarlavhasi", "bullets": ["band 1", "band 2"]}}]}}"""

    model = genai.GenerativeModel(MODEL)
    response = model.generate_content(prompt)
    text = (response.text or "").strip()
    
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("AI javobi JSON emas, qayta urinib ko'ring")
    data = json.loads(text[start : end + 1])
    if not data.get("slides"):
        raise ValueError("AI slaydlarni qaytarmadi, qayta urinib ko'ring")
    return data


# ---------- PowerPoint fayl yaratish ----------
def build_presentation(deck_title, slides, theme_name, bg_style="sodda"):
    theme = THEMES.get(theme_name, THEMES["ko'k"])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Sarlavha slaydi
    s = prs.slides.add_slide(prs.slide_layouts[0])
    paint_background(s, theme["bg"])
    add_background(s, bg_style, theme)
    s.shapes.title.text = deck_title or "Taqdimot"
    for p in s.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(48)
            r.font.bold = True
            r.font.color.rgb = rgb(theme["title"])
    s.placeholders[1].text = ""

    # Kontent slaydlari
    for item in slides:
        title = (item.get("title") or "").strip()
        bullets = [str(b).strip() for b in (item.get("bullets") or []) if str(b).strip()]
        if not title and not bullets:
            continue
        s = prs.slides.add_slide(prs.slide_layouts[1])
        paint_background(s, theme["bg"])
        add_background(s, bg_style, theme)

        s.shapes.title.text = title
        s.shapes.title.left, s.shapes.title.top = Inches(0.8), Inches(0.5)
        s.shapes.title.width, s.shapes.title.height = Inches(11.7), Inches(1.3)
        for p in s.shapes.title.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(38)
                r.font.bold = True
                r.font.color.rgb = rgb(theme["title"])

        body = s.placeholders[1]
        body.left, body.top = Inches(0.8), Inches(2.0)
        body.width, body.height = Inches(11.7), Inches(4.9)
        tf = body.text_frame
        tf.clear()
        tf.word_wrap = True
        for i, line in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.space_after = Pt(12)
            for r in p.runs:
                r.font.size = Pt(26)
                r.font.color.rgb = rgb(theme["text"])

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return PAGE


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify(error="Mavzuni kiriting"), 400
    limit_error = check_rate_limit(client_ip())
    if limit_error:
        return jsonify(error=limit_error), 429
    bg_style = data.get("bg") if data.get("bg") in BACKGROUNDS else "sodda"
    try:
        count = max(3, min(15, int(data.get("count", 7))))
    except (TypeError, ValueError):
        count = 7
    try:
        content = generate_content(topic, count, data.get("lang", "uz"))
        buf = build_presentation(
            content.get("title") or topic, content["slides"], data.get("theme", "ko'k"), bg_style
        )
    except Exception as e:
        return jsonify(error=str(e)), 500
    return send_file(
        buf,
        as_attachment=True,
        download_name="taqdimot.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


PAGE = r"""<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Slayd yaratuvchi</title>
<meta name="description" content="Mavzuni yozing, tayyor PowerPoint taqdimotni yuklab oling.">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect x='2' y='6' width='28' height='20' rx='3' fill='%231F3A5F'/%3E%3Crect x='6' y='11' width='14' height='3' rx='1.5' fill='%23FFD166'/%3E%3Crect x='6' y='17' width='20' height='2' rx='1' fill='%23E8EEF5'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Onest:wght@400..800&display=swap" rel="stylesheet">
<style>
  :root{
    --desk:#E3E8EE; --ink:#142033; --muted:#55627A; --line:rgba(20,32,51,.18); --focus:#2F6FEB;
    --stage-bg:#1F3A5F; --stage-title:#FFFFFF; --stage-text:#E8EEF5;
    --btn-bg:#1F3A5F; --btn-fg:#FFFFFF;
  }
  *{box-sizing:border-box}
  html{-webkit-text-size-adjust:100%}
  body{margin:0;background:var(--desk);color:var(--ink);
    font-family:"Onest",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;line-height:1.5}
  .wrap{max-width:960px;margin:0 auto;padding:16px 16px 56px}
  .vh{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}

  .top{display:flex;align-items:center;gap:10px;padding:10px 0 22px;font-weight:800;font-size:19px;letter-spacing:-.01em}
  .logo{position:relative;width:28px;height:19px;border-radius:3px;background:#1F3A5F}
  .logo::before{content:"";position:absolute;left:5px;top:5px;width:12px;height:3px;border-radius:2px;background:#FFD166}
  .logo::after{content:"";position:absolute;left:5px;top:11px;width:18px;height:2px;border-radius:1px;background:#E8EEF5}

  /* Slayd: sahifaning asosiy elementi */
  .stage{position:relative;overflow:hidden;display:flex;flex-direction:column;aspect-ratio:16/9;padding:clamp(20px,5.5vw,60px);
    background:var(--stage-bg);border-radius:6px;
    box-shadow:0 1px 2px rgba(20,32,51,.14),0 22px 44px -22px rgba(20,32,51,.45);
    transition:background-color .25s ease}
  .stage:focus-within{outline:3px solid var(--ink);outline-offset:5px}
  .stage textarea,.stage .bars{position:relative;z-index:1}
  /* Fon uslublari: fayldagi (.pptx) fon bilan bir xil joylashuv */
  .stage::before,.stage::after{content:"";position:absolute;pointer-events:none;display:none}
  .stage[data-bg="geometrik"]::before,.stage[data-bg="geometrik"]::after{display:block;border-radius:50%}
  .stage[data-bg="geometrik"]::before{width:48.7%;aspect-ratio:1;right:-18.75%;bottom:-38.7%;
    background:color-mix(in srgb,var(--stage-bg) 93%,var(--stage-text))}
  .stage[data-bg="geometrik"]::after{width:24%;aspect-ratio:1;right:-4.5%;top:-18.7%;
    background:color-mix(in srgb,var(--stage-bg) 88%,var(--stage-text))}
  .stage[data-bg="chiziqli"]::before{display:block;left:0;top:0;bottom:0;width:2.1%;background:var(--stage-title)}
  .stage[data-bg="chiziqli"]::after{display:block;left:0;right:0;bottom:0;height:1.6%;
    background:color-mix(in srgb,var(--stage-bg) 88%,var(--stage-text))}
  .stage textarea{display:block;width:100%;margin:0;padding:0;border:0;outline:0;resize:none;overflow:hidden;
    background:none;color:var(--stage-title);caret-color:var(--stage-title);
    font:800 clamp(26px,5.2vw,56px)/1.08 "Onest",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
    letter-spacing:-.025em;text-align:left;transition:color .25s ease}
  .stage textarea::placeholder{color:var(--stage-title);opacity:.45}
  .bars{margin-top:auto;display:grid;gap:clamp(8px,1.6vw,14px)}
  .bars i{display:block;height:clamp(7px,1.3vw,11px);border-radius:6px;background:var(--stage-text);opacity:.22;transition:background-color .25s ease}
  .bars i:nth-child(1){width:64%}.bars i:nth-child(2){width:48%}.bars i:nth-child(3){width:72%}.bars i:nth-child(4){width:40%}
  .busy .bars i{animation:fill 1.3s ease-in-out infinite}
  .busy .bars i:nth-child(2){animation-delay:.2s}.busy .bars i:nth-child(3){animation-delay:.4s}.busy .bars i:nth-child(4){animation-delay:.6s}
  @keyframes fill{0%,100%{opacity:.16}50%{opacity:.62}}

  .eg{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:18px;font-size:14px;color:var(--muted)}
  .eg button{font:inherit;color:var(--ink);background:transparent;border:1px solid var(--line);border-radius:999px;padding:5px 13px;cursor:pointer}
  .eg button:hover{border-color:var(--ink)}

  .ctl{display:flex;flex-wrap:wrap;gap:22px 36px;margin-top:30px}
  .grp{display:grid;gap:9px;align-content:start}
  .grp>span{font-size:14px;font-weight:600}
  .segs{display:flex;border:1px solid var(--line);border-radius:8px;overflow:hidden;background:rgba(255,255,255,.55)}
  .seg{position:relative}
  .seg input,.sw input{position:absolute;inset:0;width:100%;height:100%;margin:0;opacity:0;cursor:pointer}
  .seg span{display:block;padding:9px 15px;font-size:15px;border-right:1px solid var(--line)}
  .seg:last-child span{border-right:0}
  .seg input:checked+span{background:var(--ink);color:#fff}
  .seg input:focus-visible+span{outline:3px solid var(--focus);outline-offset:-3px}

  .sws{display:flex;gap:12px;padding:3px}
  .sw{position:relative}
  .sw span{position:relative;display:block;width:54px;height:36px;border-radius:5px;background:var(--b);
    box-shadow:inset 0 0 0 1px rgba(20,32,51,.2)}
  .sw span::before{content:"";position:absolute;left:8px;top:9px;width:22px;height:5px;border-radius:2px;background:var(--t)}
  .sw span::after{content:"";position:absolute;left:8px;top:19px;width:36px;height:3px;border-radius:2px;background:var(--x);opacity:.7;box-shadow:0 6px 0 var(--x)}
  .sw input:checked+span{box-shadow:inset 0 0 0 1px rgba(20,32,51,.2),0 0 0 3px var(--desk),0 0 0 5px var(--ink)}
  .sw input:focus-visible+span{outline:3px solid var(--focus);outline-offset:6px}

  .act{display:flex;flex-wrap:wrap;align-items:center;gap:14px 20px;margin-top:30px}
  .go{font:inherit;font-weight:700;font-size:17px;padding:14px 30px;border:0;border-radius:10px;cursor:pointer;
    background:var(--btn-bg);color:var(--btn-fg);transition:background-color .25s ease,color .25s ease}
  .go:active{transform:translateY(1px)}
  .go:disabled{opacity:.7;cursor:progress}
  .go:focus-visible{outline:3px solid var(--focus);outline-offset:3px}
  .note{font-size:14px;color:var(--muted);min-height:1.5em}
  .note.err{color:#B42318;font-weight:600}
  .note.ok{color:#1B6B3A;font-weight:600}
  .foot{margin:44px 0 0;font-size:13px;color:var(--muted);max-width:60ch}

  @media (max-width:560px){
    .go{width:100%}
    .ctl{gap:20px}
    .stage{aspect-ratio:auto;min-height:230px}
  }
  @media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style>
</head>
<body>
<div class="wrap">
  <header class="top"><span class="logo" aria-hidden="true"></span><span>Slayd</span></header>

  <main>
    <h1 class="vh">Slayd yaratuvchi</h1>

    <div class="stage" id="stage">
      <label class="vh" for="topic">Taqdimot mavzusi</label>
      <textarea id="topic" rows="1" maxlength="140" placeholder="Taqdimot mavzusini shu yerga yozing"></textarea>
      <div class="bars" aria-hidden="true"><i></i><i></i><i></i><i></i></div>
    </div>

    <div class="eg" id="eg"></div>

    <div class="ctl">
      <div class="grp"><span id="l-theme">Rang</span><div class="sws" id="sws" role="radiogroup" aria-labelledby="l-theme"></div></div>
      <div class="grp"><span id="l-bg">Fon</span><div class="segs" id="bgs" role="radiogroup" aria-labelledby="l-bg"></div></div>
      <div class="grp"><span id="l-count">Slaydlar soni</span><div class="segs" id="counts" role="radiogroup" aria-labelledby="l-count"></div></div>
      <div class="grp"><span id="l-lang">Til</span><div class="segs" id="langs" role="radiogroup" aria-labelledby="l-lang"></div></div>
    </div>

    <div class="act">
      <button class="go" id="go" type="button">Taqdimotni yaratish</button>
      <div class="note" id="msg" role="status" aria-live="polite"></div>
    </div>
  </main>

  <p class="foot">Sun'iy intellekt xato qilishi mumkin: raqamlar va faktlarni topshirishdan oldin tekshirib chiqing.</p>
</div>

<script>
const THEMES=[
  {id:"ko'k",name:"Ko'k",bg:"#1F3A5F",title:"#FFFFFF",text:"#E8EEF5",btn:"#1F3A5F",btnText:"#FFFFFF"},
  {id:"yashil",name:"Yashil",bg:"#1E5B4F",title:"#FFFFFF",text:"#E6F2EF",btn:"#1E5B4F",btnText:"#FFFFFF"},
  {id:"qora",name:"Qora",bg:"#1A1A1A",title:"#FFD166",text:"#EEEEEE",btn:"#1A1A1A",btnText:"#FFD166"},
  {id:"oq",name:"Oq",bg:"#FFFFFF",title:"#1F3A5F",text:"#333333",btn:"#1F3A5F",btnText:"#FFFFFF"}
];
const BGS=[{id:"sodda",name:"Sodda"},{id:"geometrik",name:"Geometrik"},{id:"chiziqli",name:"Chiziqli"}];
const COUNTS=["5","7","10","12","15"];
const LANGS=[
  {id:"uz",name:"O'zbekcha",eg:["Quyosh energiyasi","Sun'iy intellekt asoslari","Amir Temur davri"]},
  {id:"ru",name:"Русский",eg:["Солнечная энергия","Основы искусственного интеллекта","История Великого шёлкового пути"]},
  {id:"en",name:"English",eg:["Solar energy","Basics of artificial intelligence","History of the Silk Road"]}
];

const $=id=>document.getElementById(id);
const topic=$("topic"), go=$("go"), msg=$("msg"), stage=$("stage");
let theme=THEMES[0].id, lang=LANGS[0].id, count="7", bg="sodda";

function addRadio(host,cls,name,value,checked,text,vars,label){
  const l=document.createElement("label"); l.className=cls;
  const i=document.createElement("input"); i.type="radio"; i.name=name; i.value=value; i.checked=checked;
  if(label) i.setAttribute("aria-label",label);
  const s=document.createElement("span");
  if(text) s.textContent=text;
  if(vars) for(const k in vars) s.style.setProperty(k,vars[k]);
  l.append(i,s); host.append(l);
  return i;
}

function setTheme(id){
  theme=id;
  const t=THEMES.find(x=>x.id===id), r=document.documentElement.style;
  r.setProperty("--stage-bg",t.bg); r.setProperty("--stage-title",t.title); r.setProperty("--stage-text",t.text);
  r.setProperty("--btn-bg",t.btn); r.setProperty("--btn-fg",t.btnText);
}

function renderExamples(){
  const host=$("eg"); host.textContent="";
  const lab=document.createElement("span"); lab.textContent="Masalan:"; host.append(lab);
  LANGS.find(x=>x.id===lang).eg.forEach(txt=>{
    const b=document.createElement("button"); b.type="button"; b.textContent=txt;
    b.addEventListener("click",()=>{topic.value=txt; fit(); topic.focus();});
    host.append(b);
  });
}

function fit(){ topic.style.height="auto"; topic.style.height=topic.scrollHeight+"px"; }

THEMES.forEach(t=>{
  addRadio($("sws"),"sw","theme",t.id,t.id===theme,"",{"--b":t.bg,"--t":t.title,"--x":t.text},t.name)
    .addEventListener("change",()=>setTheme(t.id));
});
BGS.forEach(b=>{
  addRadio($("bgs"),"seg","bg",b.id,b.id===bg,b.name).addEventListener("change",()=>{bg=b.id; stage.dataset.bg=b.id;});
});
COUNTS.forEach(c=>{
  addRadio($("counts"),"seg","count",c,c===count,c).addEventListener("change",()=>{count=c;});
});
LANGS.forEach(l=>{
  addRadio($("langs"),"seg","lang",l.id,l.id===lang,l.name).addEventListener("change",()=>{lang=l.id; renderExamples();});
});

topic.addEventListener("input",fit);
topic.addEventListener("keydown",e=>{ if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); generate(); } });
go.addEventListener("click",generate);
stage.dataset.bg=bg; setTheme(theme); renderExamples(); fit();
if(window.matchMedia("(min-width:700px)").matches) topic.focus();

async function generate(){
  const t=topic.value.trim();
  msg.className="note"; msg.textContent="";
  if(!t){ msg.className="note err"; msg.textContent="Mavzuni yozing."; topic.focus(); return; }
  go.disabled=true; go.textContent="Slaydlar yozilmoqda..."; stage.classList.add("busy");
  msg.textContent="Bu 20-40 soniya oladi. Sahifani yopmang.";
  try{
    const r=await fetch("/generate",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({topic:t,count:count,lang:lang,theme:theme,bg:bg})});
    if(!r.ok){
      let m="Xatolik (kod "+r.status+")";
      try{ const e=await r.json(); m=e.error||m; }
      catch(_){ if([502,503,504].includes(r.status)) m="Server javob bermadi (vaqt tugadi yoki qayta ishga tushmoqda), kod "+r.status+". Qayta urinib ko'ring."; }
      throw new Error(m);
    }
    const blob=await r.blob();
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob); a.download="taqdimot.pptx";
    document.body.append(a); a.click(); a.remove();
    msg.className="note ok"; msg.textContent="Tayyor: taqdimot.pptx yuklab olindi.";
  }catch(e){
    msg.className="note err";
    msg.textContent=(e instanceof TypeError)?"Tarmoq xatosi. Internetni tekshirib, qayta urinib ko'ring.":e.message;
  }
  go.disabled=false; go.textContent="Taqdimotni yaratish"; stage.classList.remove("busy");
}
</script>
</body>
</html>
"""

if __name__ == "__main__":
    # Lokal ishga tushirish. Serverda gunicorn ishlatiladi (README ga qarang).
    # Debug faqat FLASK_DEBUG=1 bo'lganda yoqiladi.
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", 5000)),
        debug=os.environ.get("FLASK_DEBUG") == "1",
    )
