# app.py — UI web local pentru kb_mail.py (FastAPI)
# Rulare din app: uvicorn app:app --port 8000
# URL: http://127.0.0.1:8000/

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from typing import Optional
import html, os, signal, threading, time

from kb_mail import (
    acquire_token_public,
    fetch_last_messages,
    search_messages,
    extract_participants,
    generate_summary_and_reply,
    generate_search_summary_and_reply,
    create_reply_draft,
    graph_get,
    TZ_NAME,
)

APP_TITLE = "Outlook KB — UI local"
DEFAULT_LOGIN = "laurentiu@omnia.capital"
DEFAULT_LAST = 5
DEFAULT_TONE = "brief-firm"

app = FastAPI(title=APP_TITLE)

BASE_CSS = """
<style>
  :root { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, Cantarell, Noto Sans, Arial, "Apple Color Emoji", "Segoe UI Emoji"; }
  body { max-width: 980px; margin: 32px auto; padding: 0 16px; }
  header { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  h1 { font-size: 20px; margin: 0; }
  .right { display: flex; gap: 8px; align-items: center; }
  form { display: grid; gap: 12px; border: 1px solid #eee; padding: 16px; border-radius: 12px; background: #fff; }
  fieldset { border: 0; padding: 0; margin: 0; display: grid; gap: 8px; }
  .row { display: grid; grid-template-columns: 1fr 2fr; gap: 12px; align-items: center; }
  input[type=text], input[type=number], select { width: 100%; padding: 8px 10px; border-radius: 8px; border: 1px solid #ddd; }
  .actions { display: flex; gap: 12px; align-items: center; }
  button { padding: 10px 14px; border: 0; border-radius: 10px; cursor: pointer; }
  .primary { background: #0b5; color: white; }
  .danger { background: #d33; color: white; }
  .muted { color: #666; font-size: 13px; }
  .card { border: 1px solid #eee; padding: 16px; border-radius: 12px; margin-top: 16px; background: #fcfcfc; }
  pre { white-space: pre-wrap; word-break: break-word; }
  .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; font-size: 13px; }
  .ok { color: #090; }
  .warn { color: #b50; }
  .err { color: #d00; }
  details { margin-top: 8px; }
  .inline-form { display: inline; margin: 0; padding: 0; }
  .grid { display: grid; gap: 16px; grid-template-columns: 1fr 1fr; }
</style>
"""

HOME_HTML = f"""
<!doctype html><meta charset="utf-8"><title>{APP_TITLE}</title>
{BASE_CSS}
<header>
  <h1>{APP_TITLE}</h1>
  <div class="right">
    <span class="muted">local-only • token cache în .token_cache.json</span>
    <form class="inline-form" method="post" action="/shutdown" onsubmit="return confirm('Închizi UI-ul?');">
      <button class="danger" type="submit">Stop server</button>
    </form>
  </div>
</header>

<div class="grid">
  <form method="post" action="/run">
    <h3>Summarize & Draft (by Sender/Domain)</h3>
    <fieldset>
      <div class="row"><label>Login (MSA)</label><input type="text" name="login" value="{html.escape(DEFAULT_LOGIN)}" required></div>
      <div class="row">
        <label>Mod</label>
        <div style="display:flex; gap:16px; align-items:center;">
          <label><input type="radio" name="mode" value="domain" checked> Domain</label>
          <label><input type="radio" name="mode" value="sender"> Sender</label>
        </div>
      </div>
      <div class="row"><label>Valoare</label><input type="text" name="value" placeholder="firma.com sau ana@firma.com" required></div>
      <div class="row"><label>Ultimele</label><input type="number" name="last" value="{DEFAULT_LAST}" min="1" max="50"></div>
      <div class="row"><label>Ultimele N zile</label><input type="number" name="days" placeholder="ex: 30 (opțional)" min="1"></div>
      <div class="row"><label>Tone</label>
        <select name="tone">
          <option value="brief-firm" selected>brief-firm</option>
          <option value="friendly-formal">friendly-formal</option>
          <option value="very-concise">very-concise</option>
          <option value="no-nonsense">no-nonsense</option>
        </select>
      </div>
      <div class="row"><label>Propune slot (opțional)</label><input type="text" name="slot" placeholder="Thu 14:00-15:00 Europe/Bucharest"></div>
      <div class="row"><label>Creează draft reply</label><label><input type="checkbox" name="create_draft" checked> da</label></div>
    </fieldset>
    <div class="actions"><button class="primary" type="submit">Run</button></div>
  </form>

  <form method="post" action="/search">
    <h3>Search (by keyword/phrase)</h3>
    <fieldset>
      <div class="row"><label>Login (MSA)</label><input type="text" name="login" value="{html.escape(DEFAULT_LOGIN)}" required></div>
      <div class="row"><label>Fraza/Cuvinte</label><input type="text" name="q" placeholder="ex: contract cadru, oferta 12.3k, deadline vineri" required></div>
      <div class="row"><label>Max rezultate</label><input type="number" name="last" value="20" min="1" max="100"></div>
      <div class="row"><label>Ultimele N zile</label><input type="number" name="days" placeholder="ex: 60 (opțional)" min="1"></div>
      <div class="row"><label>Tone</label>
        <select name="tone">
          <option value="brief-firm" selected>brief-firm</option>
          <option value="friendly-formal">friendly-formal</option>
          <option value="very-concise">very-concise</option>
          <option value="no-nonsense">no-nonsense</option>
        </select>
      </div>
      <div class="row"><label>Creează draft reply</label><label><input type="checkbox" name="create_draft" checked> da</label></div>
    </fieldset>
    <div class="actions"><button class="primary" type="submit">Search</button></div>
  </form>
</div>
"""

RESULT_TPL = """
<!doctype html><meta charset="utf-8"><title>{title}</title>
{css}
<header>
  <h1>{title}</h1>
  <div class="right">
    <a class="muted" href="/">← Înapoi</a>
    <form class="inline-form" method="post" action="/shutdown" onsubmit="return confirm('Închizi UI-ul?');">
      <button class="danger" type="submit">Stop server</button>
    </form>
  </div>
</header>
<div class="card mono">{me_line}</div>
{body}
"""

def render_page(body_html: str, me_line: str = "", title: str = APP_TITLE) -> HTMLResponse:
    return HTMLResponse(RESULT_TPL.format(title=html.escape(title), css=BASE_CSS, me_line=me_line, body=body_html))

@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(HOME_HTML)

# ------- Summarize & Draft (by sender/domain) -------
@app.post("/run", response_class=HTMLResponse)
def run(
    request: Request,
    login: str = Form(...),
    mode: str = Form(...),
    value: str = Form(...),
    last: str = Form(str(DEFAULT_LAST)),
    days: str = Form(""),
    tone: str = Form(DEFAULT_TONE),
    slot: str = Form(""),
    create_draft: Optional[str] = Form(None),
):
    # coercie
    try: last_int = int(last)
    except Exception: last_int = DEFAULT_LAST
    days_int = None
    if days.strip():
        try: days_int = int(days)
        except Exception: days_int = None

    token = acquire_token_public(login_hint=login)
    try:
        me = graph_get("/me", headers={"Authorization": f"Bearer {token}"}, params={"$select":"userPrincipalName,mail,id,displayName"})
        me_line = f"[ME] {html.escape(me.get('userPrincipalName',''))} • {html.escape(me.get('mail','') or '')} • id={html.escape(me.get('id',''))}"
    except Exception as e:
        me_line = f'<span class="warn">[WARN] /me failed: {html.escape(str(e))}</span>'

    val = value.strip()
    if not val:
        return render_page('<p class="err">Valoarea e goală.</p>' + HOME_HTML, me_line)

    if mode == "sender":
        msgs = fetch_last_messages(token, sender=val, top=last_int, days=days_int)
        hint = val; label = f"Sender: {html.escape(val)}"
    else:
        msgs = fetch_last_messages(token, domain=val, top=last_int, days=days_int)
        hint = val; label = f"Domain: {html.escape(val)}"

    if not msgs:
        return render_page(f'<p class="warn">Nu am găsit mesaje pentru <b>{label}</b>.</p>' + HOME_HTML, me_line)

    try:
        summary_md, draft_html = generate_summary_and_reply(msgs, sender_hint=hint, tone=tone, propose_slot=slot or None, timezone_name=TZ_NAME)
    except Exception as e:
        return render_page(f'<p class="err">LLM a eșuat: {html.escape(str(e))}</p>' + HOME_HTML, me_line)

    web_link_html = ""
    if create_draft is not None:
        try:
            newest_id = msgs[0]["id"]
            draft_id = create_reply_draft(token, newest_id, draft_html)
            link = graph_get(f"/me/messages/{draft_id}", headers={"Authorization": f"Bearer {token}"}, params={"$select":"webLink"}).get("webLink")
            web_link_html = f'<p class="ok">Draft creat: <a href="{html.escape(link or "")}">{html.escape(link or "")}</a></p>' if link else '<p class="warn">Draft creat, dar fără webLink.</p>'
        except Exception as e:
            web_link_html = f'<p class="err">Eroare la creare draft: {html.escape(str(e))}</p>'

    # afișare
    summary_block = f"<div class='card'><h3>Summary</h3><pre>{html.escape(summary_md)}</pre></div>"
    draft_block = f"<div class='card'><h3>Draft</h3>{draft_html}</div>"
    body = summary_block + draft_block + (f"<div class='card'>{web_link_html}</div>" if web_link_html else "")
    return render_page(body, me_line)

# ------- NEW: Search (by keyword/phrase) -------
@app.post("/search", response_class=HTMLResponse)
def search(
    request: Request,
    login: str = Form(...),
    q: str = Form(...),
    last: str = Form("20"),
    days: str = Form(""),
    tone: str = Form(DEFAULT_TONE),
    create_draft: Optional[str] = Form(None),
):
    # coercie
    try: last_int = int(last)
    except Exception: last_int = 20
    days_int = None
    if days.strip():
        try: days_int = int(days)
        except Exception: days_int = None

    token = acquire_token_public(login_hint=login)
    try:
        me = graph_get("/me", headers={"Authorization": f"Bearer {token}"}, params={"$select":"userPrincipalName,mail,id,displayName"})
        me_line = f"[ME] {html.escape(me.get('userPrincipalName',''))} • {html.escape(me.get('mail','') or '')} • id={html.escape(me.get('id',''))}"
    except Exception as e:
        me_line = f'<span class="warn">[WARN] /me failed: {html.escape(str(e))}</span>'

    phrase = q.strip()
    if not phrase:
        return render_page('<p class="err">Fraza de căutare e goală.</p>' + HOME_HTML, me_line)

    # 1) căutare mesaje
    try:
        msgs = search_messages(token, phrase=phrase, top=last_int, days=days_int)
    except Exception as e:
        return render_page(f'<p class="err">Eroare la căutare: {html.escape(str(e))}</p>' + HOME_HTML, me_line)

    if not msgs:
        return render_page(f'<p class="warn">Nu am găsit mesaje pentru <b>{html.escape(phrase)}</b>.</p>' + HOME_HTML, me_line)

    # 2) participanți unici
    participants = extract_participants(msgs)
    participants_html = "<ul>" + "".join(f"<li>{html.escape(a)}</li>" for a in participants) + "</ul>"

    # 3) timeline (cele mai noi primele)
    timeline_html = "<ol>" + "".join(
        f"<li>{html.escape(m.get('receivedDateTime',''))} — <span class='mono'>{html.escape((m.get('from') or {}).get('emailAddress',{}).get('address','') or '')}</span> — {html.escape(m.get('subject','') or '')}</li>"
        for m in msgs
    ) + "</ol>"

    # 4) summary focalizat pe căutare + draft
    try:
        summary_md, draft_html = generate_search_summary_and_reply(msgs, query=phrase, tone=tone, timezone_name=TZ_NAME)
    except Exception as e:
        return render_page(f'<p class="err">LLM a eșuat: {html.escape(str(e))}</p>' + HOME_HTML, me_line)

    # 5) creare draft (opțional) — reply la cel mai nou mesaj găsit
    web_link_html = ""
    if create_draft is not None:
        try:
            newest_id = msgs[0]["id"]
            draft_id = create_reply_draft(token, newest_id, draft_html)
            link = graph_get(f"/me/messages/{draft_id}", headers={"Authorization": f"Bearer {token}"}, params={"$select":"webLink"}).get("webLink")
            web_link_html = f'<p class="ok">Draft creat: <a href="{html.escape(link or "")}">{html.escape(link or "")}</a></p>' if link else '<p class="warn">Draft creat, dar fără webLink.</p>'
        except Exception as e:
            web_link_html = f'<p class="err">Eroare la creare draft: {html.escape(str(e))}</p>'

    # out
    blocks = []
    blocks.append(f"<div class='card'><h3>Search query</h3><div class='mono'>{html.escape(phrase)}</div></div>")
    blocks.append(f"<div class='card'><h3>Summary (focus pe căutare)</h3><pre>{html.escape(summary_md)}</pre></div>")
    blocks.append(f"<div class='card'><h3>Participanți</h3>{participants_html}</div>")
    blocks.append(f"<div class='card'><h3>Timeline</h3>{timeline_html}</div>")
    blocks.append(f"<div class='card'><h3>Draft</h3>{draft_html}</div>")
    if web_link_html:
        blocks.append(f"<div class='card'>{web_link_html}</div>")

    return render_page("".join(blocks), me_line, title=f"{APP_TITLE} — Search")

# ------- Stop server -------
@app.post("/shutdown", response_class=HTMLResponse)
def shutdown():
    def _kill():
        time.sleep(0.2); os.kill(os.getpid(), signal.SIGINT)
    threading.Thread(target=_kill, daemon=True).start()
    return HTMLResponse("<p class='ok'>Serverul se oprește… Poți închide această fereastră.</p>" + HOME_HTML)
