# kb_mail.py — Outlook KB Agent (Graph) — single file (MSA-friendly)
# Funcții: căutare după expeditor/domeniu/sintagmă, generare summary + draft, creare draft Outlook.
# Azure App (MSA):
#   - Authentication: Mobile and desktop applications -> Redirect URI: http://localhost
#   - Supported accounts: Any org directory AND personal Microsoft accounts
#   - Allow public client flows: Yes
#   - API permissions (Delegated): Mail.Read, Mail.ReadWrite, User.Read
#
# .env:
#   CLIENT_ID=...
#   TENANT_ID=consumers
#   OPENAI_API_KEY=...
#   DEFAULT_MODEL=gpt-4.1-mini
#   TIMEZONE=Europe/Bucharest

import os, sys, json, time, argparse, re
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal, requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# ---------- env ----------
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID", "consumers")
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
DEFAULT_MODEL = os.getenv("DEFAULT_MODEL", "gpt-4.1-mini")
TZ_NAME = os.getenv("TIMEZONE", "Europe/Bucharest")
if not CLIENT_ID:
    raise RuntimeError("CLIENT_ID lipseste din .env")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "User.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_FILE = ".token_cache.json"

# ---------- LLM ----------
try:
    from openai import OpenAI
    llm = OpenAI(api_key=OPENAI_KEY) if OPENAI_KEY else None
except Exception:
    llm = None

# ---------- token cache ----------
class FileCache(msal.SerializableTokenCache):
    def __init__(self, filename):
        super().__init__()
        self.filename = filename
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                self.deserialize(f.read())
    def persist(self):
        with open(self.filename, "w", encoding="utf-8") as f:
            f.write(self.serialize())

def acquire_token_public(login_hint=None) -> str:
    cache = FileCache(TOKEN_CACHE_FILE)
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if accounts:
        res = app.acquire_token_silent(SCOPES, account=accounts[0])
        if res and "access_token" in res:
            cache.persist(); return res["access_token"]
    res = app.acquire_token_interactive(scopes=SCOPES, timeout=300, prompt="login", login_hint=login_hint)
    if res and "access_token" in res:
        cache.persist(); return res["access_token"]
    raise RuntimeError(f"Interactive flow fără token: {res}")

# ---------- Graph helpers ----------
def graph_get(path, headers=None, params=None):
    url = f"{GRAPH}{path}"
    r = requests.get(url, headers=headers, params=params)
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "3"))); r = requests.get(url, headers=headers, params=params)
    r.raise_for_status(); return r.json()

def graph_post(path, headers=None, data=None):
    url = f"{GRAPH}{path}"
    r = requests.post(url, headers=headers, json=data or {})
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "3"))); r = requests.post(url, headers=headers, json=data or {})
    r.raise_for_status(); return r.json()

def graph_patch(path, headers=None, data=None):
    url = f"{GRAPH}{path}"
    r = requests.patch(url, headers=headers, json=data or {})
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "3"))); r = requests.patch(url, headers=headers, json=data or {})
    r.raise_for_status(); return r.json() if r.text else {}

# ---------- util: parse datetime ----------
def _parse_iso_dt(s: str):
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None

# ---------- Email fetch (by sender/domain) ----------
def fetch_last_messages(token: str, sender=None, domain=None, top=5, folder_id=None, days=None):
    """
    Outlook.com/MSA:
      - $search NU merge cu $orderby; NU combina $search cu $filter => filtrăm local pe zile.
    """
    headers = {"Authorization": f"Bearer {token}"}
    path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
    params = {
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,conversationId,webLink",
    }

    search_term = None
    if sender:
        search_term = f"from:{sender}"
    elif domain:
        search_term = f"from:{domain}"

    if search_term:
        headers["ConsistencyLevel"] = "eventual"
        params["$search"] = f"\"{search_term}\""
        params["$top"] = max(top * 5, 25)  # oversample
        data = graph_get(path, headers=headers, params=params)
        items = data.get("value", [])
        if days is not None:
            cutoff = datetime.now(timezone.utc) - timedelta(days=days)
            items = [m for m in items if (dt := _parse_iso_dt(m.get("receivedDateTime",""))) and dt >= cutoff]
        items.sort(key=lambda m: m.get("receivedDateTime",""), reverse=True)
        return items[:top]

    # fără $search – putem folosi $filter + $orderby
    if days is not None:
        dt = datetime.now(timezone.utc) - timedelta(days=days)
        params["$filter"] = f"receivedDateTime ge {dt.isoformat()}"
    params["$orderby"] = "receivedDateTime desc"
    params["$top"] = top
    data = graph_get(path, headers=headers, params=params)
    return data.get("value", [])

# ---------- NEW: Search by keyword/phrase ----------
def search_messages(token: str, phrase: str, top=20, folder_id=None, days=None):
    """
    Căutare full-text cu $search="phrase". Nu combinăm cu $filter/$orderby.
    Facem oversample, apoi sortăm local + filtrăm pe zile local.
    """
    headers = {"Authorization": f"Bearer {token}", "ConsistencyLevel": "eventual"}
    path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
    params = {
        "$search": f"\"{phrase}\"",
        "$top": max(top * 5, 50),  # oversample
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,conversationId,webLink",
    }
    data = graph_get(path, headers=headers, params=params)
    items = data.get("value", [])

    if days is not None:
        cutoff = datetime.now(timezone.utc) - timedelta(days=days)
        items = [m for m in items if (dt := _parse_iso_dt(m.get("receivedDateTime",""))) and dt >= cutoff]

    items.sort(key=lambda m: m.get("receivedDateTime",""), reverse=True)
    return items[:top]

def extract_participants(messages):
    """
    Returnează setul de adrese unice implicate (from/to/cc) + listă (sortată).
    """
    addrs = set()
    def safe_add(addr_obj):
        if not addr_obj: return
        a = addr_obj.get("address") if isinstance(addr_obj, dict) else None
        if a: addrs.add(a.lower())

    for m in messages:
        frm = (m.get("from") or {}).get("emailAddress")
        safe_add(frm)
        for r in (m.get("toRecipients") or []):
            safe_add(r.get("emailAddress"))
        for r in (m.get("ccRecipients") or []):
            safe_add(r.get("emailAddress"))
    return sorted(addrs)

# ---------- email cleaning ----------
RE_SIGNATURE = re.compile(r"(?is)(--\s*\n.*$|^Sent from my iPhone.*$|^Best regards,.*$|^Cu stima,.*$)")
RE_QUOTED   = re.compile(r"(?is)(^>.*\n|^On .* wrote:.*$|^De la:.*\n|^From:.*\n.*)$")

def html_to_text(html_content: str) -> str:
    soup = BeautifulSoup(html_content or "", "html.parser")
    for bq in soup.find_all("blockquote"): bq.decompose()
    return soup.get_text("\n")

def trim_email_body(raw_html: str) -> str:
    txt = html_to_text(raw_html)
    txt = RE_QUOTED.sub("", txt); txt = RE_SIGNATURE.sub("", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt).strip()
    return txt[:8000]

# ---------- LLM prompts ----------
SYSTEM_PROMPT = """Esti un asistent de email concis, precis, fara emoticoane.
1) Rezumi ultimile emailuri (cel mai nou primul) in 4-8 bullet-uri: intentii, cerinte, blocaje, termene, cifre.
2) Redactezi un draft de raspuns business: 2 paragrafe scurte + lista next steps numerotata. Ton: ferm, politicos.
3) Daca lipsesc atasamente sau info, cere-le explicit.
Returneaza JSON cu cheile: summary (string Markdown), draft_html (string HTML).
"""

SEARCH_SYSTEM_PROMPT = """Esti un asistent de email care sintetizeaza rezultate pentru o cautare text.
1) Construieste un rezumat focalizat strict pe cuvintele/fraza data: concluzii, actiuni, decizii, cifre, termene.
2) Listeaza clar lacunele de informatie sau contradictiile aparute in firul de discutie.
3) Genereaza un draft de reply (HTML) care abordeaza tema cautarii si propune next steps concrete.
Returneaza JSON cu cheile: summary (string Markdown), draft_html (string HTML).
"""

def generate_summary_and_reply(emails, sender_hint=None, tone="brief-firm", propose_slot=None, timezone_name="Europe/Bucharest"):
    if not llm: raise RuntimeError("LLM client nu este configurat. Seteaza OPENAI_API_KEY in .env")
    items = []
    for i, m in enumerate(emails, 1):
        body_html = m.get("body", {}).get("content", "")
        items.append({
            "i": i,
            "subject": m.get("subject", ""),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address", ""),
            "to": [(r.get("emailAddress") or {}).get("address","") for r in (m.get("toRecipients") or [])],
            "cc": [(r.get("emailAddress") or {}).get("address","") for r in (m.get("ccRecipients") or [])],
            "received": m.get("receivedDateTime", ""),
            "snippet": trim_email_body(body_html)
        })
    payload = {"task":"summarize_and_draft","tone":tone,"timezone":timezone_name,"propose_slot":propose_slot,"sender_hint":sender_hint,"emails":items}
    resp = llm.chat.completions.create(
        model=DEFAULT_MODEL, temperature=0.2,
        messages=[{"role":"system","content":SYSTEM_PROMPT},{"role":"user","content":json.dumps(payload, ensure_ascii=False)}]
    )
    content = resp.choices[0].message.content
    m = re.search(r"\{.*\}\s*$", content, re.S); data = json.loads(m.group(0) if m else content)
    return data["summary"], data["draft_html"]

def generate_search_summary_and_reply(emails, query, tone="brief-firm", timezone_name="Europe/Bucharest"):
    if not llm: raise RuntimeError("LLM client nu este configurat. Seteaza OPENAI_API_KEY in .env")
    items = []
    for i, m in enumerate(emails, 1):
        body_html = m.get("body", {}).get("content", "")
        items.append({
            "i": i,
            "subject": m.get("subject", ""),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address", ""),
            "to": [(r.get("emailAddress") or {}).get("address","") for r in (m.get("toRecipients") or [])],
            "cc": [(r.get("emailAddress") or {}).get("address","") for r in (m.get("ccRecipients") or [])],
            "received": m.get("receivedDateTime", ""),
            "snippet": trim_email_body(body_html)
        })
    payload = {"task":"search_summarize_and_draft","tone":tone,"timezone":timezone_name,"query":query,"emails":items}
    resp = llm.chat.completions.create(
        model=DEFAULT_MODEL, temperature=0.2,
        messages=[{"role":"system","content":SEARCH_SYSTEM_PROMPT},{"role":"user","content":json.dumps(payload, ensure_ascii=False)}]
    )
    content = resp.choices[0].message.content
    m = re.search(r"\{.*\}\s*$", content, re.S); data = json.loads(m.group(0) if m else content)
    return data["summary"], data["draft_html"]

# ---------- Draft reply ----------
def create_reply_draft(token: str, message_id: str, reply_html: str) -> str:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    draft = graph_post(f"/me/messages/{message_id}/createReply", headers=headers, data={})
    draft_id = draft["id"]
    patch = {"body": {"contentType": "HTML", "content": reply_html}}
    graph_patch(f"/me/messages/{draft_id}", headers=headers, data=patch)
    return draft_id

# ---------- CLI (ramane util pentru terminal) ----------
def main():
    p = argparse.ArgumentParser(description="Outlook KB Agent (summarize + draft reply via Graph)")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--from-sender", help="email exact al expeditorului (ex: john@company.com)")
    g.add_argument("--from-domain", help="domeniu (ex: company.com)")
    p.add_argument("--last", type=int, default=5, help="cate mesaje luam (default 5)")
    p.add_argument("--days", type=int, default=None, help="limiteaza la ultimele N zile")
    p.add_argument("--folder-id", help="restrict la un folder anume")
    p.add_argument("--tone", default="brief-firm", help="brief-firm, friendly-formal, very-concise")
    p.add_argument("--slot", default=None, help="propune un slot: Thu 14:00-15:00 Europe/Bucharest")
    p.add_argument("--create-draft", action="store_true", help="creeaza draft reply la cel mai nou mesaj")
    p.add_argument("--login", default=None, help="login_hint (ex: nume@outlook.com)")
    args = p.parse_args()

    token = acquire_token_public(login_hint=args.login)
    try:
        who = graph_get("/me", headers={"Authorization": f"Bearer {token}"},
                        params={"$select":"userPrincipalName,mail,id,displayName,surname,givenName,preferredLanguage,ageGroup,mobilePhone,jobTitle,officeLocation,businessPhones"})
        print("[ME]", who)
    except Exception as e:
        print("[WARN] /me check failed:", e)

    if args.from_sender:
        msgs = fetch_last_messages(token, sender=args.from_sender, top=args.last, days=args.days, folder_id=args.folder_id)
        hint = args.from_sender
    else:
        msgs = fetch_last_messages(token, domain=args.from_domain, top=args.last, days=args.days, folder_id=args.folder_id)
        hint = args.from_domain

    if not msgs:
        print("Nu am gasit mesaje pe criteriile date."); return 0

    summary_md, draft_html = generate_summary_and_reply(msgs, sender_hint=hint, tone=args.tone, propose_slot=args.slot, timezone_name=TZ_NAME)

    print("\n=== SUMMARY ===\n"); print(summary_md)
    print("\n=== DRAFT (HTML) ===\n"); print(draft_html)

    if args.create_draft:
        newest = msgs[0]["id"]
        draft_id = create_reply_draft(token, newest, draft_html)
        link = graph_get(f"/me/messages/{draft_id}", headers={"Authorization": f"Bearer {token}"}, params={"$select":"webLink"}).get("webLink")
        print(f"\nDraft creat: {link or '(fara link)'}")

    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except requests.HTTPError as e:
        body = ""
        try: body = e.response.text
        except Exception: pass
        print(f"[HTTP ERROR] {e} / {body}"); sys.exit(2)
    except Exception as e:
        print(f"[ERROR] {e}"); sys.exit(1)
