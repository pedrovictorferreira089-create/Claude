"""Acesso à API do Lever Conversas (plataforma Helena) com cache em ./data/lever/.

A chave vem de LEVER_TOKEN (token permanente "pn_...", enviado como Bearer). Em ambientes cujo proxy já injeta a chave,
a variável pode ficar vazia.
"""
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://api.app.leverconversas.com.br"
TOKEN = os.environ.get("LEVER_TOKEN", "")
CACHE = os.path.join(os.environ.get("BELLE_CACHE", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")), "lever")
PAINEL_SDR = "SDRs"
FASES_VENDA = ("Convertidos", "Convertidos avulsos")


def api(path, cache=True, **q):
    os.makedirs(CACHE, exist_ok=True)
    fn = os.path.join(CACHE, (path.strip("/") + "__" + urllib.parse.urlencode(sorted(q.items()))).replace("/", "_")
                      .replace("&", "_").replace("=", "-") + ".json")
    if cache and os.path.exists(fn):
        return json.load(open(fn))
    url = BASE + path + ("?" + urllib.parse.urlencode(q, doseq=True) if q else "")
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}"} if TOKEN else {})
    for tentativa in range(5):
        try:
            data = json.loads(urllib.request.urlopen(req, timeout=120).read())
            break
        except urllib.error.HTTPError as e:
            if tentativa == 4 or (400 <= e.code < 500 and e.code != 429) or (e.code == 500 and tentativa == 1):
                raise
            time.sleep(5 * (tentativa + 1))
        except Exception:
            if tentativa == 4:
                raise
            time.sleep(5 * (tentativa + 1))
    time.sleep(0.2)
    json.dump(data, open(fn, "w"), ensure_ascii=False)
    return data


def paginas(path, **q):
    out, n = [], 1
    while True:
        d = api(path, PageSize=100, PageNumber=n, **q)
        out.extend(d["items"])
        if not d.get("hasMorePages"):
            return out
        n += 1


def painel(titulo=PAINEL_SDR):
    p = next(p for p in paginas("/crm/v1/panel") if (p.get("title") or "").strip() == titulo)
    return api(f"/crm/v1/panel/{p['id']}", IncludeDetails="Steps")


def cards(painel_det, fases=FASES_VENDA):
    """Cards das fases pedidas, com a fase e os contatos (nome e telefone) de cada um."""
    out = []
    for s in painel_det["steps"]:
        if s["title"].strip() not in fases:
            continue
        for c in paginas("/crm/v1/panel/card", PanelId=painel_det["id"], StepId=s["id"]):
            c["fase"] = s["title"].strip()
            c["contatos"] = [contato(i) for i in c.get("contactIds") or []]
            out.append(c)
    return out


def contato(i):
    """Contato pelo id; contatos apagados respondem erro 500 e voltam só com o id."""
    try:
        return api(f"/core/v1/contact/{i}")
    except urllib.error.HTTPError:
        return {"id": i, "erro": True}


def campo(card, prefixo):
    """Campo personalizado pelo começo da chave (as chaves do Lever têm sufixos gerados, ex.: m-s-de-fechamento-93)."""
    cf = card.get("customFields") or {}
    k = next((k for k in cf if k.startswith(prefixo)), None)
    v = cf.get(k) if k else None
    return (", ".join(map(str, v)) if isinstance(v, list) else str(v if v is not None else "")).strip()
