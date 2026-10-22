"""Placar do Quintou (quinta-feira comercial) a partir da API Belle Software.

Uso:
  python3 belle_quintou.py [dd/mm/aaaa] [--corte ID] [--teto ID] > quintou.json   (padrão: hoje)
  python3 belle_quintou.py --corte-agora     maior idAgendamento existente agora (rodar antes do expediente)

Para cada colaboradora do comercial, no dia informado:
  - agendamentos que ela incluiu no Belle no dia (um por cliente e data, sem desmarcados/cancelados);
  - quantos desses foram marcados para o próprio dia (encaixe no mesmo dia);
  - vendas do dia de clientes agendadas por ela, com o percentual já pago (regra dos 70%).

A API não expõe a data de inclusão do agendamento, mas o idAgendamento é sequencial: o que foi incluído
no dia tem id entre o último agendamento existente antes do expediente (corte) e o primeiro do dia seguinte
(teto). O corte exato vem de --corte-agora rodado antes do expediente, e o teto do idMaximo da última
rodada do dia. Sem eles, os dois são estimados pelo maior id dos agendamentos com data anterior (ou igual)
ao dia, tolerando alguns lançamentos retroativos. A estimativa erra quando a recepção lança atendimentos
passados durante o expediente, por isso o corte do dia deve ser medido antes de começar.
Quem vendeu vem do Lever: cada card que entrou em Convertidos (ou Reativados com venda) no dia é casado
com a venda do Belle pelo telefone da cliente, e a vendedora é a responsável do card. Sem card casado, a
venda fica sugerida para quem agendou a cliente (o agendamento do próprio dia ou, se não houver, o mais
recente dela no período consultado), para a líder confirmar no painel.
"""
import argparse
import datetime as dt
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from zoneinfo import ZoneInfo

BASE = "https://app.bellesoftware.com.br/api/release/controller/IntegracaoExterna/v1.0/"
TOKEN = os.environ.get("BELLE_TOKEN", "")
UNIDADES = {1: "Petrópolis", 2: "Lagoa Nova", 3: "Capim Macio", 6: "Zona Norte", 4: "Laser"}
# Nome no Belle -> nome no placar
EQUIPE = {
    "Juliane Hemilly Salvador Rodrigues": "Juliane",
    "Bruna Letícia do Nascimento Silva": "Bruna",
    "Samilly Rohany Câmara Silva": "Samilly",
}
# Primeiro nome do responsável no Lever -> nome no placar
EQUIPE_LEVER = {"juliane": "Juliane", "julliane": "Juliane", "bruna": "Bruna", "samilly": "Samilly"}
LEVER = "https://api.app.leverconversas.com.br/"
LEVER_TOKEN = os.environ.get("LEVER_TOKEN", "")  # sem ele, o proxy do ambiente injeta a credencial
PAINEIS = {"929d58b8-2650-4993-9deb-b7aaa6481b3b": "SDR", "8132810d-57ae-48e6-853e-eb3abcf13d2c": "Social Selling",
           "7bea81d4-9a57-4556-8219-92df72d100e4": "Reativação"}
FASES_VENDA = ("convertidos", "convertidos avulsos", "reativados com venda")
FORA = ("desmarc", "cancel")
FUSO = ZoneInfo("America/Fortaleza")
JANELA_ANTES, JANELA_DEPOIS = 60, 92  # dias consultados antes e depois do dia (a API aceita até 3 meses)
RETROATIVOS = 5  # lançamentos retroativos tolerados acima do corte estimado


def api(path, **q):
    req = urllib.request.Request(BASE + path + "?" + urllib.parse.urlencode(q), headers={"Authorization": TOKEN})
    for tentativa in range(4):
        try:
            data = json.loads(urllib.request.urlopen(req, timeout=300).read())
            break
        except urllib.error.HTTPError as e:
            if tentativa == 3 or (400 <= e.code < 500 and e.code != 429):
                raise
            time.sleep(20)
        except Exception:
            if tentativa == 3:
                raise
            time.sleep(20)
    time.sleep(1.7)
    if isinstance(data, dict) and ("erro" in data or "error" in data):
        raise RuntimeError(f"{path} {q}: {data}")
    return data


def lever(path, **q):
    headers = {"Authorization": f"Bearer {LEVER_TOKEN}"} if LEVER_TOKEN else {}
    req = urllib.request.Request(LEVER + path + ("?" + urllib.parse.urlencode(q) if q else ""), headers=headers)
    for tentativa in range(4):
        try:
            return json.loads(urllib.request.urlopen(req, timeout=120).read())
        except urllib.error.HTTPError as e:
            if tentativa == 3 or (400 <= e.code < 500 and e.code != 429):
                raise
        except Exception:
            if tentativa == 3:
                raise
        time.sleep(10)


def telefone(s):
    """Últimos 8 dígitos: casam celulares com e sem o 9 extra, DDD e código do país."""
    d = "".join(c for c in str(s or "") if c.isdigit())
    return d[-8:] if len(d) >= 8 else None


def cards_vendidos(dia):
    """Cards que entraram numa fase de venda no dia (última movimentação no dia, horário de Natal)."""
    nomes = {a["userId"]: a["name"] for a in lever("core/v1/agent", pageSize=100)}
    out = []
    for painel, funil in PAINEIS.items():
        fases = [f for f in lever(f"crm/v1/panel/{painel}", includeDetails="Steps")["steps"] if f["title"].strip().lower() in FASES_VENDA]
        for fase in fases:
            pagina = 1
            while True:
                r = lever("crm/v1/panel/card", panelId=painel, stepId=fase["id"], pageNumber=pagina, pageSize=100,
                          orderBy="updatedAt", orderDirection="DESCENDING")
                passou = False
                for c in r["items"]:
                    quando = dt.datetime.fromisoformat(c["updatedAt"].replace("Z", "+00:00")).astimezone(FUSO).date()
                    if quando < dia:
                        passou = True
                        break
                    if quando == dia:
                        resp = nomes.get(c.get("responsibleUserId"), "")
                        out.append(dict(card=c["id"], key=c["key"], titulo=c["title"], funil=funil, fase=fase["title"].strip(),
                                        valor=c.get("monetaryAmount") or 0, responsavelLever=resp,
                                        responsavel=EQUIPE_LEVER.get((resp.split() or [""])[0].lower()),
                                        contatos=c.get("contactIds") or []))
                if passou or not r.get("hasMorePages"):
                    break
                pagina += 1
    for c in out:
        fones = {telefone(c["titulo"])}
        for cid in c.pop("contatos"):
            fones.add(telefone(lever(f"core/v1/contact/{cid}").get("phoneNumber")))
        c["fones"] = sorted(f for f in fones if f)
    return out


def br(v):
    if v in (None, ""):
        return 0.0
    s = str(v)
    return float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)


def data_br(s):
    return dt.datetime.strptime(s, "%d/%m/%Y").date()


def fmt(d):
    return d.strftime("%d/%m/%Y")


def primeiro_nome(s):
    return " ".join(str(s or "").split()[:1]).title()


def dona(a):
    return EQUIPE.get(" ".join(str(a.get("nomeUsuarioInclusao") or "").split()))


def agenda(ini, fim):
    out = []
    for e in UNIDADES:
        out += [dict(a, unidade=UNIDADES[e]) for a in api("relatorios/relatorio_atendimentos", dtInicio=fmt(ini), dtFim=fmt(fim), codEstab=str(e))]
    return out


def estimar_corte(ids):
    """Maior id tolerando alguns lançamentos retroativos acima dele. Só uma aproximação: a recepção lança
    atendimentos passados ao longo do dia, e cada um deles empurra o maior id para dentro do expediente."""
    ids = sorted(ids, reverse=True)
    return ids[min(RETROATIVOS, len(ids) - 1)] if ids else 0


def corte_agora(hoje):
    antes = agenda(hoje - dt.timedelta(days=JANELA_ANTES), hoje - dt.timedelta(days=1))
    depois = agenda(hoje, hoje + dt.timedelta(days=JANELA_DEPOIS))
    return max(a["idAgendamento"] for a in antes + depois)


def montar(dia, corte=None, teto=None, hoje=None):
    hoje = hoje or dt.datetime.now(FUSO).date()
    antes = agenda(dia - dt.timedelta(days=JANELA_ANTES), dia - dt.timedelta(days=1))
    depois = agenda(dia, dia + dt.timedelta(days=JANELA_DEPOIS))
    vendas = []
    for e in UNIDADES:
        vendas += [dict(v, unidade=UNIDADES[e]) for v in api("vendas_detalhado", estab=str(e), dtInicio=fmt(dia), dtFim=fmt(dia))]

    corte_estimado = corte is None
    if corte is None:
        corte = estimar_corte([a["idAgendamento"] for a in antes])
    if teto is None and hoje > dia:
        teto = estimar_corte([a["idAgendamento"] for a in antes + depois if data_br(a["dataAgendamento"]) <= dia])

    placar = {nome: dict(agendamentos=[], vendas=[]) for nome in EQUIPE.values()}
    vistos = set()
    for a in sorted(depois, key=lambda x: x["idAgendamento"]):
        nome = dona(a)
        if not nome or a["idAgendamento"] <= corte or (teto and a["idAgendamento"] > teto):
            continue
        if any(x in str(a["statusAgendamento"]).lower() for x in FORA):
            continue
        chave = (nome, a["codigoCliente"], a["dataAgendamento"])
        if chave in vistos:  # AV + serviço avulso da mesma cliente no mesmo dia contam uma vez
            continue
        vistos.add(chave)
        placar[nome]["agendamentos"].append(dict(
            id=a["idAgendamento"], cliente=primeiro_nome(a["nomeCliente"]), data=data_br(a["dataAgendamento"]).isoformat(),
            hora=a["horarioAgendamento"], tipo=a["tipoAgendamento"], servico=(a["nomeServico"] or "").strip(),
            unidade=a["unidade"], status=a["statusAgendamento"], mesmoDia=data_br(a["dataAgendamento"]) == dia))

    # quem agendou cada cliente: prioridade ao agendamento do próprio dia, senão o mais recente
    dono_cliente = {}
    for a in sorted(antes + depois, key=lambda x: (data_br(x["dataAgendamento"]) == dia, x["idAgendamento"])):
        if dona(a):
            dono_cliente[str(a["codigoCliente"])] = dona(a)
    for v in vendas:
        nome = dono_cliente.get(str(v["cod_cliente"]))
        valor = br(v["valor_venda"])
        if not nome or valor <= 0:
            continue
        pago = sum(br(p["valorBruto"]) for p in v.get("parcelas") or [] if p.get("confirmado") == "Sim")
        placar[nome]["vendas"].append(dict(
            id=v["id_venda"], cliente=primeiro_nome((v.get("dados_cliente") or {}).get("nome")), valor=round(valor, 2),
            pago=round(pago, 2), pctPago=round(min(pago / valor, 1), 4), unidade=v["unidade"],
            itens=", ".join(sorted({i["desc_item"] for i in v.get("itens_venda") or []}))))

    # todas as vendas do dia com dona sugerida (quem agendou) e, quando houver, o card do Lever casado pelo telefone
    cards = cards_vendidos(dia)
    por_fone = {}
    for c in cards:
        for f in c["fones"]:
            por_fone.setdefault(f, []).append(c)
    todas, casados = [], set()
    for v in vendas:
        valor = br(v["valor_venda"])
        if valor <= 0:
            continue
        cli = v.get("dados_cliente") or {}
        achados = {c["card"]: c for f in {telefone(cli.get(k)) for k in ("celular", "celular2", "telefone")} if f for c in por_fone.get(f, [])}
        card = max(achados.values(), key=lambda c: (c["responsavel"] is not None, c["valor"]), default=None)
        agendou = dono_cliente.get(str(v["cod_cliente"]))
        if not agendou and not (card and card["responsavel"]):
            continue
        if card:
            casados.add(card["card"])
        pago = sum(br(p["valorBruto"]) for p in v.get("parcelas") or [] if p.get("confirmado") == "Sim")
        todas.append(dict(
            id=v["id_venda"], cliente=primeiro_nome(cli.get("nome")), valor=round(valor, 2), pago=round(pago, 2),
            pctPago=round(min(pago / valor, 1), 4), unidade=v["unidade"], agendou=agendou,
            itens=", ".join(sorted({i["desc_item"] for i in v.get("itens_venda") or []})),
            lever=card and {k: card[k] for k in ("key", "funil", "fase", "valor", "responsavel", "responsavelLever")}))
    sem_belle = [{k: c[k] for k in ("key", "funil", "fase", "valor", "responsavel", "responsavelLever")} | {"cliente": primeiro_nome(c["titulo"])}
                 for c in cards if c["card"] not in casados and c["responsavel"]]

    for p in placar.values():
        p["agendamentos"].sort(key=lambda x: x["id"])
        p["vendas"].sort(key=lambda x: x["id"])
    return dict(dia=dia.isoformat(), atualizadoEm=dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                fonte="Belle", corteId=corte, corteEstimado=corte_estimado, tetoId=teto,
                idMaximo=max((a["idAgendamento"] for a in antes + depois), default=0), equipe=placar,
                vendas=sorted(todas, key=lambda x: x["id"]), leverSemBelle=sem_belle)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("dia", nargs="?", help="dd/mm/aaaa (padrão: hoje)")
    ap.add_argument("--corte", type=int, help="maior idAgendamento existente antes do expediente")
    ap.add_argument("--teto", type=int, help="maior idAgendamento incluído até o fim do dia")
    ap.add_argument("--corte-agora", action="store_true", help="imprime o maior idAgendamento existente agora")
    args = ap.parse_args()
    hoje = dt.datetime.now(FUSO).date()
    if args.corte_agora:
        print(corte_agora(hoje))
    else:
        dia = data_br(args.dia) if args.dia else hoje
        json.dump(montar(dia, args.corte, args.teto, hoje), sys.stdout, ensure_ascii=False, indent=1)
