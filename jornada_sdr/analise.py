"""Análise da jornada de compra (rodar depois da coleta completa). Imprime os números usados no fluxograma."""
import collections
import datetime as dt
import glob
import json
import re
import statistics as st

from base import D, HOJE, UNI, br, categoria, compras_por_cliente, planos

ps = planos()
C = compras_por_cliente(ps)
ESTABS = ("1", "2", "3", "6")


def pc(x, n):
    return f"{100 * x / n:.1f}%" if n else "-"


# ------------------------------------------------------------------ vendas detalhadas (planos + avulsos) e cadastro
CAT_VENDA = {}  # id da venda -> categoria do plano (pelos serviços do plano no relatório de planos)
for f in glob.glob(D + "venda_planos__*"):
    for p in json.load(open(f)):
        if p.get("idVenda"):
            val = collections.Counter()
            for s_ in p["servicos"]:
                val[categoria(s_["nomeServico"])] += br(s_["valorTotalServico"])
            princ = [k for k, _ in val.most_common() if k != "Acessório"]
            CAT_VENDA[str(p["idVenda"])] = princ[0] if princ else "Acessório"
vendas, cad = [], {}
vistos = set()
for f in glob.glob(D + "vendas_detalhado__*"):
    e = re.search(r"estab-(\d+)", f).group(1)
    for v in json.load(open(f)):
        if v["id_venda"] in vistos or not v.get("data_venda"):
            continue
        vistos.add(v["id_venda"])
        c = v.get("cod_cliente")
        if v.get("dados_cliente"):
            cad[c] = v["dados_cliente"]
        d = dt.date.fromisoformat(v["data_venda"][:10])
        for i in v.get("itens_venda") or []:
            if "cancel" in str(i.get("status") or "").lower():
                continue
            tipo = (i.get("tipo") or "").strip()
            cat = CAT_VENDA.get(str(v["id_venda"]), "Outros") if tipo == "Plano" else ("Produto" if tipo == "Produto" else categoria(i.get("desc_item")))
            vendas.append(dict(cli=c, d=d, e=UNI.get(e, e), tipo=tipo, desc=(i.get("desc_item") or "").strip(), cat=cat,
                               v=br(i.get("valor_liquido")), bruto=br(i.get("valor_bruto"))))
print("itens de venda", len(vendas), "clientes com cadastro", len(cad), "período", min(x["d"] for x in vendas), max(x["d"] for x in vendas))
print(collections.Counter(x["tipo"] for x in vendas).most_common())

# primeira transação de cada cliente (qualquer tipo) = porta de entrada
por_cli = collections.defaultdict(list)
for x in vendas:
    por_cli[x["cli"]].append(x)
INICIO_HIST = min(x["d"] for x in vendas) + dt.timedelta(days=180)  # evita "primeira" compra que é só início do histórico

entrada = {}
for c, l in por_cli.items():
    l.sort(key=lambda x: x["d"])
    d0 = l[0]["d"]
    if d0 < INICIO_HIST or d0 > HOJE - dt.timedelta(days=365):
        continue
    dia = [x for x in l if x["d"] == d0]
    plano_dia = [x for x in dia if x["tipo"] == "Plano"]
    entrada[c] = dict(d=d0, avulso=not plano_dia, itens=dia, depois=[x for x in l if x["d"] > d0])


def isca(it):
    """Nome curto do procedimento avulso de entrada."""
    s = " + ".join(sorted({x["cat"] for x in it}))
    return s


print("\n== PORTA DE ENTRADA (1ª compra entre", INICIO_HIST, "e", HOJE - dt.timedelta(days=365), ") ==")
grp = collections.defaultdict(list)
for c, x in entrada.items():
    if x["avulso"]:
        nomes = sorted({x2["desc"].upper() for x2 in x["itens"]})
        chave = "Avulso: " + isca(x["itens"])
    else:
        chave = "Plano direto: " + max((y for y in x["itens"] if y["tipo"] == "Plano"), key=lambda y: y["v"])["cat"]
    fim = x["d"] + dt.timedelta(days=365)
    plano12 = [y for y in x["depois"] if y["tipo"] == "Plano" and y["d"] <= fim]
    plano60 = [y for y in plano12 if (y["d"] - x["d"]).days <= 60]
    ltv = sum(y["v"] for y in x["itens"]) + sum(y["v"] for y in x["depois"] if y["d"] <= fim)
    grp[chave].append((sum(y["v"] for y in x["itens"]), ltv, bool(plano60), bool(plano12)))
print(f"{'entrada':48}{'n':>6}{'1º ticket':>10}{'LTV12 méd':>10}{'LTV12 med':>10}{'plano≤60d':>10}{'plano≤12m':>10}")
for k, v in sorted(grp.items(), key=lambda kv: -len(kv[1]))[:18]:
    n = len(v)
    print(f"{k[:48]:48}{n:6}{st.median(r[0] for r in v):10.0f}{st.mean(r[1] for r in v):10.0f}{st.median(r[1] for r in v):10.0f}"
          f"{pc(sum(r[2] for r in v), n):>10}{pc(sum(r[3] for r in v), n):>10}")

# iscas específicas (serviço avulso de entrada, pelo nome)
print("\n== ISCAS AVULSAS POR SERVIÇO ==")
g2 = collections.defaultdict(list)
for c, x in entrada.items():
    if not x["avulso"]:
        continue
    for y in {y["desc"].upper() for y in x["itens"] if y["tipo"] != "Produto"}:
        fim = x["d"] + dt.timedelta(days=365)
        plano12 = [z for z in x["depois"] if z["tipo"] == "Plano" and z["d"] <= fim]
        ltv = sum(z["v"] for z in x["itens"]) + sum(z["v"] for z in x["depois"] if z["d"] <= fim)
        g2[y].append((ltv, bool(plano12), bool([z for z in plano12 if (z["d"] - x["d"]).days <= 60]),
                      sum(z["v"] for z in x["itens"] if z["desc"].upper() == y)))
print(f"{'serviço de entrada':55}{'n':>6}{'preço med':>10}{'→plano 60d':>11}{'→plano 12m':>11}{'LTV12 méd':>10}")
for k, v in sorted(g2.items(), key=lambda kv: -len(kv[1]))[:25]:
    n = len(v)
    print(f"{k[:55]:55}{n:6}{st.median(r[3] for r in v):10.0f}{pc(sum(r[2] for r in v), n):>11}{pc(sum(r[1] for r in v), n):>11}{st.mean(r[0] for r in v):10.0f}")
