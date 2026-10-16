"""Agenda: inativos por faixa, reavaliação (acompanhamento estético) e 2ª compra; poder de compra por bairro/profissão."""
import collections
import datetime as dt
import glob
import json
import re
import statistics as st

from base import D, HOJE, compras_por_cliente, data, planos

ps = planos()
C = compras_por_cliente(ps)


def pc(x, n):
    return f"{100 * x / n:.1f}%" if n else "-"


ag, vistos = [], set()
for f in glob.glob(D + "relatorios_relatorio_atendimentos__*"):
    if "codEstab-4" in f:
        continue
    for a in json.load(open(f)):
        if a["idAgendamento"] in vistos:
            continue
        vistos.add(a["idAgendamento"])
        d = data(a.get("dataAgendamento"))
        if not d or not a.get("codigoCliente"):
            continue
        ag.append(dict(cli=int(a["codigoCliente"]), d=d, st=(a.get("statusAgendamento") or "").strip(), tipo=a.get("tipoAgendamento") or "",
                       sc=str(a.get("codigoServico") or ""), serv=(a.get("nomeServico") or "").strip().upper(), orc=a.get("idOrcamento")))
print("agendamentos", len(ag), min(a["d"] for a in ag), max(a["d"] for a in ag))
INI_AG = min(a["d"] for a in ag)
atend = collections.defaultdict(list)
for a in ag:
    if a["st"] == "Atendido" and a["d"] <= HOJE:
        atend[a["cli"]].append(a["d"])
for l in atend.values():
    l.sort()

# ------------------------------------------------------------------ inativos: retorno nos 60 dias seguintes, por faixa
print("\n== INATIVOS: o que acontece nos 60 dias seguintes ==")
FX = [(30, 59), (60, 89), (90, 179), (180, 364)]
res = collections.defaultdict(lambda: [0, 0, 0, 0.0])
fotos = []
t = dt.date(INI_AG.year + (INI_AG.month + 5) // 12, (INI_AG.month + 5) % 12 + 1, 1)  # 6 meses após o início da agenda
while t <= HOJE - dt.timedelta(days=60):
    fotos.append(t)
    t = dt.date(t.year + t.month // 12, t.month % 12 + 1, 1)
import bisect  # noqa: E402

for c, l in atend.items():
    compras = C.get(c, [])
    for t in fotos:
        i = bisect.bisect_left(l, t)
        if i == 0:
            continue
        dias = (t - l[i - 1]).days
        fx = next((f for f in FX if f[0] <= dias <= f[1]), None)
        if not fx:
            continue
        voltou = i < len(l) and (l[i] - t).days <= 60
        comprou = [x for x in compras if t <= x["d"] < t + dt.timedelta(days=60)]
        r = res[fx]
        r[0] += 1
        r[1] += voltou
        r[2] += bool(comprou)
        r[3] += sum(x["v"] for x in comprou)
print(f"{'dias sem vir':14}{'casos':>8}{'voltou 60d':>11}{'comprou plano 60d':>18}{'R$/cliente':>11}")
for fx in FX:
    r = res[fx]
    print(f"{f'{fx[0]}-{fx[1]}':14}{r[0]:8}{pc(r[1], r[0]):>11}{pc(r[2], r[0]):>18}{r[3] / r[0]:11.0f}")

# ------------------------------------------------------------------ reavaliação / acompanhamento estético
print("\n== ACOMPANHAMENTO ESTÉTICO / RETORNO x COMPRA ==")
REAV = lambda a: a["sc"] == "33351798" or "ACOMPANHAM" in a["serv"] or "REAVALIA" in a["serv"] or a["tipo"] == "Retorno"  # noqa: E731
reav = [a for a in ag if REAV(a) and a["st"] == "Atendido" and a["d"] <= HOJE - dt.timedelta(days=30)]
print("acompanhamentos/retornos atendidos", len(reav))
for janela in (0, 7, 30):
    n = sum(1 for a in reav if any(0 <= (x["d"] - a["d"]).days <= janela for x in C.get(a["cli"], [])))
    print(f" compra de plano em até {janela} dias após o acompanhamento: {pc(n, len(reav))}")
# quando acontece o acompanhamento em relação à 1ª compra
pos = []
for a in reav:
    l = C.get(a["cli"], [])
    antes = [x for x in l if x["d"] < a["d"]]
    if antes:
        pos.append((a["d"] - antes[-1]["d"]).days)
print(" dias desde a última compra até o acompanhamento: mediana", st.median(pos), "quartis", st.quantiles(pos, n=4))
fx = collections.defaultdict(lambda: [0, 0])
for a in reav:
    l = C.get(a["cli"], [])
    antes = [x for x in l if x["d"] < a["d"]]
    if not antes:
        k = "sem compra antes"
    else:
        dd = (a["d"] - antes[-1]["d"]).days
        k = next(f"{lo}-{hi}d" for lo, hi in ((1, 20), (21, 40), (41, 60), (61, 90), (91, 180), (181, 99999)) if lo <= dd <= hi)
    fx[k][0] += 1
    fx[k][1] += any(0 <= (x["d"] - a["d"]).days <= 7 for x in l)
for k in ("1-20d", "21-40d", "41-60d", "61-90d", "91-180d", "181-99999d", "sem compra antes"):
    print(f"  acompanhamento {k:18} n={fx[k][0]:5}  comprou em 7d: {pc(fx[k][1], fx[k][0])}")

# clientes com 1ª compra no período da agenda: fez acompanhamento até 60d? taxa de 2ª compra em 90d
coorte = {c: l for c, l in C.items() if INI_AG + dt.timedelta(days=30) <= l[0]["d"] <= HOJE - dt.timedelta(days=90)}
com, sem = [0, 0], [0, 0]
for c, l in coorte.items():
    d0 = l[0]["d"]
print("coorte p/ efeito do acompanhamento:", len(coorte))
idx = collections.defaultdict(list)
for a in ag:
    if REAV(a) and a["st"] == "Atendido":
        idx[a["cli"]].append(a["d"])
for c, l in coorte.items():
    d0 = l[0]["d"]
    fez = any(14 <= (x - d0).days <= 60 for x in idx.get(c, []))
    seg = any(0 < (x["d"] - d0).days <= 90 for x in l[1:])
    alvo = com if fez else sem
    alvo[0] += 1
    alvo[1] += seg
print(f" fez acompanhamento entre 14 e 60 dias após a 1ª compra: n={com[0]}, 2ª compra em 90d {pc(com[1], com[0])}")
print(f" não fez: n={sem[0]}, 2ª compra em 90d {pc(sem[1], sem[0])}")

# ------------------------------------------------------------------ porta de entrada pela agenda (inclui sessões gratuitas)
print("\n== 1ª SESSÃO ATENDIDA (porta de entrada pela agenda) ==")
primeira = {}
for a in sorted(ag, key=lambda a: a["d"]):
    if a["st"] == "Atendido" and a["cli"] not in primeira:
        primeira[a["cli"]] = a
corte_ini = INI_AG + dt.timedelta(days=180)
g = collections.defaultdict(list)
dia_ag = collections.defaultdict(list)
for a in ag:
    if a["st"] == "Atendido":
        dia_ag[(a["cli"], a["d"])].append(a)
for c, a in primeira.items():
    if not corte_ini <= a["d"] <= HOJE - dt.timedelta(days=365):
        continue
    l = C.get(c, [])
    antes = [x for x in l if x["d"] < a["d"]]
    if antes:  # já era cliente de plano antes do início da agenda
        continue
    do_dia = dia_ag[(c, a["d"])]
    tem_av = any(x["tipo"] == "Avaliação" or x["sc"] == "52" for x in do_dia)
    servs = sorted({x["serv"] for x in do_dia if x["tipo"] == "Serviço" and x["sc"] != "52"})
    if tem_av and servs:
        k = "Avaliação + " + (servs[0] if len(servs) == 1 else "vários")
    elif tem_av:
        k = "Só avaliação"
    elif servs:
        k = (servs[0] if len(servs) == 1 else "vários serviços")
    else:
        k = "Retorno/consulta"
    plano_proprio = any(x["orc"] for x in do_dia)  # sessão já dentro de um plano = comprou antes/no dia
    fim = a["d"] + dt.timedelta(days=365)
    p0 = [x for x in l if x["d"] == a["d"]]
    p60 = [x for x in l if a["d"] <= x["d"] <= a["d"] + dt.timedelta(days=60)]
    p12 = [x for x in l if a["d"] <= x["d"] <= fim]
    g[k].append((bool(p0), bool(p60), bool(p12), sum(x["v"] for x in p12), plano_proprio))
print(f"{'1ª sessão':60}{'n':>6}{'plano no dia':>13}{'plano 60d':>10}{'plano 12m':>10}{'R$ planos 12m méd':>18}")
for k, v in sorted(g.items(), key=lambda kv: -len(kv[1]))[:30]:
    n = len(v)
    if n < 15:
        continue
    print(f"{k[:60]:60}{n:6}{pc(sum(r[0] for r in v), n):>13}{pc(sum(r[1] for r in v), n):>10}{pc(sum(r[2] for r in v), n):>10}{st.mean(r[3] for r in v):18.0f}")
