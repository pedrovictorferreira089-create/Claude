"""Cashback x Tabela de Descontos do Belle (tabela "Serviços", Petrópolis, prints de 30/10/2026).

Para cada 2ª compra de upgrade em até 60 dias da 1ª: desconto efetivo já dado (preço de tabela x preço final)
+ cashback proposto (10% do 1º plano, teto R$ 300), comparado com os limites da tabela.
"""
import collections
import datetime as dt
import json
import glob
import re
import statistics as st
from base import D, br, categoria, data, UNI

# codigoServico: (nome, desconto máx %, desconto máx sem senha %). Páginas 1-3 e início da 4 dos prints.
TABELA = {
    56280297: ("BIOESTIMULADOR PERFECT BUMBUM 300 ML", 27, 22), 56280296: ("BIOESTIMULADOR PERFECT BUMBUM 120 ML", 27, 22),
    56280295: ("BIOESTIMULADOR PERFECT BUMBUM 60 ML", 27, 22), 56280201: ("BOTOX DYSPORT 50 UI", 49, 42),
    55520726: ("TERMOTERAPIA", 30, 20), 24: ("TURBINA BUMBUM", 35, 25), 56280197: ("STIMULUSCULP LIPEDEMA", 35, 25),
    56261722: ("RADIOFREQUÊNCIA", 60, 50), 56252756: ("MICRO AGULHAMENTO COM GHK Cu", 35, 25),
    56227666: ("MICROROBOTICO MELLAN FREE", 62, 57), 56161623: ("PUMP UP", 35, 25), 56124704: ("ORTOSHOCK 1 HR", 40, 33),
    55558487: ("STIMULUSCULP", 35, 25), 55538446: ("METODO BBB", 30, 20), 28: ("RÁDIO FREQUENCIA FRACIONADA", 35, 30),
    15: ("MICROAGULHAMENTO CORPORAL ROBÓTICO", 29, 25), 56240772: ("LIPO S/ CORTE", 50, 40),
    56124706: ("LIPO ESCULTURA GESSADA 1 HR", 35, 30), 33353403: ("LIMPEZA DE PELE premium", 40, 20),
    34: ("HIDROLIPOCLASIA SEM CORTE", 40, 30), 56249914: ("CRIOLIFTING GORDURA", 40, 30),
    33448547: ("CRIOTURBO PEITORAL (HOMEM)", 50, 35), 33339420: ("DRENAGEM FACIAL", 50, 40),
    33337364: ("CRIOTURBO HARMONIZAÇÃO COMPLETA", 60, 50), 43: ("CRIOLIFTING FLACIDEZ", 40, 30), 17: ("DIÁSTASE OFF", 45, 35),
    10: ("CRIOTURBO HARMONIZAÇÃO PERNA", 60, 50), 9: ("CRIOTURBO HARMONIZAÇÃO BRAÇO", 60, 50), 3: ("CRIOTURBO ÁREA", 60, 50),
    56161621: ("BIOGLUTEOS", 27, 22), 56128219: ("BOTOX DYSPORT", 49, 42), 27: ("CARBOXTERAPIA PARA ESTRIAS", 30, 22),
}
UP = {"Crioturbo", "Stimulus/Glúteo", "Lipo Gessada/Modelagem", "Facial", "Diástase", "Enzimas/Injetáveis"}


def planos():
    out = {}
    for f in glob.glob(D + "venda_planos__*"):
        e = re.search(r"codEstab-(\d+)", f).group(1)
        if e == "4":
            continue
        for p in json.load(open(f)):
            if p["statusPlano"] != "Aprovado" or not p.get("dataVenda"):
                continue
            tab = cob = 0.0
            lim = collections.Counter()
            cats = set()
            for s in p["servicos"]:
                if s["cortesia"] == "Sim":
                    continue
                v = br(s["valorServico"]) * int(s.get("qtdSessoes") or 1)
                tab += v
                c = categoria(s["nomeServico"])
                if c != "Acessório":
                    cats.add(c)
                t = TABELA.get(int(s["codigoServico"] or 0))
                if t:
                    cob += v
                    lim["max"] += v * t[1]
                    lim["ss"] += v * t[2]
            out[p["codOrcamento"]] = dict(cli=int(p["cliente"].split(" - ")[0]), d=data(p["dataVenda"]), e=UNI[e],
                                          tab=tab, fim=br(p["precoFinal"]), cob=cob, lim=lim, cats=cats)
    return list(out.values())


def main():
    por = collections.defaultdict(lambda: collections.defaultdict(list))
    for p in planos():
        por[p["cli"]][p["d"]].append(p)
    casos = []
    for c, dias in por.items():
        ds = sorted(dias)
        if ds[0] < dt.date(2024, 10, 1):
            continue
        v1 = sum(p["fim"] for p in dias[ds[0]])
        if v1 <= 0:
            continue
        for d in ds[1:]:
            l = dias[d]
            if (d - ds[0]).days > 60 or not set().union(*(p["cats"] for p in l)) & UP:
                continue
            tab = sum(p["tab"] for p in l); fim = sum(p["fim"] for p in l); cob = sum(p["cob"] for p in l)
            if tab <= 0 or fim <= 0 or cob < 0.7 * tab:  # só compras com 70%+ do valor coberto pelos prints
                break
            lmax = sum(p["lim"]["max"] for p in l) / cob; lss = sum(p["lim"]["ss"] for p in l) / cob
            cb = min(0.1 * v1, 300)
            desc = 100 * (1 - fim / tab)
            casos.append(dict(desc=desc, cb=100 * cb / tab, tot=desc + 100 * cb / tab, max=lmax, ss=lss, fim=fim, tab=tab, cbv=cb))
            break
    n = len(casos)
    q = lambda a, p: sorted(a)[int(len(a) * p)]
    print(f"2ªs compras de upgrade (até 60 dias, 1ª compra desde out/24) com limites conhecidos: {n}")
    print(f"desconto efetivo já dado hoje: mediana {q([c['desc'] for c in casos], .5):.0f}%  p75 {q([c['desc'] for c in casos], .75):.0f}%")
    print(f"limite sem senha (média ponderada): mediana {q([c['ss'] for c in casos], .5):.0f}%  · máx: mediana {q([c['max'] for c in casos], .5):.0f}%")
    print(f"já acima do sem senha hoje: {100*sum(c['desc'] > c['ss'] + .5 for c in casos)/n:.0f}%  · já acima do máx: {100*sum(c['desc'] > c['max'] + .5 for c in casos)/n:.0f}%")
    print(f"cashback sobre o preço de tabela: mediana {q([c['cb'] for c in casos], .5):.1f}%")
    for nome, regra in [("10% teto 300", lambda c: c["cbv"]),
                        ("10% teto 300, máx 15% da nova", lambda c: min(c["cbv"], .15 * c["fim"])),
                        ("10% teto 300, só até o máx", lambda c: max(0, min(c["cbv"], c["tab"] * (c["max"] - c["desc"]) / 100)))]:
        tot = [c["desc"] + 100 * regra(c) / c["tab"] for c in casos]
        val = [regra(c) for c in casos]
        print(f"[{nome}] total > sem senha: {100*sum(t > c['ss'] + .5 for t, c in zip(tot, casos))/n:.0f}%  "
              f"total > máx: {100*sum(t > c['max'] + .5 for t, c in zip(tot, casos))/n:.0f}%  cashback mediano R$ {st.median(val):.0f}  custo total R$ {sum(val):,.0f}")
    print("desconto efetivo hoje por faixa:", collections.Counter(min(int(c['desc'] // 10) * 10, 60) for c in casos).most_common())


main()
