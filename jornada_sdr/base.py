"""Base comum da análise de jornada: planos aprovados e categorias de procedimento."""
import collections
import os
import datetime as dt
import glob
import json
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from belle_sabados import br  # noqa: E402

D = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data") + "/"
HOJE = dt.date(2026, 10, 16)
UNI = {"1": "Petrópolis", "2": "Lagoa Nova", "3": "Capim Macio", "6": "Norte Shopping", "4": "Laser"}

REGRAS = [  # ordem importa
    ("Acessório", r"CONSULTA NUTRI|BIOIMPED|AVALIA"),
    ("Pós-Crioturbo", r"P[ÓO]S.?(TURBO|CRIO|48)"),
    ("Crioturbo", r"CRIOTURBO|HARMONIZA[ÇC][ÃA]O CRIO|REFINAMENTO|PR[ÉE] CRIO"),
    ("Drenagem", r"DRENAGEM M[ÉE]TODO|M[ÉE]TODO DRENESSE|^DRENAGEM(?! FACIAL)|LINF[ÁA]TICA"),
    ("Stimulus/Glúteo", r"STIMUL|BUMBUM|GL[ÚU]TEO|BIOGLUTEO|GLUTEOMAX|BBB"),
    ("Diástase", r"DI[ÁA]STASE"),
    ("Lipo Gessada/Modelagem", r"LIPO|ORTOSHOCK|SLIM|ELETROLIPO|REDU[C]?T|FIRM|CELU|CAVITA|ENDERMO|TERMOTERAPIA|POWER SHAPE|PUMP UP|ESTRIAS"),
    ("Enzimas/Injetáveis", r"ENZIMA|BIOESTIMULADOR|BOTOX|PEIM|INJET"),
    ("Laser", r"LASER|PASSE LIVRE"),
    ("Facial", r"PELE|PEELING|FACIAL|FACE|MELLAN|MICRO ?AGULHA|LINE SKIN|CRIOLIFTING|CRIOFREQ|R[ÁA]DIO|PLASMA|GLOW|STRIORT|CARBOX"),
]


def categoria(nome):
    n = (nome or "").upper()
    for cat, rx in REGRAS:
        if re.search(rx, n):
            return cat
    return "Outros"


def data(s):
    return dt.datetime.strptime(s[:10], "%d/%m/%Y").date() if s else None


def planos():
    """Planos aprovados: cliente, data, unidade, valor, categoria principal, saldo de sessões."""
    out = []
    for f in glob.glob(D + "venda_planos__*"):
        e = re.search(r"codEstab-(\d+)", f).group(1)
        for p in json.load(open(f)):
            if p["statusPlano"] != "Aprovado" or not p.get("dataVenda"):
                continue
            cli = int(p["cliente"].split(" - ")[0])
            val = collections.Counter()
            for s in p["servicos"]:
                v = 0 if s["cortesia"] == "Sim" else br(s["valorTotalServico"])
                val[categoria(s["nomeServico"])] += v
            princ = [c for c, _ in val.most_common() if c != "Acessório"]
            restante = sum(int(s.get("qtdSessaoRestante") or 0) for s in p["saldoPlano"] if categoria(s["nomeServico"]) != "Acessório")
            out.append(dict(orc=p["codOrcamento"], cli=cli, d=data(p["dataVenda"]), e=UNI[e], v=br(p["precoFinal"]),
                            cat=princ[0] if princ else "Acessório", cats=set(princ), nome=p["nomePlano"],
                            restante=restante, vendedor=(p.get("vendedor") or "").strip()))
    # o mesmo plano pode vir em duas janelas: deduplica pelo número do orçamento
    return list({p["orc"]: p for p in out}.values())


def compras_por_cliente(ps):
    """Compras agrupadas por cliente e dia (vários planos no mesmo dia = uma compra)."""
    por = collections.defaultdict(lambda: collections.defaultdict(list))
    for p in ps:
        por[p["cli"]][p["d"]].append(p)
    out = {}
    for c, dias in por.items():
        out[c] = [dict(d=d, v=sum(p["v"] for p in l), cats=set().union(*(p["cats"] for p in l)),
                       cat=max(l, key=lambda p: p["v"])["cat"], e=l[0]["e"], planos=l) for d, l in sorted(dias.items())]
    return out
