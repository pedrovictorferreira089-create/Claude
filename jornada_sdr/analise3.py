"""Poder de compra: gasto de 24 meses por bairro, profissão e porta de entrada; perfil das clientes de maior valor."""
import collections
import datetime as dt
import glob
import json
import re
import statistics as st
import unicodedata

from base import D, HOJE, compras_por_cliente, planos

ps = planos()
C = compras_por_cliente(ps)
INI = HOJE - dt.timedelta(days=730)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper().strip()
    return re.sub(r"\s+", " ", s)


cad = {}
for f in glob.glob(D + "vendas_detalhado__*"):
    for v in json.load(open(f)):
        if v.get("dados_cliente") and v.get("cod_cliente"):
            cad[int(v["cod_cliente"])] = v["dados_cliente"]

gasto = {c: sum(x["v"] for x in l if x["d"] >= INI) for c, l in C.items()}
gasto = {c: v for c, v in gasto.items() if v > 0}
vals = sorted(gasto.values(), reverse=True)
CORTE_ALTO = vals[int(len(vals) * 0.2) - 1]  # top 20%
print(f"clientes com plano em 24m: {len(gasto)}; top 20% gasta ≥ R$ {CORTE_ALTO:,.0f}; mediana R$ {st.median(vals):,.0f}")

PROF = [("Saúde (médica/dentista/fisio/enfermeira)", r"MEDIC|DENTIST|ODONTO|FISIOTER|ENFERM|NUTRICION|FARMAC|BIOMED|PSICOL|FONO"),
        ("Direito/Contabilidade/Gestão", r"ADVOG|JUIZ|PROMOTOR|DELEGAD|CONTAD|ADMINISTR|GEREN|DIRETOR|ECONOMIS|AUDITOR|BANCAR"),
        ("Empresária/Autônoma", r"EMPRES|EMPREEND|AUTONOM|COMERCIANT|LOJIST|CORRETOR|VENDEDOR|CONSULTOR"),
        ("Servidora pública/Militar", r"SERVIDOR|FUNCIONARIA PUBLICA|POLICIAL|MILITAR|BOMBEIR|FISCAL|ANALISTA JUD|TECNICO JUD"),
        ("Professora/Educação", r"PROFESS|PEDAGOG|EDUCAD|COORDENADOR"),
        ("Engenharia/Arquitetura/TI", r"ENGENH|ARQUITET|ANALISTA|PROGRAMAD|DESENVOLV|TI\b"),
        ("Estética/Beleza", r"ESTETIC|CABELEIR|MANICURE|DESIGNER DE|MAQUIAD|BELEZA"),
        ("Estudante", r"ESTUDANT|UNIVERSIT"),
        ("Do lar/Aposentada", r"DO LAR|DONA DE CASA|APOSENT|PENSION"),
        ("Atendimento/Operacional", r"ATENDENT|RECEPCION|AUXILIAR|ASSISTENT|OPERADOR|CAIXA|SECRETAR|DIARISTA|COZINH|MOTORIST")]


def prof(p):
    p = norm(p)
    if not p or p in ("NULL", "-", "."):
        return "Não informada"
    return next((k for k, rx in PROF if re.search(rx, p)), "Outras")


def tabela(titulo, chave, minimo=30, top=25):
    g = collections.defaultdict(list)
    for c, v in gasto.items():
        g[chave(c)].append(v)
    print(f"\n== {titulo} ==")
    print(f"{'grupo':42}{'clientes':>9}{'gasto med':>10}{'gasto méd':>10}{'% top20':>9}{'% receita':>10}")
    tot = sum(gasto.values())
    for k, v in sorted(g.items(), key=lambda kv: -st.mean(kv[1]))[: top * 3]:
        if len(v) < minimo:
            continue
        print(f"{str(k)[:42]:42}{len(v):9}{st.median(v):10.0f}{st.mean(v):10.0f}{100 * sum(x >= CORTE_ALTO for x in v) / len(v):8.1f}%{100 * sum(v) / tot:9.1f}%")


tabela("BAIRRO (cadastro do Belle)", lambda c: norm((cad.get(c) or {}).get("bairro")) or "NÃO INFORMADO", minimo=40)
tabela("PROFISSÃO (agrupada)", lambda c: prof((cad.get(c) or {}).get("profissao")), minimo=20)
tabela("CIDADE", lambda c: norm((cad.get(c) or {}).get("cidade")) or "NÃO INFORMADA", minimo=40)
tabela("1ª CATEGORIA COMPRADA", lambda c: C[c][0]["cat"], minimo=30)
tabela("UNIDADE DA 1ª COMPRA", lambda c: C[c][0]["e"], minimo=30)


def idade(c):
    n = (cad.get(c) or {}).get("dtNascimento")
    try:
        d = dt.datetime.strptime(n[:10], "%d/%m/%Y").date() if "/" in n else dt.date.fromisoformat(n[:10])
    except Exception:
        return "Não informada"
    a = (HOJE - d).days // 365
    return next((f"{lo}-{hi}" for lo, hi in ((15, 24), (25, 34), (35, 44), (45, 54), (55, 64), (65, 120)) if lo <= a <= hi), "Não informada")


tabela("FAIXA DE IDADE", idade, minimo=30)
