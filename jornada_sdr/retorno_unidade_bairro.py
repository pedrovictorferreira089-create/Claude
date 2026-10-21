"""Quais clientes dão mais retorno por unidade, bairro e faixa de idade, e quais procedimentos compram.

Base: planos aprovados dos últimos 24 meses (até HOJE de base.py) nas 4 unidades (sem a Laser). O valor de cada plano é
rateado entre os serviços pelo preço final, para somar por procedimento. Bairro, cidade e nascimento vêm do cadastro do
Belle (Vendas Detalhado). Gera Retorno_Unidade_Bairro.xlsx (só números agregados, sem nomes de clientes).
"""
import collections
import datetime as dt
import glob
import json
import re
import statistics as st
import unicodedata

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from base import D, HOJE, UNI, br, categoria, data

INI = HOJE - dt.timedelta(days=730)
CARMIM, CREME = "C60039", "F5E3D7"
FAIXAS = ["15 a 24", "25 a 34", "35 a 44", "45 a 54", "55 a 64", "65 ou mais", "Não informada"]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper().strip()
    return re.sub(r"\s+", " ", s)


def faixa(nasc):
    try:
        d = dt.datetime.strptime(nasc[:10], "%d/%m/%Y").date() if "/" in nasc else dt.date.fromisoformat(nasc[:10])
    except Exception:
        return "Não informada"
    a = (HOJE - d).days // 365
    for lo, hi, nome in ((15, 24, "15 a 24"), (25, 34, "25 a 34"), (35, 44, "35 a 44"), (45, 54, "45 a 54"), (55, 64, "55 a 64"), (65, 120, "65 ou mais")):
        if lo <= a <= hi:
            return nome
    return "Não informada"


# ---------------------------------------------------------------- cadastro
cad = {}
for f in glob.glob(D + "vendas_detalhado__*"):
    for v in json.load(open(f)):
        if v.get("dados_cliente") and v.get("cod_cliente"):
            cad[int(v["cod_cliente"])] = v["dados_cliente"]


def bairro(c):
    dc = cad.get(c) or {}
    b, cid = norm(dc.get("bairro")), norm(dc.get("cidade"))
    if not b or b in ("NULL", "-", "."):
        return "NÃO INFORMADO"
    return f"{b} ({cid.title() if cid and cid != 'NULL' else 'cidade não informada'})"


# ---------------------------------------------------------------- linhas de serviço (rateio do preço final)
linhas, planos = [], {}
for f in glob.glob(D + "venda_planos__*"):
    e = re.search(r"codEstab-(\d+)", f).group(1)
    if e == "4":
        continue
    for p in json.load(open(f)):
        d = data(p.get("dataVenda")) if p.get("dataVenda") else None
        if p["statusPlano"] != "Aprovado" or not d or d < INI or d > HOJE or not p.get("cliente") or " - " not in p["cliente"]:
            continue
        if p["codOrcamento"] in planos:
            continue
        cli = int(p["cliente"].split(" - ")[0])
        tot = br(p["precoFinal"])
        planos[p["codOrcamento"]] = dict(cli=cli, e=UNI[e], v=tot)
        bruto = sum(br(s["valorTotalServico"]) for s in p["servicos"] if s["cortesia"] != "Sim")
        for s in p["servicos"]:
            if s["cortesia"] == "Sim" or bruto <= 0:
                continue
            v = br(s["valorTotalServico"]) * tot / bruto
            if v > 0:
                linhas.append(dict(cli=cli, e=UNI[e], cat=categoria(s["nomeServico"]), serv=(s["nomeServico"] or "").strip().upper(), v=v))

gasto = collections.Counter()
for p in planos.values():
    gasto[p["cli"]] += p["v"]
vals = sorted(gasto.values(), reverse=True)
TOP20 = vals[int(len(vals) * 0.2) - 1]
RECEITA = sum(vals)
print(f"{len(planos)} planos, {len(gasto)} clientes, receita R$ {RECEITA:,.0f}, top 20% ≥ R$ {TOP20:,.0f}")


def bloco_clientes(gastos):
    """Indicadores de um grupo de clientes a partir do gasto de cada uma no grupo."""
    v = list(gastos.values())
    if not v:
        return [0, 0, 0, 0, 0, 0]
    tot = sum(v)
    return [len(v), round(st.mean(v), 2), round(st.median(v), 2), sum(1 for x in v if x >= TOP20) / len(v), round(tot, 2), tot / RECEITA]


def top_cats(ls, n=3):
    c = collections.Counter()
    for l in ls:
        c[l["cat"]] += l["v"]
    tot = sum(c.values()) or 1
    return [f"{k} ({100 * x / tot:.0f}%)" for k, x in c.most_common() if k != "Acessório"][:n]


def top_servs(ls, n=5):
    c = collections.Counter()
    for l in ls:
        c[l["serv"]] += l["v"]
    return [k.title() for k, _ in c.most_common() if categoria(k) != "Acessório"][:n]


# ---------------------------------------------------------------- agrupamentos
por_unidade = collections.defaultdict(collections.Counter)  # unidade -> cliente -> gasto na unidade
for p in planos.values():
    por_unidade[p["e"]][p["cli"]] += p["v"]
lin_unid = collections.defaultdict(list)
for l in linhas:
    lin_unid[l["e"]].append(l)

grupo_bairro = collections.defaultdict(collections.Counter)
grupo_faixa = collections.defaultdict(collections.Counter)
for c, v in gasto.items():
    grupo_bairro[bairro(c)][c] = v
    grupo_faixa[faixa((cad.get(c) or {}).get("dtNascimento") or "")][c] = v
lin_bairro, lin_faixa = collections.defaultdict(list), collections.defaultdict(list)
for l in linhas:
    lin_bairro[bairro(l["cli"])].append(l)
    lin_faixa[faixa((cad.get(l["cli"]) or {}).get("dtNascimento") or "")].append(l)

# ---------------------------------------------------------------- planilha
wb = Workbook()
H = Font(bold=True, color="FFFFFF")
FILL = PatternFill("solid", fgColor=CARMIM)
ALT = PatternFill("solid", fgColor=CREME)
BRL = 'R$ #,##0;-R$ #,##0;"-"'
PCT = '0.0%;-0.0%;"-"'


def aba(ws, titulo, nota, cols, rows, fmts, larg):
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=14, color=CARMIM)
    ws["A2"] = nota
    ws["A2"].font = Font(italic=True, size=9, color="555555")
    ws.append([])
    ws.append(cols)
    for c in ws[4]:
        c.font, c.fill = H, FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows):
        ws.append(r)
        for j, c in enumerate(ws[ws.max_row]):
            if fmts[j]:
                c.number_format = fmts[j]
            if i % 2:
                c.fill = ALT
    for i, w in enumerate(larg, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 32
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{ws.max_row}"


NOTA = (f"Planos aprovados de {INI:%d/%m/%Y} a {HOJE:%d/%m/%Y}, unidades Petrópolis, Lagoa Nova, Capim Macio e Norte Shopping. "
        f"\"Top 20%\" = clientes que gastaram R$ {TOP20:,.0f} ou mais no período. Procedimentos pelo valor rateado de cada serviço do plano.").replace(",", ".")
COLS = ["Clientes", "Gasto médio por cliente", "Gasto mediano", "% no top 20%", "Receita no período", "% da receita"]
FM = [None, BRL, BRL, PCT, BRL, PCT]

# 1. Unidade
rows = []
for u in ("Petrópolis", "Lagoa Nova", "Capim Macio", "Norte Shopping"):
    g = por_unidade[u]
    fx = collections.Counter(faixa((cad.get(c) or {}).get("dtNascimento") or "") for c in g)
    fx.pop("Não informada", None)
    rows.append([u] + bloco_clientes(g) + [", ".join(top_cats(lin_unid[u])), ", ".join(top_servs(lin_unid[u])),
                                           ", ".join(f"{k} ({100 * n / sum(fx.values()):.0f}%)" for k, n in fx.most_common(2))])
aba(wb.active, "Retorno por unidade", NOTA + " Gasto da cliente na própria unidade.",
    ["Unidade"] + COLS + ["Categorias que mais vendem (% da receita)", "5 procedimentos que mais vendem", "Faixas de idade mais comuns"],
    rows, [None] + FM + [None, None, None], [18, 10, 14, 12, 10, 16, 10, 48, 70, 34])
wb.active.title = "Por unidade"

# 2. Faixa de idade
rows = []
for k in FAIXAS:
    g = grupo_faixa.get(k)
    if g:
        rows.append([k] + bloco_clientes(g) + [", ".join(top_cats(lin_faixa[k])), ", ".join(top_servs(lin_faixa[k]))])
aba(wb.create_sheet("Por idade"), "Retorno por faixa de idade", NOTA, ["Faixa de idade"] + COLS +
    ["Categorias que mais compram (% do gasto)", "5 procedimentos que mais compram"], rows, [None] + FM + [None, None],
    [16, 10, 14, 12, 10, 16, 10, 48, 70])

# 2b. Idade x unidade
rows = []
for k in FAIXAS[:-1]:
    r = [k]
    for u in ("Petrópolis", "Lagoa Nova", "Capim Macio", "Norte Shopping"):
        g = [v for c, v in por_unidade[u].items() if faixa((cad.get(c) or {}).get("dtNascimento") or "") == k]
        r += [len(g), round(st.mean(g), 2) if g else 0]
    rows.append(r)
aba(wb.create_sheet("Idade x unidade"), "Clientes e gasto médio por faixa de idade em cada unidade", NOTA + " Gasto da cliente na própria unidade.",
    ["Faixa de idade", "Petrópolis clientes", "Petrópolis gasto médio", "Lagoa Nova clientes", "Lagoa Nova gasto médio",
     "Capim Macio clientes", "Capim Macio gasto médio", "Norte Shopping clientes", "Norte Shopping gasto médio"],
    rows, [None, None, BRL, None, BRL, None, BRL, None, BRL], [16, 11, 13, 11, 13, 11, 13, 13, 15])

# 3. Bairro (20+ clientes)
rows = []
for b, g in sorted(grupo_bairro.items(), key=lambda kv: -sum(kv[1].values())):
    if len(g) < 20:
        continue
    fx = collections.Counter(faixa((cad.get(c) or {}).get("dtNascimento") or "") for c in g)
    fx.pop("Não informada", None)
    tot = sum(fx.values()) or 1
    mais45 = sum(n for k, n in fx.items() if k in ("45 a 54", "55 a 64", "65 ou mais")) / tot
    pred = fx.most_common(1)[0][0] if fx else "Não informada"
    rows.append([b] + bloco_clientes(g) + [mais45, pred, ", ".join(top_cats(lin_bairro[b])), ", ".join(top_servs(lin_bairro[b], 3))])
aba(wb.create_sheet("Por bairro"), "Retorno por bairro da cliente (bairros com 20 clientes ou mais)",
    NOTA + " Ordenado pela receita. Bairro e cidade do cadastro do Belle.",
    ["Bairro (cidade)"] + COLS + ["% com 45 anos ou mais", "Faixa de idade mais comum", "Categorias que mais compram (% do gasto)", "3 procedimentos que mais compram"],
    rows, [None] + FM + [PCT, None, None, None], [38, 10, 14, 12, 10, 16, 10, 12, 14, 48, 60])

# 4. Bairro x unidade (onde as clientes de cada bairro compram)
cli_unid = collections.defaultdict(collections.Counter)
for p in planos.values():
    cli_unid[p["cli"]][p["e"]] += p["v"]
rows = []
for b, g in sorted(grupo_bairro.items(), key=lambda kv: -sum(kv[1].values())):
    if len(g) < 20:
        continue
    u = collections.Counter()
    for c in g:
        u.update(cli_unid[c])
    tot = sum(u.values()) or 1
    rows.append([b] + [u[x] / tot for x in ("Petrópolis", "Lagoa Nova", "Capim Macio", "Norte Shopping")] + [round(tot, 2)])
aba(wb.create_sheet("Bairro x unidade"), "Em qual unidade as clientes de cada bairro compram",
    NOTA + " % da receita de cada bairro por unidade.",
    ["Bairro (cidade)", "Petrópolis", "Lagoa Nova", "Capim Macio", "Norte Shopping", "Receita do bairro"],
    rows, [None, PCT, PCT, PCT, PCT, BRL], [38, 13, 13, 13, 15, 16])

# 5. Procedimentos x unidade
cats = collections.Counter()
for l in linhas:
    if l["cat"] != "Acessório":
        cats[l["cat"]] += l["v"]
m = collections.defaultdict(collections.Counter)
mc = collections.defaultdict(lambda: collections.defaultdict(set))
for l in linhas:
    m[l["cat"]][l["e"]] += l["v"]
    mc[l["cat"]][l["e"]].add(l["cli"])
rows = []
for k, _ in cats.most_common():
    r = [k]
    for u in ("Petrópolis", "Lagoa Nova", "Capim Macio", "Norte Shopping"):
        r += [round(m[k][u], 2), len(mc[k][u])]
    r.append(round(sum(m[k].values()), 2))
    rows.append(r)
aba(wb.create_sheet("Procedimentos x unidade"), "Receita e clientes por procedimento e unidade", NOTA,
    ["Procedimento", "Petrópolis R$", "Petrópolis clientes", "Lagoa Nova R$", "Lagoa Nova clientes", "Capim Macio R$",
     "Capim Macio clientes", "Norte Shopping R$", "Norte Shopping clientes", "Total R$"],
    rows, [None, BRL, None, BRL, None, BRL, None, BRL, None, BRL], [26, 14, 11, 14, 11, 14, 11, 14, 11, 15])

# 6. Procedimentos x idade
m = collections.defaultdict(collections.Counter)
for l in linhas:
    m[l["cat"]][faixa((cad.get(l["cli"]) or {}).get("dtNascimento") or "")] += l["v"]
tot_fx = collections.Counter()
for k in m:
    tot_fx.update(m[k])
rows = []
for k, _ in cats.most_common():
    rows.append([k] + [m[k][f] / tot_fx[f] if tot_fx[f] else 0 for f in FAIXAS[:-1]])
aba(wb.create_sheet("Procedimentos x idade"), "Peso de cada procedimento no gasto de cada faixa de idade", NOTA + " Cada coluna soma 100%.",
    ["Procedimento"] + FAIXAS[:-1], rows, [None] + [PCT] * 6, [26, 11, 11, 11, 11, 11, 12])

wb.save("/home/user/Claude/Retorno_Unidade_Bairro.xlsx")
print("ok")
