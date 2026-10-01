"""Vendas de avaliação (clientes novos) x vendas de cabine (clientes recorrentes) – API Belle Software.

Uso: python3 belle_avaliacao.py [dd/mm/aaaa início] [dd/mm/aaaa fim] [codEstab ...]
     padrão: 01/01/2026 a 26/09/2026, todas as unidades. Gera Avaliacao_x_Cabine_Drenesse.xlsx (ou BELLE_OUT).

Avaliação = agendamento avulso (sem plano vinculado) de um dos serviços de entrada de cliente novo (AVALIACAO).
Venda de avaliação = venda a cliente novo feita até BELLE_JANELA dias depois da avaliação atendida (padrão 0: no mesmo
dia), sem plano aprovado comprado entre a avaliação e a venda. Todas as outras vendas são de cabine.
Comparecimento = atendidos ÷ (atendidos + faltas). Fechamento da avaliação = avaliações atendidas (cliente/dia) que
compraram plano ÷ avaliações atendidas; da cabine = clientes atendidos na cabine no mês que compraram plano de cabine
no mesmo mês ÷ clientes atendidos na cabine no mês.
"""
import collections
import datetime as dt
import os
import sys

from openpyxl import Workbook

from belle_auditoria import DATA, MES, NUM, PCT, base, compactar_xlsx, data_br, linha, nome, secao, titulo
from belle_sabados import BRL, FB, FI, HOJE, UNIDADES, br, cabecalho, carregar_planos, historico, larguras

# Serviços com que os agendamentos avulsos de clientes novos são lançados no Belle (informados pela Drenesse).
AVALIACAO = {
    22: "DRENAGEM MÉTODO DRENESSE",
    56210744: "SESSÃO EXPERIMENTAL DRENAGEM - MÉTODO DRENESSE",
    56260425: "DRENAGEM MÉTODO DRENESSE 98,70",
    33353403: "LIMPEZA DE PELE PREMIUM",
    56210746: "SESSÃO EXPERIMENTAL - STIMULUS",
    56210745: "SESSÃO EXPERIMENTAL - DIÁSTASE",
}
JANELA = int(os.environ.get("BELLE_JANELA", "0"))  # dias após a avaliação em que a venda ainda conta como de avaliação
ANTES = 90  # dias antes do início buscados só para achar a avaliação das vendas do começo do período
AV, CAB, SEM = "Avaliação", "Cabine", "Sem serviço"
CAB_PLANO = "Cabine (sessão de plano)"
STATUS = {"Atendido": "Atendido", "Falhou": "Falta", "Desmarcado": "Desmarcado", "Cancelado": "Cancelado"}
P_DIA = "Comprou no dia da avaliação"
P_JANELA = f"Comprou até {JANELA} dias após a avaliação"
P_NOVO = "Cliente novo – comprou " + ("depois do dia da avaliação" if JANELA == 0 else f"mais de {JANELA} dias após a avaliação")
P_REC = "Cliente recorrente (já tinha plano)"
P_SEM = "Sem avaliação registrada e sem plano anterior"
PERFIS = [(AV, P_DIA)] + ([(AV, P_JANELA)] if JANELA else []) + [(CAB, P_NOVO), (CAB, P_REC), (CAB, P_SEM)]


def codigo(x):
    s = str(x if x is not None else "").strip()
    return int(s) if s.isdigit() else s


def campo(d, *chaves):
    return next((d[k] for k in chaves if d.get(k) not in (None, "")), None)


def reais(v):
    return f"R$ {v:,.2f}".replace(",", "x").replace(".", ",").replace("x", ".")


def meses(ini, fim):
    m, out = ini.replace(day=1), []
    while m <= fim:
        out.append(m)
        m = (m + dt.timedelta(days=32)).replace(day=1)
    return out


def coletar(ini, fim, estabs):
    desde = ini - dt.timedelta(days=ANTES)
    desde = dt.date(desde.year, 3 * ((desde.month - 1) // 3) + 1, 1)  # janelas trimestrais da API
    planos = carregar_planos()
    nomes, compras, vendas = {}, collections.defaultdict(list), []
    for orc, p in planos.items():
        r = p["raw"]
        cliente = r.get("cliente") or ""
        c = codigo(campo(r, "codigoCliente", "codCliente") or cliente.split(" - ", 1)[0])
        nomes.setdefault(c, nome(cliente.split(" - ", 1)[-1]))
        d = data_br(r.get("dataVenda"))
        if not d:
            continue
        if p["status"] == "Aprovado":  # histórico completo de compras do cliente, em todas as unidades
            compras[c].append((d, p["nome"]))
        if p["estab"] in estabs and ini <= d <= fim:
            vendas.append(dict(unidade=UNIDADES[p["estab"]], data=d, cliente_cod=c, cliente=nomes[c], tipo="Plano",
                               desc=nome(p["nome"]), orc=orc, status=p["status"], vendedor=nome(r.get("vendedor")),
                               valor=br(r.get("precoFinal")), conta=p["status"] == "Aprovado",
                               obs="" if p["status"] == "Aprovado" else f"Plano {p['status']} – fora do faturamento"))
    for lst in compras.values():
        lst.sort()

    agenda, sem_servico = [], collections.Counter()
    for e in estabs:
        for a in historico("relatorios/relatorio_atendimentos", "codEstab", e, fim, desde=desde):
            d, c = data_br(a.get("dataAgendamento")), codigo(a.get("codigoCliente"))
            if not d or not desde <= d <= fim or c in ("", 0):
                continue
            sc, orc = codigo(a.get("codigoServico")), a.get("idOrcamento") or ""
            grupo = SEM if sc in ("", 0) else AV if sc in AVALIACAO and not orc else CAB
            if grupo == SEM and ini <= d:
                sem_servico[a.get("tipoAgendamento") or "(sem tipo)"] += 1
            nomes.setdefault(c, nome(a.get("nomeCliente")))
            st = (a.get("statusAgendamento") or "").strip()
            agenda.append(dict(unidade=UNIDADES[e], data=d, mes=d.replace(day=1), hora=a.get("horarioAgendamento") or "",
                               cliente_cod=c, cliente=nomes[c], servico_cod=sc,
                               servico=nome(a.get("nomeServico")) or nome(a.get("tipoAgendamento")), orc=orc,
                               status_belle=st, status=STATUS.get(st, "Outros"), prof=nome(a.get("nomeProfissional")),
                               grupo=grupo))
    agenda.sort(key=lambda a: (a["data"], a["hora"], a["unidade"], a["cliente"]))

    # Avaliação realizada = cliente + dia com avaliação atendida (dois serviços de avaliação no mesmo dia contam uma vez).
    visitas = {}
    for a in agenda:
        if a["grupo"] == AV and a["status"] == "Atendido":
            k = (a["cliente_cod"], a["data"])
            if k in visitas:
                visitas[k]["outros"].append(a["servico"])
            else:
                visitas[k] = dict(a, outros=[], vendas=[])
    por_cliente = collections.defaultdict(list)
    for v in sorted(visitas.values(), key=lambda v: v["data"]):
        por_cliente[v["cliente_cod"]].append(v)

    def classificar(c, d):
        """Grupo, perfil, avaliação ligada e dias desde ela, para uma venda ao cliente c no dia d."""
        av = next((v for v in reversed(por_cliente.get(c, [])) if v["data"] <= d), None)
        antes = [x for x, _ in compras.get(c, []) if x < d]
        if av and not any(x >= av["data"] for x in antes):  # primeira compra desde a avaliação
            dias = (d - av["data"]).days
            if dias <= JANELA:
                return AV, (P_DIA if dias == 0 else P_JANELA), av, dias
            if any(x < av["data"] for x in antes):  # já tinha plano antes da avaliação: não é cliente novo
                return CAB, P_REC, None, None
            return CAB, P_NOVO, av, dias
        return CAB, (P_REC if antes else P_SEM), None, None

    tipos, ids_planos = collections.Counter(), {str(p["raw"]["idVenda"]) for p in planos.values() if p["raw"].get("idVenda")}
    for e in estabs:
        for v in historico("vendas_detalhado", "estab", e, fim, desde=ini):
            d = dt.date.fromisoformat(v["data_venda"][:10]) if v.get("data_venda") else None
            if not d or not ini <= d <= fim or "cancel" in str(campo(v, "status", "situacao") or "").casefold():
                continue
            c = codigo(v.get("cod_cliente"))
            idv = campo(v, "id_venda", "idVenda", "cod_venda", "codigo_venda", "codVenda")
            com_plano = idv is not None and str(idv) in ids_planos
            for i in v.get("itens_venda") or []:
                tipo = (i.get("tipo") or "").strip()
                tipos[tipo or "(sem tipo)"] += 1
                if "plano" in tipo.casefold() or "pacote" in tipo.casefold():
                    continue  # o plano já entra pelo relatório de planos, com status e preço final
                repetido = com_plano and tipo == "Serviço"
                vendas.append(dict(unidade=UNIDADES[e], data=d, cliente_cod=c,
                                   cliente=nomes.get(c) or nome(campo(v, "nome_cliente", "cliente", "nomeCliente")),
                                   tipo="Serviço avulso" if tipo == "Serviço" else (tipo or "Item"), desc=nome(i.get("desc_item")),
                                   orc="", status="", vendedor=nome(campo(v, "vendedor", "nome_vendedor", "nomeVendedor")),
                                   valor=br(i.get("valor_liquido")), conta=not repetido,
                                   obs="Serviço da mesma venda de um plano – já contado no preço do plano" if repetido else ""))
    vendas.sort(key=lambda x: (x["data"], x["unidade"], x["cliente"], x["tipo"] != "Plano"))
    for x in vendas:
        x["grupo"], x["perfil"], av, x["dias"] = classificar(x["cliente_cod"], x["data"])
        x["av_data"], x["av_serv"] = (av["data"], av["servico"]) if av else (None, "")
        if av and x["grupo"] == AV and x["tipo"] == "Plano" and x["conta"]:
            av["vendas"].append(x)

    for v in visitas.values():
        cmp = compras.get(v["cliente_cod"], [])
        prox = [x for x in cmp if v["data"] <= x[0] <= fim]
        v["ja_cliente"] = any(x < v["data"] for x, _ in cmp)
        v["dias_compra"] = (prox[0][0] - v["data"]).days if prox else None
        v["primeira"] = ", ".join(n for x, n in prox if x == prox[0][0]) if prox else ""

    periodo = [a for a in agenda if a["data"] >= ini]
    cab_mes = collections.Counter((a["unidade"], a["mes"], a["cliente_cod"]) for a in periodo
                                  if a["grupo"] == CAB and a["status"] == "Atendido")
    compra_cab = collections.defaultdict(list)
    for x in vendas:
        if x["grupo"] == CAB and x["tipo"] == "Plano" and x["conta"]:
            compra_cab[(x["cliente_cod"], x["data"].replace(day=1))].append(x["valor"])
    resumo = collections.Counter((a["unidade"], a["mes"], a["grupo"], a["status"]) for a in periodo)

    diag = dict(sem_servico=sem_servico, tipos=tipos,
                aval_plano=sum(1 for a in periodo if a["servico_cod"] in AVALIACAO and a["grupo"] == CAB),
                repetidos=sum(1 for x in vendas if not x["conta"] and x["tipo"] != "Plano"))
    return dict(vendas=vendas, visitas=sorted((v for v in visitas.values() if v["data"] >= ini),
                                              key=lambda v: (v["data"], v["hora"], v["unidade"])),
                agenda_av=[a for a in periodo if a["servico_cod"] in AVALIACAO], cab_mes=cab_mes, compra_cab=compra_cab,
                resumo=resumo, diag=diag, nomes=nomes)


# ---------------------------------------------------------------- planilha
ABAS = {"V": "Vendas", "AV": "Avaliações", "AA": "Agenda avaliações", "AR": "Agenda resumo", "CC": "Clientes cabine"}
DIM = {"unidade": {"AR": "A", "AV": "A", "CC": "A", "V": "A"}, "mes": {"AR": "B", "AV": "C", "CC": "B", "V": "C"}}


def exemplos(d, fim):
    """Casos reais de cada situação, para mostrar como cada atendimento e cada venda foi classificado."""
    rec = lambda lst, n: list(reversed(lst))[:n]  # os mais recentes
    vis, ag, vd = d["visitas"], d["agenda_av"], d["vendas"]
    out = []

    def add(caso, unidade, data, cliente, servico, situacao, grupo, porque):
        out.append([caso, unidade, data, cliente, servico, situacao, grupo, porque])

    for v in rec([v for v in vis if v["vendas"]], 3):
        add("Avaliação que fechou no dia", v["unidade"], v["data"], v["cliente"],
            f"{v['servico']} → {', '.join(x['desc'] for x in v['vendas'])}",
            f"Atendido • plano {reais(sum(x['valor'] for x in v['vendas']))}", AV,
            f"Agendamento avulso com ID {v['servico_cod']} (cliente novo) e plano comprado no mesmo dia"
            + (f" (vendedor: {v['vendas'][0]['vendedor']})" if v["vendas"][0]["vendedor"] else "") + ".")
    for v in rec([v for v in vis if v["dias_compra"] is None and not v["ja_cliente"] and v["data"] <= fim - dt.timedelta(days=30)], 2):
        add("Avaliação sem fechamento", v["unidade"], v["data"], v["cliente"], v["servico"], "Atendido • não comprou plano",
            AV, f"Compareceu à avaliação e não comprou plano até {fim:%d/%m/%Y}: entra na base do fechamento, não no fechou.")
    for a in rec([a for a in ag if a["grupo"] == AV and a["status"] == "Falta"], 2):
        add("Falta na avaliação", a["unidade"], a["data"], a["cliente"], a["servico"], "Falhou (não compareceu)", AV,
            f"Agendado com ID {a['servico_cod']} e não veio: reduz a taxa de comparecimento da avaliação.")
    for x in rec([x for x in vd if x["perfil"] == P_NOVO and x["tipo"] == "Plano"], 2):
        add("Cliente novo que fechou depois", x["unidade"], x["data"], x["cliente"], x["desc"], f"Plano {reais(x['valor'])}",
            CAB, f"Fez a avaliação em {x['av_data']:%d/%m/%Y} ({x['av_serv']}) e comprou {x['dias']} dias depois, já na "
                 f"cabine: venda de cabine.")
    for x in rec([x for x in vd if x["perfil"] == P_REC and x["tipo"] == "Plano" and x["conta"]], 3):
        add("Cliente recorrente comprando na cabine", x["unidade"], x["data"], x["cliente"], x["desc"],
            f"Plano {reais(x['valor'])}", CAB, "Já tinha plano aprovado comprado antes: renovação ou novo plano na cabine.")
    for a in rec([a for a in ag if a["grupo"] == CAB], 2):
        add("ID de avaliação dentro de plano", a["unidade"], a["data"], a["cliente"], a["servico"], a["status_belle"],
            CAB_PLANO, f"Serviço com ID de avaliação, mas vinculado ao plano nº {a['orc']}: é sessão de cliente recorrente.")
    for v in rec([v for v in vis if v["ja_cliente"]], 2):
        add("Avaliação de quem já tinha plano", v["unidade"], v["data"], v["cliente"], v["servico"], "Atendido", AV,
            "Agendado avulso com ID de avaliação, mas já tinha comprado plano antes: fica na avaliação (regra dos IDs) e "
            "aparece marcado na aba Avaliações.")
    for x in rec([x for x in vd if x["perfil"] == P_SEM and x["tipo"] == "Plano" and x["conta"]], 2):
        add("Venda sem avaliação registrada", x["unidade"], x["data"], x["cliente"], x["desc"], f"Plano {reais(x['valor'])}",
            CAB, "Primeira compra do cliente sem agendamento de avaliação (nem nos 90 dias antes do período): fica na cabine.")
    for x in rec([x for x in vd if x["grupo"] == AV and x["tipo"] != "Plano" and x["conta"]], 2):
        add("Pagamento avulso no dia da avaliação", x["unidade"], x["data"], x["cliente"], x["desc"],
            f"{x['tipo']} {reais(x['valor'])}", AV, "Valor pago à parte no dia da avaliação: entra no faturamento avulso da avaliação.")
    return out


def montar(d, ini, fim, estabs, destino):
    wb = Workbook()
    unids = [UNIDADES[e] for e in estabs]
    lista_meses = meses(ini, fim)
    vendas, visitas, agenda_av = d["vendas"], d["visitas"], d["agenda_av"]
    resumo_ag = sorted({(u, m, g) for (u, m, g, _) in d["resumo"]}, key=lambda x: (unids.index(x[0]), x[1], x[2]))
    cab = sorted(d["cab_mes"].items(), key=lambda x: (unids.index(x[0][0]), x[0][1], x[0][2]))
    n = {"V": len(vendas), "AV": len(visitas), "AA": len(agenda_av), "AR": len(resumo_ag), "CC": len(cab)}

    def rg(b, col):
        return f"'{ABAS[b]}'!${col}$2:${col}${max(n[b], 1) + 1}"

    regra = "no dia da avaliação" if JANELA == 0 else f"em até {JANELA} dias após a avaliação"
    fonte = (f"Período: {ini:%d/%m/%Y} a {fim:%d/%m/%Y} • Unidades: {', '.join(unids)} • Fonte: API Belle Software, extraído em "
             f"{HOJE:%d/%m/%Y} • Venda de avaliação = cliente novo que comprou {regra}")

    # ---------------- Resumo
    rs = wb.active
    rs.title = "Resumo"
    titulo(rs, "Drenesse – Vendas de Avaliação (clientes novos) x Vendas de Cabine (clientes recorrentes)", fonte)
    secao(rs, 4, "1. Comparativo geral")
    h = 5
    cabecalho(rs, h, ["Indicador", "Avaliação (clientes novos)", "Cabine (clientes recorrentes)", "Total", "Como é calculado"])

    def grp(g):
        return {"agend": f"SUMIFS({rg('AR', 'I')},{rg('AR', 'C')},\"{g}\")",
                "atend": f"SUMIFS({rg('AR', 'D')},{rg('AR', 'C')},\"{g}\")",
                "faltas": f"SUMIFS({rg('AR', 'E')},{rg('AR', 'C')},\"{g}\")",
                "desm": f"SUMIFS({rg('AR', 'F')},{rg('AR', 'C')},\"{g}\")+SUMIFS({rg('AR', 'G')},{rg('AR', 'C')},\"{g}\")",
                "base": f"COUNTA({rg('AV', 'A')})" if g == AV else f"COUNTA({rg('CC', 'A')})",
                "fech": f"COUNTIF({rg('AV', 'K')},\"Sim\")" if g == AV else f"COUNTIF({rg('CC', 'F')},\"Sim\")",
                "planos": f"COUNTIFS({rg('V', 'F')},\"Plano\",{rg('V', 'L')},\"{g}\",{rg('V', 'Q')},\"Sim\")",
                "fatpl": f"SUMIFS({rg('V', 'K')},{rg('V', 'F')},\"Plano\",{rg('V', 'L')},\"{g}\",{rg('V', 'Q')},\"Sim\")",
                "avulso": f"SUMIFS({rg('V', 'K')},{rg('V', 'F')},\"<>Plano\",{rg('V', 'L')},\"{g}\",{rg('V', 'Q')},\"Sim\")",
                "naoapr": f"COUNTIFS({rg('V', 'F')},\"Plano\",{rg('V', 'L')},\"{g}\",{rg('V', 'Q')},\"Não\")"}
    ga, gc = grp(AV), grp(CAB)
    r1 = h + 1  # linha de cada indicador: r1 + posição
    ind = [
        ("Agendamentos", "agend", NUM, "soma", "Todos os agendamentos do período (qualquer status). Avaliação = agendamento avulso com um dos 6 IDs."),
        ("Atendidos", "atend", NUM, "soma", "Status 'Atendido' no Belle."),
        ("Faltas (não compareceram)", "faltas", NUM, "soma", "Status 'Falhou' no Belle."),
        ("Desmarcados / cancelados", "desm", NUM, "soma", "Status 'Desmarcado' ou 'Cancelado' (não entram na taxa de comparecimento)."),
        ("Taxa de comparecimento", f"=IFERROR({{c}}{r1 + 1}/({{c}}{r1 + 1}+{{c}}{r1 + 2}),0)", PCT, "taxa", "Atendidos ÷ (atendidos + faltas)."),
        ("Comparecimento sobre todos os agendamentos", f"=IFERROR({{c}}{r1 + 1}/{{c}}{r1},0)", PCT, "taxa", "Atendidos ÷ agendamentos (inclui desmarcados e cancelados)."),
        ("Base do fechamento", "base", NUM, "-", "Avaliação: avaliações atendidas (cliente/dia). Cabine: clientes atendidos na cabine em cada mês (cliente/mês)."),
        ("Fecharam (compraram plano)", "fech", NUM, "-", f"Avaliação: compraram plano {regra}. Cabine: compraram plano de cabine no mesmo mês em que foram atendidos."),
        ("Taxa de fechamento", f"=IFERROR({{c}}{r1 + 7}/{{c}}{r1 + 6},0)", PCT, "-", "Fecharam ÷ base do fechamento."),
        ("Planos vendidos (aprovados)", "planos", NUM, "soma", "Planos com status Aprovado vendidos no período (data da venda)."),
        ("Faturamento em planos (R$)", "fatpl", BRL, "soma", "Preço final dos planos aprovados."),
        ("Ticket médio do plano (R$)", f"=IFERROR({{c}}{r1 + 10}/{{c}}{r1 + 9},0)", BRL, "taxa", "Faturamento em planos ÷ planos vendidos."),
        ("Faturamento avulso (R$)", "avulso", BRL, "soma", "Sessões e produtos vendidos à parte (Vendas Detalhado, valor líquido)."),
        ("Faturamento total (R$)", f"={{c}}{r1 + 10}+{{c}}{r1 + 12}", BRL, "soma", "Planos + avulso."),
        ("% do faturamento total", f"=IFERROR({{c}}{r1 + 13}/$D${r1 + 13},0)", PCT, "taxa", "Participação de cada grupo no faturamento total."),
        ("Planos não aprovados (fora do faturamento)", "naoapr", NUM, "soma", "Planos Pendentes/Suspensos/Cancelados vendidos no período – só informativo."),
    ]
    for i, (lbl, chave, fmt, tot, expl) in enumerate(ind):
        r = r1 + i
        if chave in ga:
            vb, vc = "=" + ga[chave], "=" + gc[chave]
        else:
            vb, vc = chave.format(c="B"), chave.format(c="C")
        vd = f"=B{r}+C{r}" if tot == "soma" else chave.format(c="D") if tot == "taxa" else "—"
        linha(rs, r, [lbl, vb, vc, vd, expl], [None, fmt, fmt, fmt, None])
    r = r1 + len(ind) - 1

    r += 2
    secao(rs, r, "2. Avaliação – fechamento por prazo depois da avaliação")
    r += 1
    cabecalho(rs, r, ["Prazo", "Avaliações realizadas", "Compraram plano", "Taxa de fechamento", "Observação"])
    tb = f"COUNTA({rg('AV', 'A')})"
    for lbl, col, obs in [(f"{regra.capitalize()} (venda de avaliação)", "K", "É a taxa do comparativo acima."),
                          ("Em até 7 dias", "P", "Inclui quem comprou no dia. Compras depois do dia da avaliação contam como cabine."),
                          ("Em até 30 dias", "Q", f"Avaliações do fim do período têm menos de 30 dias até {fim:%d/%m/%Y}.")]:
        r += 1
        linha(rs, r, [lbl, f"={tb}", f"=COUNTIF({rg('AV', col)},\"Sim\")", f"=IFERROR(C{r}/B{r},0)", obs], [None, NUM, NUM, PCT, None])
    r += 1
    linha(rs, r, ["Avaliações de clientes que já tinham plano antes", f"={tb}", f"=COUNTIF({rg('AV', 'J')},\"Sim\")",
                  f"=IFERROR(C{r}/B{r},0)", "Agendados com ID de avaliação, mas já eram clientes: ficam na avaliação (regra dos IDs)."],
          [None, NUM, NUM, PCT, None])

    r += 2
    secao(rs, r, "3. De onde vêm as vendas – perfil do cliente no momento da compra")
    r += 1
    hp = r
    cabecalho(rs, r, ["Perfil da venda", "Grupo", "Planos vendidos", "Faturamento em planos (R$)", "Faturamento avulso (R$)",
                      "Faturamento total (R$)", "% do faturamento total"])
    tp = hp + len(PERFIS) + 1
    for g, perfil in PERFIS:
        r += 1
        c = f"{rg('V', 'M')},$A{r},{rg('V', 'Q')},\"Sim\""
        linha(rs, r, [perfil, g, f"=COUNTIFS({c},{rg('V', 'F')},\"Plano\")", f"=SUMIFS({rg('V', 'K')},{c},{rg('V', 'F')},\"Plano\")",
                      f"=SUMIFS({rg('V', 'K')},{c},{rg('V', 'F')},\"<>Plano\")", f"=D{r}+E{r}", f"=IFERROR(F{r}/$F${tp},0)"],
              [None, None, NUM, BRL, BRL, BRL, PCT])
    r += 1
    linha(rs, r, ["TOTAL", ""] + [f"=SUM({c}{hp + 1}:{c}{r - 1})" for c in "CDEFG"], [None, None, NUM, BRL, BRL, BRL, PCT], total=True)

    def bloco(ws, r, grupo, dim, chaves, fmt_chave):
        """Tabela de um grupo por unidade ou por mês. Devolve a linha do total."""
        aval = grupo == AV
        cabecalho(ws, r, ["Unidade" if dim == "unidade" else "Mês", "Agendamentos", "Atendidos", "Faltas", "Desmarcados / cancelados",
                          "Taxa de comparecimento",
                          "Avaliações realizadas (cliente/dia)" if aval else "Clientes atendidos (cliente/mês)",
                          "Avaliações que fecharam" if aval else "Clientes que compraram plano", "Taxa de fechamento",
                          "Planos vendidos", "Faturamento em planos (R$)", "Ticket médio do plano (R$)", "Faturamento avulso (R$)",
                          "Faturamento total (R$)"])
        ws.row_dimensions[r].height = 45
        h0, g = r, f'"{grupo}"'
        fm = [fmt_chave, NUM, NUM, NUM, NUM, PCT, NUM, NUM, PCT, NUM, BRL, BRL, BRL, BRL]
        for k in chaves:
            r += 1

            def cr(b):
                return f"{rg(b, DIM[dim][b])},$A{r}"
            ar = f"{rg('AR', 'C')},{g},{cr('AR')}"
            bf = f"COUNTIFS({cr('AV')})" if aval else f"COUNTIFS({cr('CC')})"
            ff = f"COUNTIFS({cr('AV')},{rg('AV', 'K')},\"Sim\")" if aval else f"COUNTIFS({cr('CC')},{rg('CC', 'F')},\"Sim\")"
            vv = f"{cr('V')},{rg('V', 'L')},{g},{rg('V', 'Q')},\"Sim\""
            linha(ws, r, [k, f"=SUMIFS({rg('AR', 'I')},{ar})", f"=SUMIFS({rg('AR', 'D')},{ar})", f"=SUMIFS({rg('AR', 'E')},{ar})",
                          f"=SUMIFS({rg('AR', 'F')},{ar})+SUMIFS({rg('AR', 'G')},{ar})", f"=IFERROR(C{r}/(C{r}+D{r}),0)",
                          f"={bf}", f"={ff}", f"=IFERROR(H{r}/G{r},0)", f"=COUNTIFS({vv},{rg('V', 'F')},\"Plano\")",
                          f"=SUMIFS({rg('V', 'K')},{vv},{rg('V', 'F')},\"Plano\")", f"=IFERROR(K{r}/J{r},0)",
                          f"=SUMIFS({rg('V', 'K')},{vv},{rg('V', 'F')},\"<>Plano\")", f"=K{r}+M{r}"], fm)
        r += 1
        s = {c: f"=SUM({c}{h0 + 1}:{c}{r - 1})" for c in "BCDEGHJKMN"}
        linha(ws, r, ["TOTAL", s["B"], s["C"], s["D"], s["E"], f"=IFERROR(C{r}/(C{r}+D{r}),0)", s["G"], s["H"],
                      f"=IFERROR(H{r}/G{r},0)", s["J"], s["K"], f"=IFERROR(K{r}/J{r},0)", s["M"], s["N"]],
              [None] + fm[1:], total=True)
        return r

    r += 2
    secao(rs, r, "4a. Avaliação (clientes novos) por unidade")
    r = bloco(rs, r + 1, AV, "unidade", unids, None)
    r += 2
    secao(rs, r, "4b. Cabine (clientes recorrentes) por unidade")
    r = bloco(rs, r + 1, CAB, "unidade", unids, None)
    r += 2
    notas = [
        "COMO LER",
        "• Avaliação = agendamento avulso (sem plano vinculado) com um dos 6 IDs de cliente novo (lista na aba Metodologia). Cabine = todos os outros agendamentos.",
        f"• Venda de avaliação = primeira compra do cliente novo feita {regra}. Compras dele em dias posteriores e as de clientes que já tinham plano são vendas de cabine.",
        "• Fechamento da cabine é medido por cliente/mês: dos clientes atendidos na cabine no mês, quantos compraram plano de cabine no mesmo mês.",
        "• Os números são fórmulas sobre as abas de base (Vendas, Avaliações, Agenda resumo, Clientes cabine): filtre essas abas para ver cada atendimento e cada venda.",
        "• Passo a passo completo e diagnóstico dos dados na aba Metodologia; casos reais na aba Exemplos.",
    ]
    for i, t in enumerate(notas):
        c = rs.cell(row=r + i, column=1, value=t)
        if i == 0:
            c.font = FB
    larguras(rs, [44, 17, 17, 14, 16, 15, 17, 16, 13, 12, 17, 15, 16, 17])
    rs.column_dimensions["E"].width = 16
    rs.row_dimensions[h].height = rs.row_dimensions[hp].height = 30
    for rr in range(r1, r1 + len(ind)):  # coluna de explicação do comparativo, sem quebrar as larguras dos blocos
        rs.cell(row=rr, column=5).font = FI

    # ---------------- Mensal
    wm = wb.create_sheet("Mensal")
    titulo(wm, "Avaliação x Cabine – mês a mês", fonte)
    secao(wm, 4, "A. Avaliação (clientes novos)")
    r = bloco(wm, 5, AV, "mes", lista_meses, MES)
    r += 2
    secao(wm, r, "B. Cabine (clientes recorrentes)")
    bloco(wm, r + 1, CAB, "mes", lista_meses, MES)
    larguras(wm, [11, 14, 12, 10, 14, 14, 16, 14, 13, 11, 17, 15, 16, 17])
    wm.freeze_panes = "B6"

    # ---------------- Serviços de avaliação
    wsv = wb.create_sheet("Serviços de avaliação")
    titulo(wsv, "Avaliação por serviço de entrada – comparecimento e fechamento", fonte)
    cabecalho(wsv, 4, ["ID do serviço", "Serviço", "Agendamentos", "Atendidos", "Faltas", "Desmarcados / cancelados",
                       "Taxa de comparecimento", "Avaliações realizadas (cliente/dia)", "Fecharam", "Taxa de fechamento",
                       "Planos vendidos", "Valor vendido (R$)", "Ticket médio (R$)"])
    fms = ["0", None, NUM, NUM, NUM, NUM, PCT, NUM, NUM, PCT, NUM, BRL, BRL]
    nomes_srv = {}
    for a in agenda_av:
        nomes_srv.setdefault(a["servico_cod"], a["servico"])
    r = 4
    for cod_s, nm in AVALIACAO.items():
        r += 1
        aa = f"{rg('AA', 'G')},$A{r},{rg('AA', 'L')},\"{AV}\""
        av_ = f"{rg('AV', 'F')},$A{r}"
        linha(wsv, r, [cod_s, nomes_srv.get(cod_s, nm), f"=COUNTIFS({aa})", f"=COUNTIFS({aa},{rg('AA', 'K')},\"Atendido\")",
                       f"=COUNTIFS({aa},{rg('AA', 'K')},\"Falta\")",
                       f"=COUNTIFS({aa},{rg('AA', 'K')},\"Desmarcado\")+COUNTIFS({aa},{rg('AA', 'K')},\"Cancelado\")",
                       f"=IFERROR(D{r}/(D{r}+E{r}),0)", f"=COUNTIFS({av_})", f"=COUNTIFS({av_},{rg('AV', 'K')},\"Sim\")",
                       f"=IFERROR(I{r}/H{r},0)", f"=SUMIFS({rg('AV', 'L')},{av_})", f"=SUMIFS({rg('AV', 'M')},{av_})",
                       f"=IFERROR(L{r}/K{r},0)"], fms)
    r += 1
    s = {c: f"=SUM({c}5:{c}{r - 1})" for c in "CDEFHIKL"}
    linha(wsv, r, ["", "TOTAL", s["C"], s["D"], s["E"], s["F"], f"=IFERROR(D{r}/(D{r}+E{r}),0)", s["H"], s["I"],
                   f"=IFERROR(I{r}/H{r},0)", s["K"], s["L"], f"=IFERROR(L{r}/K{r},0)"], fms, total=True)
    r += 2
    wsv.cell(row=r, column=1, value="Sessões com esses IDs vinculadas a plano (contadas como cabine, não como avaliação):").font = FB
    c = wsv.cell(row=r, column=8, value=f"=COUNTIFS({rg('AA', 'L')},\"{CAB_PLANO}\")")
    c.number_format = NUM
    wsv.cell(row=r + 1, column=1, value="Avaliação realizada = cliente + dia com avaliação atendida; quando o cliente fez dois serviços de "
             "avaliação no mesmo dia, conta no primeiro (o outro aparece em 'Outros serviços no dia').").font = FI
    larguras(wsv, [11, 46, 13, 11, 9, 14, 14, 16, 10, 13, 10, 16, 14])
    wsv.row_dimensions[4].height = 45

    # ---------------- Exemplos
    we = wb.create_sheet("Exemplos")
    titulo(we, "Exemplos reais de atendimentos e vendas – como cada caso foi classificado", fonte)
    cabecalho(we, 4, ["Caso", "Unidade", "Data", "Cliente", "Serviço / plano", "Situação no Belle", "Contado como", "Por quê"])
    r = 4
    for vals in exemplos(d, fim):
        r += 1
        linha(we, r, vals, [None, None, DATA, None, None, None, None, None])
    if r == 4:
        we.cell(row=5, column=1, value="Sem atendimentos no período.").font = FI
    larguras(we, [30, 15, 11, 32, 48, 26, 22, 90])
    we.freeze_panes = "A5"

    # ---------------- Metodologia
    wt = wb.create_sheet("Metodologia")
    titulo(wt, "Passo a passo – como os números foram obtidos", fonte)
    passos = [
        "1. Dados extraídos da API do Belle, por unidade, em janelas de até 3 meses (limite da API):",
        "   • Relatório de Atendimentos (agenda): data, cliente, serviço, status e plano vinculado de cada agendamento.",
        "   • Venda de Planos: data da venda, cliente, plano, status, vendedor e preço final de cada plano (todo o histórico, todas as unidades).",
        "   • Vendas Detalhado: itens vendidos à parte (sessões avulsas e produtos), com valor líquido.",
        f"   • A agenda também foi lida nos {ANTES} dias anteriores ao início do período, só para achar a avaliação de vendas do começo do período.",
        "2. Cada agendamento foi classificado:",
        "   • Avaliação = serviço com um dos 6 IDs abaixo E sem plano vinculado (agendamento avulso = cliente novo).",
        "   • Cabine = qualquer outro agendamento com serviço, inclusive sessões com esses IDs que estão dentro de um plano (aba Agenda avaliações, 'Contado como').",
        "   • Agendamentos sem serviço (bloqueios, compromissos) ficam fora das duas contas.",
        "3. Comparecimento = atendidos ÷ (atendidos + faltas), usando o status do Belle: Atendido, Falhou (falta), Desmarcado, Cancelado.",
        "4. Avaliação realizada = cliente + dia com avaliação atendida. Dois serviços de avaliação no mesmo dia contam como uma avaliação.",
        "5. Cada venda (plano ou item avulso) foi ligada ao cliente e à data:",
        f"   • Venda de avaliação: o cliente fez avaliação atendida e esta é a primeira compra desde ela, feita {regra}.",
        "   • Venda de cabine: todas as outras – clientes que já tinham plano (recorrentes), clientes novos que compraram em outro dia e compras sem avaliação registrada.",
        "6. Fechamento:",
        f"   • Avaliação: avaliações realizadas em que o cliente comprou plano {regra} ÷ avaliações realizadas. Também mostrado em até 7 e 30 dias.",
        "   • Cabine: para cada mês, clientes atendidos na cabine que compraram plano de cabine no mesmo mês ÷ clientes atendidos na cabine no mês.",
        "7. Faturamento = valor vendido na data da venda: preço final dos planos com status Aprovado + valor líquido dos itens avulsos.",
        "   Planos Pendentes/Suspensos/Cancelados ficam listados na aba Vendas, mas fora do faturamento e do fechamento.",
        "8. Todas as contas do Resumo e do Mensal são fórmulas sobre as abas de base: dá para conferir filtrando as abas Vendas, Avaliações e Agenda.",
    ]
    r = 4
    for t in passos:
        c = wt.cell(row=r, column=1, value=t)
        if t[0].isdigit():
            c.font = FB
        r += 1
    r += 1
    secao(wt, r, "IDs de serviço considerados avaliação (cliente novo), informados pela Drenesse")
    r += 1
    cabecalho(wt, r, ["ID", "Serviço"])
    for cod_s, nm in AVALIACAO.items():
        r += 1
        linha(wt, r, [cod_s, nm], ["0", None])
    diag = d["diag"]
    r += 2
    secao(wt, r, "Diagnóstico dos dados")
    linhas_diag = [
        ("Sessões com ID de avaliação vinculadas a plano (contadas como cabine)", diag["aval_plano"]),
        ("Itens avulsos da mesma venda de um plano (fora do faturamento para não contar 2 vezes)", diag["repetidos"]),
    ] + [(f"Agendamentos sem serviço – tipo '{k}' (fora da análise)", q) for k, q in diag["sem_servico"].most_common()] \
      + [(f"Itens no Vendas Detalhado – tipo '{k}'", q) for k, q in diag["tipos"].most_common()]
    for lbl, q in linhas_diag:
        r += 1
        linha(wt, r, [lbl, q], [None, NUM])
    larguras(wt, [130, 14])

    # ---------------- Bases
    base(wb.create_sheet(ABAS["V"]),
         ["Unidade", "Data da venda", "Mês", "Cliente", "Cód. cliente", "Tipo", "Descrição", "Nº plano", "Status do plano", "Vendedor",
          "Valor (R$)", "Grupo", "Perfil da venda", "Data da avaliação", "Dias após a avaliação", "Serviço da avaliação",
          "Conta no faturamento?", "Observação"],
         [[x["unidade"], x["data"], x["data"].replace(day=1), x["cliente"], x["cliente_cod"], x["tipo"], x["desc"], x["orc"] or None,
           x["status"] or None, x["vendedor"] or None, x["valor"], x["grupo"], x["perfil"], x["av_data"], x["dias"],
           x["av_serv"] or None, "Sim" if x["conta"] else "Não", x["obs"] or None] for x in vendas],
         [None, DATA, MES, None, "0", None, None, "0", None, None, BRL, None, None, DATA, NUM, None, None, None], "TabVendas",
         [15, 11, 10, 32, 10, 14, 34, 10, 12, 26, 13, 11, 40, 11, 10, 34, 11, 50])
    base(wb.create_sheet(ABAS["AV"]),
         ["Unidade", "Data", "Mês", "Cliente", "Cód. cliente", "ID do serviço", "Serviço de avaliação", "Outros serviços no dia",
          "Profissional", "Já tinha plano antes?", "Fechou plano na avaliação?", "Planos vendidos na avaliação",
          "Valor vendido na avaliação (R$)", "Vendedor", "Dias até a 1ª compra de plano", "Comprou em até 7 dias?",
          "Comprou em até 30 dias?", "1ª compra de plano após a avaliação"],
         [[v["unidade"], v["data"], v["mes"], v["cliente"], v["cliente_cod"], v["servico_cod"], v["servico"],
           ", ".join(v["outros"]) or None, v["prof"] or None, "Sim" if v["ja_cliente"] else "Não", "Sim" if v["vendas"] else "Não",
           len(v["vendas"]), round(sum(x["valor"] for x in v["vendas"]), 2),
           ", ".join(sorted({x["vendedor"] for x in v["vendas"] if x["vendedor"]})) or None, v["dias_compra"],
           "Sim" if v["dias_compra"] is not None and v["dias_compra"] <= 7 else "Não",
           "Sim" if v["dias_compra"] is not None and v["dias_compra"] <= 30 else "Não", v["primeira"] or None] for v in visitas],
         [None, DATA, MES, None, "0", "0", None, None, None, None, None, NUM, BRL, None, NUM, None, None, None], "TabAvaliacoes",
         [15, 11, 10, 32, 10, 11, 36, 26, 26, 10, 11, 10, 13, 26, 11, 10, 10, 34])
    base(wb.create_sheet(ABAS["AA"]),
         ["Unidade", "Data", "Mês", "Hora", "Cliente", "Cód. cliente", "ID do serviço", "Serviço", "Nº plano", "Status no Belle",
          "Status", "Contado como", "Profissional"],
         [[a["unidade"], a["data"], a["mes"], a["hora"], a["cliente"], a["cliente_cod"], a["servico_cod"], a["servico"],
           a["orc"] or None, a["status_belle"], a["status"], AV if a["grupo"] == AV else CAB_PLANO, a["prof"] or None]
          for a in agenda_av],
         [None, DATA, MES, None, None, "0", "0", None, "0", None, None, None, None], "TabAgendaAval",
         [15, 11, 10, 7, 32, 10, 11, 40, 10, 12, 11, 22, 28])
    res = d["resumo"]
    base(wb.create_sheet(ABAS["AR"]),
         ["Unidade", "Mês", "Grupo", "Atendidos", "Faltas", "Desmarcados", "Cancelados", "Outros status", "Total de agendamentos"],
         [[u, m, g, res[(u, m, g, "Atendido")], res[(u, m, g, "Falta")], res[(u, m, g, "Desmarcado")],
           res[(u, m, g, "Cancelado")], res[(u, m, g, "Outros")], f"=SUM(D{i}:H{i})"] for i, (u, m, g) in enumerate(resumo_ag, 2)],
         [None, MES, None, NUM, NUM, NUM, NUM, NUM, NUM], "TabAgendaResumo", [15, 11, 13, 11, 10, 12, 11, 11, 13])
    base(wb.create_sheet(ABAS["CC"]),
         ["Unidade", "Mês", "Cliente", "Cód. cliente", "Sessões atendidas na cabine", "Comprou plano na cabine no mês?",
          "Planos comprados no mês", "Valor comprado no mês (R$)"],
         [[u, m, d["nomes"].get(c, ""), c, q,
           "Sim" if d["compra_cab"].get((c, m)) else "Não", len(d["compra_cab"].get((c, m), [])),
           round(sum(d["compra_cab"].get((c, m), [])), 2)] for (u, m, c), q in cab],
         [None, MES, None, "0", NUM, None, NUM, BRL], "TabClientesCabine", [15, 11, 32, 10, 12, 13, 11, 14])

    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)
    compactar_xlsx(destino, {ABAS["AR"]: {"I"}})


if __name__ == "__main__":
    datas = [dt.datetime.strptime(x, "%d/%m/%Y").date() for x in sys.argv[1:] if "/" in x]
    ini = datas[0] if datas else dt.date(2026, 1, 1)
    fim = datas[1] if len(datas) > 1 else dt.date(2026, 9, 26)
    estabs = [int(x) for x in sys.argv[1:] if "/" not in x] or list(UNIDADES)
    dados = coletar(ini, fim, estabs)
    out = os.environ.get("BELLE_OUT", "Avaliacao_x_Cabine_Drenesse.xlsx")
    montar(dados, ini, fim, estabs, out)
    dg = dados["diag"]
    print(f"{len(dados['vendas'])} vendas, {len(dados['visitas'])} avaliações realizadas, {len(dados['agenda_av'])} agendamentos "
          f"com ID de avaliação -> {out}")
    print("Tipos de item no Vendas Detalhado:", dict(dg["tipos"]))
    print("Agendamentos sem serviço por tipo:", dict(dg["sem_servico"]))
    achados = {a["servico_cod"] for a in dados["agenda_av"]}
    for cod_s, nm in AVALIACAO.items():
        if cod_s not in achados:
            print(f"ATENÇÃO: nenhum agendamento com o ID {cod_s} ({nm}) no período – confira o código do serviço no Belle.")
