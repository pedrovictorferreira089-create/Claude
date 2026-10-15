"""Simulação de cashback sobre o histórico de compras de planos (API Belle Software).

Uso: python3 belle_cashback.py [--taxa 0.03 | --fixo 100] [--minimo 500] [--carencia 30] [--validade 180] [--limite 0.10]
     python3 belle_cashback.py --segunda 150 [--minimo 500] [--carencia 30] [--validade 180] [--limite 0.20]
Repete as regras do módulo de cashback do Belle sobre as compras reais (mesma definição de compra de
belle_recompra.py): cada compra a partir do valor mínimo gera cashback (taxa sobre o valor ou valor fixo), que
pode ser usado em compras feitas entre a carência e a validade, até o limite (fração da nova compra).
Mostra quanto seria devolvido a quem já compra hoje (custo sem mudança de comportamento), o ponto de
equilíbrio em clientes novos a mais com 2ª compra e os cenários de +3, +5 e +8 pontos.
Com --segunda, simula um crédito só para a 2ª compra: no dia da carência, o cliente novo (1ª compra a partir do
mínimo) que ainda não voltou recebe o valor, que vale até a validade (contada da 1ª compra) e paga até o limite
da 2ª compra. O custo é o crédito que usariam os que voltariam de qualquer jeito; quem já voltou antes não recebe.
"""
import argparse
import collections
import datetime as dt

from belle_recompra import analisar, carregar, montar_compras


def simular(por_cliente, ini, fim, taxa=0.03, fixo=None, minimo=500, carencia=30, validade=180, limite=0.10):
    """Cashback gerado e resgatado nas compras entre ini e fim, por número da compra e por unidade."""
    gerado = 0.0
    resgatado = collections.Counter()
    por_unidade = collections.Counter()
    for compras in por_cliente.values():
        creditos = []  # [data de geração, saldo]
        for x in compras:
            d, v = x["data"], x["valor"]
            teto = v * limite if limite is not None else v
            usado = 0.0
            for cr in creditos:  # usa primeiro o saldo mais antigo
                if cr[1] > 0 and carencia <= (d - cr[0]).days <= validade and usado < teto:
                    u = min(cr[1], teto - usado)
                    cr[1] -= u
                    usado += u
            if ini <= d <= fim:
                resgatado["2ª compra" if x["n"] == 2 else "3ª em diante"] += usado
                por_unidade[x["unidade"]] += usado
            if v >= minimo:
                cb = fixo if fixo is not None else taxa * v
                creditos.append([d, cb])
                gerado += cb if ini <= d <= fim else 0.0
    return gerado, resgatado, por_unidade


def credito_segunda(por_cliente, ini, fim, credito=150, minimo=500, dia=30, validade=180, limite=0.20):
    """Clientes novos entre ini e fim, quantos recebem o crédito, quantos o usariam sem mudar de comportamento e o
    custo, total e por unidade da 1ª compra."""
    novos = recebem = usos = 0
    custo = 0.0
    por_unidade = collections.Counter()
    for compras in por_cliente.values():
        compras = sorted(compras, key=lambda x: x["data"])
        p = compras[0]
        if not ini <= p["data"] <= fim:
            continue
        novos += 1
        dias = (compras[1]["data"] - p["data"]).days if len(compras) > 1 else None
        if p["valor"] < minimo or (dias is not None and dias < dia):
            continue
        recebem += 1
        if dias is not None and dias <= validade:
            u = min(credito, limite * compras[1]["valor"])
            usos += 1
            custo += u
            por_unidade[p["unidade"]] += u
    return novos, recebem, usos, custo, por_unidade


def reais(v):
    return ("−" if v < 0 else "") + f"R$ {abs(v):,.0f}".replace(",", ".")


if __name__ == "__main__":
    a = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    a.add_argument("--taxa", type=float, default=0.03)
    a.add_argument("--fixo", type=float, help="valor fixo por compra, em vez da taxa")
    a.add_argument("--minimo", type=float, default=500)
    a.add_argument("--carencia", type=int, default=30)
    a.add_argument("--validade", type=int, default=180)
    a.add_argument("--limite", type=float, help="fração da nova compra que o cashback pode pagar (padrão 0.10; 0.20 com --segunda)")
    a.add_argument("--segunda", type=float, help="crédito fixo só para a 2ª compra, dado no dia da carência a quem não voltou")
    a.add_argument("--inicio", default="01/10/2025")
    a.add_argument("--fim", default="30/09/2026")
    a.add_argument("--novos", type=int, default=1400, help="clientes novos por ano")
    a.add_argument("--valor", type=float, default=2700, help="recompras no 1º ano de cada cliente a mais que volta (R$)")
    a.add_argument("--modulo", type=float, default=3000, help="custo anual do módulo (R$)")
    o = a.parse_args()
    if o.limite is None:
        o.limite = 0.20 if o.segunda else 0.10
    ini, fim = (dt.datetime.strptime(s, "%d/%m/%Y").date() for s in (o.inicio, o.fim))

    planos, validade, sessoes, futuros, visitas, agendado, _ = carregar()
    compras, _ = montar_compras(planos, sessoes, futuros)
    analisar(compras, visitas, agendado, validade)  # numera as compras de cada cliente
    por_cliente = collections.defaultdict(list)
    for x in compras:
        por_cliente[x["cliente"]].append(x)
    if o.segunda:
        # coorte que já teve a validade inteira para voltar: 1ª compra entre (início − 1 ano) e (fim − 1 ano)
        ini2, fim2 = ini.replace(year=ini.year - 1), fim.replace(year=fim.year - 1)
        novos, recebem, usos, custo, por_unidade = credito_segunda(por_cliente, ini2, fim2, o.segunda, o.minimo, o.carencia,
                                                                   o.validade, o.limite)
        print(f"Clientes novos de {ini2:%d/%m/%Y} a {fim2:%d/%m/%Y}: {novos} • recebem o crédito de {reais(o.segunda)}"
              f" no {o.carencia}º dia: {recebem} • usariam sem mudar de comportamento: {usos}")
        print(f"Custo: {reais(custo)} por ano (crédito médio {reais(custo / max(usos, 1))})")
        for u, v in sorted(por_unidade.items(), key=lambda x: -x[1]):
            print(f"  {u}: {reais(v)}")
        equilibrio = custo / (o.valor - custo / max(usos, 1))
        pontos_eq = f"{equilibrio / novos * 100:.1f}".replace(".", ",")
        print(f"Equilíbrio: {equilibrio:.0f} clientes novos a mais com 2ª compra por ano (+{pontos_eq} pontos)")
        for pontos in (3, 5, 8):
            k = novos * pontos / 100
            print(f"  +{pontos} pontos: +{k:.0f} clientes, vendas {reais(k * o.valor)}, resultado {reais(k * (o.valor - o.segunda) - custo)}")
        raise SystemExit
    gerado, resgatado, por_unidade = simular(por_cliente, ini, fim, o.taxa, o.fixo, o.minimo, o.carencia, o.validade, o.limite)
    vendas = sum(x["valor"] for x in compras if ini <= x["data"] <= fim)
    custo = sum(resgatado.values())

    print(f"Compras de {o.inicio} a {o.fim}: {reais(vendas)} em planos")
    pct = f"{custo / vendas:.2%}".replace(".", ",")
    print(f"Cashback gerado: {reais(gerado)} • resgatado por quem já compra: {reais(custo)} ({pct} das vendas)")
    for k, v in resgatado.items():
        print(f"  {k}: {reais(v)} ({v / custo:.0%})")
    for u, v in sorted(por_unidade.items(), key=lambda x: -x[1]):
        print(f"  {u}: {reais(v)}")
    taxa_efetiva = o.taxa if o.fixo is None else 0.0
    por_cliente_a_mais = o.valor * (1 - taxa_efetiva)
    equilibrio = (custo + o.modulo) / por_cliente_a_mais
    pontos_eq = f"{equilibrio / o.novos * 100:.1f}".replace(".", ",")
    print(f"Equilíbrio: {equilibrio:.0f} clientes novos a mais com 2ª compra por ano (+{pontos_eq} pontos)")
    for pontos in (3, 5, 8):
        k = o.novos * pontos / 100
        print(f"  +{pontos} pontos: +{k:.0f} clientes, vendas {reais(k * o.valor)}, resultado {reais(k * por_cliente_a_mais - custo - o.modulo)}")
