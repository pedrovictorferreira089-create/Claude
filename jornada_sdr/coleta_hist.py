import os, sys, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from belle_sabados import historico, UNIDADES
hoje = dt.date(2026, 10, 16)
for e in UNIDADES:
    v = historico("vendas_detalhado", "estab", e, hoje)
    print("vendas", e, len(v), flush=True)
for e in UNIDADES:
    a = historico("relatorios/relatorio_atendimentos", "codEstab", e, hoje, desde=dt.date(2023, 1, 1))
    print("agenda", e, len(a), flush=True)
print("fim")
