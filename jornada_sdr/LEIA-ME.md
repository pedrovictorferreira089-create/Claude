# Jornada ideal de compra (time SDR)

Análise que alimenta o fluxograma publicado em https://claude.ai/artifact/M9gXvf9sQD1tAhZQgEDBWu (cópia em `jornada_sdr.html`).

1. `python3 coleta_hist.py`: baixa para `../data/` (ignorado pelo git) as vendas detalhadas de todo o histórico e a agenda desde 2023. Leva cerca de 1h30 sem cache.
2. Rodar de dentro desta pasta:
   - `analise.py`: porta de entrada pelas vendas pagas (avulso x plano direto) e valor em 12 meses;
   - `analise2.py`: inativos por faixa, efeito do acompanhamento estético (reavaliação) e porta de entrada pela agenda (inclui sessões gratuitas);
   - `analise3.py`: poder de compra por bairro, profissão, cidade, idade e primeiro procedimento;
   - `lista.py`: gera `Base_Reativacao_SDR.xlsx` com as inativas de 30 a 365 dias, nível de gasto, saldo de sessões e próximo passo. Tem dados pessoais: não versionar nem publicar.

`base.py` define as categorias de procedimento (regras por nome do serviço) e junta os planos aprovados por cliente e dia.
Referência das datas: `HOJE` em `base.py` (16/10/2026).
