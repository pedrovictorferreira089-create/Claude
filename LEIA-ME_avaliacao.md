# Passagem: análise Avaliação x Cabine SDR x Cabine (Drenesse / Belle Software)

Documento para quem for continuar este trabalho em outra sessão do Claude Code. A conversa original não fica
disponível; tudo o que foi decidido com a Drenesse está aqui.

## Objetivo

Separar as vendas da Drenesse em três grupos e comparar, para cada um, **faturamento, quantidade de planos, taxa de
comparecimento e taxa de fechamento**, com exemplos reais de atendimentos e o passo a passo do cálculo.

- **Período padrão:** 01/01/2026 a 26/09/2026.
- **Unidades:** Petrópolis (1), Lagoa Nova (2), Capim Macio (3) e Norte Shopping (6). A unidade **Laser (4) fica
  fora**, a pedido da Drenesse.

## Como rodar

```bash
pip install openpyxl
export BELLE_TOKEN=...          # chave da API do Belle (troca a cada 2 semanas; NUNCA gravar no repositório)
python3 belle_avaliacao.py 01/01/2026 26/09/2026          # gera Avaliacao_x_Cabine_Drenesse.xlsx
python3 belle_avaliacao.py 01/01/2026 26/09/2026 1 2      # só algumas unidades (codEstab)
```

- **`BELLE_OUT`** muda o nome do arquivo gerado.
- **`BELLE_HOJE=dd/mm/aaaa`** fixa a data de referência. A primeira janela de consulta dos planos termina em
  `HOJE`, então manter a mesma data reaproveita o cache.
- **Cache da API:** fica em `./data/`, ignorado pelo git. A API aceita no máximo 3 meses por consulta e 40
  requisições por minuto. A coleta completa leva uns 20 a 30 minutos; com o cache pronto, a planilha sai em
  segundos.
- **Recalcular as fórmulas:** para conferir os valores fora do Excel, use o LibreOffice Calc
  (`apt-get install libreoffice-calc`). Só o `libreoffice-core` não abre planilhas.

## Regras definidas pela Drenesse

### Tipos de sessão (agenda)

| Sessão | Como identificar no Belle |
|---|---|
| **Avaliação** | `tipoAgendamento = "Avaliação"` **ou** serviço **52 – AVALIAÇÃO ESTÉTICA** (confirmado pela Drenesse: as mesmas avaliadoras fazem os dois). |
| **Experimental** | Um dos 6 IDs abaixo **sem plano vinculado** (`idOrcamento` vazio = agendamento avulso). |
| **Cabine** | Qualquer outro agendamento com serviço, inclusive os 6 IDs quando estão dentro de um plano. |
| Fora da análise | Retorno, Consulta e agendamentos sem serviço. |

IDs das sessões experimentais, informados pela Drenesse:

| ID | Serviço |
|---|---|
| 22 | DRENAGEM MÉTODO DRENESSE (também é a sessão comum de plano; só conta como experimental sem plano vinculado) |
| 56210744 | SESSÃO EXPERIMENTAL DRENAGEM - MÉTODO DRENESSE |
| 56260425 | DRENAGEM MÉTODO DRENESSE 98,70 |
| 33353403 | LIMPEZA DE PELE PREMIUM |
| 56210746 | SESSÃO EXPERIMENTAL - STIMULUS |
| 56210745 | SESSÃO EXPERIMENTAL - DIÁSTASE |

### Classificação de cada venda (pelo cliente e pelo dia da compra)

1. **Avaliação:** a cliente teve sessão de avaliação com status **Atendido** no dia, com ou sem sessão experimental.
2. **Cabine SDR:**
   - avaliação do dia com **falta ou desmarcada**, mas sessão experimental atendida;
   - **só** sessão experimental no dia e a cliente **nunca teve plano aprovado** antes.
3. **Cabine:** qualquer outra compra. Isso inclui:
   - experimental de quem já tinha plano;
   - renovações;
   - compra em dia diferente da avaliação (confirmado pela Drenesse: fica em Cabine);
   - compras sem sessão no dia.

Cada sessão experimental entra na agenda e no comparecimento do grupo da venda daquele dia, pela mesma regra.

### Indicadores

- **Comparecimento:** atendidos ÷ (atendidos + faltas). Faltas são o status "Falhou"; desmarcados e cancelados ficam
  fora da taxa.
- **Fechamento da Avaliação e da Cabine SDR:** sessões atendidas (cliente/dia) com plano comprado no mesmo dia ÷
  sessões atendidas.
- **Fechamento da Cabine:** por cliente/mês, isto é, clientes atendidos na cabine no mês que compraram plano de
  cabine no mesmo mês.
- **Faturamento:** valor vendido na data da venda:
  - planos com status **Aprovado**, pelo `precoFinal` do relatório Venda de Planos;
  - serviços avulsos, pelo `valor_liquido` do Vendas Detalhado.

  Planos não aprovados ficam listados, mas fora da conta.

## Fontes da API (`IntegracaoExterna/v1.0/`)

- **`relatorios/relatorio_atendimentos`** (`dtInicio`, `dtFim`, `codEstab`): agenda com `tipoAgendamento`,
  `codigoServico`, `statusAgendamento`, `idOrcamento` e `idAgendamento`.
- **`venda_planos`** (`dtInicio`, `dtFim`, `codEstab`): planos com `dataVenda`, `statusPlano`, `precoFinal` e
  `vendedor`.
  - O campo `avaliacao` traz o `idAgendamento` do atendimento em que a ficha foi feita. Ele vem preenchido só em
    parte dos planos e é usado apenas como conferência.
- **`vendas_detalhado`** (`estab`, `dtInicio`, `dtFim`): itens vendidos.
  - Os itens `tipo = "Plano"` são ignorados, porque os planos já entram pelo `venda_planos`.
  - Itens com status cancelado são descartados.

## Resultado da última execução (01/01 a 26/09/2026, extraído em 01/10/2026)

| | Avaliação | Cabine SDR | Cabine |
|---|---|---|---|
| Agendamentos | 3.484 | 2.675 | 74.447 |
| Comparecimento | 55,5% | 51,2% | 83,5% |
| Fechamento | 32,0% (555/1.732) | 21,6% (232/1.072) | 18,6% (cliente/mês) |
| Planos aprovados | 622 | 254 | 2.199 |
| Ticket médio do plano | R$ 2.546 | R$ 2.085 | R$ 1.801 |
| Faturamento total | R$ 1.616.735 (26%) | R$ 571.938 (9%) | R$ 3.975.705 (65%) |

Conferências feitas:

- **Faturamento de planos:** bate com o Vendas Detalhado do Belle. São 3.075 planos e R$ 6,074 milhões, contra
  3.080 planos e R$ 6,088 milhões no Belle (diferença de 0,2%).
- **Vínculo plano–avaliação:** os 298 planos que o Belle liga a uma sessão de avaliação atendida no mesmo dia
  ficaram todos como Avaliação.
- **Simulação:** a lógica foi testada com dados simulados de resposta conhecida, cobrindo cada um dos casos acima.

## Estrutura da planilha gerada

- **Resumo:** comparativo dos 3 grupos, fechamento da avaliação em 0, 7 e 30 dias, origem das vendas e blocos por
  unidade.
- **Mensal:** os mesmos indicadores mês a mês.
- **Serviços SDR:** resultados de cada um dos 6 IDs experimentais.
- **Exemplos:** casos reais de cada regra.
- **Metodologia:** passo a passo e diagnóstico dos dados.
- **Abas de base:** Vendas, Avaliações, Sessões SDR, Agenda avaliação-experimental, Agenda resumo e Clientes
  cabine. Todas as contas do Resumo e do Mensal são fórmulas sobre essas abas.

## Código

- **`belle_avaliacao.py`:** esta análise.
  - `coletar()` busca os dados e aplica as regras.
  - `montar()` gera a planilha.
- **Reaproveita:**
  - de `belle_sabados.py`: `api()`, `historico()`, `carregar_planos()` e os estilos;
  - de `belle_auditoria.py`: `base()`, `linha()` e `compactar_xlsx()`.
- **Outros relatórios do repositório:**
  - `belle_sabados.py`: faturamento dos sábados;
  - `belle_servicos.py`: ranking de serviços;
  - `belle_auditoria.py`: auditoria de preços.
