# DIM0319 — Algoritmo e Programação de Computadores (Python)

Guia de referência do material do Prof. Marcus Terra (DIMAp/UFRN, 2026.2), usado como
norte para os trabalhos da disciplina.

## Avaliação
- P1, P2, P3 — 15% cada
- **Projeto de análise de dados em equipe (3–4 pessoas) — 40%** ("o coração da avaliação"):
  um tema de interesse por equipe, entrega e apresentação ao final
- Atividades — 15% · Frequência mínima 75%

## Ambiente e convenções do professor
- Tudo no **Google Colab** (`.ipynb`), executado de cima para baixo (Shift + Enter).
- Notebook = células de **texto (Markdown)** explicando o raciocínio + células de **código**.
- Cabeçalho padrão: título com emoji, linha `**DIM0319 · Aula N** · Prof. Marcus Terra · DIMAp/UFRN`,
  objetivo, "Como usar", seções numeradas (`## 1. Preparação`, `## 2. ...`).
- Código simples e didático: comentários curtos com `#`, nomes em português sem acento
  (`media`, `aprovados`, `precos`), constantes em MAIÚSCULAS (`N`, `LIM`, `SOMA_MIN`).
- Saídas com `print('Rótulo:', valor)`; f-strings só pontualmente (`f'{x:.1%}'`).
- Fechamento com "Sua vez!" (desafios) e um resumo do que foi praticado.

## Conteúdo por aula
| Aula | Tema | Pontos-chave |
|---|---|---|
| 1 | Visão geral | Python, Colab, NumPy/pandas/Matplotlib/SciPy; ciclo Coletar → Organizar → Analisar → Decidir |
| 2 | Pensamento computacional | decomposição, padrões, abstração, algoritmos; `print`, `#`, variáveis, `sum/len/max/min` |
| 3 | Tipos e coleções | `str/int/float/bool`, `type`, conversões, `input()` (sempre texto), listas (índice 0, `[-1]`, fatias, `append`, `sorted`), dicionários |
| 4 | Operadores, condicionais, laços | `+ - * / // % **`, comparações, `and/or/not`, `if/elif/else`, `for`/`range`, `while`, contadores |
| 5 | Funções | `def`, parâmetros, `return` (sem return → `None`), valor padrão, escopo local, docstring |
| 6 | Consolidação | **5 pilares**: tipos, operadores, controle, estruturas de dados, funções; estudo de caso "boletim da turma" (lista de dicionários) |
| 7 | NumPy: arrays e vetorização | `np.array`, operações vetorizadas, `np.mean/sum/max/min/std/median`; exemplos de engenharia (produção, alimentos, têxtil) e biologia |
| 8 | NumPy: indexação e seleção | fatias, máscaras booleanas `a[a >= 6]`, `&`, `|` com parênteses, `np.sum(mascara)`, arrays 2D, `axis=0/1` |
| 9 | **Projeto Mega-Sena (NumPy)** | gerar N jogos → calcular características → filtrar com máscaras → escolher palpites (`rng`, `argsort`, `diff`, `any`, broadcasting) |
| 10 | pandas: Series e DataFrame | `pd.Series`, `pd.DataFrame(dict)`, `head/shape/columns`, `df['col']`, `describe`, nova coluna, filtro `df[df['nota'] >= 6]` |
| 11 | pandas: ler CSV/Excel | `read_csv` (`sep=';'`), upload/URL/Drive, `head/tail/shape/dtypes/info/describe`, `value_counts`, `to_csv(index=False)` |
| Revisão | Fundamentos + NumPy + pandas | erros comuns: `:` e indentação, `&`/`|` em arrays, fim da fatia não entra, nome de coluna |

Próxima aula prevista: Aula 12 — pandas: selecionar e filtrar dados (ainda não disponível no Drive).

## Listas de exercícios (1–9)
Estrutura fixa: Partes A–D (aquecimento → prática → "Juntando tudo"), ~15 itens, último é
(Desafio). Cada lista tem versão aluno e versão "RESPOSTAS (versão do professor)".

## Padrão de projeto a seguir (Aula 9)
**gerar/carregar dados → calcular características → filtrar com máscaras → resumir/escolher**,
com texto explicativo entre as etapas, contagem do que sobra a cada filtro, e seção final
"Sua vez!" + resumo das ferramentas usadas.

---

# Parte 2 — pandas intermediário (Aulas 12 a 16)

Material da segunda pasta do Drive ("ALGORÍTMO E PROGRAMAÇÃO DE COMPUTADORES 2").
A partir daqui, `groupby`, `sort_values`, `.loc`, limpeza com `.str`, `merge`/`concat` **passam a ser
matéria vista** e podem ser usados nos trabalhos.

| Aula | Tema | Pontos-chave |
|---|---|---|
| 12 | Selecionar e filtrar | `df[['a','b']]` (colchetes duplos), `df[mascara]`, `&`/`|` com parênteses, `isin([...])`, filtrar + escolher colunas, `df.loc[condicao, colunas]` |
| 13 | Limpeza de dados | `isnull().sum()`, `dropna()` / `dropna(subset=[...])`, `drop(columns=[...])`, `fillna(media/mediana/'N/I')`, `fillna(df.median(numeric_only=True))`, `astype`, `pd.to_numeric(errors='coerce')`, `duplicated()` / `drop_duplicates()`, `.str.strip()/.lower()/.title()/.replace()`, `rename(columns={...})`; fluxo **ver → decidir → limpar → conferir**; sempre reatribuir (`df = df...`) |
| 14 | Novas colunas, ordenação, estatística | coluna por conta; coluna por condição (`df['situacao'] = 'Reprovado'` + `df.loc[cond, 'situacao'] = 'Aprovado'`); `sort_values('col', ascending=False).head(n)`; `sort_values(by=['a','b'], ascending=[True, False])`; `mean/median/std/min/max/sum`; `idxmax` + `.loc`; `nlargest(n, 'col')` |
| 15 | Agrupar (groupby) | `df.groupby('grupo')['valor'].mean()` (ou sum/count/max/min/median); `.size()`; `.agg(['mean','sum','max'])`; `groupby(['a','b'])`; ranking com `.sort_values(ascending=False)`; `.idxmax()` |
| 16 | Juntar tabelas | `pd.concat([a, b], ignore_index=True)` (mesmas colunas, empilha); `pd.merge(a, b, on='chave')`, `how='inner'` (padrão) x `how='left'` (NaN onde não casa); padronizar a chave antes (`.str.strip().str.lower()`); exemplos de Eng. de Produção (produção x meta por linha), Alimentos, Biologia |

Próxima aula prevista: Aula 17 — pandas: estudo de caso (fechando a unidade).

## Atividades de aula
- **Censo 2022 (IBGE)** — `censo2022_municipios.csv` (5.570 municípios; `uf`, `municipio`, `pop_coletada`,
  `pop_imputada`, `pop_total`; soma ≈ 203 milhões). Conhecer, selecionar, filtrar (RN, > 500 mil, Natal),
  nova coluna `taxa_imputacao`, top 10 com `sort_values`.
- **Estudo de Caso Integrador — Lanchonete do Zé Byte** — `vendas_lanchonete.csv` (22 linhas, `sep=';'`,
  preço com vírgula `3,50`, textos bagunçados `Cafe`/`cafe `/`BEBIDA`, faltantes, 2 linhas duplicadas).
  Roteiro em 5 partes, **modelo provável para o próximo trabalho**:
  1. ler (`sep=';'`) e conhecer; 2. limpar (duplicatas, `.str.strip().str.lower()`, vírgula→ponto +
  `to_numeric`, `fillna` com mediana, `astype(int)`); 3. enriquecer (`valor_total`, função `classificar` com
  `if/elif/else`, laço `for` para criar a coluna `faixa`); 4. NumPy (`np.mean`, `np.std`, máscara);
  5. analisar (`groupby` + `sort_values`, `value_counts`, top 3, `idxmax`).

## Bases de dados da pasta
| Arquivo | Linhas × colunas | "Sujeira" para limpar |
|---|---|---|
| `pinguins.csv` (Palmer) | 344 × 8 | 2 faltantes nas medidas, 11 em `sex` |
| `titanic.csv` | 891 × 12 | `Age` 177, `Cabin` 687, `Embarked` 2 faltantes |
| `taco_alimentos.csv` (NEPA/UNICAMP) | 597 × 32 | muitos nutrientes faltando (`colesterol_mg`, `RE_mcg`, `RAE_mcg` quase vazias), `preparo` 233 |
| `censo2022_municipios.csv` (IBGE) | 5.570 × 5 | limpa |
| `vendas_lanchonete.csv` | 22 × 7 | `;`, vírgula decimal, texto inconsistente, faltantes, duplicatas |

## Listas de exercícios 10–13
Mesmo formato das anteriores (Partes A–D, ~14–15 itens, último é desafio). Nesta pasta só vieram as
versões do aluno (sem gabarito).
- **10:** selecionar/filtrar (`[[ ]]`, máscara, `&`, `isin`, `.loc`).
- **11:** limpeza na base TACO (`isnull`, `drop`, `fillna`, `rename`, `drop_duplicates`, `.str.strip`, `to_csv`).
- **12:** boletim com `media`, `situacao` via `.loc`, ordenação, `describe`, `idxmax`, colunas `bonus`/`media_final`.
- **13:** `groupby` com `mean/sum/size/agg`, por duas colunas, ranking e categoria campeã.
