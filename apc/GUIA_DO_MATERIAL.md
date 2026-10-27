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
