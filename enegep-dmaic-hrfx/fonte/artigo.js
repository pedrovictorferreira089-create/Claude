const L = require("./lib");
const { P, H1, H2, bullet, alineas, caption, fonte, image, table, section, save, numbering, styles, D } = L;
const out = process.argv[2];

const c = []; // corpo
const add = (...x) => x.flat().forEach((e) => c.push(e));

// =============== 1. INTRODUÇÃO ===============
add(new D.Paragraph({ children: L.runs("1. Introdução", { bold: true }), spacing: { line: L.LINE, lineRule: D.LineRuleType.AUTO }, keepNext: true }));
add(P("Os serviços apresentam características que tornam a gestão da qualidade particularmente desafiadora: são intangíveis, heterogêneos, perecíveis e, em grande parte, produzidos e consumidos simultaneamente (FITZSIMMONS; FITZSIMMONS, 2014). Como não é possível inspecionar o serviço antes de sua entrega, a qualidade depende da estabilidade dos processos de preparação que o antecedem. Em micro e pequenas empresas, essa estabilidade é frequentemente comprometida pela informalidade dos procedimentos, pela concentração das decisões no proprietário e pela escassez de recursos para iniciativas estruturadas de melhoria (KUMAR; ANTONY; DOUGLAS, 2009)."));
add(P("No setor de serviços pirotécnicos para eventos, essas dificuldades assumem relevância adicional. Os artefatos utilizados são produtos controlados (BRASIL, 2019), e o espetáculo é executado ao vivo, diante do público e em um único momento, sem possibilidade de repetição. Uma falha na preparação do material, como a ausência de um detonador ou de um cabo de extensão, compromete simultaneamente a satisfação do contratante, a segurança dos presentes e a reputação da empresa."));
add(P("Este artigo tem como objeto de estudo a HRFx Efeitos Especiais, microempresa sediada em Natal/RN que executa shows pirotécnicos e efeitos especiais para casamentos e shows musicais. O diagnóstico realizado pelos autores em visitas técnicas à empresa identificou que itens são esquecidos na preparação dos *kits* de serviço em cerca de 15% dos eventos, que a separação de um *kit* de médio porte demanda, em média, 55 minutos e 18 deslocamentos no depósito, e que todo o conhecimento sobre a composição dos *kits* está concentrado no proprietário. Essas falhas geram deslocamentos emergenciais, improvisos na montagem e sobrecarga de trabalho às vésperas dos eventos."));
add(P("Diante desse cenário, o objetivo deste trabalho é aplicar o método DMAIC (*Define, Measure, Analyze, Improve, Control*) para identificar e analisar as causas das falhas no processo de preparação dos *kits* da HRFx e propor melhorias e um plano de controle fundamentados em dados. Conforme o escopo definido, as propostas são apresentadas sem implementação, acompanhadas dos resultados esperados. Especificamente, busca-se:"));
add(alineas([
  "definir o problema, o escopo e os requisitos críticos do cliente;",
  "medir o desempenho atual do processo e priorizar as causas de falha;",
  "analisar as causas raízes e os riscos associados a cada modo de falha;",
  "propor ações de melhoria priorizadas e um plano de controle prospectivo.",
]));
add(P("O artigo está organizado em cinco seções: após esta introdução, apresenta-se a fundamentação teórica; em seguida, o método e a caracterização da empresa; depois, os resultados de cada fase do DMAIC; e, por fim, as conclusões e limitações do estudo."));

// =============== 2. FUNDAMENTAÇÃO ===============
add(H1("2. Fundamentação teórica"));
add(H2("2.1. Qualidade em serviços"));
add(P("A qualidade em serviços é entendida como a diferença entre o que o cliente espera e o que ele percebe ter recebido. Parasuraman, Zeithaml e Berry (1988) operacionalizaram esse conceito no instrumento SERVQUAL, que avalia cinco dimensões: tangibilidade, confiabilidade, responsividade, garantia e empatia. Para cada dimensão, calcula-se a lacuna (*gap*) Q = P − E, em que P é a percepção e E a expectativa; valores negativos indicam que o serviço ficou aquém do esperado. Entre as dimensões, a confiabilidade, definida como a capacidade de executar o serviço prometido de forma precisa, costuma ser a mais valorizada pelos clientes, o que a torna especialmente crítica em serviços que não admitem correção após a entrega."));
add(H2("2.2. Seis Sigma e o método DMAIC"));
add(P("O Seis Sigma é uma abordagem estruturada de melhoria que busca reduzir a variabilidade e os defeitos dos processos por meio de decisões baseadas em dados (PANDE; NEUMAN; CAVANAGH, 2001). Schroeder *et al.* (2008) o caracterizam como uma estrutura organizada que combina um método de solução de problemas, especialistas em melhoria e métricas de desempenho. Seu principal método é o DMAIC, composto por cinco fases encadeadas (WERKEMA, 2012; MONTGOMERY, 2016):"));
add(alineas([
  "**Definir**: delimitar o problema, a meta, o escopo e os clientes do processo;",
  "**Medir**: quantificar o desempenho atual e localizar o foco do problema;",
  "**Analisar**: identificar e confirmar as causas raízes;",
  "**Melhorar**: propor, priorizar e testar soluções para as causas identificadas;",
  "**Controlar**: padronizar o processo e monitorar os resultados para que o problema não retorne.",
]));
add(P("Embora tenha origem na manufatura, o Seis Sigma tem sido aplicado a processos de serviço, nos quais os principais desafios são a dificuldade de medir resultados intangíveis e a menor disponibilidade de dados históricos (ANTONY, 2006). Em pequenas empresas, a adoção do método exige adaptar ferramentas e estrutura à disponibilidade limitada de recursos e de pessoal (KUMAR; ANTONY; DOUGLAS, 2009), o que justifica o uso de ferramentas simples e de baixo custo, como as empregadas neste estudo."));
add(H2("2.3. Ferramentas da qualidade aplicadas"));
add(P("Na fase Definir, o *Project Charter* formaliza o problema, a meta, o escopo e a equipe do projeto, enquanto o SIPOC (*Suppliers, Inputs, Process, Outputs, Customers*) representa o processo em alto nível, delimitando suas fronteiras (WERKEMA, 2012). Na fase Medir, o Diagrama de Pareto ordena as causas de um problema por frequência e evidencia os poucos fatores vitais responsáveis pela maior parte das ocorrências (JURAN; GRYNA, 1993)."));
add(P("Na fase Analisar, o Diagrama de Causa e Efeito organiza as possíveis causas de um efeito em categorias, usualmente os 6M: método, mão de obra, máquina, material, medição e meio ambiente (ISHIKAWA, 1993). A técnica dos 5 Porquês complementa o diagrama ao aprofundar uma cadeia causal até a causa raiz (OHNO, 1997). A Análise de Modo e Efeito de Falha (FMEA, do inglês *Failure Mode and Effects Analysis*) é uma ferramenta preventiva que identifica como cada etapa pode falhar, quais os efeitos e as causas da falha e quais controles existem. Cada modo de falha recebe notas de 1 a 10 para severidade (S), ocorrência (O) e detecção (D), e o Número de Prioridade de Risco é calculado como NPR = S × O × D (STAMATIS, 2003). Como combinações distintas podem gerar o mesmo NPR, o manual AIAG e VDA (2019) recomenda priorizar modos de falha de severidade elevada independentemente do NPR."));
add(P("Nas fases Melhorar e Controlar, o plano de ação 5W2H detalha o que será feito, por que, onde, quando, por quem, como e a que custo (CARPINETTI, 2016). A análise de custos da qualidade classifica os gastos em prevenção, avaliação, falhas internas e falhas externas, e permite demonstrar que investir em prevenção tende a ser mais barato do que arcar com as falhas (FEIGENBAUM, 1994). Por fim, práticas do Sistema Toyota de Produção, como o 5S, o trabalho padronizado e o nivelamento da carga de trabalho (*heijunka*), sustentam a estabilidade dos processos e funcionam como base para a melhoria contínua (OHNO, 1997; LIKER, 2005)."));

// =============== 3. MÉTODO ===============
add(H1("3. Método"));
add(H2("3.1. Caracterização da empresa"));
add(P("A HRFx Efeitos Especiais foi fundada em [COMPLETAR: ano de fundação] por Felipe Reis, sócio-proprietário que acumula as funções de gestão administrativa e comercial e a responsabilidade técnica pelos espetáculos como *blaster*. [COMPLETAR: breve histórico de como a empresa surgiu]. A empresa atende cerca de 100 eventos por ano, principalmente casamentos e eventos sociais e shows musicais, com demanda concentrada nos fins de semana. Seu portfólio inclui shows pirotécnicos com acionamento elétrico, efeitos de chafariz (*gerbs*) e efeitos com cilindros de CO₂ [COMPLETAR: demais serviços e efeitos oferecidos]."));
add(P("A estrutura operacional compreende um depósito principal em Monte Alegre/RN, onde são armazenados os insumos pirotécnicos e equipamentos, e um ponto de apoio no bairro do Alecrim, em Natal. Os principais fornecedores são uma fábrica de artefatos pirotécnicos em Minas Gerais, lojas especializadas em Natal e uma distribuidora de CO₂. Os clientes são noivos, cerimonialistas e produtoras de eventos [COMPLETAR: confirmar perfil e principais clientes]. A equipe é formada pelo proprietário e por [COMPLETAR: número] auxiliares de apoio operacional. A missão da empresa é [COMPLETAR: missão], e sua visão é [COMPLETAR: visão]. Sua estratégia competitiva baseia-se na personalização dos espetáculos e na segurança da execução [COMPLETAR: confirmar estratégia], e seus principais concorrentes em Natal são [COMPLETAR: concorrentes]. Por operar com produtos controlados, a empresa depende de registro junto ao Exército Brasileiro (BRASIL, 2019) e de autorizações do Corpo de Bombeiros para cada evento."));
add(H2("3.2. Procedimentos da pesquisa"));
add(P("A pesquisa é aplicada, de natureza exploratória e descritiva, com abordagem quali-quantitativa, e foi conduzida como um estudo de caso único (GIL, 2017; YIN, 2015). Os dados foram obtidos em seis visitas técnicas à empresa, realizadas em 2025 no âmbito de um projeto extensionista, por meio de:"));
add(alineas([
  "entrevistas semiestruturadas com o proprietário sobre o fluxo de informações, a composição dos *kits* e as dificuldades operacionais;",
  "observação direta do processo de separação no depósito, com cronometragem do tempo de preparação e contagem dos deslocamentos da equipe;",
  "análise documental de contratos, ordens de serviço e registros informais de ocorrências, que totalizaram 41 falhas nos doze meses anteriores ao estudo;",
  "aplicação do questionário SERVQUAL, em escala Likert de sete pontos, a cinco clientes atendidos nos doze meses anteriores.",
]));
add(P("As notas de severidade, ocorrência e detecção da FMEA foram atribuídas de forma consensual pelos autores com base nessas evidências. Os custos foram estimados a partir de valores informados pelo proprietário. O Quadro 1 sintetiza a aplicação do DMAIC. Como as melhorias não foram implementadas, os resultados das fases Melhorar e Controlar são apresentados como metas e estimativas."));
add(caption("Quadro 1 - Etapas do DMAIC, ferramentas e produtos do estudo"));
add(table([1500, 3300, 4270], ["Fase", "Ferramentas", "Produto"], [
  ["Definir", "*Project Charter*; SIPOC; SERVQUAL", "Problema, meta, escopo e requisitos críticos do cliente"],
  ["Medir", "Indicadores de linha de base; Diagrama de Pareto; custos da qualidade", "Desempenho atual e causas prioritárias"],
  ["Analisar", "Diagrama de Ishikawa; 5 Porquês; FMEA", "Causas raízes e modos de falha priorizados por NPR"],
  ["Melhorar", "Plano de ação 5W2H; FMEA revisada", "Propostas de melhoria e NPR esperado"],
  ["Controlar", "Indicadores de desempenho; POP; auditoria 5S", "Plano de controle prospectivo"],
], { boldFirstCol: true }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));

// =============== 4. RESULTADOS ===============
add(H1("4. Resultados e discussão"));
add(H2("4.1. Definir"));
add(P("A fase Definir iniciou-se com a elaboração do *Project Charter* (Quadro 2), que delimitou o projeto ao processo de preparação dos *kits*, desde a confirmação do evento até a carga do veículo. A execução do espetáculo no local e o processo de compras ficaram fora do escopo."));
add(caption("Quadro 2 - Termo de abertura do projeto (*Project Charter*)"));
add(table([2100, 6970], ["Elemento", "Descrição"], [
  ["Problema", "Itens esquecidos em cerca de 15% dos eventos e separação lenta dos *kits* (55 min e 18 deslocamentos por *kit* de médio porte), gerando deslocamentos emergenciais, improvisos e sobrecarga às vésperas dos eventos"],
  ["Objetivo e metas", "Reduzir a proporção de eventos com item faltante de 15% para menos de 5%; reduzir o tempo médio de separação para até 35 min e os deslocamentos para até 8 por *kit*"],
  ["Escopo", "Dentro: programação da preparação, definição da lista de materiais, separação, conferência e carga. Fora: execução do show, compras e atendimento comercial"],
  ["Cliente e CTQ", "Contratante e equipe de montagem; requisitos críticos: *kit* completo e preparado no prazo"],
  ["Benefício esperado", "Redução de custos de falha e de riscos operacionais; menor dependência do proprietário"],
  ["Equipe", "Autores do estudo; proprietário como dono do processo"],
], { boldFirstCol: true, shadeFirstCol: true }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("O SIPOC (Quadro 3) mostra que o processo recebe informações do contratante (data, local e roteiro do espetáculo) e materiais de três fornecedores, e entrega um *kit* completo à equipe de montagem, que é o cliente interno imediato. Esse mapeamento evidenciou que não existe, entre a confirmação do contrato e a separação, nenhum documento que traduza o roteiro do show em uma lista de materiais: essa tradução é feita de memória pelo proprietário."));
add(caption("Quadro 3 - SIPOC do processo de preparação dos *kits*"));
add(table([1650, 1900, 2170, 1750, 1600], ["Fornecedores", "Entradas", "Processo", "Saídas", "Clientes"], [
  ["Fábrica de artefatos (MG); lojas especializadas (Natal); distribuidora de CO₂; contratante", "Contrato e roteiro do show; insumos pirotécnicos; detonadores; cabos; bases; ferramentas; extintor e EPI",
    "1. Confirmar evento\n2. Definir lista de materiais\n3. Separar itens no depósito\n4. Conferir e testar\n5. Carregar o veículo", "*Kit* completo e conferido; veículo carregado no prazo", "Equipe de montagem (interno); contratante e público (externos)"],
], {}));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Para captar a voz do cliente, aplicou-se o SERVQUAL (Tabela 1). Todas as dimensões apresentaram lacunas negativas, com maior severidade em confiabilidade (−1,9) e responsividade (−1,3). A confiabilidade combina a maior expectativa (6,5) com uma das menores percepções (4,6), o que é coerente com os esquecimentos de itens relatados. A garantia foi a dimensão mais bem avaliada (−0,2), sustentada pela competência técnica reconhecida do proprietário. Assim, o problema não está no conhecimento técnico, mas na capacidade do processo de entregar esse conhecimento de forma consistente. Com base nesse resultado, definiram-se dois requisitos críticos para a qualidade (CTQ): a completude do *kit* e o cumprimento do prazo de preparação."));
add(caption("Tabela 1 - Resultados do SERVQUAL aplicado a clientes da HRFx (n = 5)"));
add(table([2470, 1650, 1650, 1650, 1650], ["Dimensão", "Expectativa (E)", "Percepção (P)", "Lacuna (P − E)", "Prioridade"], [
  ["Confiabilidade", "6,5", "4,6", "−1,9", "1ª"],
  ["Responsividade", "6,1", "4,8", "−1,3", "2ª"],
  ["Empatia", "5,9", "5,0", "−0,9", "3ª"],
  ["Tangibilidade", "5,8", "5,2", "−0,6", "4ª"],
  ["Garantia", "6,3", "6,1", "−0,2", "5ª"],
], { centerCols: [1, 2, 3, 4], boldFirstCol: true }));
add(fonte("Fonte: Dados da pesquisa (2026)"));

add(H2("4.2. Medir"));
add(P("A fase Medir estabeleceu a linha de base do processo (Tabela 2). Considerando cerca de 100 eventos por ano, a taxa de 15% corresponde a aproximadamente 15 eventos anuais com algum item faltante. Além disso, a distribuição da carga de trabalho mostrou-se desequilibrada: a jornada ultrapassa frequentemente 12 horas às sextas-feiras, enquanto às terças-feiras a ocupação é inferior a 4 horas, porque a preparação é concentrada na véspera dos eventos."));
add(caption("Tabela 2 - Indicadores de linha de base do processo de preparação dos *kits*"));
add(table([3570, 1750, 2000, 1750], ["Indicador", "Linha de base", "Fonte", "Meta"], [
  ["Eventos com item faltante (%)", "15%", "Registros e entrevistas", "< 5%"],
  ["Tempo médio de separação por *kit* (min)", "55", "Cronometragem", "≤ 35"],
  ["Deslocamentos por *kit* (nº)", "18", "Observação direta", "≤ 8"],
  ["Jornada às sextas-feiras (h)", "> 12", "Entrevistas", "≤ 8"],
  ["Custo anual de falhas (R$)", "3.900", "Estimativa", "≤ 1.050"],
], { centerCols: [1, 2, 3] }));
add(fonte("Fonte: Dados da pesquisa (2026)"));
add(P("O custo anual de falhas foi estimado em duas parcelas. A falha externa corresponde aos itens esquecidos que exigem deslocamento emergencial até o local do evento: cada ocorrência custa cerca de R$ 60,00 de transporte e duas horas do proprietário, avaliadas em R$ 50,00 por hora, totalizando R$ 160,00. Para 15 ocorrências anuais, o custo é de R$ 2.400,00. A falha interna, relativa ao retrabalho de busca e nova separação de itens ainda no depósito, foi estimada em R$ 1.500,00 por ano. O custo total das falhas é, portanto, de R$ 3.900,00 por ano, sem considerar o dano à imagem da empresa, que não foi quantificado."));
add(P("Para localizar o foco do problema, as 41 ocorrências de falha registradas nos doze meses anteriores foram classificadas por causa e organizadas no Diagrama de Pareto (Figura 1)."));
add(caption("Figura 1 - Diagrama de Pareto das causas de falha na preparação dos *kits* (41 ocorrências em 12 meses)"));
add(image("pareto.png", 560, 280));
add(fonte("Fonte: Elaborado pelos autores com base em registros da empresa (2026)"));
add(P("A ausência de *checklist* formalizado (34,1%) e a falta de endereçamento visual no depósito (29,3%) concentram 63,4% das ocorrências. Somadas à sobrecarga do proprietário (17,1%), as três primeiras causas respondem por 80,5% das falhas, confirmando o princípio de Pareto. Os resultados indicam que o problema é sistêmico, ligado ao método e ao ambiente de trabalho, e não a falhas individuais de atenção. Cabe registrar que os registros de ocorrência eram informais e que não foi realizada análise do sistema de medição, o que é tratado como limitação na seção 5."));

add(H2("4.3. Analisar"));
add(P("A partir das causas priorizadas, construiu-se o Diagrama de Ishikawa (Figura 2), tendo como efeito as falhas na preparação dos *kits*. As causas de maior peso concentram-se nas categorias método (separação de memória e ausência de procedimento) e meio ambiente (depósito sem endereçamento e com itens obsoletos misturados), o que confirma o resultado do Pareto. As categorias mão de obra e medição revelam fatores que sustentam o problema ao longo do tempo: o conhecimento tácito concentrado no proprietário e a inexistência de indicadores, que impede a empresa de perceber a frequência real das falhas."));
add(caption("Figura 2 - Diagrama de Ishikawa das falhas na preparação dos *kits*"));
add(image("ishikawa.png", 560, 348));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Para aprofundar a principal causa do Pareto, aplicou-se a técnica dos 5 Porquês (Quadro 4). A cadeia causal mostra que a ausência de *checklist* é sintoma de uma causa mais profunda: a falta de trabalho padronizado em uma empresa que trata cada evento como um projeto único, embora a maior parte dos *kits* se repita entre eventos de mesmo tipo."));
add(caption("Quadro 4 - Aplicação dos 5 Porquês à principal causa de falha"));
add(table([2600, 6470], ["Pergunta", "Resposta"], [
  ["1. Por que itens são esquecidos?", "Porque a separação é feita de memória, sem conferência sistemática"],
  ["2. Por que a separação é feita de memória?", "Porque não existe lista formal de materiais por tipo de evento"],
  ["3. Por que não existe lista formal?", "Porque cada evento é tratado como um projeto único, montado a partir da experiência do proprietário"],
  ["4. Por que cada evento é tratado como único?", "Porque o conhecimento sobre os *kits* nunca foi documentado e permanece tácito"],
  ["5. Por que o conhecimento não foi documentado?", "**Causa raiz:** ausência de trabalho padronizado e de gestão do processo de preparação"],
], { boldFirstCol: false }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Por fim, aplicou-se a FMEA de processo às sete etapas da preparação (Tabela 3), o que permitiu avaliar não apenas a frequência das falhas, mas também sua gravidade e a capacidade dos controles atuais de detectá-las."));
add(caption("Tabela 3 - FMEA do processo de preparação dos *kits* (situação atual)"));
add(table([1250, 1500, 1550, 400, 1650, 400, 1290, 400, 630],
  ["Etapa", "Modo de falha", "Efeito", "S", "Causa", "O", "Controle atual", "D", "NPR"], [
  ["Definição da lista", "Lista de materiais incompleta", "Falta de item no evento; show incompleto ou atrasado", "8", "Lista feita de memória, sem padrão por tipo de evento", "7", "Conferência verbal do proprietário", "7", "392"],
  ["Separação de acessórios", "Esquecimento de item pequeno (fitas, conectores, ferramentas)", "Improviso ou deslocamento emergencial", "6", "Itens sem local fixo e sem *checklist*", "7", "Nenhum controle formal", "8", "336"],
  ["Programação", "Preparação concentrada na véspera", "Pressa, fadiga e aumento de erros", "6", "Agenda centralizada e sem nivelamento", "8", "Agenda pessoal do proprietário", "6", "288"],
  ["Localização dos itens", "Item não localizado ou separado errado", "Tempo de separação elevado; troca de material", "5", "Depósito sem endereçamento; obsoletos misturados", "8", "Busca visual", "5", "200"],
  ["Separação de insumos", "Quantidade insuficiente de *gerbs* ou detonadores", "Efeito não executado durante o show", "9", "Sem lista por roteiro e sem estoque mínimo", "4", "Contagem do proprietário", "5", "180"],
  ["Conferência e carga", "Equipamento com defeito carregado", "Falha no disparo durante o show", "9", "Ausência de teste funcional antes da carga", "3", "Teste apenas no local do evento", "6", "162"],
  ["Itens de segurança", "Extintor ou EPI não carregado", "Risco sem meio de resposta em caso de incidente", "10", "Itens não vinculados ao *kit*", "2", "Lembrança do *blaster*", "6", "120"],
], { size: 16, centerCols: [3, 5, 7, 8] }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Os três maiores NPR referem-se à definição da lista de materiais (392), ao esquecimento de itens pequenos (336) e à preparação concentrada na véspera (288). Esse resultado converge com o Pareto e com o Ishikawa, pois os dois primeiros modos de falha decorrem da ausência de padronização e o terceiro está ligado à sobrecarga do proprietário. A FMEA, porém, acrescenta uma informação que as demais ferramentas não captaram: a ausência de itens de segurança tem NPR baixo (120), mas severidade máxima (S = 10), e a falha de equipamento no disparo tem severidade 9. Seguindo a recomendação da AIAG e VDA (2019), esses modos de falha foram tratados como prioritários independentemente do NPR, porque um único evento pode causar dano físico ao público ou à equipe."));

add(H2("4.4. Melhorar"));
add(P("As ações de melhoria foram desenhadas para atacar as causas raízes confirmadas e reduzir o NPR dos modos de falha. O Quadro 5 apresenta o plano de ação 5W2H, com o modo de falha tratado e o NPR atual e esperado de cada ação."));
add(caption("Quadro 5 - Plano de ação 5W2H das propostas de melhoria"));
add(table([1700, 1500, 1350, 1150, 2570, 800], ["O quê", "Por quê (NPR atual → esperado)", "Onde / Quando", "Quem", "Como", "Quanto"], [
  ["*Checklists* padronizados por tipo de evento", "Lista incompleta (392 → 72) e item pequeno esquecido (336 → 36)", "Área de carga; antes do próximo evento", "Proprietário e equipe", "Converter o conhecimento tácito em listas plastificadas por família de evento, com quantidades, fotos e campo de assinatura; caixa fixa de acessórios", "R$ 10,00"],
  ["Programa 5S com endereçamento visual", "Item não localizado (200 → 60)", "Depósito de Monte Alegre; mutirão de um dia", "Equipe e proprietário", "Descartar obsoletos, agrupar por família, posicionar itens de alto giro à altura dos olhos e etiquetar locais", "R$ 50,00"],
  ["Quadro *Kanban* de programação (*heijunka*)", "Preparação na véspera (288 → 72)", "Base operacional; revisão às segundas-feiras", "Proprietário", "Cartão por evento confirmado, alocado de segunda a quinta-feira; sexta-feira apenas para conferência e carga", "R$ 60,00"],
  ["Teste funcional antes da carga", "Equipamento com defeito (162 → 54)", "Depósito; na véspera", "Auxiliar", "Testar máquinas de disparo e continuidade dos cabos; registrar no *checklist*", "Sem custo"],
  ["Itens de segurança como bloqueio de saída", "Extintor ou EPI ausente (120 → 20)", "Área de carga; a cada evento", "Auxiliar", "Veículo só é liberado com extintor e EPI conferidos e assinados", "Sem custo"],
  ["Controle de estoque com ponto de pedido", "Insumos insuficientes (180 → 54)", "Escritório; a cada retorno de evento", "Proprietário", "Planilha com estoque mínimo; baixa conforme o *checklist* do evento", "Sem custo"],
], { size: 16 }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Os valores de NPR esperado foram obtidos mantendo a severidade, que só se altera com mudanças no próprio serviço, e reduzindo a ocorrência e a detecção de acordo com o efeito de cada ação. A padronização por *checklists* reduz a ocorrência porque elimina a dependência da memória, e a conferência assinada melhora a detecção. Com o investimento total de R$ 120,00 em materiais, espera-se atingir as metas do *Project Charter*: proporção de eventos com item faltante inferior a 5%, tempo de separação de até 35 minutos e até 8 deslocamentos por *kit*. A redução de 20 minutos por *kit* representa cerca de 33 horas de trabalho poupadas por ano. A Tabela 4 compara os custos da qualidade atuais com os esperados após as melhorias."));
add(caption("Tabela 4 - Custos da qualidade atuais e esperados após as melhorias (R$/ano)"));
add(table([3870, 1700, 1700, 1800], ["Categoria", "Atual", "Esperado", "Variação"], [
  ["Prevenção (*checklists*, 5S e *Kanban*)", "0,00", "120,00", "+120,00"],
  ["Avaliação (auditorias mensais 5S)", "0,00", "300,00", "+300,00"],
  ["Falhas internas (retrabalho na separação)", "1.500,00", "250,00", "−1.250,00"],
  ["Falhas externas (deslocamentos emergenciais)", "2.400,00", "800,00", "−1.600,00"],
  ["**Total**", "**3.900,00**", "**1.470,00**", "**−2.430,00**"],
], { centerCols: [1, 2, 3] }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("O custo esperado de falhas externas considera até cinco ocorrências anuais, coerente com a meta de menos de 5% dos eventos. Mesmo com a criação de custos de prevenção e avaliação, a redução esperada é de R$ 2.430,00 por ano (62%), o que confirma o argumento de Feigenbaum (1994) de que prevenir é mais econômico do que corrigir. O ganho mais relevante, entretanto, não é financeiro: a redução dos modos de falha de severidade 9 e 10 diminui o risco de acidentes em espetáculos realizados diante do público."));

add(H2("4.5. Controlar"));
add(P("Para que os ganhos esperados se mantenham, propõe-se o plano de controle do Quadro 6, que define indicadores, metas, frequência de medição, responsáveis e a reação prevista quando um indicador sair da meta."));
add(caption("Quadro 6 - Plano de controle prospectivo do processo de preparação dos *kits*"));
add(table([2200, 2300, 900, 1200, 1000, 1470], ["Indicador", "Cálculo", "Meta", "Frequência", "Responsável", "Reação"], [
  ["Eventos com item faltante", "Eventos com falta ÷ eventos realizados", "< 5%", "Mensal", "Proprietário", "Revisar o *checklist* do tipo de evento"],
  ["Tempo médio de separação", "Média dos tempos registrados no *checklist*", "≤ 35 min", "Semanal", "Auxiliar", "Verificar endereçamento e itens fora do lugar"],
  ["Deslocamentos emergenciais", "Nº de viagens extras a eventos", "0", "Mensal", "Proprietário", "Análise de causa com 5 Porquês"],
  ["*Kits* preparados com antecedência", "*Kits* prontos até quinta ÷ eventos da semana", "≥ 80%", "Semanal", "Proprietário", "Realocar cartões no *Kanban*"],
  ["Conformidade do *checklist*", "*Checklists* assinados ÷ eventos", "100%", "Por evento", "Auxiliar", "Bloquear a saída do veículo"],
  ["Nota da auditoria 5S", "Média das notas dos 5 sensos (1 a 5)", "≥ 4,5", "Mensal", "Proprietário", "Mutirão de reorganização"],
], { size: 16, centerCols: [2, 3] }));
add(fonte("Fonte: Elaborado pelos autores (2026)"));
add(P("Complementarmente, propõe-se um Procedimento Operacional Padrão (POP) de preparação dos *kits*, com as seguintes etapas:"));
add(alineas([
  "na segunda-feira, alocar os eventos confirmados no quadro *Kanban*;",
  "selecionar o *checklist* correspondente ao tipo de evento e ajustar as quantidades ao roteiro;",
  "separar os itens nos locais endereçados, registrando o tempo de início e de término;",
  "testar máquinas de disparo e cabos e registrar o resultado;",
  "conferir os itens de segurança e assinar o *checklist*, condição para a liberação do veículo;",
  "no retorno do evento, dar baixa no estoque e registrar qualquer ocorrência.",
]));
add(P("À medida que os registros se acumularem, recomenda-se acompanhar o tempo de separação por meio de uma carta de controle de valores individuais (I-MR) e a proporção de eventos com falha por meio de uma carta p (MONTGOMERY, 2016). Essas cartas permitirão distinguir variações aleatórias de causas especiais e verificar estatisticamente se as melhorias produziram o efeito esperado."));

// =============== 5. CONCLUSÕES ===============
add(H1("5. Conclusões"));
add(P("Este estudo aplicou o método DMAIC ao processo de preparação dos *kits* de uma microempresa de serviços pirotécnicos e mostrou que as falhas observadas têm origem sistêmica. Na fase Definir, o SERVQUAL indicou a confiabilidade como a dimensão mais crítica (−1,9). Na fase Medir, o Pareto mostrou que três causas concentram 80,5% das ocorrências. Na fase Analisar, o Ishikawa e os 5 Porquês apontaram a ausência de trabalho padronizado como causa raiz, e a FMEA priorizou a definição da lista de materiais (NPR 392) e revelou riscos de alta severidade que não apareciam nas análises de frequência."));
add(P("As propostas formuladas, de baixo custo e compatíveis com a realidade de uma microempresa, devem reduzir a proporção de eventos com item faltante de 15% para menos de 5%, o tempo de separação de 55 para até 35 minutos e o custo anual de falhas em cerca de R$ 2.430,00. O plano de controle proposto busca sustentar esses resultados e reduzir a dependência do conhecimento tácito do proprietário. Do ponto de vista teórico, o estudo reforça que o DMAIC pode ser adaptado a serviços de pequeno porte com ferramentas simples, e que a FMEA complementa as análises de frequência ao incorporar a severidade, aspecto decisivo em atividades de risco."));
add(P("O estudo apresenta limitações. As melhorias não foram implementadas, de modo que os resultados das fases Melhorar e Controlar são estimativas. O SERVQUAL foi aplicado a apenas cinco clientes, o que impede a generalização estatística. Os registros de ocorrência eram informais, não houve análise do sistema de medição, e as notas da FMEA e os custos dependem do julgamento dos autores e de informações do proprietário. Além disso, o número de visitas foi restringido por problemas de saúde do proprietário. Como trabalhos futuros, recomenda-se implementar as ações propostas, coletar dados suficientes para a aplicação das cartas de controle e repetir o SERVQUAL com uma amostra maior, a fim de verificar o efeito das melhorias sobre a percepção dos clientes."));

// =============== REFERÊNCIAS ===============
add(L.blank());
add(new D.Paragraph({ children: L.runs("REFERÊNCIAS", { bold: true }), spacing: { line: L.LINE, lineRule: D.LineRuleType.AUTO }, keepNext: true }));
const refs = [
  "AIAG; VDA. **FMEA Handbook**: design FMEA, process FMEA, supplemental FMEA for monitoring & system response. Southfield: Automotive Industry Action Group, 2019.",
  "ANTONY, J. Six Sigma for service processes. **Business Process Management Journal**, v. 12, n. 2, p. 234-248, 2006.",
  "BRASIL. Decreto nº 10.030, de 30 de setembro de 2019. Aprova o Regulamento de Produtos Controlados. **Diário Oficial da União**, Brasília, DF, 1 out. 2019.",
  "CARPINETTI, L. C. R. **Gestão da qualidade**: conceitos e técnicas. 3. ed. São Paulo: Atlas, 2016.",
  "FEIGENBAUM, A. V. **Controle da qualidade total**. São Paulo: Makron Books, 1994.",
  "FITZSIMMONS, J. A.; FITZSIMMONS, M. J. **Administração de serviços**: operações, estratégia e tecnologia da informação. 7. ed. Porto Alegre: AMGH, 2014.",
  "GIL, A. C. **Como elaborar projetos de pesquisa**. 6. ed. São Paulo: Atlas, 2017.",
  "ISHIKAWA, K. **Controle de qualidade total**: à maneira japonesa. Rio de Janeiro: Campus, 1993.",
  "JURAN, J. M.; GRYNA, F. M. **Controle da qualidade handbook**: conceitos, políticas e filosofia da qualidade. São Paulo: Makron Books, 1993.",
  "KUMAR, M.; ANTONY, J.; DOUGLAS, A. Does size matter for Six Sigma implementation? Findings from the survey in UK SMEs. **The TQM Journal**, v. 21, n. 6, p. 623-635, 2009.",
  "LIKER, J. K. **O modelo Toyota**: 14 princípios de gestão do maior fabricante do mundo. Porto Alegre: Bookman, 2005.",
  "MONTGOMERY, D. C. **Introdução ao controle estatístico da qualidade**. 7. ed. Rio de Janeiro: LTC, 2016.",
  "OHNO, T. **O Sistema Toyota de Produção**: além da produção em larga escala. Porto Alegre: Bookman, 1997.",
  "PANDE, P. S.; NEUMAN, R. P.; CAVANAGH, R. R. **Estratégia Seis Sigma**: como a GE, a Motorola e outras grandes empresas estão aguçando seu desempenho. Rio de Janeiro: Qualitymark, 2001.",
  "PARASURAMAN, A.; ZEITHAML, V. A.; BERRY, L. L. SERVQUAL: a multiple-item scale for measuring consumer perceptions of service quality. **Journal of Retailing**, v. 64, n. 1, p. 12-40, 1988.",
  "SCHROEDER, R. G. *et al*. Six Sigma: definition and underlying theory. **Journal of Operations Management**, v. 26, n. 4, p. 536-554, 2008.",
  "STAMATIS, D. H. **Failure mode and effect analysis**: FMEA from theory to execution. 2. ed. Milwaukee: ASQ Quality Press, 2003.",
  "WERKEMA, C. **Criando a cultura Lean Seis Sigma**. 3. ed. Rio de Janeiro: Elsevier, 2012.",
  "YIN, R. K. **Estudo de caso**: planejamento e métodos. 5. ed. Porto Alegre: Bookman, 2015.",
];
refs.forEach((r, i) => {
  add(new D.Paragraph({ children: L.runs(r, { size: 20 }), alignment: D.AlignmentType.LEFT, spacing: { line: L.LINE, lineRule: D.LineRuleType.AUTO, before: 0, after: 0 } }));
  if (i < refs.length - 1) add(new D.Paragraph({ children: [new D.TextRun({ text: "", size: 20 })], spacing: { line: L.LINE, lineRule: D.LineRuleType.AUTO, before: 0, after: 0 } }));
});

const doc = new D.Document({
  creator: "", title: "", description: "", lastModifiedBy: "",
  styles, numbering, sections: [section(c)],
});
save(doc, out).then(() => console.log("ok", out));
