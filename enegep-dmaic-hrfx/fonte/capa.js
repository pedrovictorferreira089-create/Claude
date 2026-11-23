const L = require("./lib");
const { D, runs, section, save, numbering, styles } = L;
const out = process.argv[2];
const AUTO = D.LineRuleType.AUTO;
const c = [];
const p = (text, o = {}) => new D.Paragraph({ children: runs(text, o), alignment: o.align || D.AlignmentType.JUSTIFIED, spacing: { line: L.LINE, lineRule: AUTO, before: o.before || 0, after: o.after || 0 } });
const center = (text, o = {}) => p(text, { ...o, align: D.AlignmentType.CENTER });

c.push(center("Aplicação do método DMAIC para redução de falhas na preparação de *kits* em uma microempresa de serviços pirotécnicos", { bold: true, size: 28, after: 240 }));
const autores = ["Pedro Victor Ferreira", "Maria Laura Lisboa", "Luiz Eduardo Câmara", "Maria Carolina Romano", "Kauã Victor Melo"];
autores.forEach((a) => c.push(center(`${a} (UFRN) [COMPLETAR: e-mail]`, { size: 20 })));
c.push(L.blank());
c.push(p("**Resumo:** Este artigo aplica o método DMAIC ao processo de preparação dos *kits* de serviço de uma microempresa de serviços pirotécnicos para eventos, localizada em Natal/RN, na qual itens eram esquecidos em cerca de 15% dos eventos e a separação de cada *kit* demandava, em média, 55 minutos. Por meio de um estudo de caso, foram aplicados o *Project Charter*, o SIPOC, o SERVQUAL, o Diagrama de Pareto, o Diagrama de Ishikawa, os 5 Porquês, a FMEA e o plano de ação 5W2H. Os resultados indicaram a confiabilidade como a dimensão de serviço mais crítica, três causas responsáveis por 80,5% das falhas e a ausência de trabalho padronizado como causa raiz. A FMEA priorizou a definição da lista de materiais (NPR 392) e evidenciou riscos de alta severidade não captados pelas análises de frequência. As propostas de melhoria e o plano de controle devem reduzir a proporção de eventos com falha para menos de 5% e o custo anual de falhas em cerca de 62%."));
c.push(p("**Palavras-chave:** DMAIC; FMEA; Qualidade em serviços; Microempresa; Serviços pirotécnicos."));
c.push(L.blank());
c.push(center("*Application of the DMAIC method to reduce failures in kit preparation at a pyrotechnic services micro-enterprise*", { bold: true, size: 24, after: 120 }));
c.push(p("**Abstract:** *This paper applies the DMAIC method to the service kit preparation process of a pyrotechnic event services micro-enterprise located in Natal, Brazil, where items were forgotten in about 15% of events and each kit took, on average, 55 minutes to prepare. Through a case study, the Project Charter, SIPOC, SERVQUAL, Pareto chart, Ishikawa diagram, 5 Whys, FMEA and the 5W2H action plan were applied. The results identified reliability as the most critical service dimension, three causes accounting for 80.5% of failures, and the lack of standardized work as the root cause. The FMEA prioritized the definition of the materials list (RPN 392) and revealed high-severity risks not captured by frequency analyses. The proposed improvements and control plan are expected to reduce the share of events with failures to less than 5% and the annual cost of failures by about 62%.*"));
c.push(p("**Keywords:** *DMAIC; FMEA; Service quality; Micro-enterprise; Pyrotechnic services.*"));

module.exports = c;
if (require.main === module) {
const doc = new D.Document({ creator: "", title: "", description: "", lastModifiedBy: "", styles, numbering, sections: [section(c)] });
save(doc, out).then(() => console.log("ok", out));
}
