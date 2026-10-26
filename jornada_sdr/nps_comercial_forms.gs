/**
 * NPS comercial da Drenesse no Google Forms.
 *
 * Como usar:
 * 1. Abra o formulário no modo de edição.
 * 2. Menu ⋮ (três pontos) > "Editor de Apps Script".
 * 3. Apague o que estiver lá, cole este arquivo inteiro e salve.
 * 4. Escolha a função "montarNpsComercial" e clique em "Executar".
 *    Na primeira vez o Google pede autorização: aceite com a conta dona do formulário.
 *
 * O script apaga as perguntas que já existirem no formulário e cria as do NPS comercial,
 * as mesmas da resposta rápida "NPS comercial" do Lever, para o teste comparar só o canal.
 */
const FORM_ID = '1JAV6lFH24Sd98N7at7KU7T53lAJrZg53YI1lWIdtO-c';

const UNIDADES = ['Petrópolis', 'Lagoa Nova', 'Capim Macio', 'Norte Shopping'];

function montarNpsComercial() {
  const form = FormApp.openById(FORM_ID);

  form.getItems().forEach(item => form.deleteItem(item));

  form.setTitle('Como foi sua avaliação na Drenesse? 💖')
    .setDescription('Obrigada por ter vindo! São 4 perguntinhas rápidas, leva menos de 1 minuto e nos ajuda a cuidar melhor de você.')
    .setCollectEmail(false)
    .setLimitOneResponsePerUser(false)
    .setAllowResponseEdits(false)
    .setProgressBar(false)
    .setConfirmationMessage('Recebemos sua resposta, obrigada! 💖 Sua consultora vai falar com você em breve.');

  // Identificação: é o que permite ligar a resposta ao card da cliente no Lever.
  form.addTextItem()
    .setTitle('Seu nome')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Seu WhatsApp (com DDD)')
    .setRequired(true);

  form.addListItem()
    .setTitle('Em qual unidade você fez a avaliação?')
    .setChoiceValues(UNIDADES)
    .setRequired(true);

  // P1: nota NPS
  form.addScaleItem()
    .setTitle('1. De 0 a 10, o quanto você indicaria a avaliação da Drenesse para uma amiga?')
    .setBounds(0, 10)
    .setLabels('Não indicaria', 'Indicaria com certeza')
    .setRequired(true);

  // P2: o plano
  form.addMultipleChoiceItem()
    .setTitle('2. O plano que a especialista te apresentou faz sentido para o que você busca?')
    .setChoiceValues(['Sim, totalmente', 'Em parte', 'Não'])
    .setRequired(true);

  // P3: motivo de não começar
  form.addMultipleChoiceItem()
    .setTitle('3. O que mais pesou para você não começar hoje?')
    .setChoiceValues([
      'Valor ou forma de pagamento',
      'Quero pensar ou conversar com alguém antes',
      'Horários e rotina',
      'Dúvida se funciona no meu caso',
      'Já fechei meu tratamento 🎉',
    ])
    .setRequired(true);

  // P4: intenção
  form.addMultipleChoiceItem()
    .setTitle('4. Se a gente resolvesse isso, você começaria nos próximos 7 dias?')
    .setChoiceValues(['Sim', 'Talvez', 'Não'])
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Quer contar mais alguma coisa sobre o seu atendimento ou sobre o que você precisa para começar?')
    .setRequired(false);

  // Uso interno: a SDR preenche pelo link pré-preenchido, a cliente não precisa mexer.
  form.addTextItem()
    .setTitle('Consultora que enviou (não precisa preencher)')
    .setRequired(false);

  Logger.log('Pronto. Link para enviar: ' + form.getPublishedUrl());
}
