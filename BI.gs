
var COOPSPORTES_BI_INSIGHTS_CONFIG_ = Object.freeze({
  GUIA_RAW: 'Respostas ao formulário 1',
  GUIA_CONFIG: 'CONFIG_BI',
  TITULO_MENU: 'BI',
  FUNCAO_GATILHO: 'coopsportesBiInsightsAoAbrir_',
  META_GERAL_PADRAO: 1000,
  MAXIMO_MODALIDADE_PADRAO: 150,
  DATA_INICIO_PADRAO: new Date(2026, 7, 17),
  DATA_FIM_PADRAO: new Date(2026, 8, 12),
  LARGURA_MODAL: 1160,
  ALTURA_MODAL: 720
});

var COOPSPORTES_BI_MODALIDADES_ = Object.freeze([
  { chave: 'clash-royale', nome: 'Clash Royale', termos: ['clash royale'] },
  { chave: 'counter-strike-2', nome: 'Counter-Strike 2', termos: ['counter-strike 2', 'counter strike 2'] },
  { chave: 'ea-fc-26', nome: 'EA FC 26', termos: ['ea fc 26'] },
  { chave: 'fortnite', nome: 'Fortnite', termos: ['fortnite'] },
  { chave: 'free-fire', nome: 'Free Fire', termos: ['free fire'] },
  { chave: 'league-of-legends', nome: 'League of Legends', termos: ['league of legends'] },
  { chave: 'sinuca', nome: 'Sinuca', termos: ['sinuca'] },
  { chave: 'valorant', nome: 'VALORANT', termos: ['valorant'] },
  { chave: 'truco', nome: 'Truco', termos: ['truco'] },
  { chave: 'xadrez', nome: 'Xadrez', termos: ['xadrez'] },
  { chave: 'dama', nome: 'Dama', termos: ['dama'] }
]);

function instalarMenuBiInsightsCoopsportes() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  if (!planilha) {
    throw new Error('Abra a planilha vinculada antes de instalar o menu BI / INSIGHTS.');
  }

  var gatilhos = ScriptApp.getProjectTriggers();
  var jaExiste = gatilhos.some(function(gatilho) {
    if (gatilho.getHandlerFunction() !== COOPSPORTES_BI_INSIGHTS_CONFIG_.FUNCAO_GATILHO) {
      return false;
    }

    try {
      return gatilho.getTriggerSourceId() === planilha.getId();
    } catch (erro) {
      return true;
    }
  });

  if (!jaExiste) {
    ScriptApp
      .newTrigger(COOPSPORTES_BI_INSIGHTS_CONFIG_.FUNCAO_GATILHO)
      .forSpreadsheet(planilha)
      .onOpen()
      .create();
  }

  coopsportesBiInsightsCriarMenu_();

  SpreadsheetApp.getUi().alert(
    'BI / INSIGHTS',
    jaExiste
      ? 'O menu já estava instalado. Recarregue a planilha se ele ainda não estiver visível.'
      : 'Menu instalado. Recarregue a planilha para confirmar a abertura automática.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function abrirBiInsightsCoopsportes() {
  try {
    var dados = coopsportesBiInsightsAnalisar_();
    var html = coopsportesBiInsightsMontarHtml_(dados)
      .setWidth(COOPSPORTES_BI_INSIGHTS_CONFIG_.LARGURA_MODAL)
      .setHeight(COOPSPORTES_BI_INSIGHTS_CONFIG_.ALTURA_MODAL);

    SpreadsheetApp.getUi().showModalDialog(html, 'DATA ANALYTICS');
  } catch (erro) {
    coopsportesBiInsightsExibirErro_(erro);
  }
}

function atualizarDadosBiInsightsCoopsportes() {
  return coopsportesBiInsightsAnalisar_();
}

function coopsportesBiInsightsAoAbrir_() {
  coopsportesBiInsightsCriarMenu_();
}

function coopsportesBiInsightsCriarMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(COOPSPORTES_BI_INSIGHTS_CONFIG_.TITULO_MENU)
    .addItem('Abrir painel executivo', 'abrirBiInsightsCoopsportes')
    .addToUi();
}

function coopsportesBiInsightsAnalisar_() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  if (!planilha) {
    throw new Error('Não foi possível identificar a planilha vinculada.');
  }

  var guiaRaw = planilha.getSheetByName(COOPSPORTES_BI_INSIGHTS_CONFIG_.GUIA_RAW);

  if (!guiaRaw) {
    throw new Error(
      'A guia RAW "' + COOPSPORTES_BI_INSIGHTS_CONFIG_.GUIA_RAW + '" não foi encontrada.'
    );
  }

  var valores = guiaRaw.getDataRange().getValues();

  if (!valores.length) {
    throw new Error('A guia RAW está vazia.');
  }

  var cabecalhos = valores[0];
  var indices = coopsportesBiInsightsLocalizarColunas_(cabecalhos);

  if (
    indices.data < 0 ||
    indices.modalidade < 0 ||
    indices.idade < 0 ||
    indices.faixaEtaria < 0
  ) {
    var ausentes = [];
    if (indices.data < 0) ausentes.push('Carimbo de data/hora');
    if (indices.modalidade < 0) ausentes.push('Qual modalidade você vai jogar?');
    if (indices.idade < 0) ausentes.push('Idade');
    if (indices.faixaEtaria < 0) ausentes.push('Faixa etária');

    throw new Error('Cabeçalho(s) obrigatório(s) não encontrado(s): ' + ausentes.join(', ') + '.');
  }

  var registros = [];
  var numerosLinhas = [];

  valores.slice(1).forEach(function(linha, indiceRelativo) {
    if (!coopsportesBiInsightsVazio_(linha[indices.data])) {
      registros.push(linha);
      numerosLinhas.push(indiceRelativo + 2);
    }
  });

  var parametros = coopsportesBiInsightsLerParametros_(planilha);
  var fuso = planilha.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var datasValidas = registros
    .map(function(linha) {
      return coopsportesBiInsightsConverterData_(linha[indices.data]);
    })
    .filter(function(data) {
      return data !== null;
    });

  var menorData = datasValidas.length
    ? new Date(Math.min.apply(null, datasValidas.map(function(data) { return data.getTime(); })))
    : null;
  var maiorData = datasValidas.length
    ? new Date(Math.max.apply(null, datasValidas.map(function(data) { return data.getTime(); })))
    : null;

  var hoje = coopsportesBiInsightsInicioDia_(new Date());
  var inicioCampanha = coopsportesBiInsightsInicioDia_(parametros.inicio);
  var fimCampanha = coopsportesBiInsightsInicioDia_(parametros.fim);
  var maiorDia = maiorData ? coopsportesBiInsightsInicioDia_(maiorData) : null;
  var dadosFuturos = Boolean(maiorDia && maiorDia.getTime() > hoje.getTime());
  var proporcaoTeste = coopsportesBiInsightsProporcaoTeste_(registros, indices.nome);
  var modoTeste = dadosFuturos || proporcaoTeste >= 0.5;
  var referencia;

  if (dadosFuturos) {
    referencia = maiorDia;
  } else if (hoje.getTime() < inicioCampanha.getTime()) {
    referencia = maiorDia && maiorDia.getTime() >= inicioCampanha.getTime()
      ? maiorDia
      : inicioCampanha;
  } else if (hoje.getTime() > fimCampanha.getTime()) {
    referencia = fimCampanha;
  } else {
    referencia = hoje;
  }

  if (referencia.getTime() < inicioCampanha.getTime()) referencia = inicioCampanha;
  if (referencia.getTime() > fimCampanha.getTime()) referencia = fimCampanha;

  var campanhaEncerrada = hoje.getTime() > fimCampanha.getTime();
  var fimUltimos7 = campanhaEncerrada
    ? referencia
    : coopsportesBiInsightsAdicionarDias_(referencia, -1);
  var inicioUltimos7 = coopsportesBiInsightsAdicionarDias_(fimUltimos7, -6);
  var fim7Anteriores = coopsportesBiInsightsAdicionarDias_(inicioUltimos7, -1);
  var inicio7Anteriores = coopsportesBiInsightsAdicionarDias_(fim7Anteriores, -6);
  var inicioAtualEfetivo = inicioUltimos7.getTime() < inicioCampanha.getTime()
    ? inicioCampanha
    : inicioUltimos7;
  var inicioAnteriorEfetivo = inicio7Anteriores.getTime() < inicioCampanha.getTime()
    ? inicioCampanha
    : inicio7Anteriores;
  var diasJanelaAtual = fimUltimos7.getTime() >= inicioAtualEfetivo.getTime()
    ? coopsportesBiInsightsDiferencaDias_(inicioAtualEfetivo, fimUltimos7) + 1
    : 0;
  var diasJanelaAnterior = fim7Anteriores.getTime() >= inicioAnteriorEfetivo.getTime()
    ? coopsportesBiInsightsDiferencaDias_(inicioAnteriorEfetivo, fim7Anteriores) + 1
    : 0;

  var mapaModalidades = coopsportesBiInsightsInicializarModalidades_();
  var mapaCooperativas = {};
  var mapaCidades = {};
  var mapaGeneros = {};
  var mapaFaixasEtarias = {
    '0 Até 15 anos': 0,
    '16 a 20 anos': 0,
    '21 a 26 anos': 0,
    '27 a 35 anos': 0,
    '36 anos ou mais': 0,
    'Não informado': 0
  };
  var contagemDiaria = {};
  var ultimos7 = 0;
  var seteAnteriores = 0;
  var datasInvalidas = 0;

  registros.forEach(function(linha) {
    var data = coopsportesBiInsightsConverterData_(linha[indices.data]);
    var dia = data ? coopsportesBiInsightsInicioDia_(data) : null;
    var noPeriodoAtual = Boolean(
      dia &&
      dia.getTime() >= inicioUltimos7.getTime() &&
      dia.getTime() <= fimUltimos7.getTime()
    );
    var noPeriodoAnterior = Boolean(
      dia &&
      dia.getTime() >= inicio7Anteriores.getTime() &&
      dia.getTime() <= fim7Anteriores.getTime()
    );

    if (dia) {
      var chaveDia = coopsportesBiInsightsChaveDia_(dia, fuso);
      contagemDiaria[chaveDia] = (contagemDiaria[chaveDia] || 0) + 1;
    } else {
      datasInvalidas += 1;
    }

    if (noPeriodoAtual) ultimos7 += 1;
    if (noPeriodoAnterior) seteAnteriores += 1;

    var modalidade = coopsportesBiInsightsResolverModalidade_(linha[indices.modalidade]);

    if (!mapaModalidades[modalidade.chave]) {
      mapaModalidades[modalidade.chave] = {
        chave: modalidade.chave,
        nome: modalidade.nome,
        total: 0,
        ultimos7: 0,
        anteriores7: 0
      };
    }

    mapaModalidades[modalidade.chave].total += 1;
    if (noPeriodoAtual) mapaModalidades[modalidade.chave].ultimos7 += 1;
    if (noPeriodoAnterior) mapaModalidades[modalidade.chave].anteriores7 += 1;

    coopsportesBiInsightsSomarGrupo_(mapaCooperativas, indices.cooperativa >= 0 ? linha[indices.cooperativa] : '');
    coopsportesBiInsightsSomarGrupo_(mapaCidades, indices.cidade >= 0 ? linha[indices.cidade] : '');
    coopsportesBiInsightsSomarGrupo_(mapaGeneros, indices.genero >= 0 ? linha[indices.genero] : '');

    var faixa = coopsportesBiInsightsResolverFaixaEtaria_(
      linha[indices.faixaEtaria]
    );
    mapaFaixasEtarias[faixa] = (mapaFaixasEtarias[faixa] || 0) + 1;
  });

  var total = registros.length;
  var quantidadeModalidades = COOPSPORTES_BI_MODALIDADES_.length;
  var capacidade = quantidadeModalidades * parametros.maximoModalidade;
  var totalDiasCampanha = Math.max(
    1,
    coopsportesBiInsightsDiferencaDias_(inicioCampanha, fimCampanha) + 1
  );
  var diasDecorridos = Math.max(
    0,
    Math.min(
      totalDiasCampanha,
      coopsportesBiInsightsDiferencaDias_(inicioCampanha, referencia) + 1
    )
  );
  var diasRestantes = Math.max(0, coopsportesBiInsightsDiferencaDias_(referencia, fimCampanha));
  var crescimentoSemanal = diasJanelaAtual === 7 && diasJanelaAnterior === 7 && seteAnteriores > 0
    ? (ultimos7 - seteAnteriores) / seteAnteriores
    : null;
  var mediaRecente = diasJanelaAtual > 0 ? ultimos7 / diasJanelaAtual : 0;
  var projecao = total + mediaRecente * diasRestantes;
  var ritmoNecessario = diasRestantes > 0
    ? Math.max(0, parametros.metaGeral - total) / diasRestantes
    : 0;
  var esperadoAteAgora = parametros.metaGeral * (diasDecorridos / totalDiasCampanha);
  var gapRitmo = total - esperadoAteAgora;

  var modalidades = Object.keys(mapaModalidades).map(function(chave) {
    var item = mapaModalidades[chave];
    var ocupacao = parametros.maximoModalidade > 0
      ? item.total / parametros.maximoModalidade
      : 0;
    var comparacaoValida = diasJanelaAtual === 7 && diasJanelaAnterior === 7;
    var crescimento = comparacaoValida && item.anteriores7 > 0
      ? (item.ultimos7 - item.anteriores7) / item.anteriores7
      : null;
    var ritmoRecente = diasJanelaAtual > 0 ? item.ultimos7 / diasJanelaAtual : 0;
    var projecaoModalidade = item.total + ritmoRecente * diasRestantes;
    var referenciaMedia = quantidadeModalidades > 0
      ? Math.ceil(parametros.metaGeral / quantidadeModalidades)
      : 0;
    var gapReferencia = Math.max(0, referenciaMedia - item.total);

    return {
      chave: item.chave,
      nome: item.nome,
      total: item.total,
      ultimos7: item.ultimos7,
      anteriores7: item.anteriores7,
      crescimento: crescimento,
      comparacaoValida: comparacaoValida,
      ritmoRecente: ritmoRecente,
      projecao: projecaoModalidade,
      referenciaMedia: referenciaMedia,
      gapReferencia: gapReferencia,
      ritmoParaReferencia: diasRestantes > 0 ? gapReferencia / diasRestantes : 0,
      ocupacao: ocupacao,
      status: coopsportesBiInsightsStatusCapacidade_(ocupacao)
    };
  }).sort(function(a, b) {
    return b.total - a.total || a.nome.localeCompare(b.nome);
  });

  var serieDiaria = coopsportesBiInsightsMontarSerieDiaria_(
    inicioCampanha,
    referencia,
    contagemDiaria,
    fuso
  );
  var cooperativas = coopsportesBiInsightsOrdenarGrupos_(mapaCooperativas, total);
  var cidades = coopsportesBiInsightsOrdenarGrupos_(mapaCidades, total);
  var cidadesMapa = coopsportesBiInsightsPrepararCidadesMapa_(cidades, total);
  var generos = coopsportesBiInsightsOrdenarGrupos_(mapaGeneros, total);
  var faixasEtarias = Object.keys(mapaFaixasEtarias).map(function(nome) {
    return {
      nome: nome,
      total: mapaFaixasEtarias[nome],
      participacao: total > 0 ? mapaFaixasEtarias[nome] / total : 0
    };
  });
  var qualidade = coopsportesBiInsightsAvaliarQualidade_(
    registros,
    indices,
    datasInvalidas,
    referencia,
    numerosLinhas
  );

  var metricas = {
    total: total,
    metaGeral: parametros.metaGeral,
    maximoModalidade: parametros.maximoModalidade,
    quantidadeModalidades: quantidadeModalidades,
    capacidade: capacidade,
    gapCapacidade: parametros.metaGeral - capacidade,
    modalidadesMinimas: parametros.maximoModalidade > 0
      ? Math.ceil(parametros.metaGeral / parametros.maximoModalidade)
      : 0,
    vagasMediasNecessarias: quantidadeModalidades > 0
      ? Math.ceil(parametros.metaGeral / quantidadeModalidades)
      : 0,
    atingimentoMeta: parametros.metaGeral > 0 ? total / parametros.metaGeral : 0,
    utilizacaoCapacidade: capacidade > 0 ? total / capacidade : 0,
    ultimos7: ultimos7,
    seteAnteriores: seteAnteriores,
    crescimentoSemanal: crescimentoSemanal,
    mediaRecente: mediaRecente,
    diasRestantes: diasRestantes,
    diasDecorridos: diasDecorridos,
    totalDiasCampanha: totalDiasCampanha,
    projecao: projecao,
    projecaoAtingimento: parametros.metaGeral > 0 ? projecao / parametros.metaGeral : 0,
    ritmoNecessario: ritmoNecessario,
    esperadoAteAgora: esperadoAteAgora,
    gapRitmo: gapRitmo,
    cooperativasAtivas: cooperativas.filter(function(item) { return item.nome !== 'Não informado'; }).length,
    cidadesAtivas: cidades.filter(function(item) { return item.nome !== 'Não informado'; }).length
  };

  var contexto = {
    inicio: coopsportesBiInsightsFormatarData_(inicioCampanha, fuso),
    fim: coopsportesBiInsightsFormatarData_(fimCampanha, fuso),
    referencia: coopsportesBiInsightsFormatarData_(referencia, fuso),
    menorRegistro: menorData ? coopsportesBiInsightsFormatarDataHora_(menorData, fuso) : 'Sem data',
    maiorRegistro: maiorData ? coopsportesBiInsightsFormatarDataHora_(maiorData, fuso) : 'Sem data',
    modoTeste: modoTeste,
    dadosFuturos: dadosFuturos,
    rotuloModo: modoTeste ? 'CENÁRIO DE TESTE' : 'DADOS DE PRODUÇÃO'
  };

  var insights = coopsportesBiInsightsGerarInsights_(
    metricas,
    modalidades,
    cooperativas,
    qualidade,
    contexto
  );
  var insightsTemporais = coopsportesBiInsightsGerarInsightsTemporais_(
    metricas,
    modalidades,
    cooperativas,
    qualidade,
    contexto,
    serieDiaria
  );

  return {
    geradoEm: Utilities.formatDate(new Date(), fuso, 'dd/MM/yyyy HH:mm:ss'),
    planilha: planilha.getName(),
    contexto: contexto,
    metricas: metricas,
    serieDiaria: serieDiaria,
    modalidades: modalidades,
    cooperativas: cooperativas.slice(0, 8),
    cidades: cidades.slice(0, 8),
    cidadesMapa: cidadesMapa,
    generos: generos,
    faixasEtarias: faixasEtarias,
    qualidade: qualidade,
    insights: insights,
    insightsTemporais: insightsTemporais
  };
}

function coopsportesBiInsightsLerParametros_(planilha) {
  var parametros = {
    inicio: new Date(COOPSPORTES_BI_INSIGHTS_CONFIG_.DATA_INICIO_PADRAO.getTime()),
    fim: new Date(COOPSPORTES_BI_INSIGHTS_CONFIG_.DATA_FIM_PADRAO.getTime()),
    metaGeral: COOPSPORTES_BI_INSIGHTS_CONFIG_.META_GERAL_PADRAO,
    maximoModalidade: COOPSPORTES_BI_INSIGHTS_CONFIG_.MAXIMO_MODALIDADE_PADRAO
  };
  var guia = planilha.getSheetByName(COOPSPORTES_BI_INSIGHTS_CONFIG_.GUIA_CONFIG);

  if (!guia) return parametros;

  var ultimaLinha = guia.getLastRow();
  if (ultimaLinha < 1) return parametros;

  var valores = guia.getRange(1, 1, ultimaLinha, 2).getValues();

  valores.forEach(function(linha) {
    var chave = coopsportesBiInsightsNormalizar_(linha[0]);
    var valor = linha[1];

    if (chave === 'inicio das inscricoes') {
      parametros.inicio = coopsportesBiInsightsConverterData_(valor) || parametros.inicio;
    } else if (chave === 'encerramento das inscricoes') {
      parametros.fim = coopsportesBiInsightsConverterData_(valor) || parametros.fim;
    } else if (chave === 'meta geral') {
      parametros.metaGeral = coopsportesBiInsightsNumero_(valor, parametros.metaGeral);
    } else if (chave === 'maximo por modalidade') {
      parametros.maximoModalidade = coopsportesBiInsightsNumero_(valor, parametros.maximoModalidade);
    }
  });

  if (parametros.fim.getTime() < parametros.inicio.getTime()) {
    throw new Error('A data final da campanha é anterior à data inicial na guia CONFIG_BI.');
  }

  if (parametros.metaGeral <= 0 || parametros.maximoModalidade <= 0) {
    throw new Error('A meta geral e o máximo por modalidade devem ser maiores que zero.');
  }

  return parametros;
}

function coopsportesBiInsightsLocalizarColunas_(cabecalhos) {
  return {
    data: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Carimbo de data/hora'),
    vinculoParticipante: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Você é:'),
    modalidade: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Qual modalidade você vai jogar?'),
    nome: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Nome completo do(a) jogador(a):'),
    telefone: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Número do WhatsApp do(a) jogador(a):'),
    email: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'E-mail do(a) jogador(a):'),
    nascimento: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Data de nascimento do(a) jogador(a):'),
    idade: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Idade'),
    faixaEtaria: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Faixa etária'),
    cpf: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'CPF do(a) jogador(a):'),
    termoAutorizacao: coopsportesBiInsightsIndiceCabecalho_(
      cabecalhos,
      'Termo de Autorização do Responsável (somente se for menor de idade):'
    ),
    nomeResponsavel: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Nome completo do cooperado/funcionário:'),
    telefoneResponsavel: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Número do WhatsApp do cooperado/funcionário:'),
    cpfResponsavel: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'CPF do cooperado/funcionário:'),
    cooperativaResponsavel: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Cooperativa do cooperado/funcionário:'),
    genero: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Gênero do(a) jogador(a):'),
    cidade: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Cidade:'),
    escolaridade: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Nível de escolaridade do(a) jogador(a):'),
    deficiencia: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Possui alguma deficiência? Se sim, qual?'),
    cooperativa: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, 'Cooperativa do jogador:'),
    id: coopsportesBiInsightsIndiceCabecalho_(cabecalhos, '_COOPSPORTES_ID')
  };
}

function coopsportesBiInsightsIndiceCabecalho_(cabecalhos, esperado) {
  var alvo = coopsportesBiInsightsNormalizar_(esperado);

  for (var indice = 0; indice < cabecalhos.length; indice += 1) {
    if (coopsportesBiInsightsNormalizar_(cabecalhos[indice]) === alvo) return indice;
  }

  return -1;
}

function coopsportesBiInsightsInicializarModalidades_() {
  var mapa = {};

  COOPSPORTES_BI_MODALIDADES_.forEach(function(modalidade) {
    mapa[modalidade.chave] = {
      chave: modalidade.chave,
      nome: modalidade.nome,
      total: 0,
      ultimos7: 0,
      anteriores7: 0
    };
  });

  return mapa;
}

function coopsportesBiInsightsResolverModalidade_(valor) {
  var normalizado = coopsportesBiInsightsNormalizar_(valor);

  for (var indice = 0; indice < COOPSPORTES_BI_MODALIDADES_.length; indice += 1) {
    var modalidade = COOPSPORTES_BI_MODALIDADES_[indice];
    var encontrou = modalidade.termos.some(function(termo) {
      return normalizado.indexOf(termo) >= 0;
    });

    if (encontrou) return modalidade;
  }

  var nome = coopsportesBiInsightsVazio_(valor) ? 'Não informado' : String(valor);

  return {
    chave: 'outra-' + coopsportesBiInsightsNormalizar_(nome).replace(/[^a-z0-9]+/g, '-'),
    nome: nome
  };
}

function coopsportesBiInsightsSomarGrupo_(mapa, valor) {
  var nome = coopsportesBiInsightsVazio_(valor) ? 'Não informado' : String(valor).trim();
  var chave = coopsportesBiInsightsNormalizar_(nome) || 'nao-informado';

  if (!mapa[chave]) mapa[chave] = { nome: nome, total: 0 };
  mapa[chave].total += 1;
}

function coopsportesBiInsightsOrdenarGrupos_(mapa, totalGeral) {
  return Object.keys(mapa).map(function(chave) {
    return {
      nome: mapa[chave].nome,
      total: mapa[chave].total,
      participacao: totalGeral > 0 ? mapa[chave].total / totalGeral : 0
    };
  }).sort(function(a, b) {
    return b.total - a.total || a.nome.localeCompare(b.nome);
  });
}

function coopsportesBiInsightsMontarSerieDiaria_(inicio, fim, contagem, fuso) {
  var serie = [];
  var cursor = coopsportesBiInsightsInicioDia_(inicio);
  var limite = coopsportesBiInsightsInicioDia_(fim);

  while (cursor.getTime() <= limite.getTime()) {
    var chave = coopsportesBiInsightsChaveDia_(cursor, fuso);
    serie.push({
      dataISO: chave,
      data: Utilities.formatDate(cursor, fuso, 'dd/MM'),
      dataCompleta: Utilities.formatDate(cursor, fuso, 'dd/MM/yyyy'),
      total: contagem[chave] || 0
    });
    cursor = coopsportesBiInsightsAdicionarDias_(cursor, 1);
  }

  return serie;
}

function coopsportesBiInsightsAvaliarQualidade_(
  registros,
  indices,
  datasInvalidas,
  referencia,
  numerosLinhas
) {
  var campos = [
    { nome: 'Nome', indice: indices.nome },
    { nome: 'Telefone', indice: indices.telefone },
    { nome: 'E-mail', indice: indices.email },
    { nome: 'Nascimento', indice: indices.nascimento },
    { nome: 'Idade', indice: indices.idade },
    { nome: 'Faixa etária', indice: indices.faixaEtaria },
    { nome: 'CPF', indice: indices.cpf },
    { nome: 'Gênero', indice: indices.genero },
    { nome: 'Cidade', indice: indices.cidade },
    { nome: 'Cooperativa', indice: indices.cooperativa },
    { nome: 'Modalidade', indice: indices.modalidade }
  ];
  var faltantesPorCampo = {};
  var preenchidos = 0;
  var possiveis = registros.length * campos.length;
  var linhasIncompletas = 0;
  var mapaProblemasPorLinha = {};

  function obterNumeroLinha(indiceRegistro) {
    return numerosLinhas && numerosLinhas[indiceRegistro]
      ? numerosLinhas[indiceRegistro]
      : indiceRegistro + 2;
  }

  function registrarProblema(indiceRegistro, nivel, codigo, descricao) {
    var numeroLinha = obterNumeroLinha(indiceRegistro);
    var chaveLinha = String(numeroLinha);
    if (!mapaProblemasPorLinha[chaveLinha]) {
      mapaProblemasPorLinha[chaveLinha] = {
        linha: numeroLinha,
        problemas: []
      };
    }

    var jaRegistrado = mapaProblemasPorLinha[chaveLinha].problemas.some(function(item) {
      return item.codigo === codigo;
    });
    if (!jaRegistrado) {
      mapaProblemasPorLinha[chaveLinha].problemas.push({
        nivel: nivel,
        codigo: codigo,
        descricao: descricao
      });
    }
  }

  campos.forEach(function(campo) {
    faltantesPorCampo[campo.nome] = 0;
  });

  registros.forEach(function(linha, indiceRegistro) {
    var incompleta = false;
    var camposAusentes = [];

    campos.forEach(function(campo) {
      var ausente = campo.indice < 0 || coopsportesBiInsightsVazio_(linha[campo.indice]);

      if (ausente) {
        faltantesPorCampo[campo.nome] += 1;
        incompleta = true;
        camposAusentes.push(campo.nome);
      } else {
        preenchidos += 1;
      }
    });

    if (incompleta) {
      linhasIncompletas += 1;
      registrarProblema(
        indiceRegistro,
        'warning',
        'campos-incompletos',
        'Campos críticos incompletos: ' + camposAusentes.join(', ')
      );
    }
  });

  var chavesDuplicidade = {};
  var cpfs = {};
  var emails = {};
  var telefones = {};
  var linhasPorCadastro = {};
  var linhasPorCpf = {};
  var linhasPorEmail = {};
  var linhasPorTelefone = {};
  var cpfInvalido = 0;
  var emailInvalido = 0;
  var telefoneInvalido = 0;
  var nascimentoInvalido = 0;
  var idadeInconsistente = 0;
  var menoresTotal = 0;
  var menoresViaResponsavel = 0;
  var menoresSujeitosAoTermo = 0;
  var menoresComAutorizacao = 0;
  var menoresSemAutorizacao = 0;
  var termosNoFluxoResponsavel = 0;
  var adultosComAutorizacao = 0;
  var menoresSujeitosPorModalidade = {};

  function adicionarLinhaAoGrupo(mapa, chave, indiceRegistro) {
    if (!chave) return;
    if (!mapa[chave]) mapa[chave] = [];
    mapa[chave].push({
      indiceRegistro: indiceRegistro,
      linha: obterNumeroLinha(indiceRegistro)
    });
  }

  registros.forEach(function(linha, indiceRegistro) {
    var nome = indices.nome >= 0 ? coopsportesBiInsightsNormalizar_(linha[indices.nome]) : '';
    var telefone = indices.telefone >= 0 ? coopsportesBiInsightsSomenteDigitos_(linha[indices.telefone]) : '';
    var email = indices.email >= 0 ? String(linha[indices.email] || '').trim().toLowerCase() : '';
    var cpf = indices.cpf >= 0 ? coopsportesBiInsightsSomenteDigitos_(linha[indices.cpf]) : '';
    var chave = nome + '|' + telefone + '|' + email;

    if (nome || telefone || email) {
      chavesDuplicidade[chave] = (chavesDuplicidade[chave] || 0) + 1;
      adicionarLinhaAoGrupo(linhasPorCadastro, chave, indiceRegistro);
    }
    if (cpf) {
      cpfs[cpf] = (cpfs[cpf] || 0) + 1;
      adicionarLinhaAoGrupo(linhasPorCpf, cpf, indiceRegistro);
    }
    if (email) {
      emails[email] = (emails[email] || 0) + 1;
      adicionarLinhaAoGrupo(linhasPorEmail, email, indiceRegistro);
    }
    if (telefone) {
      telefones[telefone] = (telefones[telefone] || 0) + 1;
      adicionarLinhaAoGrupo(linhasPorTelefone, telefone, indiceRegistro);
    }

    if (cpf && !coopsportesBiInsightsCpfValido_(cpf)) {
      cpfInvalido += 1;
      registrarProblema(indiceRegistro, 'warning', 'cpf-invalido', 'CPF com dígito verificador inválido');
    }
    if (email && !coopsportesBiInsightsEmailValido_(email)) {
      emailInvalido += 1;
      registrarProblema(indiceRegistro, 'warning', 'email-invalido', 'E-mail com formato inválido');
    }
    if (telefone && !coopsportesBiInsightsTelefoneValido_(telefone)) {
      telefoneInvalido += 1;
      registrarProblema(indiceRegistro, 'warning', 'telefone-invalido', 'Telefone fora do padrão de 10 ou 11 dígitos');
    }

    if (indices.data >= 0 && !coopsportesBiInsightsConverterData_(linha[indices.data])) {
      registrarProblema(indiceRegistro, 'warning', 'data-inscricao-invalida', 'Carimbo de data/hora inválido');
    }

    var nascimento = indices.nascimento >= 0
      ? coopsportesBiInsightsConverterData_(linha[indices.nascimento])
      : null;
    var idadeCalculada = coopsportesBiInsightsCalcularIdade_(nascimento, referencia);
    var autorizacao = indices.termoAutorizacao >= 0
      ? linha[indices.termoAutorizacao]
      : '';
    var possuiAutorizacao = !coopsportesBiInsightsVazio_(autorizacao);

    if (idadeCalculada === null || idadeCalculada < 0 || idadeCalculada > 100) {
      nascimentoInvalido += 1;
      registrarProblema(indiceRegistro, 'warning', 'nascimento-invalido', 'Nascimento ausente, futuro ou fora da faixa de 0 a 100 anos');
    } else {
      if (indices.idade >= 0 && !coopsportesBiInsightsVazio_(linha[indices.idade])) {
        var idadeInformada = Number(linha[indices.idade]);
        if (!isFinite(idadeInformada) || Math.abs(idadeInformada - idadeCalculada) > 0) {
          idadeInconsistente += 1;
          registrarProblema(indiceRegistro, 'warning', 'idade-inconsistente', 'Idade divergente da data de nascimento');
        }
      }

      if (idadeCalculada < 18) {
        menoresTotal += 1;
        var fluxoResponsavel = coopsportesBiInsightsFluxoResponsavel_(linha, indices);

        if (fluxoResponsavel) {
          menoresViaResponsavel += 1;
          if (possuiAutorizacao) termosNoFluxoResponsavel += 1;
        } else {
          menoresSujeitosAoTermo += 1;
          if (possuiAutorizacao) {
            menoresComAutorizacao += 1;
          } else {
            menoresSemAutorizacao += 1;
            registrarProblema(
              indiceRegistro,
              'critical',
              'menor-direto-sem-termo',
              'Menor em inscrição direta sem termo de autorização'
            );
          }
          var modalidadeMenor = indices.modalidade >= 0
            ? coopsportesBiInsightsResolverModalidade_(linha[indices.modalidade]).nome
            : 'Não informado';
          menoresSujeitosPorModalidade[modalidadeMenor] =
            (menoresSujeitosPorModalidade[modalidadeMenor] || 0) + 1;
        }
      } else if (possuiAutorizacao) {
        adultosComAutorizacao += 1;
      }
    }
  });

  var duplicidadeComposta = coopsportesBiInsightsResumirDuplicidades_(chavesDuplicidade);
  var duplicidadeCpf = coopsportesBiInsightsResumirDuplicidades_(cpfs);
  var duplicidadeEmail = coopsportesBiInsightsResumirDuplicidades_(emails);
  var duplicidadeTelefone = coopsportesBiInsightsResumirDuplicidades_(telefones);

  function marcarGruposRepetidos(mapa, nivel, codigo, rotulo, complemento) {
    Object.keys(mapa).forEach(function(chave) {
      var ocorrencias = mapa[chave];
      if (ocorrencias.length <= 1) return;
      var linhasGrupo = ocorrencias.map(function(item) {
        return item.linha;
      }).sort(function(a, b) {
        return a - b;
      });
      var descricao = rotulo + ' nas linhas ' + linhasGrupo.join(', ') +
        (complemento ? '. ' + complemento : '');
      ocorrencias.forEach(function(item) {
        registrarProblema(item.indiceRegistro, nivel, codigo, descricao);
      });
    });
  }

  marcarGruposRepetidos(
    linhasPorCadastro,
    'warning',
    'duplicidade-cadastro',
    'Possível duplicidade forte',
    ''
  );
  marcarGruposRepetidos(
    linhasPorCpf,
    'info',
    'cpf-repetido',
    'CPF repetido',
    ''
  );
  marcarGruposRepetidos(
    linhasPorEmail,
    'info',
    'email-repetido',
    'E-mail repetido',
    ''
  );
  marcarGruposRepetidos(
    linhasPorTelefone,
    'info',
    'telefone-repetido',
    'Telefone repetido',
    ''
  );

  var ids = {};
  var linhasPorId = {};

  if (indices.id >= 0) {
    registros.forEach(function(linha, indiceRegistro) {
      var id = String(linha[indices.id] || '');
      if (id) {
        ids[id] = (ids[id] || 0) + 1;
        adicionarLinhaAoGrupo(linhasPorId, id, indiceRegistro);
      }
    });
  }

  var gruposIdDuplicado = 0;
  var linhasComIdReutilizado = 0;

  Object.keys(ids).forEach(function(id) {
    if (ids[id] > 1) {
      gruposIdDuplicado += 1;
      linhasComIdReutilizado += ids[id];
    }
  });
  marcarGruposRepetidos(
    linhasPorId,
    'warning',
    'id-reutilizado',
    'Identificador interno reutilizado',
    ''
  );

  var faltantes = Object.keys(faltantesPorCampo).map(function(nome) {
    return { nome: nome, total: faltantesPorCampo[nome] };
  }).sort(function(a, b) {
    return b.total - a.total;
  });
  var rankingMenores = Object.keys(menoresSujeitosPorModalidade).map(function(nome) {
    return { nome: nome, total: menoresSujeitosPorModalidade[nome] };
  }).sort(function(a, b) {
    return b.total - a.total || a.nome.localeCompare(b.nome);
  });
  var pendenciasCriticas = menoresSemAutorizacao + cpfInvalido + nascimentoInvalido + datasInvalidas;
  var prioridadeNivel = { critical: 0, warning: 1, info: 2 };
  var problemasPorLinha = Object.keys(mapaProblemasPorLinha).map(function(chaveLinha) {
    var itemLinha = mapaProblemasPorLinha[chaveLinha];
    itemLinha.problemas.sort(function(a, b) {
      return prioridadeNivel[a.nivel] - prioridadeNivel[b.nivel] ||
        a.descricao.localeCompare(b.descricao);
    });
    var nivel = itemLinha.problemas.length ? itemLinha.problemas[0].nivel : 'info';
    return {
      linha: itemLinha.linha,
      nivel: nivel,
      problemas: itemLinha.problemas
    };
  }).sort(function(a, b) {
    return prioridadeNivel[a.nivel] - prioridadeNivel[b.nivel] || a.linha - b.linha;
  });
  var linhasPorNivel = { critical: 0, warning: 0, info: 0 };
  problemasPorLinha.forEach(function(item) {
    linhasPorNivel[item.nivel] += 1;
  });

  return {
    completude: possiveis > 0 ? preenchidos / possiveis : 1,
    linhasIncompletas: linhasIncompletas,
    datasInvalidas: datasInvalidas,
    gruposDuplicados: duplicidadeComposta.grupos,
    excedentesDuplicados: duplicidadeComposta.excedentes,
    linhasEmDuplicidade: duplicidadeComposta.linhas,
    gruposIdDuplicado: gruposIdDuplicado,
    linhasComIdReutilizado: linhasComIdReutilizado,
    faltantes: faltantes,
    pendenciasCriticas: pendenciasCriticas,
    totalLinhasComProblema: problemasPorLinha.length,
    totalLinhasComPendencia: linhasPorNivel.critical + linhasPorNivel.warning,
    linhasPorNivel: linhasPorNivel,
    problemasPorLinha: problemasPorLinha,
    validacoes: {
      cpfInvalido: cpfInvalido,
      emailInvalido: emailInvalido,
      telefoneInvalido: telefoneInvalido,
      nascimentoInvalido: nascimentoInvalido,
      idadeInconsistente: idadeInconsistente
    },
    menores: {
      total: menoresTotal,
      viaResponsavel: menoresViaResponsavel,
      sujeitosAoTermo: menoresSujeitosAoTermo,
      comAutorizacao: menoresComAutorizacao,
      semAutorizacao: menoresSemAutorizacao,
      termosNoFluxoResponsavel: termosNoFluxoResponsavel,
      adultosComAutorizacao: adultosComAutorizacao,
      conformidade: menoresSujeitosAoTermo > 0
        ? menoresComAutorizacao / menoresSujeitosAoTermo
        : 1,
      colunaDisponivel: indices.termoAutorizacao >= 0,
      porModalidade: rankingMenores
    },
    duplicidades: {
      cadastro: duplicidadeComposta,
      cpf: duplicidadeCpf,
      email: duplicidadeEmail,
      telefone: duplicidadeTelefone
    }
  };
}

function coopsportesBiInsightsSomenteDigitos_(valor) {
  if (typeof valor === 'number' && isFinite(valor)) return String(Math.round(valor));
  return String(valor === null || valor === undefined ? '' : valor).replace(/\D/g, '');
}

function coopsportesBiInsightsCpfValido_(valor) {
  var cpf = coopsportesBiInsightsSomenteDigitos_(valor);
  if (cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) return false;

  function digito(limite) {
    var soma = 0;
    for (var indice = 0; indice < limite; indice += 1) {
      soma += Number(cpf.charAt(indice)) * (limite + 1 - indice);
    }
    var resultado = (soma * 10) % 11;
    return resultado === 10 ? 0 : resultado;
  }

  return digito(9) === Number(cpf.charAt(9)) && digito(10) === Number(cpf.charAt(10));
}

function coopsportesBiInsightsEmailValido_(valor) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(valor || '').trim());
}

function coopsportesBiInsightsTelefoneValido_(valor) {
  var telefone = coopsportesBiInsightsSomenteDigitos_(valor);
  return telefone.length === 10 || telefone.length === 11;
}

function coopsportesBiInsightsCalcularIdade_(nascimento, referencia) {
  if (!(nascimento instanceof Date) || isNaN(nascimento.getTime())) return null;
  var dataReferencia = referencia instanceof Date && !isNaN(referencia.getTime())
    ? referencia
    : new Date();
  var idade = dataReferencia.getFullYear() - nascimento.getFullYear();
  var aniversarioAindaNaoOcorreu =
    dataReferencia.getMonth() < nascimento.getMonth() ||
    (dataReferencia.getMonth() === nascimento.getMonth() && dataReferencia.getDate() < nascimento.getDate());
  return idade - (aniversarioAindaNaoOcorreu ? 1 : 0);
}

function coopsportesBiInsightsFluxoResponsavel_(linha, indices) {
  var vinculo = indices.vinculoParticipante >= 0
    ? coopsportesBiInsightsNormalizar_(linha[indices.vinculoParticipante])
    : '';
  var vinculoDeDependente = vinculo.indexOf('filho ou conjuge') >= 0;
  var nomeInformado = indices.nomeResponsavel >= 0 &&
    !coopsportesBiInsightsVazio_(linha[indices.nomeResponsavel]);
  var contatoOuVinculoInformado = (
    indices.telefoneResponsavel >= 0 &&
    !coopsportesBiInsightsVazio_(linha[indices.telefoneResponsavel])
  ) || (
    indices.cpfResponsavel >= 0 &&
    !coopsportesBiInsightsVazio_(linha[indices.cpfResponsavel])
  ) || (
    indices.cooperativaResponsavel >= 0 &&
    !coopsportesBiInsightsVazio_(linha[indices.cooperativaResponsavel])
  );

  return vinculoDeDependente && nomeInformado && contatoOuVinculoInformado;
}

function coopsportesBiInsightsResumirDuplicidades_(mapa) {
  var resumo = { grupos: 0, excedentes: 0, linhas: 0 };
  Object.keys(mapa).forEach(function(chave) {
    var total = mapa[chave];
    if (total > 1) {
      resumo.grupos += 1;
      resumo.excedentes += total - 1;
      resumo.linhas += total;
    }
  });
  return resumo;
}

function coopsportesBiInsightsGerarInsights_(metricas, modalidades, cooperativas, qualidade, contexto) {
  var insights = [];

  function adicionar(nivel, titulo, texto, acao) {
    insights.push({
      nivel: nivel,
      titulo: titulo,
      texto: texto,
      acao: acao
    });
  }

  if (contexto.modoTeste) {
    adicionar(
      'info',
      'Leitura em modo de teste',
      'Os indicadores usam o último registro como data de referência para permitir a simulação temporal.',
      'Quando as inscrições reais começarem, o painel passará a usar a data corrente automaticamente.'
    );
  }

  if (metricas.gapCapacidade > 0) {
    adicionar(
      'critical',
      'Meta maior que a capacidade disponível',
      'A meta é de ' + coopsportesBiInsightsInteiro_(metricas.metaGeral) +
        ', mas ' + metricas.quantidadeModalidades + ' modalidades com ' +
        coopsportesBiInsightsInteiro_(metricas.maximoModalidade) + ' vagas comportam apenas ' +
        coopsportesBiInsightsInteiro_(metricas.capacidade) + ' inscrições.',
      'Adotar pelo menos ' + metricas.modalidadesMinimas +
        ' modalidades ou elevar a média para ' + metricas.vagasMediasNecessarias + ' vagas por modalidade.'
    );
  }

  if (metricas.diasRestantes > 0 && metricas.projecao < metricas.metaGeral) {
    adicionar(
      metricas.projecaoAtingimento < 0.5 ? 'critical' : 'warning',
      'Projeção abaixo da meta',
      'Mantendo ' + coopsportesBiInsightsDecimal_(metricas.mediaRecente, 1) +
        ' inscrições por dia, a projeção é de ' +
        coopsportesBiInsightsInteiro_(Math.round(metricas.projecao)) + ' até o encerramento.',
      'O ritmo necessário para a meta seria ' +
        coopsportesBiInsightsDecimal_(metricas.ritmoNecessario, 1) + ' inscrições por dia.'
    );
  }

  if (metricas.crescimentoSemanal === null) {
    adicionar(
      'info',
      'Semana sem base comparável',
      'Não há inscrições suficientes no período anterior para calcular uma variação percentual confiável.',
      'Acompanhar novamente após completar duas janelas consecutivas de sete dias.'
    );
  } else if (metricas.crescimentoSemanal >= 0.1) {
    adicionar(
      'positive',
      'Aceleração semanal positiva',
      'Os últimos sete dias cresceram ' +
        coopsportesBiInsightsPercentual_(metricas.crescimentoSemanal) + ' em relação aos sete dias anteriores.',
      'Identificar quais ações e canais geraram o aumento e reforçá-los na próxima semana.'
    );
  } else if (metricas.crescimentoSemanal <= -0.1) {
    adicionar(
      'warning',
      'Desaceleração nas inscrições',
      'Os últimos sete dias recuaram ' +
        coopsportesBiInsightsPercentual_(Math.abs(metricas.crescimentoSemanal)) + '.',
      'Reforçar divulgação e contato com cooperativas antes que a queda comprometa a projeção.'
    );
  } else {
    adicionar(
      'info',
      'Ritmo semanal estável',
      'A variação semanal ficou dentro da faixa de estabilidade de ±10%.',
      'Monitorar o volume absoluto, pois estabilidade não significa necessariamente ritmo suficiente para a meta.'
    );
  }

  var semInscricoes = modalidades.filter(function(item) { return item.total === 0; });

  if (semInscricoes.length) {
    adicionar(
      'warning',
      'Modalidades sem inscrições',
      semInscricoes.map(function(item) { return item.nome; }).join(', ') +
        (semInscricoes.length === 1 ? ' ainda não recebeu inscrição.' : ' ainda não receberam inscrições.'),
      'Criar uma ação de divulgação direcionada e validar se formulário, regulamento e comunicação estão claros.'
    );
  }

  var pertoDoLimite = modalidades.filter(function(item) { return item.ocupacao >= 0.8; });

  if (pertoDoLimite.length) {
    adicionar(
      'warning',
      'Capacidade próxima do limite',
      pertoDoLimite.map(function(item) {
        return item.nome + ' (' + coopsportesBiInsightsPercentual_(item.ocupacao) + ')';
      }).join(', ') + '.',
      'Definir lista de espera e regras de encerramento antes do preenchimento total.'
    );
  }

  if (qualidade.menores.semAutorizacao > 0) {
    adicionar(
      'critical',
      'Menores em inscrição direta sem termo',
      qualidade.menores.semAutorizacao + ' de ' + qualidade.menores.sujeitosAoTermo +
        ' menor(es) no fluxo de inscrição direta não possuem termo informado. ' +
        qualidade.menores.viaResponsavel + ' menor(es) foram cadastrados pelo responsável e não entram como pendência.',
      'Regularizar somente as inscrições diretas sem termo antes de confirmar a participação.'
    );
  }

  if (qualidade.validacoes.cpfInvalido > 0 || qualidade.validacoes.nascimentoInvalido > 0) {
    adicionar(
      'warning',
      'Dados cadastrais potencialmente incorretos',
      qualidade.validacoes.cpfInvalido + ' CPF(s) falharam na validação dos dígitos e ' +
        qualidade.validacoes.nascimentoInvalido + ' nascimento(s) ficaram fora da faixa válida.',
      'Conferir esses registros na fonte; o painel apenas sinaliza e não corrige dados automaticamente.'
    );
  }

  if (qualidade.gruposIdDuplicado > 0) {
    adicionar(
      'warning',
      'Identificador interno reutilizado',
      qualidade.gruposIdDuplicado + ' identificador interno aparece em ' +
        qualidade.linhasComIdReutilizado + ' registros.',
      'Revisar a geração do _COOPSPORTES_ID para impedir conflitos nas atualizações das modalidades.'
    );
  }

  if (qualidade.excedentesDuplicados > 0) {
    adicionar(
      'warning',
      'Possíveis inscrições duplicadas',
      qualidade.gruposDuplicados + ' grupo(s) repetem nome, telefone e e-mail, com ' +
        qualidade.excedentesDuplicados + ' registro(s) excedente(s) para conferência.',
      'Revisar os casos antes de confirmar as chaves. O BI apenas sinaliza e não exclui registros.'
    );
  }

  var mediaPorModalidade = modalidades.length > 0
    ? metricas.total / modalidades.length
    : 0;
  var menorTracao = modalidades.filter(function(item) {
    return item.total < mediaPorModalidade * 0.7;
  }).slice(-3);

  if (metricas.diasRestantes > 0 && menorTracao.length) {
    adicionar(
      'warning',
      'Modalidades com menor tração',
      menorTracao.map(function(item) {
        return item.nome + ' (' + item.total + ')';
      }).join(', ') + ' estão abaixo de 70% da média atual por modalidade.',
      'Fazer divulgação segmentada dessas modalidades e revisar se nickname, plataforma ou regras estão criando atrito no cadastro.'
    );
  }

  if (qualidade.linhasIncompletas > 0 || qualidade.datasInvalidas > 0) {
    var principalFaltante = qualidade.faltantes.filter(function(item) {
      return item.total > 0;
    })[0];
    adicionar(
      'warning',
      'Qualidade cadastral exige atenção',
      qualidade.linhasIncompletas + ' registros têm ao menos um campo crítico ausente e ' +
        qualidade.datasInvalidas + ' têm data inválida.' +
        (principalFaltante
          ? ' Principal ocorrência: ' + principalFaltante.nome + ' (' + principalFaltante.total + ').'
          : ''),
      'Corrigir primeiro os campos com maior ausência antes da montagem das chaves.'
    );
  } else {
    adicionar(
      'positive',
      'Boa completude cadastral',
      'Os campos críticos analisados apresentam ' +
        coopsportesBiInsightsPercentual_(qualidade.completude) + ' de preenchimento.',
      'Manter as validações atuais do formulário durante toda a campanha.'
    );
  }

  if (cooperativas.length && cooperativas[0].participacao >= 0.25) {
    adicionar(
      'warning',
      'Concentração elevada em uma cooperativa',
      cooperativas[0].nome + ' responde por ' +
        coopsportesBiInsightsPercentual_(cooperativas[0].participacao) + ' das inscrições.',
      'Reforçar mobilização nas demais cooperativas para reduzir dependência de uma única origem.'
    );
  }

  return insights;
}

function coopsportesBiInsightsPrepararCidadesMapa_(cidades, totalGeral) {
  var armazenamento = null;
  var propriedadesCache = {};
  var propriedadesPendentes = {};
  var geocoder = null;
  var geocoderInicializado = false;
  var prefixo = 'COOPSPORTES_BI_GEO_V2_';
  var agora = Date.now();
  var ttlFalha = 24 * 60 * 60 * 1000;

  try {
    armazenamento = PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
    propriedadesCache = armazenamento ? armazenamento.getProperties() : {};
  } catch (erroPropriedades) {
    armazenamento = null;
    propriedadesCache = {};
  }

  function obterGeocoder() {
    if (geocoderInicializado) return geocoder;
    geocoderInicializado = true;
    try {
      geocoder = Maps.newGeocoder()
        .setRegion('br')
        .setLanguage('pt-BR')
        .setBounds(-34.0, -74.2, 5.6, -33.8);
    } catch (erroGeocoder) {
      geocoder = null;
    }
    return geocoder;
  }

  function chaveCache(nome) {
    var normalizado = coopsportesBiInsightsNormalizar_(nome)
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .slice(0, 110);
    return prefixo + (normalizado || 'cidade');
  }

  function dentroDoBrasil(lat, lng) {
    return isFinite(lat) && isFinite(lng) && lat >= -34.1 && lat <= 5.7 && lng >= -74.3 && lng <= -33.7;
  }

  function resultadoEhBrasil(resultado) {
    if (!resultado || !resultado.geometry || !resultado.geometry.location) return false;
    var endereco = String(resultado.formatted_address || '').toLowerCase();
    var local = resultado.geometry.location;
    return dentroDoBrasil(Number(local.lat), Number(local.lng)) &&
      (endereco.indexOf('brasil') >= 0 || endereco.indexOf('brazil') >= 0 || !endereco);
  }

  var resultado = cidades
    .filter(function(item) {
      return item && item.nome && item.nome !== 'Não informado';
    })
    .map(function(item, indice) {
      var chave = chaveCache(item.nome);
      var cached = null;

      try {
        var textoCache = propriedadesCache[chave] || null;
        cached = textoCache ? JSON.parse(textoCache) : null;
      } catch (erroCache) {
        cached = null;
      }

      var cacheValido = cached && (
        cached.resolvida ||
        (cached.atualizadoEm && agora - Number(cached.atualizadoEm) < ttlFalha)
      );

      if (!cacheValido) {
        var resolucao = {
          resolvida: false,
          latitude: null,
          longitude: null,
          atualizadoEm: agora
        };
        var geocoderAtual = obterGeocoder();

        if (geocoderAtual) {
          try {
            var resposta = geocoderAtual.geocode(String(item.nome).trim() + ', Brasil');
            var resultados = resposta && resposta.results ? resposta.results : [];
            var melhor = null;

            for (var r = 0; r < resultados.length; r += 1) {
              if (resultadoEhBrasil(resultados[r])) {
                melhor = resultados[r];
                break;
              }
            }

            if (melhor) {
              resolucao.resolvida = true;
              resolucao.latitude = Number(melhor.geometry.location.lat);
              resolucao.longitude = Number(melhor.geometry.location.lng);
            }
          } catch (erroGeocode) {
            resolucao.resolvida = false;
          }
        }

        cached = resolucao;
        var serializado = JSON.stringify(cached);
        propriedadesCache[chave] = serializado;
        propriedadesPendentes[chave] = serializado;
      }

      return {
        nome: item.nome,
        total: item.total,
        participacao: totalGeral > 0 ? item.total / totalGeral : 0,
        ranking: indice + 1,
        latitude: cached && cached.resolvida ? Number(cached.latitude) : null,
        longitude: cached && cached.resolvida ? Number(cached.longitude) : null,
        resolvida: Boolean(cached && cached.resolvida)
      };
    });

  if (armazenamento && Object.keys(propriedadesPendentes).length) {
    try {
      armazenamento.setProperties(propriedadesPendentes, false);
    } catch (erroPersistencia) {
    }
  }

  return resultado;
}

function coopsportesBiInsightsGerarInsightsTemporais_(
  metricas,
  modalidades,
  cooperativas,
  qualidade,
  contexto,
  serieDiaria
) {
  var temporal = { passado: [], atual: [], futuro: [] };

  function adicionar(periodo, nivel, titulo, texto, acao) {
    temporal[periodo].push({
      nivel: nivel,
      titulo: titulo,
      texto: texto,
      acao: acao
    });
  }

  var serieHistorica = serieDiaria.length > 1
    ? serieDiaria.slice(0, serieDiaria.length - 1)
    : serieDiaria.slice();
  var pico = serieHistorica.slice().sort(function(a, b) {
    return b.total - a.total;
  })[0];

  if (pico) {
    adicionar(
      'passado',
      'info',
      'Pico de inscrições',
      pico.dataCompleta + ' concentrou ' + pico.total + ' inscrições, o maior volume diário observado.',
      'Cruzar esse dia com as ações de comunicação realizadas e documentar o canal com maior potencial de repetição.'
    );
  }

  if (modalidades.length) {
    adicionar(
      'passado',
      'positive',
      'Liderança histórica por modalidade',
      modalidades[0].nome + ' lidera o acumulado com ' + modalidades[0].total + ' inscrições.',
      'Preservar o canal que sustenta essa procura e usar a mensagem vencedora como teste nas modalidades de menor tração.'
    );
  }

  if (metricas.crescimentoSemanal === null) {
    adicionar(
      'passado',
      'info',
      'Histórico ainda curto para tendência semanal',
      'Ainda não existem duas janelas completas de sete dias para medir aceleração ou desaceleração com segurança.',
      'Usar o volume diário como sinal provisório e reavaliar a tendência quando a segunda semana estiver completa.'
    );
  } else {
    adicionar(
      'passado',
      metricas.crescimentoSemanal >= 0 ? 'positive' : 'warning',
      'Variação das duas últimas semanas',
      'A semana mais recente variou ' +
        coopsportesBiInsightsPercentual_(metricas.crescimentoSemanal) + ' frente à anterior.',
      metricas.crescimentoSemanal >= 0
        ? 'Identificar e repetir as ações associadas ao ganho de volume.'
        : 'Reforçar os canais que historicamente geraram os maiores picos.'
    );
  }

  var diferencaEsperado = Math.round(metricas.gapRitmo);
  adicionar(
    'atual',
    diferencaEsperado >= 0 ? 'positive' : 'warning',
    diferencaEsperado >= 0 ? 'Volume atual acima da linha esperada' : 'Volume atual abaixo da linha esperada',
    'O acumulado está ' + Math.abs(diferencaEsperado) +
      (diferencaEsperado >= 0 ? ' inscrição(ões) acima' : ' inscrição(ões) abaixo') +
      ' da trajetória linear para a meta.',
    diferencaEsperado >= 0
      ? 'Manter o ritmo e monitorar se a distribuição entre modalidades continua equilibrada.'
      : 'Acionar imediatamente as cooperativas e modalidades com menor participação.'
  );

  adicionar(
    'atual',
    qualidade.linhasPorNivel.critical > 0
      ? 'critical'
      : (qualidade.totalLinhasComPendencia > 0 ? 'warning' : 'positive'),
    qualidade.totalLinhasComPendencia > 0 ? 'Fila de regularização cadastral' : 'Cadastro sem pendências prioritárias',
    qualidade.totalLinhasComPendencia + ' linha(s) exigem tratamento: ' +
      qualidade.linhasPorNivel.critical + ' crítica(s) e ' + qualidade.linhasPorNivel.warning + ' em atenção.',
    qualidade.totalLinhasComPendencia > 0
      ? 'Tratar primeiro inscrições diretas de menores sem termo e documentos que falharam na validação.'
      : 'Manter as validações e revisar novas entradas diariamente.'
  );

  adicionar(
    'atual',
    'info',
    'Cobertura ativa da campanha',
    metricas.cooperativasAtivas + ' cooperativas e ' + metricas.cidadesAtivas + ' cidades já participam da base.',
    'Comparar o ranking de origem com a lista de cooperativas esperadas e atuar nas regiões ainda ausentes.'
  );

  var gapProjecao = Math.max(0, Math.ceil(metricas.metaGeral - metricas.projecao));
  adicionar(
    'futuro',
    gapProjecao > 0 ? 'warning' : 'positive',
    gapProjecao > 0 ? 'Projeção indica déficit ao encerramento' : 'Projeção alcança a meta',
    'A projeção atual é de ' + coopsportesBiInsightsInteiro_(Math.round(metricas.projecao)) +
      ' inscrições, ' + (gapProjecao > 0 ? gapProjecao + ' abaixo da meta.' : 'em linha com a meta.'),
    gapProjecao > 0
      ? 'Converter o déficit em metas diárias por cooperativa e acompanhar a execução até o encerramento.'
      : 'Proteger o ritmo atual e preparar controle de capacidade nas modalidades mais demandadas.'
  );

  var aceleracaoNecessaria = metricas.mediaRecente > 0
    ? metricas.ritmoNecessario / metricas.mediaRecente - 1
    : null;
  adicionar(
    'futuro',
    aceleracaoNecessaria !== null && aceleracaoNecessaria > 0 ? 'warning' : 'positive',
    'Ritmo necessário para o restante da campanha',
    'São necessárias ' + coopsportesBiInsightsDecimal_(metricas.ritmoNecessario, 1) +
      ' inscrições por dia frente ao ritmo recente de ' +
      coopsportesBiInsightsDecimal_(metricas.mediaRecente, 1) + '.',
    aceleracaoNecessaria !== null && aceleracaoNecessaria > 0
      ? 'Planejar aumento de ' + coopsportesBiInsightsPercentual_(aceleracaoNecessaria) +
        ' no ritmo diário, com responsáveis e verificação a cada 48 horas.'
      : 'Manter o ritmo e deslocar esforço para modalidades com baixa projeção.'
  );

  var prioridades = modalidades.slice().sort(function(a, b) {
    return a.projecao - b.projecao;
  }).slice(0, 3);

  if (prioridades.length) {
    adicionar(
      'futuro',
      'warning',
      'Prioridades de mobilização por modalidade',
      prioridades.map(function(item) {
        return item.nome + ' (projeção ' + coopsportesBiInsightsInteiro_(Math.round(item.projecao)) + ')';
      }).join(', ') + '.',
      'Criar campanhas segmentadas para essas modalidades e acompanhar o ganho de ritmo separadamente.'
    );
  }

  return temporal;
}

function coopsportesBiInsightsMontarHtml_(dados) {
  var json = JSON.stringify(dados).replace(/</g, '\\u003c');
  var html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root {
      --bg: #f4f4f5;
      --panel: #ffffff;
      --panel-2: #fafafa;
      --panel-3: #f1f1f2;
      --border: #dedee2;
      --border-soft: #ececef;
      --text: #11120f;
      --muted: #71717a;
      --blue: #c9ef72;
      --accent: #c9ef72;
      --accent-ink: #60751c;
      --warning: #f2c80f;
      --warning-ink: #6b5600;
      --warning-soft: rgba(242,200,15,.13);
      --warning-border: rgba(242,200,15,.62);
      --info-soft: rgba(9,9,11,.045);
      --green: #16a34a;
      --green-ink: #0f7a37;
      --green-soft: rgba(22,163,74,.10);
      --red: #dc2626;
      --red-soft: rgba(220,38,38,.09);
      --shadow: 0 8px 28px rgba(32,38,29,.06);
      --radius-lg: 18px;
      --radius-md: 14px;
      --radius-sm: 10px;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #ececef;
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, Arial, sans-serif;
      font-size: 13px;
      line-height: 1.4;
      padding: 6px;
    }
    button { font: inherit; }
    .shell {
      height: calc(100vh - 12px);
      min-height: 620px;
      padding: 0;
      background: var(--bg);
      border: 1px solid #dedee2;
      border-radius: 20px;
      box-shadow: 0 18px 48px rgba(9,9,11,.10);
      display: grid;
      grid-template-columns: 152px minmax(0, 1fr);
      overflow: hidden;
    }
    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 18px;
      margin-bottom: 20px;
      padding: 2px 2px 18px;
      color: var(--text);
      border-bottom: 1px solid var(--border-soft);
    }
    .eyebrow {
      color: var(--blue);
      font-size: 11px;
      font-weight: 800;
      letter-spacing: .14em;
      text-transform: uppercase;
      margin-bottom: 5px;
    }
    h1 { margin: 0; font-size: 24px; line-height: 1.12; letter-spacing: -.025em; font-weight: 680; }
    .subtitle { color: var(--muted); margin-top: 7px; }
    .actions { display: flex; gap: 9px; align-items: center; }
    .btn {
      border: 1px solid var(--border);
      background: var(--panel-2);
      color: var(--text);
      border-radius: var(--radius-sm);
      padding: 9px 13px;
      cursor: pointer;
      font-weight: 700;
    }
    .btn:hover { border-color: #4a4a52; background: var(--panel-3); }
    .btn:focus-visible { outline: 2px solid var(--blue); outline-offset: 2px; }
    .btn:disabled { opacity: .55; cursor: wait; }
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border: 0;
      color: #52525b;
      background: transparent;
      border-radius: 0;
      padding: 0;
      font-size: 9px;
      font-weight: 800;
      letter-spacing: .02em;
    }
    .banner {
      border: 1px solid var(--border-soft);
      border-left: 2px solid var(--blue);
      background: #101014;
      padding: 11px 14px;
      border-radius: 4px var(--radius-sm) var(--radius-sm) 4px;
      color: #c9c9d1;
      margin-bottom: 18px;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(6, minmax(150px, 1fr));
      gap: 10px;
      margin-bottom: 20px;
      background: transparent;
    }
    .card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: var(--radius-lg);
      padding: 18px;
      box-shadow: var(--shadow);
    }
    .kpis .card {
      border: 1px solid #09090B;
      border-radius: var(--radius-md);
      box-shadow: none;
      padding: 16px 15px;
    }
    .kpi-label { color: var(--muted); font-size: 11px; font-weight: 700; }
    .kpi-value { color: var(--text); font-size: 25px; font-weight: 700; letter-spacing: -.025em; margin: 8px 0 5px; }
    .kpi-detail { color: var(--muted); font-size: 11px; line-height: 1.35; }
    .tone-neutral .kpi-value, .tone-cyan .kpi-value, .tone-purple .kpi-value { color: var(--text); }
    .tone-green .kpi-value { color: #09090B; }
    .tone-red .kpi-value { color: #09090B; }
    .tone-green { border-color: #09090B !important; background: #fff; }
    .tone-red { border-color: #09090B !important; background: #fff; }
    .insights-hero {
      border: 1px solid var(--border);
      background: var(--panel);
      box-shadow: none;
      padding: 20px;
      margin-bottom: 20px;
    }
    .section-title {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 12px;
    }
    .section-title h2 { margin: 0; font-size: 17px; letter-spacing: -.015em; color: var(--text); font-weight: 650; }
    .muted { color: var(--muted); }
    .insights-summary {
      display: flex;
      gap: 7px;
      align-items: center;
      flex-wrap: wrap;
    }
    .insights-summary span {
      border: 0;
      background: transparent;
      border-radius: 0;
      padding: 0 0 0 9px;
      color: var(--muted);
      font-size: 9px;
      font-weight: 750;
      position: relative;
    }
    .insights-summary .critical-count { color: var(--red); }
    .insights-summary .warning-count { color: #9a7900; }
    .insights {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .insight {
      border: 1px solid var(--border);
      border-left-width: 3px;
      background: var(--panel);
      padding: 17px 18px;
      border-radius: 13px;
      min-height: 112px;
    }
    .insight.critical { border-left-color: var(--red); border-color: rgba(220,38,38,.52); background: rgba(220,38,38,.08); }
    .insight.warning { border-left-color: var(--warning); border-color: var(--warning-border); background: rgba(242,200,15,.13); }
    .insight.positive { border-left-color: var(--green); border-color: rgba(22,163,74,.42); background: rgba(22,163,74,.09); }
    .insight.info { border-left-color: #8b8b94; background: var(--info-soft); }
    .insight-title { color: var(--text); font-weight: 680; font-size: 14px; margin-bottom: 7px; }
    .insight-text, .insight-action { color: var(--muted); line-height: 1.45; font-size: 12px; }
    .insight-action {
      color: #3f3f46;
      margin-top: 8px;
      padding-top: 8px;
      border-top: 1px solid var(--border-soft);
    }
    .trend-strip { margin-top: 18px; padding: 12px 16px 7px; background: var(--panel); box-shadow: none; }
    .trend-strip .section-title { margin-bottom: 0; }
    .mini-chart { height: 96px; overflow: hidden; }
    .grid-bottom {
      display: grid;
      grid-template-columns: minmax(420px, 1.3fr) minmax(260px, .72fr) minmax(260px, .72fr);
      gap: 12px;
    }
    .rows { display: grid; gap: 8px; }
    .metric-row {
      display: grid;
      grid-template-columns: 130px minmax(100px, 1fr) 72px 78px;
      gap: 9px;
      align-items: center;
    }
    .metric-name { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 700; }
    .bar { height: 7px; background: #f1f1f2; border-radius: 999px; overflow: hidden; }
    .bar > span { display: block; height: 100%; border-radius: inherit; background: var(--accent); }
    .metric-number { text-align: right; font-variant-numeric: tabular-nums; }
    .trend { text-align: right; font-size: 11px; font-weight: 750; }
    .trend.up { color: var(--green); }
    .trend.down { color: var(--red); }
    .trend.flat { color: var(--muted); }
    .simple-row { display: grid; grid-template-columns: minmax(0, 1fr) 50px; gap: 10px; margin-bottom: 9px; }
    .simple-row strong { text-align: right; }
    .mini-bar { height: 5px; background: #f1f1f2; border-radius: 999px; margin-top: 4px; overflow: hidden; }
    .mini-bar span { display: block; height: 100%; background: var(--blue); border-radius: inherit; }
    .quality-score { font-size: 29px; font-weight: 700; color: var(--green); margin-bottom: 5px; }
    .quality-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 12px; }
    .quality-item { background: var(--panel-2); border: 1px solid var(--border-soft); border-radius: var(--radius-sm); padding: 10px; }
    .quality-item strong { display: block; font-size: 17px; }
    .quality-item span { color: var(--muted); font-size: 10px; }
    .footer { color: var(--muted); font-size: 10px; margin-top: 12px; text-align: right; }
    .dv-side {
      background: #fafafa;
      border-right: 1px solid var(--border);
      padding: 16px 12px;
      display: flex;
      flex-direction: column;
      min-width: 0;
    }
    .dv-brand {
      width: 38px;
      height: 38px;
      display: grid;
      place-items: center;
      border-radius: 50%;
      background: #11120f;
      color: #fff;
      font-weight: 900;
      letter-spacing: -.04em;
      margin: 0 4px 20px;
    }
    .dv-nav { display: grid; gap: 5px; }
    .dv-nav a, .dv-nav-btn {
      color: #71717a;
      text-decoration: none;
      padding: 9px 10px;
      border-radius: 9px;
      font-size: 11px;
      font-weight: 700;
      border: 0;
      background: transparent;
      text-align: left;
      width: 100%;
      cursor: pointer;
    }
    .dv-nav a:hover, .dv-nav a.active, .dv-nav-btn:hover, .dv-nav-btn.active { background: #ededf0; color: #11120f; }
    .dv-nav a.active, .dv-nav-btn.active { box-shadow: inset 3px 0 0 #11120f; }
    .dv-nav-btn:focus-visible, .dv-period-btn:focus-visible { outline: 2px solid #60751c; outline-offset: 2px; }
    .dv-watermark {
      margin-top: auto;
      padding: 0 4px 5px;
      color: #11120f;
      opacity: .16;
      font-size: 8px;
      font-weight: 700;
      letter-spacing: .08em;
      white-space: nowrap;
      user-select: none;
      pointer-events: none;
    }
    .dv-main { overflow: auto; padding: 16px 18px 12px; scroll-behavior: smooth; }
    .topbar { margin-bottom: 12px; padding: 0 0 11px; align-items: center; border-bottom: 0; }
    h1 { font-size: 22px; font-weight: 800; }
    .eyebrow { color: #60751c; font-size: 9px; letter-spacing: .12em; }
    .subtitle { margin-top: 4px; font-size: 10px; }
    .page-heading { min-width: 0; display: flex; align-items: center; }
    .page-heading h1 { font-size: 21px; line-height: 1.08; font-weight: 820; letter-spacing: .018em; text-transform: uppercase; color: #09090B; }
    .btn { background: #fff; padding: 7px 10px; font-size: 10px; border-radius: 999px; }
    .btn:hover { background: #f4f4f5; border-color: #d4d4d8; }
    .badge { background: transparent; color: #52525b; border: 0; padding: 0; border-radius: 0; display:inline-flex; align-items:center; gap:6px; font-size:9px; letter-spacing:.02em; }
    .badge::before { content:""; width:6px; height:6px; border-radius:1px; background:var(--green); flex:0 0 auto; }
    .banner { background: #fff; color: #52525b; padding: 8px 11px; margin-bottom: 10px; font-size: 10px; }
    .dv-hero-grid {
      display: grid;
      grid-template-columns: minmax(0, 1.8fr) minmax(190px, .62fr);
      gap: 10px;
      margin-bottom: 10px;
    }
    .dv-chart-card { min-height: 230px; }
    .dv-chart-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 8px; }
    .dv-big-number { font-size: 34px; font-weight: 850; letter-spacing: -.05em; line-height: 1; }
    .dv-caption { color: var(--muted); font-size: 10px; margin-top: 4px; }
    .dv-stat-stack { display: grid; gap: 10px; }
    .dv-stat {
      background: #fff;
      border: 1px solid #09090B;
      border-radius: var(--radius-lg);
      padding: 15px;
      min-height: 68px;
    }
    .dv-stat span { display: block; color: var(--muted); font-size: 9px; font-weight: 700; }
    .dv-stat strong { display: block; font-size: 22px; line-height: 1; margin: 7px 0 4px; }
    .dv-stat small { color: var(--muted); font-size: 9px; }
    .dv-stat.alert { background: #fff; border-color: #09090B; }
    .dv-stat.alert strong { color: #09090B; }
    .dv-bars {
      height: 142px;
      display: flex;
      align-items: flex-end;
      gap: 7px;
      padding: 8px 2px 0;
      border-bottom: 1px solid var(--border);
    }
    .dv-bar-col { flex: 1; min-width: 0; height: 100%; display: flex; flex-direction: column; justify-content: flex-end; align-items: stretch; gap: 5px; }
    .dv-bar-value { text-align: center; color: #71717a; font-size: 8px; }
    .dv-bar-track { flex: 1; display: flex; align-items: flex-end; background: #f4f4f5; border-radius: 12px 12px 3px 3px; overflow: hidden; }
    .dv-bar-fill { width: 100%; min-height: 3px; background: #09090B; border-radius: 12px 12px 3px 3px; transform-origin: bottom; animation: dv-rise .58s ease both; }
    .dv-bar-col:last-child .dv-bar-fill { background: #09090B; }
    .dv-bar-label { text-align: center; color: var(--muted); font-size: 8px; white-space: nowrap; }
    @keyframes dv-rise { from { transform: scaleY(.05); opacity: .35; } to { transform: scaleY(1); opacity: 1; } }
    .kpis { grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-bottom: 10px; }
    .kpis .card { padding: 13px 14px; min-height: 86px; }
    .kpi-label { font-size: 9px; }
    .kpi-value { font-size: 22px; margin: 6px 0 3px; }
    .kpi-detail { font-size: 9px; }
    .kpi-value-row, .dv-stat-value-row { display:flex; align-items:baseline; gap:3px; }
    .kpi-value-row .kpi-value { margin-right:0; }
    .kpi-arrow { display:inline-block; font-size:9px; line-height:1; font-weight:900; transform:translateY(-1px); }
    .kpi-arrow.up { color:#16A34A; }
    .kpi-arrow.down { color:#DC2626; }
    .dv-stat-value-row { margin:7px 0 4px; }
    .dv-stat-value-row strong { margin:0 !important; }
    .tone-green { border-color: #09090B !important; background: #fff; }
    .tone-red { border-color: #09090B !important; background: #fff; }
    .dv-content {
      display: grid;
      grid-template-columns: minmax(0, 1.25fr) minmax(300px, .75fr);
      gap: 10px;
    }
    .card { box-shadow: var(--shadow); padding: 14px; }
    .section-title { margin-bottom: 9px; }
    .section-title h2 { font-size: 14px; font-weight: 800; }
    .rows { gap: 6px; }
    .metric-row { grid-template-columns: 112px minmax(70px, 1fr) 62px 58px; gap: 7px; font-size: 10px; }
    .bar, .mini-bar { background: #f1f1f2; }
    .bar > span { background: var(--accent); }
    .dv-side-panels { display: grid; gap: 10px; }
    .insights-hero { margin: 0; padding: 14px; background: #fff; border-color: var(--border); box-shadow: var(--shadow); }
    .insights { display: grid; grid-template-columns: 1fr; gap: 6px; }
    .insight { padding: 10px 11px; min-height: 0; border-left-width: 3px; border-radius: var(--radius-sm); }
    .insight-title { font-size: 11px; margin-bottom: 3px; }
    .insight-text, .insight-action { font-size: 9px; line-height: 1.35; }
    .insight-action { color: #3f3f46; margin-top: 5px; padding-top: 5px; }
    .insight.critical { background: rgba(220,38,38,.08); border-color: rgba(220,38,38,.48); border-left-color: var(--red); }
    .insight.warning { background: var(--warning-soft); border-color: var(--warning-border); border-left-color: var(--warning); }
    .insight.positive { background: var(--green-soft); border-color: rgba(22,163,74,.40); border-left-color: var(--green); }
    .insight.info { background: var(--info-soft); border-left-color: #8b8b94; }
    .insights-summary span { background: transparent; border:0; border-radius:0; font-size:8px; padding:0 0 0 9px; position:relative; }
    .insights-summary span::before { content:""; position:absolute; left:0; top:50%; width:5px; height:5px; border-radius:1px; transform:translateY(-50%); background:#a1a1aa; }
    .insights-summary .critical-count::before { background:var(--red); }
    .insights-summary .warning-count::before { background:var(--warning); }
    .dv-quality { background: #fff; color: var(--text); border-color: var(--border); }
    .dv-quality .section-title h2, .dv-quality .quality-item strong { color: var(--text); }
    .dv-quality .muted, .dv-quality .quality-item span { color: var(--muted); }
    .dv-quality .quality-score { color: var(--green); font-size: 25px; }
    .dv-quality .quality-score.alert { color: var(--red); }
    .dv-quality .quality-item { background: var(--panel-2); border-color: var(--border-soft); padding: 8px; }
    .dv-quality .quality-item strong { font-size: 14px; }
    .dv-panel[hidden], .dv-temporal-panel[hidden] { display: none !important; }
    .dv-page-kpis { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-bottom: 10px; }
    .dv-page-kpis .card { min-height: 82px; padding: 12px 13px; }
    .dv-analysis-grid { display: grid; grid-template-columns: minmax(0, 1.45fr) minmax(260px, .55fr); gap: 10px; }
    .dv-table-head, .dv-mod-row {
      display: grid;
      grid-template-columns: 122px minmax(150px, 1fr) 54px 66px 76px;
      gap: 8px;
      align-items: center;
    }
    .dv-table-head { color: var(--muted); font-size: 8px; font-weight: 800; text-transform: uppercase; padding: 0 7px 7px; }
    .dv-mod-row { padding: 7px; border-top: 1px solid var(--border-soft); font-size: 10px; }
    .dv-mod-row:first-child { border-top: 0; }
    .dv-mod-row.watch { background: var(--warning-soft); }
    .dv-mod-row.capacity { background: rgba(220,38,38,.065); }
    .dv-mod-row.ok { background: rgba(22,163,74,.065); }
    .dv-mod-name { font-weight: 800; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .dv-proj-track { position: relative; height: 9px; border-radius: 999px; background: #e4e4e7; overflow: hidden; }
    .dv-proj-future, .dv-proj-current { position: absolute; left: 0; top: 0; height: 100%; border-radius: inherit; }
    .dv-proj-future { background: #d4d4d8; }
    .dv-proj-current { background: var(--blue); height: 5px; top: 2px; }
    .dv-status { justify-self: end; border-radius: 4px; padding: 3px 6px; font-size: 8px; font-weight: 800; white-space: nowrap; }
    .dv-status.ok { background: var(--green-soft); color: var(--green-ink); }
    .dv-status.watch { background: var(--warning-soft); color: var(--warning-ink); border: 1px solid var(--warning-border); }
    .dv-status.capacity { background: var(--red-soft); color: var(--red); }
    .dv-action-stack { display: grid; gap: 7px; }
    .dv-action-row { border: 1px solid var(--warning-border); background: var(--warning-soft); border-radius: var(--radius-sm); padding: 10px; }
    .dv-action-row strong { display: block; font-size: 11px; margin-bottom: 3px; }
    .dv-action-row span { color: var(--muted); font-size: 9px; line-height: 1.4; }
    .dv-action-row b { color: var(--red); }
    .dv-legend { display: flex; gap: 12px; align-items: center; color: var(--muted); font-size: 8px; }
    .dv-legend i { display: inline-block; width: 16px; height: 5px; border-radius: 999px; margin-right: 4px; vertical-align: middle; background: var(--blue); }
    .dv-legend i.future { background: #d4d4d8; height: 9px; }
    .dv-period-tabs { display: inline-flex; gap: 2px; background: #f1f1f2; border: 1px solid var(--border); border-radius: 6px; padding: 3px; }
    .dv-period-tabs-top { margin: 0 0 10px; }
    .dv-period-btn { border: 0; background: transparent; color: #71717a; border-radius: 4px; padding: 7px 14px; font-size: 10px; font-weight: 800; cursor: pointer; }
    .dv-period-btn.active { background: #11120f; color: #fff; }
    .dv-insight-hero { display: grid; grid-template-columns: minmax(0, 1fr) 240px; gap: 10px; margin-bottom: 10px; }
    .dv-time-summary { background: #fff; color: var(--text); border-top: 3px solid var(--accent); }
    .dv-time-summary .eyebrow { color: var(--accent-ink); }
    .dv-time-summary strong { display: block; font-size: 30px; line-height: 1; margin: 10px 0 7px; }
    .dv-time-summary p { color: var(--muted); font-size: 10px; margin: 0; }
    .dv-timeline { display: grid; gap: 8px; }
    .dv-temporal-card { display: grid; grid-template-columns: 10px minmax(0, 1fr) minmax(210px, .58fr); gap: 11px; align-items: start; border: 1px solid var(--border); background: #fff; border-radius: 8px; padding: 12px; }
    .dv-temporal-card.critical { background: var(--red-soft); border-color: rgba(220,38,38,.42); }
    .dv-temporal-card.warning { background: var(--warning-soft); border-color: var(--warning-border); }
    .dv-temporal-card.positive { background: var(--green-soft); border-color: rgba(22,163,74,.36); }
    .dv-temporal-card.info { background: var(--info-soft); }
    .dv-temporal-dot { width: 9px; height: 9px; border-radius: 50%; margin-top: 4px; background: #11120f; }
    .dv-temporal-card.critical .dv-temporal-dot { background: var(--red); }
    .dv-temporal-card.warning .dv-temporal-dot { background: var(--warning); }
    .dv-temporal-card.positive .dv-temporal-dot { background: var(--green); }
    .dv-temporal-copy strong { display: block; font-size: 12px; margin-bottom: 4px; }
    .dv-temporal-copy span { color: var(--muted); font-size: 10px; }
    .dv-next-step { border-left: 1px solid var(--border-soft); padding-left: 11px; color: #3f3f46; font-size: 10px; line-height: 1.45; }
    .dv-next-step b { display: block; font-size: 8px; color: #60751c; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 3px; }
    .dv-quality-grid { display: grid; grid-template-columns: minmax(0, 1.2fr) minmax(280px, .8fr); gap: 10px; }
    .dv-compliance { background: #fff; color: var(--text); border-color: var(--border); }
    .dv-compliance h2, .dv-compliance strong { color: var(--text); }
    .dv-compliance .muted { color: var(--muted); }
    .dv-progress { height: 10px; border-radius: 3px; background: #ededf0; overflow: hidden; margin: 12px 0 7px; }
    .dv-progress span { display: block; height: 100%; background: var(--accent); border-radius: inherit; }
    .dv-compliance-number { font-size: 34px; font-weight: 900; line-height: 1; color: var(--accent-ink); }
    .dv-compliance-split { display: grid; grid-template-columns: repeat(3, 1fr); gap: 7px; margin-top: 12px; }
    .dv-compliance-split div { background: var(--panel-2); border: 1px solid var(--border-soft); border-radius: 6px; padding: 8px; }
    .dv-compliance-split strong { display: block; font-size: 16px; }
    .dv-compliance-split span { color: var(--muted); font-size: 8px; }
    .dv-issue-list { display: grid; gap: 6px; }
    .dv-issue-row { display: grid; grid-template-columns: 42px minmax(0, 1fr); gap: 9px; align-items: start; border: 1px solid var(--border-soft); border-radius: 7px; padding: 9px; background: #fff; }
    .dv-issue-row.critical { background: var(--red-soft); border-color: rgba(220,38,38,.42); }
    .dv-issue-row.warning { background: var(--warning-soft); border-color: var(--warning-border); }
    .dv-issue-row.ok { background: var(--green-soft); border-color: rgba(22,163,74,.32); }
    .dv-issue-count { font-size: 19px; font-weight: 900; line-height: 1; }
    .dv-issue-row.critical .dv-issue-count { color: var(--red); }
    .dv-issue-row.warning .dv-issue-count { color: var(--warning-ink); }
    .dv-issue-row.ok .dv-issue-count { color: var(--green); }
    .dv-issue-copy strong { display: block; font-size: 10px; margin-bottom: 2px; }
    .dv-issue-copy span { color: var(--muted); font-size: 8px; line-height: 1.35; }
    .dv-issue-action { display: block; color: #3f3f46 !important; margin-top: 3px; }
    .dv-dup-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 7px; }
    .dv-dup-item { background: #fafafa; border: 1px solid var(--border-soft); border-radius: 10px; padding: 9px; }
    .dv-dup-item strong { display: block; font-size: 17px; }
    .dv-dup-item span { color: var(--muted); font-size: 8px; }
    .dv-line-audit { margin-top: 10px; }
    .dv-line-summary { display: flex; gap: 6px; flex-wrap: wrap; }
    .dv-line-summary span { border: 1px solid var(--border); background: #fafafa; border-radius: 4px; padding: 4px 7px; font-size: 8px; font-weight: 800; }
    .dv-line-summary .critical { color: var(--red); border-color: rgba(220,38,38,.42); background: var(--red-soft); }
    .dv-line-summary .info { color: #71717a; }
    .dv-line-note { margin-bottom: 8px; }
    .dv-line-scroll { max-height: 315px; overflow: auto; border: 1px solid var(--border-soft); border-radius: 11px; }
    .dv-line-head, .dv-line-row { display: grid; grid-template-columns: 62px 72px minmax(260px, 1fr) 150px; gap: 9px; align-items: start; }
    .dv-line-head { position: sticky; top: 0; z-index: 2; background: #f4f4f5; color: var(--muted); padding: 7px 9px; font-size: 8px; font-weight: 850; text-transform: uppercase; letter-spacing: .04em; }
    .dv-line-row { padding: 8px 9px; background: #fff; border-top: 1px solid var(--border-soft); font-size: 9px; }
    .dv-line-row.critical { background: rgba(220,38,38,.065); }
    .dv-line-row.warning { background: var(--warning-soft); }
    .dv-line-row.info { background: #fff; }
    .dv-line-number { font-size: 12px; font-weight: 900; font-variant-numeric: tabular-nums; }
    .dv-line-level { display: inline-flex; justify-content: center; border-radius: 4px; padding: 3px 6px; font-size: 8px; font-weight: 850; }
    .dv-line-level.critical { color: var(--red); background: var(--red-soft); border:1px solid rgba(229,72,77,.28); }
    .dv-line-level.warning { color: var(--warning-ink); background: var(--warning-soft); border: 1px solid var(--warning-border); }
    .dv-line-level.info { color: #52525b; background: #f1f1f2; border: 1px solid #dedee2; }
    .dv-problem-list { display: grid; gap: 3px; }
    .dv-problem-item { color: #3f3f46; line-height: 1.35; }
    .dv-problem-item::before { content: '•'; color: #8b8b94; margin-right: 5px; }
    .dv-line-action { color: var(--muted); line-height: 1.4; }
    .dv-line-empty { padding: 16px; color: var(--muted); text-align: center; }
    .dv-minor-bars { display: grid; gap: 7px; margin-top: 10px; }
    .dv-mini-label { color: #a1a1aa; font-size: 8px; font-weight: 800; text-transform: uppercase; letter-spacing: .06em; margin-top: 11px; }
    .dv-minor-row { display: grid; grid-template-columns: 105px 1fr 24px; align-items: center; gap: 7px; font-size: 9px; }
    .dv-minor-row .bar { height: 6px; }
    .dv-compliance .dv-minor-row .bar { background: #ededf0; }
    .footer { margin-top: 8px; padding-bottom: 2px; }
    .error { border: 1px solid var(--red); background: var(--red-soft); color: #b4232f; padding: 16px; border-radius: var(--radius-md); }
    @media (max-width: 980px) {
      .shell { grid-template-columns: 58px minmax(0, 1fr); }
      .dv-side { padding: 12px 7px; }
      .dv-brand { margin-left: 2px; width: 34px; height: 34px; }
      .dv-nav a, .dv-nav-btn { font-size: 0; padding: 9px; text-align: center; }
      .dv-nav a::first-letter, .dv-nav-btn::first-letter { font-size: 13px; }
      .dv-watermark { padding: 0 1px 3px; font-size: 6.5px; letter-spacing: .02em; text-align: center; white-space: normal; }
      .dv-hero-grid, .dv-content, .dv-analysis-grid, .dv-insight-hero, .dv-quality-grid { grid-template-columns: 1fr; }
      .kpis, .dv-page-kpis { grid-template-columns: repeat(2, 1fr); }
      .dv-temporal-card { grid-template-columns: 10px 1fr; }
      .dv-next-step { grid-column: 2; border-left: 0; border-top: 1px solid var(--border-soft); padding: 7px 0 0; }
      .dv-line-head, .dv-line-row { grid-template-columns: 58px 68px minmax(180px, 1fr); }
      .dv-line-action { grid-column: 3; }
    }
    @media (prefers-reduced-motion: reduce) { .dv-bar-fill { animation: none; } }


    .dv-time-filter { display:flex; align-items:center; gap:2px; padding:2px; background:#f4f4f5; border:1px solid var(--border); border-radius:6px; }
    .dv-time-btn { border:0; background:transparent; color:#71717a; border-radius:4px; padding:4px 8px; font-size:8px; font-weight:800; cursor:pointer; transition:background .12s ease,color .12s ease; }
    .dv-time-btn:hover { background:#e9ece6; color:#11120f; }
    .dv-time-btn.active { background:#11120f; color:#fff; }
    .dv-echart-main { width:100%; height:142px; display:block; padding:0; border-bottom:1px solid var(--border); }
    .dv-map-kpi { position:relative; overflow:hidden; padding:0 !important; background:#0b0b0d !important; border-color:#26262b !important; cursor:pointer; min-height:86px; }
    .dv-map-kpi:hover { border-color:#3b3b42 !important; }
    .metric-row, .dv-mod-row, .dv-issue-row, .dv-line-row, .insight { transition: background-color .12s ease, border-color .12s ease; }
    .metric-row:hover { background: rgba(17,18,15,.025); }
    .dv-mod-row:hover, .dv-issue-row:hover, .dv-line-row:hover { border-color: #d4d4d8; }
    .insight:hover { border-color:#d4d4d8; }
    .dv-map-kpi-copy { position:absolute; z-index:2; left:12px; top:9px; color:#fff; pointer-events:none; }
    .dv-map-kpi-copy span { display:block; color:#a9a9b1; font-size:8px; font-weight:800; letter-spacing:.06em; }
    .dv-map-kpi-copy strong { display:block; color:#fff; font-size:11px; margin-top:2px; }
    .dv-map-preview { position:absolute; inset:0; width:100%; height:100%; }
    .dv-map-layout { display:grid; grid-template-columns:minmax(0,1.55fr) minmax(250px,.45fr); gap:10px; }
    .dv-map-card { background:#09090B; border-color:#26262b; color:#f4f4f5; padding:0; overflow:hidden; }
    .dv-map-card-head { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; padding:14px 15px 0; }
    .dv-map-card-head h2 { color:#f4f4f5; margin:0; font-size:15px; }
    .dv-map-card-head p { color:#96969f; font-size:9px; margin:3px 0 0; }
    .dv-map-reset { border:1px solid #303036; background:#151517; color:#c8c8ce; border-radius:8px; padding:6px 8px; font-size:9px; cursor:pointer; }
    .dv-map-reset:hover { background:#1d1d20; }
    .dv-map-full { height:470px; width:100%; }
    .dv-city-rank { max-height:470px; overflow:auto; }
    .dv-city-row { display:grid; grid-template-columns:24px minmax(0,1fr) 42px 52px; gap:7px; align-items:center; padding:8px 10px; border-bottom:1px solid var(--border-soft); font-size:10px; }
    .dv-city-row:last-child { border-bottom:0; }
    .dv-city-row .rank { color:#a1a1aa; }
    .dv-city-row strong { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .dv-city-row .num, .dv-city-row .share { text-align:right; font-variant-numeric:tabular-nums; }
    .dv-map-empty { min-height:120px; display:grid; place-items:center; text-align:center; color:#8d8d95; font-size:10px; padding:20px; }
    @media (max-width:900px) { .dv-map-layout { grid-template-columns:1fr; } .dv-map-full { height:380px; } }
  
    .dv-line-summary .warning { color: var(--warning-ink); border-color: var(--warning-border); background: var(--warning-soft); }
    .dv-mod-row.watch { box-shadow: inset 3px 0 0 var(--warning); }
    .dv-mod-row.capacity { box-shadow: inset 3px 0 0 var(--red); }
    .dv-mod-row.ok { box-shadow: inset 3px 0 0 var(--green); }

    .kpis .card, .dv-stat { background:#fff; border-color:#09090B !important; }
    .kpis .tone-red, .kpis .tone-green { background:#fff !important; border-color:#09090B !important; }
    .kpis .tone-red .kpi-value, .kpis .tone-green .kpi-value, .dv-stat.alert strong { color:#09090B !important; }
    .dv-stat.alert { background:#fff !important; border-color:#09090B !important; }
    .insight.critical, .dv-temporal-card.critical, .dv-issue-row.critical { background:rgba(220,38,38,.08) !important; border-color:rgba(220,38,38,.50) !important; }
    .insight.warning, .dv-temporal-card.warning, .dv-issue-row.warning, .dv-mod-row.watch, .dv-action-row { background:rgba(242,200,15,.13) !important; border-color:rgba(242,200,15,.62) !important; }
    .insight.positive, .dv-temporal-card.positive, .dv-issue-row.ok, .dv-mod-row.ok { background:rgba(22,163,74,.09) !important; border-color:rgba(22,163,74,.40) !important; }
    .insight.critical, .dv-temporal-card.critical { border-left-color:#dc2626 !important; }
    .insight.warning, .dv-temporal-card.warning { border-left-color:#f2c80f !important; }
    .insight.positive, .dv-temporal-card.positive { border-left-color:#16a34a !important; }


    .dv-dark-feature,
    .dv-action-plan-card,
    .dv-quality,
    .dv-time-summary,
    .dv-compliance {
      background: #09090B !important;
      color: #FAFAFA !important;
      border-color: #27272A !important;
      box-shadow: none !important;
    }

    .dv-dark-feature .section-title h2,
    .dv-action-plan-card .section-title h2,
    .dv-quality .section-title h2,
    .dv-compliance h2,
    .dv-quality .quality-item strong,
    .dv-compliance strong,
    .dv-action-plan-card .dv-action-row strong,
    .dv-time-summary strong {
      color: #FAFAFA !important;
    }

    .dv-dark-feature .eyebrow,
    .dv-action-plan-card .eyebrow,
    .dv-time-summary .eyebrow,
    .dv-compliance .eyebrow,
    .dv-quality .eyebrow {
      color: #C9EF72 !important;
    }

    .dv-action-plan-card .dv-action-row {
      background: #111113 !important;
      border-color: #2A2A2E !important;
      color: #FAFAFA !important;
    }
    .dv-action-plan-card .dv-action-row span,
    .dv-action-plan-card .footer,
    .dv-quality .muted,
    .dv-quality .quality-item span,
    .dv-time-summary p,
    .dv-compliance .muted,
    .dv-compliance .footer,
    .dv-compliance-split span {
      color: #A1A1AA !important;
    }
    .dv-action-plan-card .dv-action-row b {
      color: #C9EF72 !important;
    }

    .dv-quality .quality-score,
    .dv-quality .quality-score.alert {
      color: #FAFAFA !important;
    }
    .dv-quality .quality-item {
      background: #111113 !important;
      border-color: #2A2A2E !important;
    }

    .dv-time-summary {
      border-top: 3px solid #C9EF72 !important;
    }

    .dv-compliance-number {
      color: #FAFAFA !important;
    }
    .dv-progress {
      background: #27272A !important;
    }
    .dv-progress span {
      background: #C9EF72 !important;
    }
    .dv-compliance-split div {
      background: #111113 !important;
      border-color: #2A2A2E !important;
    }
    .dv-minor-row,
    .dv-minor-row span,
    .dv-minor-row strong,
    .dv-compliance .dv-mini-label {
      color: #FAFAFA !important;
    }
    .dv-compliance .bar,
    .dv-compliance .mini-bar {
      background: #27272A !important;
    }
    .dv-compliance .bar > span,
    .dv-compliance .mini-bar > span {
      background: #C9EF72 !important;
    }
</style>
</head>
<body>
  <div id="app" class="shell"></div>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.6.0/dist/echarts.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/echarts-gl@2.0.9/dist/echarts-gl.min.js"></script>
  <script>
    var INITIAL_DATA = ${json};
    var CURRENT_DATA = INITIAL_DATA;
    var PAINEL_ATIVO = 'visao';
    var PERIODO_INSIGHT_ATIVO = 'futuro';
    var PERIODO_GRAFICO = 12;
    var CHARTS = {};
    var TOKEN_ANIMACAO_LINHA = 0;
    var GEO_JSON = null;
    var GEO_PROMISE = null;
    var BRAZIL_GEOJSON_URLS = [
      'https://cdn.jsdelivr.net/gh/henriquemalvar/br-geojson@main/dist/estados.geojson',
      'https://raw.githubusercontent.com/henriquemalvar/br-geojson/main/dist/estados.geojson'
    ];

    function esc(valor) {
      return String(valor === null || valor === undefined ? '' : valor)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    function inteiro(valor) {
      return Math.round(Number(valor || 0)).toLocaleString('pt-BR');
    }

    function decimal(valor, casas) {
      return Number(valor || 0).toLocaleString('pt-BR', {
        minimumFractionDigits: casas,
        maximumFractionDigits: casas
      });
    }

    function percentual(valor) {
      return Number(valor || 0).toLocaleString('pt-BR', {
        style: 'percent',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1
      });
    }

    function clamp(valor, minimo, maximo) {
      return Math.max(minimo, Math.min(maximo, valor));
    }

    function obterGrafico(id) {
      var dom = document.getElementById(id);
      if (!dom || typeof echarts === 'undefined') return null;
      var existente = echarts.getInstanceByDom(dom);
      if (existente) return existente;
      var grafico = echarts.init(dom, null, { renderer: 'canvas' });
      CHARTS[id] = grafico;
      return grafico;
    }

    function destruirGraficos() {
      Object.keys(CHARTS).forEach(function(id) {
        try { if (CHARTS[id] && !CHARTS[id].isDisposed()) CHARTS[id].dispose(); } catch (e) {}
      });
      CHARTS = {};
    }

    function serieGrafico(dados) {
      var serie = (dados && dados.serieDiaria ? dados.serieDiaria : []).slice();
      if (PERIODO_GRAFICO !== 'tudo') serie = serie.slice(-Number(PERIODO_GRAFICO || 12));
      return serie;
    }

    function mediaMovel(valores, janela) {
      return valores.map(function(valor, indice) {
        if (indice < janela - 1) return null;
        var recorte = valores.slice(indice - janela + 1, indice + 1);
        var soma = recorte.reduce(function(total, item) { return total + Number(item || 0); }, 0);
        return Number((soma / janela).toFixed(2));
      });
    }

    function renderizarEvolucao(dados, reiniciarAnimacao) {
      var grafico = obterGrafico('chart-evolution');
      if (!grafico) return;
      if (reiniciarAnimacao) { try { grafico.clear(); } catch (e) {} }
      var serie = serieGrafico(dados);
      var valores = serie.map(function(item) { return Number(item.total || 0); });
      var rotulos = serie.map(function(item) { return item.data; });
      var janela = Math.min(5, Math.max(3, valores.length));
      var media = mediaMovel(valores, janela);
      var totalAcumulado = dados && dados.metricas ? Number(dados.metricas.total || 0) : 0;
      var meta = dados && dados.metricas ? Number(dados.metricas.metaGeral || 0) : 0;
      var numero = document.getElementById('dv-chart-total');
      var legenda = document.getElementById('dv-chart-caption');
      if (numero) numero.textContent = inteiro(totalAcumulado);
      if (legenda) legenda.textContent = 'inscrições acumuladas' + (meta > 0 ? ' · ' + percentual(totalAcumulado / meta) + ' da meta' : '');

      var tokenLinha = ++TOKEN_ANIMACAO_LINHA;
      var mediaInicial = reiniciarAnimacao ? media.map(function(valor) { return valor === null ? null : null; }) : media;
      grafico.setOption({
        animation: true,
        animationDuration: 300,
        animationDurationUpdate: 180,
        animationEasing: 'cubicOut',
        animationEasingUpdate: 'cubicOut',
        grid: { left: 4, right: 4, top: 12, bottom: 18, containLabel: false },
        tooltip: {
          trigger: 'axis',
          backgroundColor: '#11120f',
          borderColor: '#3f3f46',
          textStyle: { color: '#fff', fontSize: 10 },
          formatter: function(params) {
            var indice = params && params.length ? params[0].dataIndex : 0;
            var item = serie[indice] || {};
            var html = '<b>' + esc(item.dataCompleta || item.data || '') + '</b>';
            params.forEach(function(p) {
              if (p.value === null || p.value === undefined) return;
              html += '<br>' + esc(p.seriesName) + ': <b>' + decimal(p.value, p.seriesName === 'Média móvel' ? 1 : 0) + '</b>';
            });
            return html;
          }
        },
        xAxis: {
          type: 'category', data: rotulos, boundaryGap: true,
          axisTick: { show: false }, axisLine: { lineStyle: { color: '#e4e4e7' } },
          axisLabel: { color: '#71717a', fontSize: 8, interval: 0, hideOverlap: true, margin: 7 }
        },
        yAxis: { type: 'value', min: 0, show: false },
        series: [
          {
            name: 'Inscrições', type: 'bar', data: valores, barMaxWidth: 52, barCategoryGap: '10%', showBackground: true,
            backgroundStyle: { color: '#f4f4f5', borderRadius: [11,11,3,3] },
            itemStyle: {
              color: '#09090B',
              borderRadius: [11,11,3,3]
            },
            label: { show: true, position: 'top', distance: 5, color: '#71717a', fontSize: 8.5, formatter: function(p) { return inteiro(p.value); } },
            animationDuration: 280,
            emphasis: { itemStyle: { color: '#25252a' } }, z: 2
          },
          {
            name: 'Média móvel', type: 'line', data: reiniciarAnimacao ? mediaInicial : media, connectNulls: false, smooth: .22, showSymbol: false,
            lineStyle: { color: '#c9ef72', width: 2.35 }, itemStyle: { color: '#c9ef72' },
            animationDuration: 360, animationDelay: 0, animationEasing: 'cubicOut', z: 5
          }
        ]
      }, true);

      if (reiniciarAnimacao && media.length) {
        var primeiroValido = media.findIndex(function(valor) { return valor !== null; });
        if (primeiroValido >= 0) {
          var revelados = media.map(function() { return null; });
          revelados[primeiroValido] = media[primeiroValido];
          var indiceRevelado = primeiroValido + 1;
          (function revelarProximo() {
            if (tokenLinha !== TOKEN_ANIMACAO_LINHA || indiceRevelado >= media.length) return;
            if (media[indiceRevelado] !== null) revelados[indiceRevelado] = media[indiceRevelado];
            grafico.setOption({
              animationDurationUpdate: 55,
              animationEasingUpdate: 'linear',
              series: [{}, { data: revelados.slice() }]
            }, false);
            indiceRevelado += 1;
            window.setTimeout(revelarProximo, 22);
          })();
        }
      }
    }

    function mudarPeriodoGrafico(periodo) {
      PERIODO_GRAFICO = periodo === 'tudo' ? 'tudo' : Number(periodo);
      document.querySelectorAll('[data-chart-period]').forEach(function(botao) {
        botao.className = 'dv-time-btn' + (String(botao.getAttribute('data-chart-period')) === String(PERIODO_GRAFICO) ? ' active' : '');
      });
      renderizarEvolucao(CURRENT_DATA, true);
    }

    function carregarGeoJsonBrasil() {
      if (GEO_JSON) return Promise.resolve(GEO_JSON);
      if (GEO_PROMISE) return GEO_PROMISE;
      GEO_PROMISE = (function tentar(indice) {
        if (indice >= BRAZIL_GEOJSON_URLS.length) return Promise.reject(new Error('GeoJSON do Brasil indisponível'));
        return fetch(BRAZIL_GEOJSON_URLS[indice], { cache: 'force-cache' })
          .then(function(resposta) { if (!resposta.ok) throw new Error('HTTP ' + resposta.status); return resposta.json(); })
          .catch(function() { return tentar(indice + 1); });
      })(0).then(function(geo) {
        (geo.features || []).forEach(function(feature) {
          feature.properties = feature.properties || {};
          if (!feature.properties.name) feature.properties.name = feature.properties.nome || feature.properties.sigla || 'Estado';
        });
        GEO_JSON = geo;
        echarts.registerMap('COOPSPORTES_BRASIL', geo);
        return geo;
      }).catch(function(erro) {
        GEO_PROMISE = null;
        throw erro;
      });
      return GEO_PROMISE;
    }

    function dadosCidadesMapa() {
      var cidades = (CURRENT_DATA && CURRENT_DATA.cidadesMapa ? CURRENT_DATA.cidadesMapa : [])
        .filter(function(cidade) { return cidade.resolvida && isFinite(cidade.latitude) && isFinite(cidade.longitude) && Number(cidade.total || 0) > 0; })
        .slice();
      cidades.sort(function(a,b) { return b.total - a.total || a.nome.localeCompare(b.nome); });
      cidades.forEach(function(cidade, indice) { cidade.rankingVisual = indice + 1; });
      var maximo = Math.max.apply(null, cidades.map(function(cidade) { return Number(cidade.total || 0); }).concat([1]));
      return cidades.map(function(cidade) {
        var altura = 1.2 + Math.sqrt(Number(cidade.total || 0) / maximo) * 8.5;
        return { name: cidade.nome, value: [Number(cidade.longitude), Number(cidade.latitude), altura, Number(cidade.total || 0), Number(cidade.participacao || 0), cidade.rankingVisual] };
      });
    }

    function renderizarMapa3D(id, compacto) {
      var dom = document.getElementById(id);
      if (!dom || typeof echarts === 'undefined') return;
      function desenhar() {
        var grafico = obterGrafico(id);
        if (!grafico) return;
        var cidades = dadosCidadesMapa();
        grafico.setOption({
          backgroundColor: 'transparent',
          tooltip: {
            show: !compacto,
            backgroundColor: '#18181b', borderColor: '#343439', borderWidth: 1,
            textStyle: { color: '#f4f4f5', fontSize: 10 },
            formatter: function(p) {
              var v = p.value || [];
              if (p.seriesType === 'bar3D' || p.seriesType === 'scatter3D') {
                return '<b>' + esc(p.name) + '</b><br>Inscritos: ' + inteiro(v[3]) + '<br>Participação: ' + percentual(v[4] || 0) + '<br>Ranking: #' + inteiro(v[5]);
              }
              return esc(p.name || 'Brasil');
            }
          },
          geo3D: {
            map: 'COOPSPORTES_BRASIL', shading: 'lambert', boxWidth: 92, boxHeight: compacto ? 10 : 14,
            regionHeight: compacto ? 1.7 : 2.2, environment: '#09090B',
            itemStyle: { color: '#1d1d20', opacity: 1, borderWidth: .55, borderColor: '#505056' },
            emphasis: { itemStyle: { color: '#29292d' } }, label: { show: false }, groundPlane: { show: false },
            light: { main: { intensity: 1.45, shadow: !compacto, shadowQuality: 'medium', alpha: 42, beta: 18 }, ambient: { intensity: .28 } },
            postEffect: { enable: !compacto, SSAO: { enable: !compacto, radius: 2, intensity: 1.3 } },
            temporalSuperSampling: { enable: !compacto },
            viewControl: { projection: 'perspective', alpha: compacto ? 50 : 47, beta: -8, distance: compacto ? 118 : 97, minDistance: 65, maxDistance: 150, rotateSensitivity: compacto ? 0 : 1, zoomSensitivity: compacto ? 0 : 1, panSensitivity: compacto ? 0 : .5, autoRotate: false }
          },
          series: [
            { name: 'Cidades', type: 'bar3D', coordinateSystem: 'geo3D', shading: 'lambert', barSize: .38, minHeight: .6, bevelSize: .08, data: cidades, itemStyle: { color: '#c9ef72', opacity: .94 }, emphasis: { itemStyle: { color: '#f4f4f5' } }, label: { show: false } },
            { name: 'Pontos', type: 'scatter3D', coordinateSystem: 'geo3D', data: cidades.map(function(x) { return { name: x.name, value: [x.value[0],x.value[1],x.value[2]+.35,x.value[3],x.value[4],x.value[5]] }; }), symbolSize: function(v) { return clamp(4 + Math.sqrt(Number(v[3] || 0)) * 1.2, 5, 14); }, itemStyle: { color: '#f4f4f5', opacity: .92 }, emphasis: { itemStyle: { color: '#e5484d' } } }
          ]
        }, true);
      }
      if (GEO_JSON) { desenhar(); return; }
      dom.innerHTML = '<div class="dv-map-empty">Carregando mapa 3D…</div>';
      carregarGeoJsonBrasil().then(desenhar).catch(function(erro) {
        dom.innerHTML = '<div class="dv-map-empty">Não foi possível carregar o mapa.<br>' + esc(erro && erro.message ? erro.message : erro) + '</div>';
      });
    }

    function resetarMapa() {
      var grafico = CHARTS['chart-map-full'];
      if (grafico) { try { grafico.dispose(); } catch(e) {} delete CHARTS['chart-map-full']; }
      renderizarMapa3D('chart-map-full', false);
    }

    function painelMapa(dados) {
      var cidades = (dados.cidadesMapa || []).filter(function(item) { return item.resolvida; });
      var linhas = cidades.slice().sort(function(a,b){ return b.total-a.total || a.nome.localeCompare(b.nome); }).map(function(item, indice) {
        return '<div class="dv-city-row"><span class="rank">#' + (indice + 1) + '</span><strong>' + esc(item.nome) + '</strong><span class="num">' + inteiro(item.total) + '</span><span class="share">' + percentual(item.participacao) + '</span></div>';
      }).join('');
      return '<div class="dv-map-layout"><div class="card dv-map-card"><div class="dv-map-card-head"><div><h2>Brasil</h2><p>Arraste para rotacionar e use a roda do mouse para zoom.</p></div><button class="dv-map-reset" onclick="resetarMapa()">Resetar visualização</button></div><div id="chart-map-full" class="dv-map-full"></div></div>' +
        '<div class="card"><div class="section-title"><h2>Cidades</h2><span class="muted">ranking</span></div><div class="dv-city-rank">' + (linhas || '<div class="dv-map-empty">Nenhuma cidade localizada.</div>') + '</div></div></div>';
    }

    function setaKpi(tom) {
      if (tom === 'green') return '<span class="kpi-arrow up" aria-label="positivo">▲</span>';
      if (tom === 'red') return '<span class="kpi-arrow down" aria-label="negativo">▼</span>';
      return '';
    }

    function kpi(rotulo, valor, detalhe, tom, mostrarSeta) {
      var indicador = mostrarSeta ? setaKpi(tom) : '';
      return '<div class="card tone-' + tom + '">' +
        '<div class="kpi-label">' + esc(rotulo) + '</div>' +
        '<div class="kpi-value-row"><div class="kpi-value">' + esc(valor) + '</div>' + indicador + '</div>' +
        '<div class="kpi-detail">' + esc(detalhe) + '</div>' +
      '</div>';
    }

    function graficoLinha(serie) {
      if (!serie || !serie.length) return '<div class="muted">Sem datas válidas.</div>';

      var largura = 760;
      var altura = 96;
      var margemX = 34;
      var margemY = 14;
      var maximo = Math.max.apply(null, serie.map(function(item) { return item.total; }).concat([1]));
      var larguraUtil = largura - margemX * 2;
      var alturaUtil = altura - margemY * 2;
      var pontos = serie.map(function(item, indice) {
        var x = margemX + (serie.length === 1 ? larguraUtil / 2 : indice * larguraUtil / (serie.length - 1));
        var y = margemY + alturaUtil - (item.total / maximo) * alturaUtil;
        return { x: x, y: y, item: item };
      });
      var linha = pontos.map(function(ponto) { return ponto.x + ',' + ponto.y; }).join(' ');
      var area = margemX + ',' + (margemY + alturaUtil) + ' ' + linha + ' ' +
        (margemX + larguraUtil) + ',' + (margemY + alturaUtil);
      var passoRotulo = Math.max(1, Math.ceil(serie.length / 7));
      var rotulos = pontos.map(function(ponto, indice) {
        if (indice % passoRotulo !== 0 && indice !== pontos.length - 1) return '';
        return '<text x="' + ponto.x + '" y="' + (altura - 4) + '" fill="#8e98aa" font-size="9" text-anchor="middle">' +
          esc(ponto.item.data) + '</text>';
      }).join('');
      var circulos = pontos.map(function(ponto) {
        return '<circle cx="' + ponto.x + '" cy="' + ponto.y + '" r="2.3" fill="#c9ef72">' +
          '<title>' + esc(ponto.item.dataCompleta) + ': ' + inteiro(ponto.item.total) + '</title></circle>';
      }).join('');
      var grades = [0, .5, 1].map(function(fracao) {
        var y = margemY + alturaUtil - fracao * alturaUtil;
        var valor = Math.round(fracao * maximo);
        return '<line x1="' + margemX + '" y1="' + y + '" x2="' + (largura - margemX) + '" y2="' + y +
          '" stroke="#29292e" stroke-width="1" />' +
          '<text x="5" y="' + (y + 3) + '" fill="#8e98aa" font-size="9">' + valor + '</text>';
      }).join('');

      return '<svg viewBox="0 0 ' + largura + ' ' + altura + '" width="100%" height="96" role="img" aria-label="Inscrições por dia">' +
        '<defs><linearGradient id="areaGrad" x1="0" y1="0" x2="0" y2="1">' +
          '<stop offset="0%" stop-color="#c9ef72" stop-opacity=".30" />' +
          '<stop offset="100%" stop-color="#c9ef72" stop-opacity="0" />' +
        '</linearGradient></defs>' + grades +
        '<polygon points="' + area + '" fill="url(#areaGrad)" />' +
        '<polyline points="' + linha + '" fill="none" stroke="#c9ef72" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" />' +
        circulos + rotulos + '</svg>';
    }

    function graficoBarras(serie) {
      if (!serie || !serie.length) return '<div class="muted">Sem datas válidas.</div>';
      var pontos = serie.slice(-12);
      var maximo = Math.max.apply(null, pontos.map(function(item) {
        return item.total;
      }).concat([1]));

      return '<div class="dv-bars" role="img" aria-label="Inscrições diárias">' +
        pontos.map(function(item) {
          var altura = Math.max(3, item.total / maximo * 100);
          return '<div class="dv-bar-col" title="' + esc(item.dataCompleta) + ': ' + inteiro(item.total) + '">' +
            '<div class="dv-bar-value">' + inteiro(item.total) + '</div>' +
            '<div class="dv-bar-track"><div class="dv-bar-fill" style="height:' + altura + '%"></div></div>' +
            '<div class="dv-bar-label">' + esc(item.data) + '</div>' +
          '</div>';
        }).join('') +
      '</div>';
    }

    function rotuloTendencia(item) {
      if (!item.comparacaoValida) return { texto: 'sem base', classe: 'flat' };
      if (item.anteriores7 === 0 && item.ultimos7 > 0) return { texto: 'novo', classe: 'up' };
      if (item.anteriores7 === 0) return { texto: '-', classe: 'flat' };
      if (item.crescimento > .05) return { texto: '▲ ' + percentual(item.crescimento), classe: 'up' };
      if (item.crescimento < -.05) return { texto: '▼ ' + percentual(Math.abs(item.crescimento)), classe: 'down' };
      return { texto: 'estável', classe: 'flat' };
    }

    function linhasModalidades(dados) {
      var maximo = Math.max.apply(null, dados.modalidades.map(function(item) { return item.total; }).concat([1]));

      return dados.modalidades.map(function(item) {
        var tendencia = rotuloTendencia(item);
        var largura = item.total / maximo * 100;
        return '<div class="metric-row" title="Ocupação: ' + esc(percentual(item.ocupacao)) + '">' +
          '<div class="metric-name">' + esc(item.nome) + '</div>' +
          '<div class="bar"><span style="width:' + largura + '%"></span></div>' +
          '<div class="metric-number">' + inteiro(item.total) + ' / ' + inteiro(dados.metricas.maximoModalidade) + '</div>' +
          '<div class="trend ' + tendencia.classe + '">' + esc(tendencia.texto) + '</div>' +
        '</div>';
      }).join('');
    }

    function listaGrupos(itens, cor) {
      if (!itens || !itens.length) return '<div class="muted">Sem dados.</div>';
      var maximo = Math.max.apply(null, itens.map(function(item) { return item.total; }).concat([1]));

      return itens.map(function(item) {
        return '<div class="simple-row">' +
          '<div><div>' + esc(item.nome) + '</div>' +
          '<div class="mini-bar"><span style="width:' + (item.total / maximo * 100) + '%;background:' + cor + '"></span></div></div>' +
          '<strong>' + inteiro(item.total) + '</strong>' +
        '</div>';
      }).join('');
    }

    function cardsInsights(insights) {
      var prioridade = { critical: 0, warning: 1, positive: 2, info: 3 };
      var ordenados = insights.slice().sort(function(a, b) {
        var prioridadeA = Object.prototype.hasOwnProperty.call(prioridade, a.nivel) ? prioridade[a.nivel] : 9;
        var prioridadeB = Object.prototype.hasOwnProperty.call(prioridade, b.nivel) ? prioridade[b.nivel] : 9;
        return prioridadeA - prioridadeB;
      });

      return ordenados.slice(0, 4).map(function(item) {
        return '<div class="insight ' + esc(item.nivel) + '">' +
          '<div class="insight-title">' + esc(item.titulo) + '</div>' +
          '<div class="insight-text">' + esc(item.texto) + '</div>' +
          '<div class="insight-action"><strong>Próxima ação:</strong> ' + esc(item.acao) + '</div>' +
        '</div>';
      }).join('');
    }

    function statusModalidade(item) {
      if (item.ocupacao >= .8) return { classe: 'capacity', texto: 'capacidade' };
      if (item.projecao < item.referenciaMedia * .75) return { classe: 'watch', texto: 'priorizar' };
      return { classe: 'ok', texto: 'em curso' };
    }

    function linhasModalidadesDetalhadas(dados) {
      return dados.modalidades.map(function(item) {
        var status = statusModalidade(item);
        var larguraAtual = Math.min(100, item.total / dados.metricas.maximoModalidade * 100);
        var larguraProjetada = Math.min(100, item.projecao / dados.metricas.maximoModalidade * 100);
        return '<div class="dv-mod-row ' + status.classe + '">' +
          '<div class="dv-mod-name">' + esc(item.nome) + '</div>' +
          '<div class="dv-proj-track" title="Atual ' + inteiro(item.total) + ' · projeção ' + inteiro(item.projecao) + '">' +
            '<span class="dv-proj-future" style="width:' + larguraProjetada + '%"></span>' +
            '<span class="dv-proj-current" style="width:' + larguraAtual + '%"></span>' +
          '</div>' +
          '<div>' + inteiro(item.total) + '</div>' +
          '<div>' + inteiro(item.projecao) + '</div>' +
          '<div class="dv-status ' + status.classe + '">' + status.texto + '</div>' +
        '</div>';
      }).join('');
    }

    function acoesModalidades(dados) {
      var prioridades = dados.modalidades.slice().sort(function(a, b) {
        return b.gapReferencia - a.gapReferencia || a.projecao - b.projecao;
      }).slice(0, 4);

      return prioridades.map(function(item, indice) {
        return '<div class="dv-action-row"><strong>' + (indice + 1) + '. ' + esc(item.nome) + '</strong>' +
          '<span>Faltam <b>' + inteiro(item.gapReferencia) + '</b> para a referência média de ' +
          inteiro(item.referenciaMedia) + '. Ritmo sugerido: ' + decimal(item.ritmoParaReferencia, 1) + '/dia.</span></div>';
      }).join('');
    }

    function painelModalidades(dados) {
      var m = dados.metricas;
      var lider = dados.modalidades[0] || { nome: 'Sem dados', total: 0 };
      var menor = dados.modalidades[dados.modalidades.length - 1] || { nome: 'Sem dados', total: 0 };
      var mediaAtual = dados.modalidades.length ? m.total / dados.modalidades.length : 0;      return '<div class="dv-page-kpis">' +
          kpi('Líder atual', lider.nome, inteiro(lider.total) + ' inscrições', 'green') +
          kpi('Menor volume', menor.nome, inteiro(menor.total) + ' inscrições', 'red') +
          kpi('Média por modalidade', decimal(mediaAtual, 1), 'referência analítica: ' + inteiro(m.vagasMediasNecessarias), 'neutral') +
          kpi('Ocupação da capacidade', percentual(m.utilizacaoCapacidade), inteiro(m.total) + ' de ' + inteiro(m.capacidade), 'neutral') +
        '</div>' +
        '<div class="dv-analysis-grid">' +
          '<div class="card"><div class="section-title"><h2>Ranking e projeção até o encerramento</h2>' +
          '<span class="muted">limite operacional: ' + inteiro(m.maximoModalidade) + '</span></div>' +
          '<div class="dv-table-head"><span>Modalidade</span><span>Atual x projeção</span><span>Atual</span><span>Projeção</span><span>Status</span></div>' +
          '<div>' + linhasModalidadesDetalhadas(dados) + '</div></div>' +
          '<div class="card dv-dark-feature dv-action-plan-card"><div class="section-title"><div><div class="eyebrow">PLANO DE AÇÃO</div><h2>Onde mobilizar primeiro</h2></div></div>' +
          '<div class="dv-action-stack">' + acoesModalidades(dados) + '</div>' +
          '<div class="footer">Prioridade combina distância da referência média e projeção de encerramento.</div></div>' +
        '</div>';
    }

    function cardsInsightsTemporais(itens) {
      if (!itens || !itens.length) return '<div class="card muted">Sem evidências suficientes para este período.</div>';
      return '<div class="dv-timeline">' + itens.map(function(item) {
        return '<article class="dv-temporal-card ' + esc(item.nivel) + '">' +
          '<span class="dv-temporal-dot" aria-hidden="true"></span>' +
          '<div class="dv-temporal-copy"><strong>' + esc(item.titulo) + '</strong><span>' + esc(item.texto) + '</span></div>' +
          '<div class="dv-next-step"><b>Decisão recomendada</b>' + esc(item.acao) + '</div>' +
        '</article>';
      }).join('') + '</div>';
    }

    function painelInsightsTemporais(dados) {
      var m = dados.metricas;
      var gap = Math.max(0, Math.ceil(m.metaGeral - m.projecao));
      return '<div class="dv-period-tabs dv-period-tabs-top" role="tablist" aria-label="Período dos insights">' +
          '<button role="tab" id="dv-period-passado" class="dv-period-btn" onclick="mudarPeriodoInsight(\\'passado\\')">Passado</button>' +
          '<button role="tab" id="dv-period-atual" class="dv-period-btn" onclick="mudarPeriodoInsight(\\'atual\\')">Presente</button>' +
          '<button role="tab" id="dv-period-futuro" class="dv-period-btn" onclick="mudarPeriodoInsight(\\'futuro\\')">Futuro</button>' +
        '</div>' +
        '<div class="dv-insight-hero">' +
          '<div class="card"><div class="section-title"><h2>Leitura executiva</h2><span class="muted">referência ' + esc(dados.contexto.referencia) + '</span></div>' +
          '<div class="dv-page-kpis" style="margin:0">' +
            kpi('Acumulado', inteiro(m.total), percentual(m.atingimentoMeta) + ' da meta', 'neutral') +
            kpi('Projeção final', inteiro(m.projecao), gap ? inteiro(gap) + ' abaixo da meta' : 'meta alcançada', gap ? 'red' : 'green') +
            kpi('Ritmo recente', decimal(m.mediaRecente, 1) + '/dia', 'necessário ' + decimal(m.ritmoNecessario, 1) + '/dia', m.mediaRecente >= m.ritmoNecessario ? 'green' : 'red') +
            kpi('Prazo restante', inteiro(m.diasRestantes) + ' dias', esc(dados.contexto.fim), 'neutral') +
          '</div></div>' +
          '<div class="card dv-time-summary"><div class="eyebrow">FOCO RECOMENDADO</div>' +
          '<strong>' + (gap ? inteiro(gap) : '0') + '</strong><p>' +
          (gap ? 'inscrições adicionais precisam ser geradas além da projeção atual.' : 'déficit projetado para a meta geral.') +
          '</p></div></div>' +
        '<div id="dv-temporal-passado" class="dv-temporal-panel">' + cardsInsightsTemporais(dados.insightsTemporais.passado) + '</div>' +
        '<div id="dv-temporal-atual" class="dv-temporal-panel">' + cardsInsightsTemporais(dados.insightsTemporais.atual) + '</div>' +
        '<div id="dv-temporal-futuro" class="dv-temporal-panel">' + cardsInsightsTemporais(dados.insightsTemporais.futuro) + '</div>';
    }

    function linhaQualidade(nivel, total, titulo, descricao, acao) {
      return '<div class="dv-issue-row ' + nivel + '"><div class="dv-issue-count">' + inteiro(total) + '</div>' +
        '<div class="dv-issue-copy"><strong>' + esc(titulo) + '</strong><span>' + esc(descricao) + '</span>' +
        '<span class="dv-issue-action"><b>Ação:</b> ' + esc(acao) + '</span></div></div>';
    }

    function rankingMenores(qualidade) {
      var itens = qualidade.menores.porModalidade.slice(0, 6);
      if (!itens.length) return '<div class="muted">Nenhum menor identificado.</div>';
      var maximo = Math.max.apply(null, itens.map(function(item) { return item.total; }).concat([1]));
      return '<div class="dv-minor-bars">' + itens.map(function(item) {
        return '<div class="dv-minor-row"><span>' + esc(item.nome) + '</span><div class="bar"><span style="width:' +
          (item.total / maximo * 100) + '%"></span></div><strong>' + inteiro(item.total) + '</strong></div>';
      }).join('') + '</div>';
    }

    function mapaLinhasQualidade(qualidade) {
      var itens = qualidade.problemasPorLinha || [];
      if (!itens.length) return '<div class="dv-line-empty">Nenhuma linha com problema ou sinal de revisão.</div>';
      var rotulos = { critical: 'Crítico', warning: 'Atenção', info: 'Validar' };
      var acoes = {
        critical: 'Regularizar antes de confirmar a participação.',
        warning: 'Conferir e corrigir a linha indicada.',
        info: 'Validar se a repetição é permitida; não excluir automaticamente.'
      };

      return '<div class="dv-line-scroll"><div class="dv-line-head">' +
        '<span>Linha</span><span>Nível</span><span>Problemas encontrados</span><span>Ação</span></div>' +
        itens.map(function(item) {
          return '<div class="dv-line-row ' + esc(item.nivel) + '">' +
            '<div class="dv-line-number">' + inteiro(item.linha) + '</div>' +
            '<div><span class="dv-line-level ' + esc(item.nivel) + '">' + esc(rotulos[item.nivel] || 'Validar') + '</span></div>' +
            '<div class="dv-problem-list">' + item.problemas.map(function(problema) {
              return '<div class="dv-problem-item">' + esc(problema.descricao) + '</div>';
            }).join('') + '</div>' +
            '<div class="dv-line-action">' + esc(acoes[item.nivel] || acoes.info) + '</div>' +
          '</div>';
        }).join('') + '</div>';
    }

    function painelQualidade(dados) {
      var q = dados.qualidade;
      var v = q.validacoes;
      var d = q.duplicidades;
      var faltantePrincipal = q.faltantes.filter(function(item) { return item.total > 0; })[0];      return '<div class="dv-page-kpis">' +
          kpi('Linhas com pendência', inteiro(q.totalLinhasComPendencia), inteiro(q.linhasPorNivel.critical) + ' críticas · ' + inteiro(q.linhasPorNivel.warning) + ' atenção', q.totalLinhasComPendencia ? 'red' : 'green') +
          kpi('Inscrições diretas regularizadas', percentual(q.menores.conformidade), inteiro(q.menores.comAutorizacao) + ' de ' + inteiro(q.menores.sujeitosAoTermo) + ' que exigem termo', q.menores.semAutorizacao ? 'red' : 'green') +
          kpi('Duplicidades fortes', inteiro(q.excedentesDuplicados), inteiro(q.gruposDuplicados) + ' grupos para revisar', q.excedentesDuplicados ? 'red' : 'green') +
          kpi('Completude crítica', percentual(q.completude), inteiro(q.linhasIncompletas) + ' linha(s) incompleta(s)', q.linhasIncompletas ? 'red' : 'green') +
        '</div>' +
        '<div class="dv-quality-grid">' +
          '<div class="card"><div class="section-title"><h2>Fila de tratamento recomendada</h2><span class="muted">ordem de risco</span></div>' +
          '<div class="dv-issue-list">' +
            linhaQualidade(q.menores.semAutorizacao ? 'critical' : 'ok', q.menores.semAutorizacao, 'Inscrição direta de menor sem termo', 'Apenas menores fora do fluxo comprovado do responsável entram nesta fila.', 'Regularizar antes de confirmar participação.') +
            linhaQualidade(v.cpfInvalido ? 'warning' : 'ok', v.cpfInvalido, 'CPF com dígito verificador inválido', 'Falha algorítmica; pode representar erro de digitação.', 'Conferir com o documento original.') +
            linhaQualidade(v.nascimentoInvalido + q.datasInvalidas ? 'warning' : 'ok', v.nascimentoInvalido + q.datasInvalidas, 'Datas fora da faixa válida', 'Inclui nascimento inválido e carimbo de inscrição não reconhecido.', 'Confirmar as datas e recalcular os campos derivados.') +
            linhaQualidade(q.excedentesDuplicados ? 'warning' : 'ok', q.excedentesDuplicados, 'Possíveis duplicidades cadastrais', 'Nome, telefone e e-mail coincidem no mesmo grupo.', 'Validar antes de qualquer exclusão.') +
            linhaQualidade(v.emailInvalido + v.telefoneInvalido ? 'warning' : 'ok', v.emailInvalido + v.telefoneInvalido, 'Contato com formato inválido', 'E-mail sem estrutura válida ou telefone fora de 10–11 dígitos.', 'Corrigir para preservar a capacidade de contato.') +
            linhaQualidade(v.idadeInconsistente ? 'warning' : 'ok', v.idadeInconsistente, 'Idade divergente do nascimento', 'A idade calculada não coincide com o campo informado.', 'Recalcular a fórmula ou corrigir a data de nascimento.') +
            linhaQualidade(q.linhasIncompletas ? 'warning' : 'ok', q.linhasIncompletas, 'Campos críticos incompletos', faltantePrincipal ? 'Maior ausência: ' + faltantePrincipal.nome + '.' : 'Nenhuma ausência crítica encontrada.', 'Corrigir o campo de maior incidência primeiro.') +
          '</div></div>' +
          '<div class="dv-action-stack">' +
            '<div class="card dv-compliance"><div class="eyebrow">TERMO NO FLUXO DIRETO</div>' +
              '<div class="dv-compliance-number">' + percentual(q.menores.conformidade) + '</div><div class="muted">conformidade documental</div>' +
              '<div class="dv-progress"><span style="width:' + Math.min(100, q.menores.conformidade * 100) + '%"></span></div>' +
              '<div class="dv-compliance-split"><div><strong>' + inteiro(q.menores.sujeitosAoTermo) + '</strong><span>exigem termo</span></div>' +
              '<div><strong>' + inteiro(q.menores.comAutorizacao) + '</strong><span>com termo</span></div>' +
              '<div><strong>' + inteiro(q.menores.semAutorizacao) + '</strong><span>sem termo</span></div></div>' +
              '<div class="dv-mini-label">Menores no fluxo direto por modalidade</div>' + rankingMenores(q) +
              '<div class="footer">' + inteiro(q.menores.viaResponsavel) +
              ' menor(es) no fluxo do responsável não entram como pendência. ' +
              inteiro(q.menores.termosNoFluxoResponsavel) + ' termo(s) foram anexados opcionalmente nesse fluxo e ' +
              inteiro(q.menores.adultosComAutorizacao) + ' por participantes adultos.</div></div>' +
            '<div class="card"><div class="section-title"><h2>Conflitos de identidade</h2><span class="muted">grupos repetidos</span></div>' +
              '<div class="dv-dup-grid"><div class="dv-dup-item"><strong>' + inteiro(d.cadastro.grupos) + '</strong><span>cadastro completo</span></div>' +
              '<div class="dv-dup-item"><strong>' + inteiro(d.cpf.grupos) + '</strong><span>CPF repetido</span></div>' +
              '<div class="dv-dup-item"><strong>' + inteiro(d.email.grupos) + '</strong><span>e-mail repetido</span></div>' +
              '<div class="dv-dup-item"><strong>' + inteiro(d.telefone.grupos) + '</strong><span>telefone repetido</span></div></div>' +
              '<div class="footer">Repetição de identificador é indício de revisão, não confirmação de duplicidade.</div></div>' +
          '</div>' +
        '</div>' +
        '<div class="card dv-line-audit"><div class="section-title"><div><div class="eyebrow">RASTREABILIDADE</div>' +
          '<h2>Mapa de linhas com problema ou sinal de revisão</h2></div>' +
          '<div class="dv-line-summary"><span class="critical">' + inteiro(q.linhasPorNivel.critical) + ' críticas</span>' +
          '<span class="warning">' + inteiro(q.linhasPorNivel.warning) + ' atenção</span>' +
          '<span class="info">' + inteiro(q.linhasPorNivel.info) + ' para validar</span></div></div>' +
          '<div class="muted dv-line-note">Use o número exibido para localizar a ocorrência na guia Respostas ao formulário 1. Linhas informativas representam repetição, não erro confirmado.</div>' +
          mapaLinhasQualidade(q) + '</div>';
    }


    function executarQuandoOcioso(fn) {
      if (typeof window.requestIdleCallback === 'function') {
        window.requestIdleCallback(function() { fn(); }, { timeout: 450 });
      } else {
        window.setTimeout(fn, 120);
      }
    }

    function mudarPainel(nome) {
      var configuracoes = {
        visao: { titulo: 'VISÃO GERAL DAS INSCRIÇÕES' },
        modalidades: { titulo: 'ANÁLISE DE MODALIDADES' },
        insights: { titulo: 'INSIGHTS E DECISÕES' },
        qualidade: { titulo: 'QUALIDADE DOS DADOS' },
        mapa: { titulo: 'COBERTURA GEOGRÁFICA' }
      };
      if (!configuracoes[nome]) nome = 'visao';
      PAINEL_ATIVO = nome;
      Object.keys(configuracoes).forEach(function(chave) {
        var painel = document.getElementById('dv-panel-' + chave);
        var botao = document.getElementById('dv-nav-' + chave);
        if (painel) painel.hidden = chave !== nome;
        if (botao) {
          botao.className = 'dv-nav-btn' + (chave === nome ? ' active' : '');
          botao.setAttribute('aria-selected', chave === nome ? 'true' : 'false');
        }
      });
      var titulo = document.getElementById('dv-page-title');
      var principal = document.getElementById('dv-main');
      if (titulo) titulo.textContent = configuracoes[nome].titulo;
      if (principal) principal.scrollTop = 0;
      window.setTimeout(function() {
        if (nome === 'visao') {
          renderizarEvolucao(CURRENT_DATA, false);
          executarQuandoOcioso(function() { if (PAINEL_ATIVO === 'visao') renderizarMapa3D('chart-map-preview', true); });
          document.querySelectorAll('[data-chart-period]').forEach(function(botao) {
            botao.className = 'dv-time-btn' + (String(botao.getAttribute('data-chart-period')) === String(PERIODO_GRAFICO) ? ' active' : '');
          });
        } else if (nome === 'mapa') {
          renderizarMapa3D('chart-map-full', false);
        }
      }, 0);
    }

    function mudarPeriodoInsight(periodo) {
      var periodos = ['passado', 'atual', 'futuro'];
      if (periodos.indexOf(periodo) < 0) periodo = 'futuro';
      PERIODO_INSIGHT_ATIVO = periodo;
      periodos.forEach(function(chave) {
        var painel = document.getElementById('dv-temporal-' + chave);
        var botao = document.getElementById('dv-period-' + chave);
        if (painel) painel.hidden = chave !== periodo;
        if (botao) {
          botao.className = 'dv-period-btn' + (chave === periodo ? ' active' : '');
          botao.setAttribute('aria-selected', chave === periodo ? 'true' : 'false');
        }
      });
    }

    function renderizar(dados) {
      destruirGraficos();
      CURRENT_DATA = dados;
      var m = dados.metricas;
      var crescimento = m.crescimentoSemanal === null ? 'sem base' : percentual(m.crescimentoSemanal);
      var detalheCrescimento = m.crescimentoSemanal === null
        ? 'aguardando duas semanas completas'
        : inteiro(m.ultimos7) + ' vs. ' + inteiro(m.seteAnteriores);
      var tomInscricoes = m.total >= m.esperadoAteAgora ? 'green' : 'red';
      var tomCrescimento = m.crescimentoSemanal === null ? 'neutral' : (m.crescimentoSemanal >= 0 ? 'green' : 'red');
      var tomRitmo = m.mediaRecente >= m.ritmoNecessario ? 'green' : 'red';
      var totalCriticos = dados.insights.filter(function(item) { return item.nivel === 'critical'; }).length;
      var totalAtencao = dados.insights.filter(function(item) { return item.nivel === 'warning'; }).length;
      var modo = dados.contexto.modoTeste
        ? '<div class="banner"><strong>Cenário de teste:</strong> referência em ' + esc(dados.contexto.referencia) +
          '. O painel usa o último registro para simular passado, presente e projeção futura.</div>'
        : '';

      document.getElementById('app').innerHTML =
        '<aside class="dv-side">' +
          '<div class="dv-brand">TDJ</div>' +
          '<nav class="dv-nav" role="tablist" aria-label="Painéis do dashboard">' +
            '<button role="tab" id="dv-nav-visao" class="dv-nav-btn" onclick="mudarPainel(\\'visao\\')">Visão geral</button>' +
            '<button role="tab" id="dv-nav-modalidades" class="dv-nav-btn" onclick="mudarPainel(\\'modalidades\\')">Modalidades</button>' +
            '<button role="tab" id="dv-nav-insights" class="dv-nav-btn" onclick="mudarPainel(\\'insights\\')">Insights</button>' +
            '<button role="tab" id="dv-nav-qualidade" class="dv-nav-btn" onclick="mudarPainel(\\'qualidade\\')">Qualidade</button>' +
            '<button role="tab" id="dv-nav-mapa" class="dv-nav-btn" onclick="mudarPainel(\\'mapa\\')">Mapa</button>' +
          '</nav>' +
          '<div class="dv-watermark" aria-hidden="true">Ezequiel R</div>' +
        '</aside>' +
        '<main class="dv-main" id="dv-main">' +
          '<div class="topbar">' +
            '<div class="page-heading"><h1 id="dv-page-title">VISÃO GERAL DAS INSCRIÇÕES</h1></div>' +
            '<div class="actions">' +
            '<button id="btnAtualizar" class="btn" onclick="atualizar()">Atualizar</button>' +
            '<button class="btn" onclick="google.script.host.close()">Fechar</button></div>' +
          '</div>' + modo +
          '<section id="dv-panel-visao" class="dv-panel" role="tabpanel">' +
          '<section class="dv-hero-grid">' +
            '<div class="card dv-chart-card">' +
              '<div class="dv-chart-head"><div><div class="eyebrow">EVOLUÇÃO DIÁRIA</div>' +
              '<div id="dv-chart-total" class="dv-big-number">' + inteiro(m.total) + '</div>' +
              '<div id="dv-chart-caption" class="dv-caption">inscrições no período · ' + percentual(m.atingimentoMeta) + ' da meta</div></div>' +
              '<div class="dv-time-filter" aria-label="Filtro de período do gráfico">' +
                '<button class="dv-time-btn" data-chart-period="7" onclick="mudarPeriodoGrafico(7)">7d</button>' +
                '<button class="dv-time-btn" data-chart-period="12" onclick="mudarPeriodoGrafico(12)">12d</button>' +
                '<button class="dv-time-btn" data-chart-period="tudo" onclick="mudarPeriodoGrafico(\\'tudo\\')">Tudo</button>' +
              '</div></div>' +
              '<div id="chart-evolution" class="dv-bars dv-echart-main"></div>' +
            '</div>' +
            '<div class="dv-stat-stack">' +
              '<div class="dv-stat ' + (m.projecao < m.metaGeral ? 'alert' : '') + '"><span>PROJEÇÃO FINAL</span>' +
                '<div class="dv-stat-value-row"><strong>' + inteiro(m.projecao) + '</strong>' + setaKpi(m.projecao >= m.metaGeral ? 'green' : 'red') + '</div><small>' + percentual(m.projecaoAtingimento) + ' da meta</small></div>' +
              '<div class="dv-stat ' + (m.mediaRecente < m.ritmoNecessario ? 'alert' : '') + '"><span>RITMO RECENTE</span>' +
                '<div class="dv-stat-value-row"><strong>' + decimal(m.mediaRecente, 1) + '/dia</strong>' + setaKpi(tomRitmo) + '</div><small>necessário: ' + decimal(m.ritmoNecessario, 1) + '/dia</small></div>' +
              '<div class="dv-stat"><span>PRAZO RESTANTE</span><strong>' + inteiro(m.diasRestantes) + ' dias</strong>' +
                '<small>referência em ' + esc(dados.contexto.referencia) + '</small></div>' +
            '</div>' +
          '</section>' +
          '<section class="kpis">' +
            kpi('Atingimento da meta', percentual(m.atingimentoMeta), inteiro(m.total) + ' de ' + inteiro(m.metaGeral), tomInscricoes, true) +
            kpi('Crescimento semanal', crescimento, detalheCrescimento, tomCrescimento, true) +
            kpi('Linhas com pendência', inteiro(dados.qualidade.totalLinhasComPendencia), percentual(dados.qualidade.completude) + ' de completude', dados.qualidade.totalLinhasComPendencia ? 'red' : 'green', false) +
            '<div class="card dv-map-kpi" role="button" tabindex="0" onclick="mudarPainel(\\'mapa\\')" onkeydown="if(event.key===\\'Enter\\'||event.key===\\' \\'){event.preventDefault();mudarPainel(\\'mapa\\')}"><div class="dv-map-kpi-copy"><span>MAPA / GEO 3D</span><strong>Visão rápida</strong></div><div id="chart-map-preview" class="dv-map-preview"></div></div>' +
          '</section>' +
          '<section class="dv-content">' +
            '<div class="card"><div class="section-title"><h2>Desempenho por modalidade</h2>' +
              '<span class="muted">inscritos / limite · tendência 7d</span></div>' +
              '<div class="rows">' + linhasModalidades(dados) + '</div></div>' +
            '<div class="dv-side-panels">' +
              '<div class="card insights-hero">' +
                '<div class="section-title"><div><h2>Insights prioritários</h2></div>' +
                '<div class="insights-summary"><span class="critical-count">' + inteiro(totalCriticos) + ' críticos</span>' +
                '<span class="warning-count">' + inteiro(totalAtencao) + ' atenção</span></div></div>' +
                '<div class="insights">' + cardsInsights(dados.insights) + '</div>' +
              '</div>' +
              '<div class="card dv-quality">' +
                '<div class="section-title"><h2>Qualidade dos dados</h2><span class="muted">prioridade</span></div>' +
                '<div class="quality-score ' + (dados.qualidade.totalLinhasComPendencia ? 'alert' : '') + '">' + inteiro(dados.qualidade.totalLinhasComPendencia) + '</div>' +
                '<div class="muted">linhas com pendência para conferência</div>' +
                '<div class="quality-grid">' +
                  '<div class="quality-item"><strong>' + inteiro(dados.qualidade.menores.semAutorizacao) + '</strong><span>menores diretos sem termo</span></div>' +
                  '<div class="quality-item"><strong>' + inteiro(dados.qualidade.excedentesDuplicados) + '</strong><span>duplicidades</span></div>' +
                  '<div class="quality-item"><strong>' + inteiro(dados.qualidade.validacoes.cpfInvalido) + '</strong><span>CPF inválido</span></div>' +
                  '<div class="quality-item"><strong>' + inteiro(dados.qualidade.linhasIncompletas) + '</strong><span>incompletas</span></div>' +
                '</div>' +
              '</div>' +
            '</div>' +
          '</section>' +
          '</section>' +
          '<section id="dv-panel-modalidades" class="dv-panel" role="tabpanel" hidden>' + painelModalidades(dados) + '</section>' +
          '<section id="dv-panel-insights" class="dv-panel" role="tabpanel" hidden>' + painelInsightsTemporais(dados) + '</section>' +
          '<section id="dv-panel-qualidade" class="dv-panel" role="tabpanel" hidden>' + painelQualidade(dados) + '</section>' +
          '<section id="dv-panel-mapa" class="dv-panel" role="tabpanel" hidden>' + painelMapa(dados) + '</section>' +
          '<div class="footer">Análise agregada · sem exposição de dados pessoais · planilha somente leitura</div>' +
        '</main>';
      mudarPainel(PAINEL_ATIVO);
      mudarPeriodoInsight(PERIODO_INSIGHT_ATIVO);
    }

    function atualizar() {
      var botao = document.getElementById('btnAtualizar');
      if (botao) {
        botao.disabled = true;
        botao.textContent = 'Atualizando…';
      }

      google.script.run
        .withSuccessHandler(function(dados) {
          renderizar(dados);
        })
        .withFailureHandler(function(erro) {
          document.getElementById('app').innerHTML = '<div class="error"><strong>Erro ao atualizar:</strong> ' +
            esc(erro && erro.message ? erro.message : erro) + '</div>';
        })
        .atualizarDadosBiInsightsCoopsportes();
    }

    var RESIZE_TIMER = null;
    window.addEventListener('resize', function() {
      window.clearTimeout(RESIZE_TIMER);
      RESIZE_TIMER = window.setTimeout(function() {
        Object.keys(CHARTS).forEach(function(id) { try { CHARTS[id].resize(); } catch(e) {} });
      }, 90);
    });

    renderizar(INITIAL_DATA);
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html);
}

function coopsportesBiInsightsStatusCapacidade_(ocupacao) {
  if (ocupacao >= 1) return 'Lotada';
  if (ocupacao >= 0.9) return 'Crítica';
  if (ocupacao >= 0.75) return 'Atenção';
  return 'Disponível';
}

function coopsportesBiInsightsResolverFaixaEtaria_(valor) {
  var normalizado = coopsportesBiInsightsNormalizar_(valor);
  var faixas = {
    '0 ate 15 anos': '0 Até 15 anos',
    '16 a 20 anos': '16 a 20 anos',
    '21 a 26 anos': '21 a 26 anos',
    '27 a 35 anos': '27 a 35 anos',
    '36 anos ou mais': '36 anos ou mais'
  };

  return faixas[normalizado] || 'Não informado';
}

function coopsportesBiInsightsProporcaoTeste_(registros, indiceNome) {
  if (indiceNome < 0 || !registros.length) return 0;

  var testes = registros.filter(function(linha) {
    return coopsportesBiInsightsNormalizar_(linha[indiceNome]).indexOf('teste') >= 0;
  }).length;

  return testes / registros.length;
}

function coopsportesBiInsightsConverterData_(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return new Date(valor.getTime());
  }

  if (typeof valor === 'string' && valor.trim()) {
    var texto = valor.trim();
    var brasileira = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);

    if (brasileira) {
      var dataBr = new Date(
        Number(brasileira[3]),
        Number(brasileira[2]) - 1,
        Number(brasileira[1]),
        Number(brasileira[4] || 0),
        Number(brasileira[5] || 0),
        Number(brasileira[6] || 0)
      );
      return isNaN(dataBr.getTime()) ? null : dataBr;
    }

    var dataTexto = new Date(texto);
    return isNaN(dataTexto.getTime()) ? null : dataTexto;
  }

  return null;
}

function coopsportesBiInsightsInicioDia_(data) {
  var copia = new Date(data.getTime());
  copia.setHours(0, 0, 0, 0);
  return copia;
}

function coopsportesBiInsightsAdicionarDias_(data, dias) {
  var copia = coopsportesBiInsightsInicioDia_(data);
  copia.setDate(copia.getDate() + dias);
  return copia;
}

function coopsportesBiInsightsDiferencaDias_(inicio, fim) {
  var inicioUtc = Date.UTC(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());
  var fimUtc = Date.UTC(fim.getFullYear(), fim.getMonth(), fim.getDate());
  return Math.round((fimUtc - inicioUtc) / 86400000);
}

function coopsportesBiInsightsChaveDia_(data, fuso) {
  return Utilities.formatDate(data, fuso, 'yyyy-MM-dd');
}

function coopsportesBiInsightsFormatarData_(data, fuso) {
  return Utilities.formatDate(data, fuso, 'dd/MM/yyyy');
}

function coopsportesBiInsightsFormatarDataHora_(data, fuso) {
  return Utilities.formatDate(data, fuso, 'dd/MM/yyyy HH:mm:ss');
}

function coopsportesBiInsightsNumero_(valor, padrao) {
  if (typeof valor === 'number' && isFinite(valor)) return valor;

  var numero = Number(String(valor || '').replace(/\./g, '').replace(',', '.'));
  return isFinite(numero) ? numero : padrao;
}

function coopsportesBiInsightsNormalizar_(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function coopsportesBiInsightsVazio_(valor) {
  if (valor === null || valor === undefined || valor === '') return true;
  if (typeof valor !== 'string') return false;
  var texto = valor.trim();
  if (!texto) return true;
  return /^#(NUM!|N\/A|VALUE!|REF!|DIV\/0!|NAME\?)$/i.test(texto);
}

function coopsportesBiInsightsInteiro_(valor) {
  return Math.round(Number(valor || 0)).toLocaleString('pt-BR');
}

function coopsportesBiInsightsDecimal_(valor, casas) {
  return Number(valor || 0).toLocaleString('pt-BR', {
    minimumFractionDigits: casas,
    maximumFractionDigits: casas
  });
}

function coopsportesBiInsightsPercentual_(valor) {
  return Number(valor || 0).toLocaleString('pt-BR', {
    style: 'percent',
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  });
}

function coopsportesBiInsightsExibirErro_(erro) {
  var mensagem = erro && erro.message ? erro.message : String(erro);

  SpreadsheetApp.getUi().alert(
    'Erro no BI / INSIGHTS',
    mensagem,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
