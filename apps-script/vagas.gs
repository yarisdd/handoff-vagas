/**
 * Handoff Vagas — conferência diária das vagas do site.
 *
 * Vai como um ARQUIVO NOVO (vagas.gs) no Apps Script da planilha Handoff — Base.
 * Tudo aqui tem o prefixo "vagas" para não colidir com o código das ligas,
 * e o script só mexe na aba "Vagas do site".
 *
 * Instalação (uma vez só):
 *   1. Na Handoff — Base: Extensões > Apps Script > "+" ao lado de Arquivos > Script > nome: vagas
 *   2. Cole este conteúdo e salve.
 *   3. No topo, escolha a função vagasInstalar e clique em Executar. Autorize quando pedir.
 *   Isso cria a aba "Vagas do site" com as vagas atuais, o menu "Handoff Vagas",
 *   o gatilho diário das 7h, e faz a primeira conferência.
 *
 * Rotina:
 *   - Todo dia às 7h, vagasConferir() abre cada link, marca o que saiu do ar ou mudou
 *     e te manda um e-mail só quando há algo para olhar.
 *   - Você corrige a linha, seleciona a linha e usa o menu
 *     Handoff Vagas > Marcar linhas selecionadas como conferidas.
 *   - Coluna "publicar" = não esconde a vaga do site sem apagar.
 */

var VAGAS = {
  ABA: 'Vagas do site',
  TZ: 'America/Sao_Paulo',
  HORA: 7,
  DIAS_SEM_ACESSO: 14, // sites que bloqueiam a checagem voltam para a lista a cada 14 dias
  // Colunas que você edita (o site lê estas)
  COLS: ['publicar', 'titulo', 'instituicao', 'categoria', 'pais', 'cidade', 'modalidade', 'publico', 'custo',
    'inscricao_abre', 'inscricao_fecha', 'periodo_programa', 'recorrente', 'ciclo', 'resumo', 'link', 'conferido_em'],
  // Colunas preenchidas pelo script
  MON: ['conferir', 'motivo', 'situacao_link', 'pagina_mudou_em', 'datas_na_pagina', 'ultima_checagem', 'assinatura'],
  DATAS: ['inscricao_abre', 'inscricao_fecha', 'conferido_em', 'pagina_mudou_em', 'ultima_checagem'],
  OPCOES: {
    publicar: ['sim', 'não'],
    recorrente: ['sim', 'não'],
    categoria: ['Eletiva clínica', 'Curso de verão/inverno', 'Estágio/Observership', 'Pesquisa', 'Intercâmbio',
      'Evento científico', 'Bolsa', 'Curso online'],
    modalidade: ['Presencial', 'Online', 'Híbrido'],
    custo: ['Gratuito', 'Pago', 'Com bolsa/auxílio', 'Varia']
  },
  // Só os trechos que falam de prazo/inscrição entram na assinatura, para notícias e banners não dispararem alerta
  CHAVES: /(inscri|prazo|edital|sele[cç][aã]o|vagas?\b|deadline|applica|apply|registration|regist|submission|abstract|closing|closes|opens|abert|encerr|janeiro|fevereiro|mar[cç]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro|january|february|march|april|may|june|july|august|september|october|november|december|20(2[6-9]|3\d))/i,
  MESES: { janeiro: 1, fevereiro: 2, 'março': 3, marco: 3, abril: 4, maio: 5, junho: 6, julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12,
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    jan: 1, feb: 2, fev: 2, mar: 3, apr: 4, abr: 4, jun: 6, jul: 7, aug: 8, ago: 8, sep: 9, sept: 9, set: 9, oct: 10, out: 10, nov: 11, dec: 12, dez: 12 }
};

// ---------------------------------------------------------------- menu (gatilho próprio, não substitui o onOpen das ligas)

function vagasMenu() {
  SpreadsheetApp.getUi().createMenu('Handoff Vagas')
    .addItem('Conferir agora', 'vagasConferir')
    .addItem('Marcar linhas selecionadas como conferidas', 'vagasMarcarConferidas')
    .addSeparator()
    .addItem('Reinstalar gatilhos das vagas', 'vagasCriarGatilhos')
    .addToUi();
}

// ---------------------------------------------------------------- instalação

function vagasInstalar() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(VAGAS.ABA);
  if (sh && sh.getLastRow() > 1) {
    throw new Error('A aba "' + VAGAS.ABA + '" já tem vagas. Para reinstalar do zero, apague só essa aba e rode vagasInstalar() de novo.');
  }
  if (!sh) sh = ss.insertSheet(VAGAS.ABA, ss.getSheets().length); // entra no fim, sem mudar a ordem das outras abas

  var header = VAGAS.COLS.concat(VAGAS.MON);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#2B45D6').setFontColor('#FFFFFF');
  sh.getRange(1, VAGAS.COLS.length + 1, 1, VAGAS.MON.length).setBackground('#D9DEEA').setFontColor('#14206B');
  sh.getRange(1, VAGAS.COLS.length + 1).setNote('Daqui para a direita, as colunas são preenchidas pelo script todo dia.');

  var linhas = VAGAS_DADOS.map(function (o) {
    return VAGAS.COLS.map(function (c) {
      var v = o[c];
      if (c === 'publicar') return 'sim';
      if (c === 'recorrente') return v ? 'sim' : 'não';
      if (VAGAS.DATAS.indexOf(c) >= 0) return vagas_paraData(v);
      return v == null ? '' : v;
    }).concat(VAGAS.MON.map(function () { return ''; }));
  });
  if (linhas.length) sh.getRange(2, 1, linhas.length, header.length).setValues(linhas);

  var n = Math.max(linhas.length + 200, 300);
  VAGAS.DATAS.forEach(function (c) { sh.getRange(2, header.indexOf(c) + 1, n, 1).setNumberFormat('dd/mm/yyyy'); });
  Object.keys(VAGAS.OPCOES).forEach(function (c) {
    var regra = SpreadsheetApp.newDataValidation().requireValueInList(VAGAS.OPCOES[c], true).setAllowInvalid(false).build();
    sh.getRange(2, header.indexOf(c) + 1, n, 1).setDataValidation(regra);
  });
  var conf = sh.getRange(2, header.indexOf('conferir') + 1, n, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SIM').setBackground('#FBE7D6').setFontColor('#A84C06').setRanges([conf]).build()
  ]);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.setColumnWidth(header.indexOf('titulo') + 1, 280);
  sh.setColumnWidth(header.indexOf('resumo') + 1, 360);
  sh.setColumnWidth(header.indexOf('ciclo') + 1, 300);
  sh.getRange(2, 1, n, header.length).setWrap(false).setVerticalAlignment('top');

  vagasCriarGatilhos();
  vagasMenu();
  vagasConferir();
}

function vagasCriarGatilhos() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'vagasConferir' || f === 'vagasMenu') ScriptApp.deleteTrigger(t); // só os gatilhos das vagas
  });
  ScriptApp.newTrigger('vagasConferir').timeBased().everyDays(1).atHour(VAGAS.HORA).inTimezone(VAGAS.TZ).create();
  ScriptApp.newTrigger('vagasMenu').forSpreadsheet(ss).onOpen().create();
}

// ---------------------------------------------------------------- conferência diária

function vagasConferir() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(VAGAS.ABA);
  if (!sh) throw new Error('Aba "' + VAGAS.ABA + '" não encontrada. Rode vagasInstalar() primeiro.');
  var valores = sh.getDataRange().getValues();
  var head = valores[0].map(function (h) { return String(h).trim(); });
  var ix = {};
  head.forEach(function (h, i) { ix[h] = i; });
  VAGAS.MON.concat(['link', 'titulo']).forEach(function (c) { if (ix[c] == null) throw new Error('Coluna "' + c + '" não encontrada na linha 1 da aba ' + VAGAS.ABA + '.'); });

  var hoje = vagas_inicioDoDia(new Date());
  var rel = { fora: [], semAcesso: [], mudou: [], fecha: [], passou: [] };
  var saida = [];

  for (var r = 1; r < valores.length; r++) {
    var linha = valores[r];
    var mon = VAGAS.MON.map(function (c) { return linha[ix[c]]; });
    var m = {};
    VAGAS.MON.forEach(function (c, i) { m[c] = mon[i]; });
    var titulo = String(linha[ix.titulo] || '').trim();
    var link = String(linha[ix.link] || '').trim();
    if (!titulo) { saida.push(mon); continue; }
    var publicada = String(linha[ix.publicar] || 'sim').toLowerCase().indexOf('n') !== 0;
    var item = { titulo: titulo, linha: r + 1, link: link };

    // 1. Abre o link
    if (/^https?:\/\//i.test(link)) {
      var res = vagas_abrir(link);
      m.ultima_checagem = new Date();
      if (res.ok) {
        m.situacao_link = 'OK';
        if (m.assinatura && m.assinatura !== res.assinatura) m.pagina_mudou_em = new Date();
        m.assinatura = res.assinatura;
        m.datas_na_pagina = res.datas;
      } else if (res.codigo === 404 || res.codigo === 410 || res.codigo === -1) {
        m.situacao_link = 'Fora do ar' + (res.codigo > 0 ? ' (' + res.codigo + ')' : ' (site não responde)');
      } else {
        m.situacao_link = 'Não deu para checar (' + res.codigo + ')';
      }
    } else {
      m.situacao_link = 'Sem link';
    }

    // 2. Decide se precisa de conferência humana
    var conferido = vagas_paraData(linha[ix.conferido_em], true);
    var mudou = vagas_paraData(m.pagina_mudou_em, true);
    var fecha = vagas_paraData(linha[ix.inscricao_fecha]);
    var abre = vagas_paraData(linha[ix.inscricao_abre]);
    var recorrente = String(linha[ix.recorrente] || '').toLowerCase().indexOf('s') === 0;
    var motivos = [];
    if (m.situacao_link.indexOf('Fora do ar') === 0 || m.situacao_link === 'Sem link') { motivos.push('link fora do ar'); rel.fora.push(item); }
    else if (m.situacao_link.indexOf('Não deu') === 0 && (!conferido || (hoje - conferido) / 86400000 >= VAGAS.DIAS_SEM_ACESSO)) {
      // site que bloqueia robô: pede conferência manual a cada 14 dias, não todo dia
      motivos.push('site bloqueou a checagem: abrir na mão'); rel.semAcesso.push(item);
    }
    if (mudou && (!conferido || mudou > conferido)) { motivos.push('página mudou em ' + vagas_fmt(mudou)); rel.mudou.push(item); }
    if (fecha && fecha < hoje && recorrente && (!conferido || conferido <= fecha)) { motivos.push('prazo passou: procurar o próximo ciclo'); rel.passou.push(item); }
    if (publicada && fecha && fecha >= hoje && (!abre || abre <= hoje)) {
      var d = Math.round((fecha - hoje) / 86400000);
      if (d <= 7) { item.d = d; rel.fecha.push(item); }
    }
    m.conferir = motivos.length ? 'SIM' : '';
    m.motivo = motivos.join('; ');
    saida.push(VAGAS.MON.map(function (c) { return m[c] == null ? '' : m[c]; }));
  }

  if (saida.length) sh.getRange(2, ix[VAGAS.MON[0]] + 1, saida.length, VAGAS.MON.length).setValues(saida);
  vagas_enviarResumo(ss, sh, rel);
}

function vagas_abrir(url) {
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; HandoffVagas/1.0; +https://www.instagram.com/handoff.med/)' }
    });
    var codigo = resp.getResponseCode();
    if (codigo < 200 || codigo >= 400) return { ok: false, codigo: codigo };
    var tipo = String((resp.getHeaders() || {})['Content-Type'] || '').toLowerCase();
    if (tipo.indexOf('pdf') >= 0 || tipo.indexOf('octet-stream') >= 0 || /\.pdf($|\?)/i.test(url)) {
      return { ok: true, codigo: codigo, assinatura: vagas_hash(resp.getContent()), datas: '(PDF: checa só se o arquivo mudou)' };
    }
    var texto = vagas_limpar(resp.getContentText());
    var trechos = vagas_trechos(texto);
    return { ok: true, codigo: codigo, assinatura: vagas_hash(trechos || texto), datas: vagas_datas(trechos || texto) };
  } catch (e) {
    return { ok: false, codigo: -1 };
  }
}

// ---------------------------------------------------------------- e-mail

function vagas_enviarResumo(ss, sh, rel) {
  var total = rel.fora.length + rel.semAcesso.length + rel.mudou.length + rel.fecha.length + rel.passou.length;
  if (!total) return;
  var base = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A';
  function lista(titulo, itens, extra) {
    if (!itens.length) return '';
    return '<h3 style="margin:18px 0 6px;font-family:Arial;color:#14206B">' + titulo + ' (' + itens.length + ')</h3><ul style="margin:0;padding-left:18px;font-family:Arial;font-size:14px">' +
      itens.map(function (it) {
        return '<li style="margin:4px 0"><a href="' + base + it.linha + '">' + vagas_esc(it.titulo) + '</a>' + (extra ? extra(it) : '') +
          (it.link ? ' · <a href="' + vagas_esc(it.link) + '">abrir página</a>' : '') + '</li>';
      }).join('') + '</ul>';
  }
  var html = '<div style="font-family:Arial;color:#0E1738">' +
    '<p style="font-size:15px">Conferência das vagas do site em ' + vagas_fmt(new Date()) + '. Abra cada uma, corrija a linha na aba <b>' + VAGAS.ABA + '</b> e use <b>Handoff Vagas &gt; Marcar linhas selecionadas como conferidas</b>.</p>' +
    lista('Fecham em até 7 dias (bom para divulgar)', rel.fecha, function (it) { return ' — ' + (it.d === 0 ? 'fecha hoje' : 'D-' + it.d); }) +
    lista('Página mudou: conferir prazo e regras', rel.mudou) +
    lista('Prazo passou: procurar o próximo ciclo', rel.passou) +
    lista('Link fora do ar', rel.fora) +
    lista('Site bloqueou a checagem automática (abrir na mão)', rel.semAcesso) +
    '</div>';
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Handoff Vagas · ' + (rel.mudou.length + rel.passou.length + rel.fora.length) + ' para conferir, ' + rel.fecha.length + ' fechando',
    htmlBody: html
  });
}

// ---------------------------------------------------------------- marcar como conferida

function vagasMarcarConferidas() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();
  if (sh.getName() !== VAGAS.ABA) { ss.toast('Abra a aba "' + VAGAS.ABA + '" e selecione as linhas.'); return; }
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var cConf = head.indexOf('conferido_em') + 1, cFlag = head.indexOf('conferir') + 1, cMot = head.indexOf('motivo') + 1;
  var n = 0;
  sh.getActiveRangeList().getRanges().forEach(function (rg) {
    for (var r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) {
      if (r < 2) continue;
      sh.getRange(r, cConf).setValue(new Date());
      sh.getRange(r, cFlag).setValue('');
      sh.getRange(r, cMot).setValue('');
      n++;
    }
  });
  ss.toast(n + (n === 1 ? ' linha marcada' : ' linhas marcadas') + ' como conferida' + (n === 1 ? '' : 's') + ' hoje.');
}

// ---------------------------------------------------------------- utilidades (prefixo vagas_ para não colidir)

function vagas_limpar(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<(br|\/p|\/li|\/h\d|\/tr|\/div)\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(+n); })
    .replace(/[ \t\r]+/g, ' ')
    .replace(/\n\s*/g, '\n')
    .trim();
}
function vagas_trechos(texto) {
  return texto.split(/\n|[.!?]\s+/).map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 8 && s.length < 400 && VAGAS.CHAVES.test(s); })
    .join('\n');
}
function vagas_datas(texto) {
  var achadas = [], visto = {};
  var anoMin = new Date().getFullYear();
  function add(d, m, a) {
    d = +d; m = +m; a = +a;
    if (a < 100) a += 2000;
    if (!d || !m || m > 12 || d > 31 || a < anoMin || a > anoMin + 2) return;
    var s = ('0' + d).slice(-2) + '/' + ('0' + m).slice(-2) + '/' + a;
    if (!visto[s]) { visto[s] = 1; achadas.push(s); }
  }
  var x, re1 = /\b(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})\b/g;
  while ((x = re1.exec(texto))) add(x[1], x[2], x[3]);
  var re2 = /\b(\d{4})-(\d{2})-(\d{2})\b/g;
  while ((x = re2.exec(texto))) add(x[3], x[2], x[1]);
  var nomes = Object.keys(VAGAS.MESES).sort(function (a, b) { return b.length - a.length; }).join('|');
  var re3 = new RegExp('\\b(\\d{1,2})(?:º|st|nd|rd|th)?\\s+(?:de\\s+)?(' + nomes + ')\\.?\\s+(?:de\\s+)?(\\d{4})\\b', 'gi');
  while ((x = re3.exec(texto))) add(x[1], VAGAS.MESES[x[2].toLowerCase()], x[3]);
  var re4 = new RegExp('\\b(' + nomes + ')\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})\\b', 'gi');
  while ((x = re4.exec(texto))) add(x[2], VAGAS.MESES[x[1].toLowerCase()], x[3]);
  return achadas.slice(0, 8).join('; ');
}
function vagas_hash(conteudo) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, conteudo)).slice(0, 16);
}
function vagas_paraData(v, manterHora) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? '' : (manterHora ? v : vagas_inicioDoDia(v));
  var s = String(v).trim(), m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return '';
}
function vagas_inicioDoDia(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function vagas_fmt(d) { return Utilities.formatDate(d, VAGAS.TZ, 'dd/MM/yyyy'); }
function vagas_esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// ---------------------------------------------------------------- vagas iniciais (usadas só pelo vagasInstalar)

var VAGAS_DADOS = [
 {
  "publicar": "sim",
  "titulo": "Curso Teórico-Prático de Pediatria Hospitalar",
  "instituicao": "Instituto PENSI / Sabará Hospital Infantil",
  "categoria": "Estágio/Observership",
  "pais": "Brasil",
  "cidade": "São Paulo/SP",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina de qualquer faculdade reconhecida pelo MEC, a partir do 5º período",
  "custo": "Pago",
  "inscricao_abre": "2026-09-21",
  "inscricao_fecha": "2026-10-25",
  "periodo_programa": "jan 2027 (Turma A 04–15/01; Turma B 18–29/01)",
  "recorrente": true,
  "ciclo": "Semestral, nas férias de janeiro e julho. Taxa de inscrição de R$ 45 e matrícula de R$ 900; são 30 vagas em 2 turmas de 15.",
  "resumo": "Imersão de cerca de duas semanas num hospital pediátrico de referência, com aulas teóricas e acompanhamento observacional supervisionado em várias áreas do hospital. Dá certificado.",
  "link": "https://institutopensi.org.br/ensino/curso-tp-de-pediatria-hospitalar",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Programa Bolsas de Verão (PBV)",
  "instituicao": "Centro Nacional de Pesquisa em Energia e Materiais (CNPEM)",
  "categoria": "Pesquisa",
  "pais": "Brasil",
  "cidade": "Campinas/SP",
  "modalidade": "Presencial",
  "publico": "Graduandos de ciências biológicas e da saúde, exatas e engenharias da América Latina e Caribe",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": "2026-08-10",
  "inscricao_fecha": "2026-09-08",
  "periodo_programa": "06/01–26/02/2027",
  "recorrente": true,
  "ciclo": "Anual. Inscrições em agosto e setembro, com resultado em outubro.",
  "resumo": "Oito semanas de pesquisa nos laboratórios do CNPEM com orientador. O programa paga passagem, hospedagem, alimentação, transporte e seguro. É muito concorrido: mais de mil inscritos para cerca de 30 vagas.",
  "link": "https://pages.cnpem.br/bolsasdeverao/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "XII Curso de Verão do PPG em Ciências Biológicas",
  "instituicao": "Universidade Federal de Goiás (UFG) – ICB",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Goiânia/GO",
  "modalidade": "Presencial",
  "publico": "Graduandos e graduados das ciências biológicas, da saúde e afins (10 vagas para graduação)",
  "custo": "Pago",
  "inscricao_abre": "2026-09-14",
  "inscricao_fecha": "2026-09-30",
  "periodo_programa": "18–22/01/2027",
  "recorrente": true,
  "ciclo": "Anual, com inscrições em setembro. Taxa de R$ 60 para graduandos.",
  "resumo": "Uma semana de palestras e workshops pela manhã e aulas práticas de laboratório à tarde, para quem quer conhecer a pesquisa básica em biologia e saúde.",
  "link": "https://files.cercomp.ufg.br/weby/up/101/o/Edital_XII_Curso_de_Ver%C3%A3o_do_PPGCB_2027.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Verão em Genética (CVGen)",
  "instituicao": "Faculdade de Medicina de Ribeirão Preto (FMRP-USP)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Ribeirão Preto/SP",
  "modalidade": "Presencial",
  "publico": "Graduandos nos 2 últimos anos ou recém-formados, aprovados em ao menos uma disciplina de Genética",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "jan–fev (2026: 26/01–04/02)",
  "recorrente": true,
  "ciclo": "Anual. As inscrições costumam fechar entre agosto e setembro. A edição de 2027 ainda não foi publicada no site. Seguro de vida é obrigatório e a hospedagem fica por conta do aluno.",
  "resumo": "Um dos cursos de verão mais tradicionais do país: cerca de 40 alunos de todo o Brasil passam duas semanas em laboratórios de genética e fazem minicursos. A seleção é pelo currículo Lattes.",
  "link": "https://veraogen.fmrp.usp.br/inscricoes/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Verão em Oncologia do INCA",
  "instituicao": "Instituto Nacional de Câncer (INCA)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Rio de Janeiro/RJ",
  "modalidade": "Híbrido",
  "publico": "Graduandos de biologia e saúde de faculdades públicas e privadas",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "jan–fev (2026: 25/01–06/02)",
  "recorrente": true,
  "ciclo": "Anual. As inscrições costumam fechar entre agosto e setembro (14/09/2025 para a edição de 2026). O prazo da edição de 2027 ainda não saiu.",
  "resumo": "Duas semanas de imersão em pesquisa oncológica nas unidades do INCA, com cerca de 40 vagas presenciais em minicursos. As aulas teóricas são transmitidas pelo YouTube.",
  "link": "https://www.gov.br/inca/pt-br/assuntos/eventos/xvi-curso-de-verao-em-oncologia-do-inca",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Cursos de Férias do IOC – Verão",
  "instituicao": "Instituto Oswaldo Cruz (IOC/Fiocruz)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Rio de Janeiro/RJ",
  "modalidade": "Presencial",
  "publico": "Graduandos matriculados; vários cursos aceitam a área da saúde (alguns a partir do 2º ou 3º período)",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de jan–início de fev",
  "recorrente": true,
  "ciclo": "Edições de verão e de inverno. O verão de 2026 teve inscrições de 24 a 28/11/2025 pelo Campus Virtual Fiocruz, com taxa de R$ 20. Se o padrão se mantiver, as inscrições do verão de 2027 abrem em novembro de 2026.",
  "resumo": "Cursos práticos de uma a duas semanas no campus de Manguinhos, em temas como diagnóstico de doenças parasitárias, bioinformática e proteínas recombinantes.",
  "link": "https://www.ioc.fiocruz.br/sites/default/files/03._chamada_externa_-_cursos_de_ferias_verao_2026_-_ver_07_11_25.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Verão em Farmacologia da Unicamp",
  "instituicao": "Faculdade de Ciências Médicas da Unicamp",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Campinas/SP",
  "modalidade": "Presencial",
  "publico": "Graduandos e recém-formados com interesse em farmacologia e pós-graduação",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de janeiro (2026: 26–30/01)",
  "recorrente": true,
  "ciclo": "Anual. As inscrições costumam ser entre setembro e outubro. O prazo da edição de 2027 ainda não foi publicado.",
  "resumo": "Curso de extensão gratuito de uma semana para conhecer a pesquisa em farmacologia, com turma pequena de cerca de 15 selecionados.",
  "link": "https://farmacounicamp.wixsite.com/curso",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Inverno em Farmacologia da UFSC",
  "instituicao": "Universidade Federal de Santa Catarina (UFSC)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Florianópolis/SC",
  "modalidade": "Híbrido",
  "publico": "Graduandos e formados nas áreas da saúde, biológicas e afins",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de julho (2026: 27–31/07)",
  "recorrente": true,
  "ciclo": "Anual, com inscrições entre abril e junho. Em 2024 a taxa foi de R$ 100, paga só pelos aprovados. A edição de 2027 ainda não foi publicada.",
  "resumo": "Semana de aulas e atividades em farmacologia com a pós-graduação da UFSC, com cerca de 40 vagas presenciais e opção online.",
  "link": "https://cursodeinvernofarmacologia.ufsc.br/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Inverno em Microbiologia e Biologia Molecular Aplicada",
  "instituicao": "Instituto de Ciências Biomédicas da USP (ICB-USP)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "São Paulo/SP",
  "modalidade": "Híbrido",
  "publico": "Graduandos que já cursaram microbiologia e recém-formados (até 2 anos)",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de julho (2026: 27–31/07)",
  "recorrente": true,
  "ciclo": "Anual, com inscrições entre março e abril (2026: 09/03 a 10/04). Deslocamento e hospedagem ficam por conta do aluno.",
  "resumo": "Curso intensivo de uma semana sobre microbiologia e biologia molecular aplicada, com 40 vagas presenciais e 60 online.",
  "link": "https://posbmm.icb.usp.br/wp-content/uploads/sites/951/2026/03/Edital-XI-Curso-de-Inverno-em-Microbiologia-e-Biologia-Molecular-Aplicada-2026.docx.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Inverno em Neurociências da UFMG",
  "instituicao": "Universidade Federal de Minas Gerais (UFMG)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Belo Horizonte/MG",
  "modalidade": "Presencial",
  "publico": "Graduandos e recém-graduados interessados em neurociências",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "início de julho (2026: 06–10/07)",
  "recorrente": true,
  "ciclo": "Realizado em julho. Em 2026 as inscrições foram de 28/04 a 15/05, com taxa de R$ 150 e 8 vagas gratuitas para candidatos de baixa renda. A edição de 2027 ainda não foi publicada.",
  "resumo": "Uma semana de aulas sobre neurociências na UFMG, com 80 vagas presenciais.",
  "link": "https://www.ufmg.br//app/uploads/2026/05/Edital-Curso-de-Inverno-em-Neurociencias-da-UFMG-2026-FINAL.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Curso de Verão em Virologia do IEC",
  "instituicao": "Instituto Evandro Chagas (IEC)",
  "categoria": "Curso de verão/inverno",
  "pais": "Brasil",
  "cidade": "Ananindeua/PA",
  "modalidade": "Presencial",
  "publico": "Graduandos em fase final e recém-graduados de biológicas, saúde e agrárias",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de julho (2026: 27–31/07)",
  "recorrente": true,
  "ciclo": "Anual, com inscrições gratuitas em junho. A edição de 2027 ainda não foi publicada.",
  "resumo": "Uma semana de virologia num dos principais institutos de pesquisa em doenças infecciosas da Amazônia, com 80 vagas.",
  "link": "https://www.gov.br/iec/pt-br/instituto-evandro-chagas-abre-inscricoes-para-o-xii-curso-de-verao-em-virologia",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "VER-SUS Florianópolis – Vivências no SUS",
  "instituicao": "Rede Unida, Ministério da Saúde, OPAS, UFSC e SMS Florianópolis",
  "categoria": "Estágio/Observership",
  "pais": "Brasil",
  "cidade": "Florianópolis/SC",
  "modalidade": "Presencial",
  "publico": "Graduandos e residentes em saúde, maiores de 18 anos, com matrícula ativa",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fevereiro (2026: 22–28/02)",
  "recorrente": true,
  "ciclo": "Edições regionais ao longo do ano. Em Florianópolis, as inscrições de 2026 foram de 19/12/2025 a 09/01/2026. A próxima edição ainda não foi publicada.",
  "resumo": "Uma semana de imersão em tempo integral na rede do SUS, visitando serviços e discutindo gestão e controle social. Hospedagem, alimentação e transporte interno são cobertos.",
  "link": "https://www.redeunida.org.br/media/ckeditor_files/2025/12/19/edital-florianopolis.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Ciclo de Práticas Médicas – Med+Férias",
  "instituicao": "Complexo do Hospital de Clínicas da UFPR (CHC-UFPR/Ebserh)",
  "categoria": "Estágio/Observership",
  "pais": "Brasil",
  "cidade": "Curitiba/PR",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina selecionados por edital (requisitos do próximo ciclo a confirmar)",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "férias de janeiro (2026: 26–30/01)",
  "recorrente": true,
  "ciclo": "Ciclo curto nas férias de janeiro. O edital do ciclo de janeiro de 2027 ainda não foi localizado.",
  "resumo": "Prática médica supervisionada de uma semana em diferentes serviços do hospital universitário da UFPR, durante as férias.",
  "link": "https://www.gov.br/hubrasil/pt-br/hospitais-universitarios/regiao-sul/chc-ufpr/ensino-e-pesquisa/atividades-de-graduacao-e-ensino-tecnico/atividades-de-graduacao-ensino-tecnico-e-extensao/Edital_Final_MedMFrias_2025.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Imersão em Cardiologia",
  "instituicao": "Instituto Dante Pazzanese de Cardiologia",
  "categoria": "Estágio/Observership",
  "pais": "Brasil",
  "cidade": "São Paulo/SP",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina do 4º ao 6º ano",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "módulos de 160 h",
  "recorrente": true,
  "ciclo": "Oferecido em módulos. Datas, valores e inscrições saem no portal de cursos do instituto, na categoria Curso de Imersão.",
  "resumo": "Imersão prática em cardiologia num dos maiores hospitais cardiológicos públicos do país, para alunos dos anos finais.",
  "link": "https://cursos.idpc.org.br/image/catalog/pdf/IMERSAO_PS/Imersao_em_Cardiologia_Dante_programa_novo_160.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "PIBIC – Iniciação Científica",
  "instituicao": "CNPq, pelas instituições participantes",
  "categoria": "Pesquisa",
  "pais": "Brasil",
  "cidade": "Na sua faculdade",
  "modalidade": "Presencial",
  "publico": "Graduandos com currículo Lattes e orientador pesquisador; requisitos de cada instituição",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "12 meses, de setembro a agosto",
  "recorrente": true,
  "ciclo": "O aluno não se inscreve no CNPq: a seleção é feita pelo edital interno de cada faculdade, normalmente no 1º semestre.",
  "resumo": "Principal programa nacional de bolsas de iniciação científica: você desenvolve um projeto com orientador e recebe bolsa mensal. É a porta de entrada mais comum para a pesquisa na graduação.",
  "link": "https://www.gov.br/cnpq/pt-br/acesso-a-informacao/acoes-e-programas/programas/programas-ict/pibic",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Projeto Rondon",
  "instituicao": "Ministério da Defesa, com as faculdades participantes",
  "categoria": "Intercâmbio",
  "pais": "Brasil",
  "cidade": "Interior do país (2027: Piauí)",
  "modalidade": "Presencial",
  "publico": "Graduandos selecionados pela própria faculdade participante",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "janeiro de 2027 (cerca de 15 dias)",
  "recorrente": true,
  "ciclo": "Operações periódicas em diferentes regiões. Cada faculdade selecionada abre seu próprio edital para escolher alunos: procure a pró-reitoria de extensão da sua.",
  "resumo": "Projeto de extensão em que equipes universitárias passam cerca de duas semanas em municípios do interior fazendo ações de saúde e cidadania, com apoio logístico das Forças Armadas.",
  "link": "https://www.gov.br/defesa/pt-br/assuntos/programas-sociais/copy_of_projeto-rondon/downloads/operacao-carnauba/dou_edital_op-carnauba.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "65º Congresso Brasileiro de Educação Médica (COBEM 2027)",
  "instituicao": "Associação Brasileira de Educação Médica (ABEM)",
  "categoria": "Evento científico",
  "pais": "Brasil",
  "cidade": "Rio de Janeiro/RJ",
  "modalidade": "Presencial",
  "publico": "Estudantes, docentes e pesquisadores de educação médica",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "07–09/10/2027",
  "recorrente": true,
  "ciclo": "Anual. Em 2026, a submissão de trabalhos foi de março a maio. Os prazos de 2027 ainda não saíram. A submissão é gratuita, mas um autor de cada trabalho aprovado precisa se inscrever.",
  "resumo": "Principal congresso de educação médica do país, com apresentação de trabalhos e grande participação de estudantes.",
  "link": "https://website.abem-educmed.org.br/congressos/cobem/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "COBREM – Congresso Brasileiro de Estudantes de Medicina",
  "instituicao": "Direção Executiva Nacional dos Estudantes de Medicina (DENEM)",
  "categoria": "Evento científico",
  "pais": "Brasil",
  "cidade": "Itinerante",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina de todo o país",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "janeiro",
  "recorrente": true,
  "ciclo": "Anual, com sede que muda a cada edição, em geral em janeiro. Datas e sede da próxima edição ainda não foram divulgadas.",
  "resumo": "Encontro de formação do movimento estudantil de medicina, com debates sobre formação médica, SUS e equidade.",
  "link": "https://doity.com.br/35cobrem",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Congresso Médico Acadêmico da Unicamp (CoMAU)",
  "instituicao": "Alunos da Faculdade de Ciências Médicas da Unicamp",
  "categoria": "Evento científico",
  "pais": "Brasil",
  "cidade": "Campinas/SP",
  "modalidade": "Presencial",
  "publico": "Estudantes da saúde; alguns workshops exigem a partir do 2º ano de medicina",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "fim de setembro ou início de outubro",
  "recorrente": true,
  "ciclo": "Anual desde 1991. As datas da próxima edição ainda não foram confirmadas em página oficial.",
  "resumo": "Congresso tradicional organizado por alunos da FCM-Unicamp, com módulos temáticos, workshops práticos e apresentação de trabalhos.",
  "link": "https://www.fcm.unicamp.br/comau/sites/default/files/2024-09/MANUAL%20DA%20PROGRAMA%C3%87%C3%83O.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Congresso Médico Universitário do ABC (COMUABC)",
  "instituicao": "Alunos do Centro Universitário FMABC",
  "categoria": "Evento científico",
  "pais": "Brasil",
  "cidade": "Santo André/SP",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina e da saúde",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "agosto (2026: 10–14/08)",
  "recorrente": true,
  "ciclo": "Anual, em agosto. A edição de 2027 ainda não foi divulgada.",
  "resumo": "Um dos congressos médicos universitários mais antigos do país, organizado pelos alunos da FMABC, com palestras, debates e trabalhos científicos e mais de 800 participantes.",
  "link": "https://abcdoabc.com.br/comuabc-inicia-foco-inovacao-medicina/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Congresso do Departamento Universitário da AMRIGS",
  "instituicao": "Associação Médica do Rio Grande do Sul (AMRIGS)",
  "categoria": "Evento científico",
  "pais": "Brasil",
  "cidade": "Porto Alegre/RS",
  "modalidade": "Presencial",
  "publico": "Acadêmicos de medicina",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "agosto",
  "recorrente": true,
  "ciclo": "Em 2026 (7ª edição), a submissão de trabalhos foi de abril a julho, com apresentações em agosto. A próxima edição ainda não foi divulgada.",
  "resumo": "Congresso acadêmico focado em temas livres, com apresentações orais, pôsteres digitais e premiação dos melhores trabalhos.",
  "link": "https://www.amrigs.org.br/wp-content/uploads/2026/04/Edital-VII-CDU-AMRIGS.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Cursos online UNA-SUS",
  "instituicao": "Universidade Aberta do SUS (UNA-SUS/Ministério da Saúde)",
  "categoria": "Curso online",
  "pais": "Brasil",
  "cidade": "Online",
  "modalidade": "Online",
  "publico": "Profissionais de saúde; vários cursos são de livre acesso (confira os pré-requisitos)",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "No seu ritmo",
  "recorrente": true,
  "ciclo": "Ofertas contínuas, cada uma com seu prazo de matrícula.",
  "resumo": "Catálogo de cursos a distância gratuitos e com certificado, sobre temas do SUS e da atenção primária. Bom para somar horas complementares.",
  "link": "https://www.unasus.gov.br/cursos/sobre",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Visiting Medical Student Clerkships",
  "instituicao": "Mayo Clinic Alix School of Medicine",
  "categoria": "Eletiva clínica",
  "pais": "Estados Unidos",
  "cidade": "Rochester (MN) / Phoenix-Scottsdale (AZ)",
  "modalidade": "Presencial",
  "publico": "Último ano/internato, com cerca de 48 semanas de prática clínica concluídas",
  "custo": "Pago",
  "inscricao_abre": "2026-12-01",
  "inscricao_fecha": "2026-12-31",
  "periodo_programa": "abr–jun 2027 (janela de dezembro)",
  "recorrente": true,
  "ciclo": "Quatro janelas fixas por ano: 1–31 dez (rotações abr–jun), 1–31 mar (jul–set), 1–30 jun (out–dez) e 1–30 set (jan–mar). A resposta sai cerca de 6 semanas após o prazo.",
  "resumo": "Estágio clínico na Mayo aberto a escolas médicas internacionais nos campi de Minnesota e Arizona. Exige USMLE Step 1 aprovado, TOEFL de pelo menos 104 se o curso não for em inglês, seguro saúde e taxa de US$ 350. A inscrição é só pelo VSLO, então a faculdade de origem precisa estar cadastrada nele.",
  "link": "https://college.mayo.edu/academics/visiting-medical-student-clerkships/application-process/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Elective Program – International Visiting Students",
  "instituicao": "Cleveland Clinic",
  "categoria": "Eletiva clínica",
  "pais": "Estados Unidos",
  "cidade": "Cleveland (OH)",
  "modalidade": "Presencial",
  "publico": "Último ano, com todas as rotações obrigatórias concluídas",
  "custo": "Pago",
  "inscricao_abre": "2027-01-25",
  "inscricao_fecha": null,
  "periodo_programa": "Blocos de 4 semanas ao longo do ano letivo",
  "recorrente": true,
  "ciclo": "As inscrições para 2026/27 estão abertas e fecham cerca de 4 meses antes do início de cada bloco. Para 2027/28 abrem em 25/01/2027 (blocos 1–2), 15/02/2027 (blocos 3–8) e 16/08/2027 (blocos 10–13).",
  "resumo": "Eletivas de 4 semanas para estudantes do último ano de escolas estrangeiras, com inscrição pelo portal internacional da Cleveland Clinic. Exige USMLE Step 1, OET ou TOEFL, visto B1 e checagem de antecedentes, com taxa de US$ 400 por eletiva. Quem não cumpre os requisitos pode pedir um observership, sem crédito acadêmico.",
  "link": "https://my.clevelandclinic.org/departments/elective-program/apply",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Medical Student Elective Program (candidatos internacionais)",
  "instituicao": "Memorial Sloan Kettering Cancer Center (MSK)",
  "categoria": "Eletiva clínica",
  "pais": "Estados Unidos",
  "cidade": "Nova York (NY)",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina com USMLE Step 1 feito",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": "2026-11-13",
  "periodo_programa": "jan–jun 2027",
  "recorrente": true,
  "ciclo": "Anual para candidatos internacionais. A inscrição é feita no site da MSK, não pelo VSLO, até meados de novembro, para eletivas entre janeiro e junho do ano seguinte.",
  "resumo": "Número limitado de eletivas clínicas em oncologia para estudantes internacionais. Pede carta de intenção, carta do diretor da faculdade, ao menos uma carta de recomendação, CV, histórico e relatório do USMLE Step 1. Mandar CV direto aos departamentos desclassifica o candidato.",
  "link": "https://www.mskcc.org/hcp-education-training/medical-student-opportunities/elective/apply",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Elective Placement in Clinical Medicine",
  "instituicao": "King's College London",
  "categoria": "Eletiva clínica",
  "pais": "Reino Unido",
  "cidade": "Londres",
  "modalidade": "Presencial",
  "publico": "Penúltimo ou último ano do curso",
  "custo": "Pago",
  "inscricao_abre": "2026-10-26",
  "inscricao_fecha": "2026-12-06",
  "periodo_programa": "4 semanas; inícios em 10 mai, 7 jun e 5 jul 2027",
  "recorrente": true,
  "ciclo": "Janelas fixas para cada data de início: 10/05/2027 (26/10 a 06/12/2026), 07/06/2027 (26/10/2026 a 03/01/2027) e 05/07/2027 (09/11/2026 a 24/01/2027).",
  "resumo": "Estágio clínico de até 4 semanas, numa única especialidade, nos hospitais ligados ao King's. Exige IELTS 6.5 no geral e em cada componente, duas referências e autorização da faculdade. As taxas são £ 270 de inscrição e £ 480 de estágio.",
  "link": "https://www.kcl.ac.uk/elective-placement-in-clinical-medicine-course",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Visiting Student Electives",
  "instituicao": "UCL Medical School (University College London)",
  "categoria": "Eletiva clínica",
  "pais": "Reino Unido",
  "cidade": "Londres",
  "modalidade": "Presencial",
  "publico": "Pelo menos 3 anos de curso concluídos, incluindo 12 meses de prática clínica",
  "custo": "Pago",
  "inscricao_abre": "2026-09-28",
  "inscricao_fecha": "2026-10-30",
  "periodo_programa": "13 set–17 dez 2027 (Bloco 1 de 2027/28)",
  "recorrente": true,
  "ciclo": "Três blocos por ano. Para 2027/28: Bloco 1 (28/09 a 30/10/2026), Bloco 2 para jan–abr 2028 (18/01 a 12/02/2027) e Bloco 3 para abr–jul 2028 (19/04 a 14/05/2027).",
  "resumo": "Eletiva de 4 a 8 semanas na UCL, aberta a estudantes internacionais. Exige IELTS Academic de pelo menos 6.5 e que a faculdade de origem também receba alunos da UCL em eletivas. As taxas começam em £ 100 de inscrição mais £ 925 para 4 semanas.",
  "link": "https://www.ucl.ac.uk/medical-sciences/divisions/medical-school/study/undergraduate/visiting-student-electives",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Elective Placements – School of Clinical Medicine",
  "instituicao": "University of Cambridge",
  "categoria": "Eletiva clínica",
  "pais": "Reino Unido",
  "cidade": "Cambridge",
  "modalidade": "Presencial",
  "publico": "Últimos 18 meses do curso; a eletiva precisa ser obrigatória no currículo de origem",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Blocos de até 6 semanas ao longo de 2028",
  "recorrente": true,
  "ciclo": "Anual. O programa de 2027 já fechou. As inscrições para 2028 abrem em fevereiro de 2027, sem data exata publicada. As vagas são sorteadas, com no máximo 8 alunos por bloco.",
  "resumo": "Eletiva de até 6 semanas nos hospitais de Cambridge. Exige IELTS 7.0 ou outra rota de comprovação de inglês. As taxas são £ 66 de inscrição e £ 622 de administração, mais £ 990 de acomodação.",
  "link": "https://www.medschl.cam.ac.uk/elective-placements",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Clinical electives for visiting undergraduate students",
  "instituicao": "University of Bristol Medical School",
  "categoria": "Eletiva clínica",
  "pais": "Reino Unido",
  "cidade": "Bristol",
  "modalidade": "Presencial",
  "publico": "Penúltimo ou último ano do curso",
  "custo": "Pago",
  "inscricao_abre": "2026-11-01",
  "inscricao_fecha": "2026-11-30",
  "periodo_programa": "mai–ago 2027 (4 a 8 semanas)",
  "recorrente": true,
  "ciclo": "No Bristol Royal Infirmary, os pedidos são aceitos por e-mail só entre 1 e 30 de novembro, para estágios de maio a agosto do ano seguinte. É preciso ter um consultor supervisor até 30/11.",
  "resumo": "Estágio de 4 a 8 semanas na rede de Bristol, com 30 vagas no Bristol Royal Infirmary. Exige comprovar inglês, checagem de antecedentes e avaliação de saúde ocupacional. A taxa administrativa é de £ 125.",
  "link": "https://www.bristol.ac.uk/medical-school/undergraduate-study/clinical-electives-for-visiting-students/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "International Student Clinical Elective Programme",
  "instituicao": "NUS Yong Loo Lin School of Medicine",
  "categoria": "Eletiva clínica",
  "pais": "Singapura",
  "cidade": "Singapura",
  "modalidade": "Presencial",
  "publico": "Fase clínica, com pelo menos 1 ano de estágios clínicos concluído",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "2 a 12 semanas, em datas flexíveis",
  "recorrente": true,
  "ciclo": "Fluxo contínuo. A inscrição é feita de 12 a 6 meses antes da data de início desejada.",
  "resumo": "Eletivas clínicas e de pesquisa nos hospitais de ensino de Singapura. Não exige certificado de inglês, mas é preciso conversar bem no idioma. A taxa é de S$ 152,80 por pedido mais S$ 625 por semana.",
  "link": "https://medicine.nus.edu.sg/current-students/clinical-electives-programme-for-overseas-visiting-student/eligibility",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Visiting Medical Elective Program",
  "instituicao": "The University of Queensland Medical School",
  "categoria": "Eletiva clínica",
  "pais": "Austrália",
  "cidade": "Brisbane",
  "modalidade": "Presencial",
  "publico": "Pelo menos 1 semestre de prática clínica concluído",
  "custo": "Pago",
  "inscricao_abre": "2026-11-30",
  "inscricao_fecha": "2026-12-14",
  "periodo_programa": "8 mar–30 abr 2027 (Termo 2)",
  "recorrente": true,
  "ciclo": "Quatro termos em 2027. Termo 2: inscrições de 30/11 a 14/12/2026. Termo 3 (17/05 a 09/07/2027): de 22/02 a 08/03/2027. Termo 4 (12/07 a 03/09/2027): de 12/04 a 26/04/2027.",
  "resumo": "Eletiva de 4 a 8 semanas em hospitais de Queensland. Exige IELTS 7.0 em todas as bandas ou declaração de proficiência da faculdade. As taxas são AUD 100 de inscrição e AUD 1.000 de estágio, e a UQ não patrocina visto.",
  "link": "https://medical-school.uq.edu.au/current-students/placements/inbound-clinical-placements",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Programa de Estudiantes Visitantes – Rotaciones Clínicas",
  "instituicao": "Universidad Autónoma de Madrid – Facultad de Medicina",
  "categoria": "Eletiva clínica",
  "pais": "Espanha",
  "cidade": "Madri",
  "modalidade": "Presencial",
  "publico": "5º ou 6º ano de Medicina",
  "custo": "Pago",
  "inscricao_abre": "2026-10-01",
  "inscricao_fecha": "2026-10-31",
  "periodo_programa": "1 mai–31 ago 2027 (1 a 4 meses)",
  "recorrente": true,
  "ciclo": "Anual: as inscrições vão de 1 a 31 de outubro do ano anterior às rotações. São até 10 alunos por universidade, aceitos por ordem de chegada.",
  "resumo": "Rotações clínicas nos hospitais ligados à UAM para estudantes de fora dos convênios de intercâmbio. Exige espanhol B1 certificado, seguro saúde e certidão de antecedentes. A taxa fica em torno de € 831 por mês.",
  "link": "https://www.uam.es/uam/media/doc/1606852196101/proced.-visitantes-rotaciones--26-27-def-ok-medicina.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Clinical Electives – Free Movers Scheme",
  "instituicao": "Università Cattolica del Sacro Cuore – Faculdade de Medicina",
  "categoria": "Eletiva clínica",
  "pais": "Itália",
  "cidade": "Roma",
  "modalidade": "Presencial",
  "publico": "A partir do 3º ano, já na fase de prática clínica",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Semestre de outono (set–fev) ou de primavera (mar–jul); 1 a 3 meses",
  "recorrente": true,
  "ciclo": "Duas janelas por ano, por ordem de chegada: abril (para o outono) e novembro (para a primavera). A primavera de 2027 já lotou; o outono de 2027 abre em abril de 2027.",
  "resumo": "Estágios clínicos de 1 a 3 meses para alunos de universidades sem convênio com a Cattolica. Exige inglês ou italiano B2, visto de estudo e seguro privado. A taxa é de € 500 de matrícula mais € 500 por mês.",
  "link": "https://international.unicatt.it/ucscinternational-clinical-internships",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Programa Free Mover – Estágios Clínicos",
  "instituicao": "Faculdade de Medicina da Universidade de Coimbra (FMUC)",
  "categoria": "Eletiva clínica",
  "pais": "Portugal",
  "cidade": "Coimbra",
  "modalidade": "Presencial",
  "publico": "A partir do 3º ano (pelo menos 2 anos curriculares concluídos)",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": "2026-11-15",
  "periodo_programa": "2º semestre 2026/27",
  "recorrente": true,
  "ciclo": "Pré-seleção por e-mail até 15/jun (1º semestre) ou 15/nov (2º semestre), e candidatura formal até 15/jul ou 15/dez. São 6 vagas de estágio clínico por ano. Os prazos vêm do regulamento de 2022: confirme com o gabinete internacional.",
  "resumo": "Caminho para alunos de faculdades sem convênio com a FMUC fazerem estágios clínicos em Coimbra, em português. Exige autorização da faculdade de origem. A taxa é de € 80 por crédito ECTS.",
  "link": "https://pages.uc.pt/site/assets/files/726110/freemover.pdf",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Professional Exchanges (SCOPE) – Intercâmbio clínico",
  "instituicao": "IFMSA, pela IFMSA Brazil no Brasil",
  "categoria": "Intercâmbio",
  "pais": "Vários países",
  "cidade": "Pelo comitê local da IFMSA Brazil",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina, pelo comitê local da IFMSA Brazil na sua faculdade",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "4 semanas (datas variam por país)",
  "recorrente": true,
  "ciclo": "Anual. As vagas são negociadas entre as organizações nacionais da IFMSA, e o edital é aberto pela IFMSA Brazil e seus comitês locais.",
  "resumo": "Estágio clínico de 4 semanas em hospital universitário de outro país, com médico supervisor. A taxa de participação cobre alojamento e pelo menos uma refeição por dia. Cerca de 13 mil estudantes de mais de 100 países participam por ano.",
  "link": "https://ifmsa.org/professional-exchanges/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Research Exchanges (SCORE) – Intercâmbio de pesquisa",
  "instituicao": "IFMSA, pela IFMSA Brazil no Brasil",
  "categoria": "Intercâmbio",
  "pais": "Vários países",
  "cidade": "Pelo comitê local da IFMSA Brazil",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina, pelo comitê local da IFMSA Brazil na sua faculdade",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Varia por projeto",
  "recorrente": true,
  "ciclo": "Anual, com vagas e editais abertos pela IFMSA Brazil e seus comitês locais.",
  "resumo": "Intercâmbio num projeto de pesquisa básica ou clínica, com mentor, em faculdade de medicina estrangeira. No fim, pode ser exigido relatório escrito ou apresentação oral.",
  "link": "https://ifmsa.org/research-exchanges/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "WHO Internship Programme",
  "instituicao": "Organização Mundial da Saúde (OMS)",
  "categoria": "Estágio/Observership",
  "pais": "Suíça",
  "cidade": "Genebra e escritórios regionais",
  "modalidade": "Presencial",
  "publico": "A partir de 20 anos e 3 anos de graduação concluídos",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "6 a 24 semanas, ao longo do ano",
  "recorrente": true,
  "ciclo": "Fluxo contínuo. As vagas são publicadas no portal de carreiras da OMS, cada uma com prazo próprio.",
  "resumo": "Estágio em unidades técnicas da OMS para estudantes de saúde pública, medicina e áreas ligadas. A OMS oferece seguro e auxílio de subsistência a quem precisa de apoio financeiro. Cada pessoa só pode participar uma vez.",
  "link": "https://www.who.int/careers/internship-programme",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "OIST Research Internship",
  "instituicao": "Okinawa Institute of Science and Technology (OIST)",
  "categoria": "Pesquisa",
  "pais": "Japão",
  "cidade": "Okinawa",
  "modalidade": "Presencial",
  "publico": "Últimos 2 anos da graduação, com aprovação da instituição de origem",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": null,
  "inscricao_fecha": "2027-04-15",
  "periodo_programa": "4 a 6 meses entre out 2027 e mar 2028",
  "recorrente": true,
  "ciclo": "Duas rodadas por ano: até 15/abr (estágios de out a mar) e até 15/out (estágios de abr a set).",
  "resumo": "Estágio de pesquisa de 4 a 6 meses em áreas como neurociência e biologia celular, pensado para quem planeja pós-graduação. A OIST paga ¥ 2.400 por dia útil, passagem e moradia. Medicina não aparece na lista de áreas, que é só ilustrativa.",
  "link": "https://www.oist.jp/admissions/research-internship/apply-research-internship",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Vienna BioCenter Summer School",
  "instituicao": "Vienna BioCenter (VBC)",
  "categoria": "Pesquisa",
  "pais": "Áustria",
  "cidade": "Viena",
  "modalidade": "Presencial",
  "publico": "Graduandos com pelo menos 2 anos de curso em ciências da vida",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "9 semanas no verão europeu",
  "recorrente": true,
  "ciclo": "Anual. A página diz que a chamada abre em 15 de novembro, sem informar o prazo final.",
  "resumo": "Programa de pesquisa de 9 semanas para cerca de 30 alunos do mundo todo. Oferece bolsa de € 1.400, moradia gratuita, seguro e auxílio de viagem. A página não cita estudantes de medicina, então confirme a elegibilidade.",
  "link": "https://www.training.vbc.ac.at/summer-school",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Internships for undergraduate students",
  "instituicao": "European Molecular Biology Laboratory (EMBL)",
  "categoria": "Pesquisa",
  "pais": "Alemanha",
  "cidade": "Heidelberg e outras sedes",
  "modalidade": "Presencial",
  "publico": "Graduandos em ciências naturais, incluindo medicina",
  "custo": "Varia",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Datas combinadas com o grupo anfitrião",
  "recorrente": true,
  "ciclo": "Fluxo contínuo. A candidatura vai direto ao líder do grupo de pesquisa escolhido, sem prazo fixo.",
  "resumo": "Estágio de pesquisa num grupo do EMBL, combinado diretamente com o líder do grupo, com carta de apoio de um professor da sua faculdade. O pagamento e a moradia ficam a critério do laboratório.",
  "link": "https://www.embl.org/about/info/undergraduates/how-to-apply/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Emerging Leaders in the Americas Program (ELAP)",
  "instituicao": "Governo do Canadá (EduCanada)",
  "categoria": "Bolsa",
  "pais": "Canadá",
  "cidade": "Instituição anfitriã",
  "modalidade": "Presencial",
  "publico": "Brasileiros matriculados em tempo integral numa faculdade do Brasil",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "4 meses de estudo ou pesquisa",
  "recorrente": true,
  "ciclo": "Anual. No ciclo 2026/27, as instituições canadenses enviaram as candidaturas até 31/03/2026. Os termos de 2027/28 ainda não saíram.",
  "resumo": "Bolsa de CAD 8.600 para estudantes brasileiros fazerem estudo ou pesquisa curta numa instituição canadense. Quem envia a candidatura é a instituição canadense, então é preciso ter um contato ou parceria por lá.",
  "link": "https://www.educanada.ca/scholarships-bourses/can/institutions/elap-pfla.aspx?lang=eng",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "ISCOMS Travel Grant",
  "instituicao": "ISCOMS / University Medical Center Groningen",
  "categoria": "Bolsa",
  "pais": "Países Baixos",
  "cidade": "Groningen",
  "modalidade": "Presencial",
  "publico": "Participantes do ISCOMS que precisam de apoio financeiro",
  "custo": "Com bolsa/auxílio",
  "inscricao_abre": "2027-02-15",
  "inscricao_fecha": "2027-03-01",
  "periodo_programa": "Para o ISCOMS de 7–10 jun 2027",
  "recorrente": true,
  "ciclo": "Anual. Os pedidos abrem em 15 de fevereiro e fecham em 1º de março.",
  "resumo": "Auxílio de viagem de até € 700 para participar do ISCOMS. Pede uma carta curta justificando o pedido e uma carta do orientador confirmando a necessidade financeira.",
  "link": "https://iscoms.com/travel-grant/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Immuno-oncology (Utrecht Summer School)",
  "instituicao": "Utrecht University – Faculty of Medical Sciences",
  "categoria": "Curso de verão/inverno",
  "pais": "Países Baixos",
  "cidade": "Utrecht",
  "modalidade": "Presencial",
  "publico": "Fim da graduação, com base em medicina ou ciências da vida",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "2 semanas em julho",
  "recorrente": true,
  "ciclo": "Curso do catálogo anual da Utrecht Summer School. Em 2026, o prazo foi 03/07, com desconto para quem pagou até 01/04. A edição de 2027 ainda não foi publicada.",
  "resumo": "Curso intensivo de 2 semanas sobre imunologia do câncer. Em 2026 custou € 2.699 (€ 1.999 com desconto antecipado), mais moradia. Pede carta de motivação, CV e histórico escolar.",
  "link": "https://utrechtsummerschool.nl/courses/life-sciences/immuno-oncology",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "Leiden Oxford Transplantation Summer School (LOTS)",
  "instituicao": "Leiden University Medical Center e University of Oxford",
  "categoria": "Curso de verão/inverno",
  "pais": "Países Baixos",
  "cidade": "Leiden",
  "modalidade": "Presencial",
  "publico": "Principalmente graduandos de medicina e biomedicina",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "4 dias em julho",
  "recorrente": true,
  "ciclo": "Anual. Em 2025, as inscrições foram por e-mail até 1º de maio. As datas de 2027 ainda não estão na página.",
  "resumo": "Curso curto sobre transplante de órgãos, aberto a estudantes internacionais. A taxa de € 100 cobre hospedagem e alimentação; a inscrição é feita com formulário e CV de 2 páginas.",
  "link": "https://www.universiteitleiden.nl/en/education/study-programmes/summer-schools/leiden-oxford-transplantation-summer-school-lots",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "ISCOMS 2027 – International Student Congress Of (bio)Medical Sciences",
  "instituicao": "ISCOMS / University Medical Center Groningen",
  "categoria": "Evento científico",
  "pais": "Países Baixos",
  "cidade": "Groningen",
  "modalidade": "Presencial",
  "publico": "Graduandos e recém-formados que fizeram a pesquisa como estudantes",
  "custo": "Pago",
  "inscricao_abre": "2026-10-12",
  "inscricao_fecha": "2027-01-13",
  "periodo_programa": "7–10 jun 2027",
  "recorrente": true,
  "ciclo": "Anual, em junho. O prazo acima é o de envio de resumos; o resultado sai em fevereiro. A inscrição antecipada vai de 1º de dezembro a 31 de janeiro.",
  "resumo": "Congresso estudantil de pesquisa biomédica com apresentações orais e de pôster em inglês. Não aceita relatos de caso nem revisões narrativas. Cobra € 10 por resumo e oferece auxílios de viagem de até € 700.",
  "link": "https://iscoms.com/abstract-submission/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "LIMSC 2027 – Leiden International (Bio)Medical Student Conference",
  "instituicao": "LIMSC (Leiden)",
  "categoria": "Evento científico",
  "pais": "Países Baixos",
  "cidade": "Leiden",
  "modalidade": "Presencial",
  "publico": "Graduandos, pós-graduandos e jovens pesquisadores",
  "custo": "Varia",
  "inscricao_abre": "2026-09-01",
  "inscricao_fecha": "2026-11-15",
  "periodo_programa": "17–20 mar 2027",
  "recorrente": true,
  "ciclo": "Anual, em março. O prazo acima é o de envio de resumos; o aceite sai em 05/01/2027 e quem apresenta precisa se inscrever até 01/02/2027.",
  "resumo": "Congresso estudantil com mais de 700 participantes de mais de 70 países. Aceita pesquisa original, relato de caso e revisões, em inglês. Os prêmios chegam a € 1.000 para apresentação oral.",
  "link": "https://limsc.nl/submission-guidelines.php",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "YES Meeting",
  "instituicao": "Estudantes da Faculdade de Medicina da Universidade do Porto",
  "categoria": "Evento científico",
  "pais": "Portugal",
  "cidade": "Porto",
  "modalidade": "Presencial",
  "publico": "Graduandos e recém-formados da área biomédica",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Setembro",
  "recorrente": true,
  "ciclo": "Anual, em setembro (2026 foi em 17–20/09). As datas e os prazos de 2027 ainda não foram divulgados.",
  "resumo": "Conferência internacional de ciências biomédicas organizada por estudantes da FMUP, aberta a estudantes do mundo todo que queiram apresentar pesquisa.",
  "link": "https://yesmeeting.org",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "AIMS Meeting 2027",
  "instituicao": "Associação de Estudantes da Faculdade de Medicina de Lisboa",
  "categoria": "Evento científico",
  "pais": "Portugal",
  "cidade": "Lisboa",
  "modalidade": "Presencial",
  "publico": "Estudantes de medicina e de ciências biomédicas",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "Primavera de 2027 (data não divulgada)",
  "recorrente": true,
  "ciclo": "Anual, na primavera europeia (2026 foi em abril). O ingresso de 2027 para estudantes externos custa € 40, mas as inscrições ainda não abriram.",
  "resumo": "Encontro internacional em Lisboa com cerca de mil participantes por ano, palestras e competição científica. A competição clínica e as masterclasses são para falantes de português.",
  "link": "https://www.aimsmeeting.org/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "AMEE 2027 – The AMEE Conference",
  "instituicao": "AMEE – International Association for Health Professions Education",
  "categoria": "Evento científico",
  "pais": "Finlândia",
  "cidade": "Helsinque",
  "modalidade": "Presencial",
  "publico": "Educadores e estudantes das profissões da saúde",
  "custo": "Pago",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "28 ago–1 set 2027",
  "recorrente": true,
  "ciclo": "Anual, no fim de agosto. O envio de resumos para 2027 já está aberto, mas a página não informa o prazo.",
  "resumo": "Congresso internacional de educação em saúde que também recebe estudantes. Bom lugar para apresentar trabalhos sobre ensino médico, ligas acadêmicas e currículo.",
  "link": "https://amee.org/events/amee-2027/",
  "conferido_em": "2026-10-18"
 },
 {
  "publicar": "sim",
  "titulo": "OpenWHO – cursos online da OMS",
  "instituicao": "Organização Mundial da Saúde (OMS)",
  "categoria": "Curso online",
  "pais": "Online",
  "cidade": "Online",
  "modalidade": "Online",
  "publico": "Qualquer pessoa, incluindo estudantes",
  "custo": "Gratuito",
  "inscricao_abre": null,
  "inscricao_fecha": null,
  "periodo_programa": "No seu ritmo, o ano todo",
  "recorrente": true,
  "ciclo": "Fluxo contínuo, com novos cursos ao longo do ano.",
  "resumo": "Plataforma gratuita da OMS com cursos sobre saúde pública e emergências sanitárias, em vários idiomas, com parte do catálogo em português. Alguns cursos dão certificado.",
  "link": "https://openwho.org/",
  "conferido_em": "2026-10-18"
 }
];
