/**
 * Handoff — radar de oportunidades médicas em Santa Catarina.
 * Busca publicações recentes de perfis profissionais do Instagram autorizados
 * pela API da Meta e fontes RSS públicas. Tudo entra em uma fila de revisão;
 * nada é publicado no site automaticamente.
 *
 * Grupos do Facebook são cadastrados como fontes manuais: cole o link e o
 * texto do post na aba de fontes e use o menu para importar. A leitura direta
 * dos grupos não é feita por este script.
 */
var VMSC = {
  FONTES: 'Vagas Médicas SC — Fontes',
  REVISAO: 'Vagas Médicas SC — Revisão',
  HORA: 4,
  GRAPH_VERSION: 'v26.0',
  HEAD_FONTES: ['ativo', 'plataforma', 'nome da fonte', 'perfil ou URL', 'termos', 'última busca', 'novos itens', 'situação', 'link manual Facebook', 'texto manual Facebook'],
  HEAD_REVISAO: ['status', 'plataforma', 'fonte', 'perfil', 'publicado em', 'texto para revisar', 'link original', 'capturado em', 'resultado da revisão', 'observações'],
  TERMOS_VAGA: /(vaga|vagas|plant[aã]o|contrata|contratando|oportunidade|m[eé]dic[oa]|cl[ií]nic[oa]|pronto atendimento|urg[eê]ncia|escala|corpo cl[ií]nico|credenciamento|processo seletivo)/i,
  TERMOS_SC: /(santa catarina|\bsc\b|florian[oó]polis|s[aã]o jos[eé]|palho[cç]a|bigua[cç]u|itajai|balne[aá]rio cambori[uú]|navegantes|brusque|blumenau|joinville|jaragu[aá] do sul|chapec[oó]|conc[oó]rdia|crici[uú]ma|tubar[aã]o|lages|rio do sul|joa[cç]aba|maf[r]?a|ca[cç]ador|s[aã]o bento do sul|ararangu[aá]|itapema|ind[aá]ial|garopaba)/i,
  MAX_POR_FONTE: 25,
  MAX_TEXTO: 2200
};

/** Instale uma vez no Apps Script da planilha Handoff — Base. */
function vagasMedicasSC_instalar() {
  var ss = SpreadsheetApp.getActive();
  var fontes = vmsc_aba_(ss, VMSC.FONTES, VMSC.HEAD_FONTES, '#2B45D6', '#FFFFFF');
  var revisao = vmsc_aba_(ss, VMSC.REVISAO, VMSC.HEAD_REVISAO, '#14206B', '#FFFFFF');
  fontes.setFrozenRows(1); revisao.setFrozenRows(1);
  [180, 150, 220, 360, 240, 150, 110, 260, 360, 500].forEach(function (w, i) { fontes.setColumnWidth(i + 1, w); });
  [120, 150, 220, 180, 150, 520, 360, 150, 170, 340].forEach(function (w, i) { revisao.setColumnWidth(i + 1, w); });
  fontes.getRange('A1').setNote('Use instagram para perfis profissionais públicos, rss para feeds públicos e facebook_grupo para captura manual. O script só lê fontes configuradas nesta aba.');
  fontes.getRange('I1').setNote('Para grupos do Facebook, cole aqui o link do post público que você quer revisar. A API de Grupos foi encerrada pela Meta; não há leitura automática de posts de grupos.');
  vmsc_validar_(fontes);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'vagasMedicasSC_buscar' || t.getHandlerFunction() === 'vagasMedicasSC_menu') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('vagasMedicasSC_buscar').timeBased().everyHours(VMSC.HORA).create();
  ScriptApp.newTrigger('vagasMedicasSC_menu').forSpreadsheet(ss).onOpen().create();
  vagasMedicasSC_menu();
  ss.toast('Radar de vagas médicas em SC instalado. Adicione perfis e feeds na aba de fontes.');
}

function vagasMedicasSC_menu() {
  SpreadsheetApp.getUi().createMenu('Vagas médicas SC')
    .addItem('Buscar novas oportunidades agora', 'vagasMedicasSC_buscar')
    .addItem('Importar link manual selecionado', 'vagasMedicasSC_importarSelecionada')
    .addToUi();
}

/** Pesquisa somente perfis/feeds ativos configurados em Vagas Médicas SC — Fontes. */
function vagasMedicasSC_buscar() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(VMSC.FONTES), fila = ss.getSheetByName(VMSC.REVISAO);
  if (!sh || !fila) throw new Error('Rode vagasMedicasSC_instalar() primeiro.');
  var ultima = sh.getLastRow(); if (ultima < 2) return;
  var rows = sh.getRange(2, 1, ultima - 1, VMSC.HEAD_FONTES.length).getValues();
  var vistos = vmsc_linksExistentes_(fila), inicio = Date.now();
  rows.forEach(function (r, i) {
    if (String(r[0]).toLowerCase().indexOf('n') === 0) return;
    var row = i + 2, plataforma = vmsc_norm_(r[1]), nome = String(r[2] || r[3] || 'Fonte'), ref = String(r[3] || '').trim(), novos = 0, erro = '';
    if (plataforma === 'facebook_grupo' || plataforma === 'facebook' || plataforma === 'grupo_facebook') {
      sh.getRange(row, 8).setValue('Captura manual: cole link e texto do post nas colunas I–J.');
      sh.getRange(row, 6).setValue(new Date()); return;
    }
    try {
      var itens = [];
      if (plataforma === 'instagram') itens = vmsc_instagram_(ref);
      else if (plataforma === 'rss') itens = vmsc_rss_(ref);
      else { sh.getRange(row, 8).setValue('Plataforma não suportada. Use instagram, rss ou facebook_grupo.'); return; }
      itens.slice(0, VMSC.MAX_POR_FONTE).forEach(function (p) {
        if (Date.now() - inicio > 270000) return;
        var trecho = String(p.texto || '').replace(/\s+/g, ' ').trim().slice(0, VMSC.MAX_TEXTO);
        if (!VMSC.TERMOS_VAGA.test(trecho) || !VMSC.TERMOS_SC.test(trecho)) return;
        var url = String(p.link || '').trim(); if (!url || vistos[url]) return;
        vmsc_gravar_(fila, vistos, { plataforma: plataforma, fonte: nome, perfil: p.perfil || ref, data: p.data, texto: trecho, link: url });
        novos++;
      });
      sh.getRange(row, 6).setValue(new Date()); sh.getRange(row, 7).setValue(novos); sh.getRange(row, 8).setValue('OK');
    } catch (e) {
      erro = String(e.message || e).slice(0, 240); sh.getRange(row, 6).setValue(new Date()); sh.getRange(row, 8).setValue(erro);
    }
  });
}

/** Importa o link e texto manual de um perfil/grupo selecionado na aba de fontes. */
function vagasMedicasSC_importarSelecionada() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet(), fila = ss.getSheetByName(VMSC.REVISAO);
  if (!sh || sh.getName() !== VMSC.FONTES || !fila) throw new Error('Selecione uma linha na aba "' + VMSC.FONTES + '".');
  var row = sh.getActiveRange().getRow(); if (row < 2) throw new Error('Selecione uma linha de fonte, abaixo do cabeçalho.');
  var r = sh.getRange(row, 1, 1, VMSC.HEAD_FONTES.length).getValues()[0];
  var plataforma = vmsc_norm_(r[1]), link = String(r[8] || '').trim(), texto = String(r[9] || '').trim();
  if (!/facebook|grupo/.test(plataforma)) throw new Error('Esta ação é para grupos do Facebook ou outras fontes de captura manual.');
  if (!/^https?:\/\//i.test(link) || !texto) throw new Error('Preencha o link e o texto do post nas colunas I e J.');
  var vistos = vmsc_linksExistentes_(fila);
  if (!vistos[link]) vmsc_gravar_(fila, vistos, { plataforma: 'facebook_grupo', fonte: String(r[2] || 'Grupo do Facebook'), perfil: String(r[3] || ''), texto: texto.slice(0, VMSC.MAX_TEXTO), link: link, data: '' });
  sh.getRange(row, 8).setValue('Importado para revisão em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm'));
  sh.getRange(row, 9, 1, 2).clearContent();
  ss.toast('Link adicionado à fila de revisão.');
}

function vmsc_instagram_(username) {
  username = String(username || '').replace(/^@/, '').trim();
  if (!/^[A-Za-z0-9._]{1,30}$/.test(username)) throw new Error('Informe o @perfil do Instagram na coluna D.');
  var props = PropertiesService.getScriptProperties(), token = props.getProperty('META_TOKEN');
  var igId = props.getProperty('VAGAS_MEDICAS_SC_IG_USER_ID');
  if (!token || !igId) throw new Error('Configure META_TOKEN e VAGAS_MEDICAS_SC_IG_USER_ID nas propriedades do Apps Script.');
  var version = props.getProperty('META_GRAPH_API_VERSION') || VMSC.GRAPH_VERSION;
  var fields = 'business_discovery.username(' + username + '){username,media.limit(25){caption,permalink,timestamp}}';
  var url = 'https://graph.facebook.com/' + encodeURIComponent(version) + '/' + encodeURIComponent(igId) + '?fields=' + encodeURIComponent(fields) + '&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var body = JSON.parse(resp.getContentText() || '{}');
  if (resp.getResponseCode() >= 300 || body.error) throw new Error('Instagram API: ' + String((body.error && body.error.message) || resp.getResponseCode()).slice(0, 180));
  var bd = body.business_discovery || {}, media = bd.media && bd.media.data || [];
  return media.map(function (m) { return { perfil: '@' + (bd.username || username), texto: m.caption || '', link: m.permalink || '', data: m.timestamp || '' }; });
}

function vmsc_rss_(url) {
  if (!/^https?:\/\//i.test(url)) throw new Error('Informe o endereço RSS na coluna D.');
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() >= 300) throw new Error('Feed RSS respondeu HTTP ' + resp.getResponseCode());
  var xml = XmlService.parse(resp.getContentText()), root = xml.getRootElement(), channel = root.getName() === 'rss' ? root.getChild('channel') : root;
  var items = channel ? channel.getChildren('item') : [];
  return items.map(function (it) {
    var title = vmsc_child_(it, 'title'), desc = vmsc_child_(it, 'description'), link = vmsc_child_(it, 'link'), date = vmsc_child_(it, 'pubDate');
    return { texto: title + ' ' + desc, link: link, data: date };
  });
}

function vmsc_child_(el, name) { var c = el && el.getChild(name); return c ? c.getText() : ''; }
function vmsc_aba_(ss, name, header, bg, fg) {
  var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name, ss.getSheets().length);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground(bg).setFontColor(fg);
  return sh;
}
function vmsc_validar_(sh) {
  var n = 500;
  sh.getRange(2, 1, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['sim', 'não'], true).build());
  sh.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['instagram', 'rss', 'facebook_grupo'], true).build());
  sh.getRange(2, 6, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
}
function vmsc_linksExistentes_(sh) {
  var out = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 7, last - 1, 1).getValues().forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = 1; });
  return out;
}
function vmsc_gravar_(sh, vistos, p) {
  var link = String(p.link || '').trim(); if (!link || vistos[link]) return false;
  sh.appendRow(['REVISAR', p.plataforma || '', p.fonte || '', p.perfil || '', p.data || '', p.texto || '', link, new Date(), '', 'Confirme cidade, vínculo, CRM/requisitos, remuneração e validade antes de divulgar.']);
  vistos[link] = 1; return true;
}
function vmsc_norm_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim().replace(/[ -]+/g, '_'); }
