/**
 * Handoff Vagas — busca automática de vagas novas, sem aprovação manual.
 *
 * Vai como mais um arquivo (VagasBusca.gs) no Apps Script da Handoff — Base, junto do vagas.gs.
 * Usa a mesma chave do Gemini que o Publicações.gs já usa (propriedade GEMINI_API_KEY) e o modelo PUB.GEMINI_MODELO.
 *
 * Instalação: depois de colar e salvar, rode vagasBuscaInstalar() uma vez. Ela cria as abas
 * "Fontes de vagas" (de onde a busca lê) e "Vagas vistas" (o que já foi olhado) e o gatilho a cada 4 horas.
 *
 * De onde vêm as vagas (aba "Fontes de vagas"):
 *   - tipo "rss":    alertas do Google entregues como feed RSS (cole o link do feed na coluna url);
 *   - tipo "pagina": páginas que listam vários programas (cursos de férias, summer schools, estágios...).
 *     Quando a página muda, a IA lista os programas e cada um vira uma candidata.
 *
 * Cada candidata é aberta e lida pelo Gemini. Ela só entra no site (publicar = sim) se passar em TODAS as checagens:
 *   é uma oportunidade com inscrição, aceita estudante de medicina de outras faculdades, a categoria é uma das do site,
 *   o prazo de inscrição é futuro e está ESCRITO na própria página, o título e a instituição aparecem no texto,
 *   a página é oficial (notícia é trocada pelo link oficial que ela citar) e a vaga ainda não está na planilha.
 * O que não passa vai para "Vagas vistas" com o motivo e nunca é olhado de novo.
 */

var VBUSCA = {
  ABA_FONTES: 'Fontes de vagas',
  ABA_VISTAS: 'Vagas vistas',
  CABECALHO_FONTES: ['ativo', 'tipo', 'nome', 'url', 'ultima_leitura', 'assinatura', 'vagas_novas', 'erro'],
  CABECALHO_VISTAS: ['url', 'visto_em', 'resultado', 'detalhe', 'fonte'],
  MAX_CANDIDATAS: 15,          // páginas lidas pela IA por rodada (a rodada é a cada 4 horas)
  MAX_ITENS_POR_PAGINA: 15,
  TEMPO_MAX_MS: 270000,        // para antes dos 6 minutos do Apps Script; o resto fica para a próxima rodada
  MAX_TEXTO: 14000,
  DIAS_MAX_PRAZO: 400,
  EMAIL: true,                 // e-mail curto (só informativo) quando entram vagas novas
  BLOQUEADOS: /(^|\.)(instagram|facebook|fb|linkedin|twitter|x|tiktok|youtube|youtu|whatsapp|wa|t|telegram|pinterest|reddit)\.(com|me|be|org)$/i,
};

// ---------------------------------------------------------------- instalação

function vagasBuscaInstalar() {
  var ss = SpreadsheetApp.getActive();
  var f = ss.getSheetByName(VBUSCA.ABA_FONTES);
  if (!f) {
    f = ss.insertSheet(VBUSCA.ABA_FONTES, ss.getSheets().length);
    f.getRange(1, 1, 1, VBUSCA.CABECALHO_FONTES.length).setValues([VBUSCA.CABECALHO_FONTES])
      .setFontWeight('bold').setBackground('#2B45D6').setFontColor('#FFFFFF');
    var linhas = VBUSCA_FONTES_INICIAIS.map(function (x) { return ['sim', x.tipo, x.nome, x.url, '', '', 0, '']; });
    if (linhas.length) f.getRange(2, 1, linhas.length, VBUSCA.CABECALHO_FONTES.length).setValues(linhas);
    var n = Math.max(linhas.length + 100, 150);
    f.getRange(2, 1, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['sim', 'não'], true).build());
    f.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['rss', 'pagina'], true).build());
    f.getRange(2, 5, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    f.setFrozenRows(1);
    f.setColumnWidth(3, 260);
    f.setColumnWidth(4, 420);
  }
  var v = ss.getSheetByName(VBUSCA.ABA_VISTAS);
  if (!v) {
    v = ss.insertSheet(VBUSCA.ABA_VISTAS, ss.getSheets().length);
    v.getRange(1, 1, 1, VBUSCA.CABECALHO_VISTAS.length).setValues([VBUSCA.CABECALHO_VISTAS])
      .setFontWeight('bold').setBackground('#D9DEEA').setFontColor('#14206B');
    v.getRange(2, 2, 2000, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    v.setFrozenRows(1);
  }
  vagasCriarGatilhos(); // do vagas.gs: conferência às 7h + busca a cada 4 horas + menu
  ss.toast('Busca automática instalada. A primeira rodada começa sozinha em até 4 horas, ou pelo menu Handoff Vagas.');
}

// ---------------------------------------------------------------- rodada de busca

function vagasBuscar() {
  var inicio = Date.now();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var ss = SpreadsheetApp.getActive();
    var abaVagas = ss.getSheetByName(VAGAS.ABA);
    var abaFontes = ss.getSheetByName(VBUSCA.ABA_FONTES);
    var abaVistas = ss.getSheetByName(VBUSCA.ABA_VISTAS);
    if (!abaVagas || !abaFontes || !abaVistas) throw new Error('Rode vagasBuscaInstalar() primeiro.');
    var ixV = vagas_garantirColunas(abaVagas);

    var existentes = vbusca_existentes(abaVagas, ixV);
    var vistas = {};
    if (abaVistas.getLastRow() > 1) {
      abaVistas.getRange(2, 1, abaVistas.getLastRow() - 1, 1).getValues().forEach(function (r) { vistas[vbusca_normUrl(r[0])] = 1; });
    }
    var candidatas = [], naFila = {};
    var empurrar = function (url, fonte) {
      var k = vbusca_normUrl(url);
      if (!k || vistas[k] || existentes.links[k] || naFila[k] || vbusca_bloqueado(url)) return;
      naFila[k] = 1;
      candidatas.push({ url: url, fonte: fonte });
    };

    // 1. Lê as fontes
    var fontes = abaFontes.getLastRow() > 1 ? abaFontes.getRange(2, 1, abaFontes.getLastRow() - 1, VBUSCA.CABECALHO_FONTES.length).getValues() : [];
    fontes.forEach(function (f, i) {
      if (String(f[0]).toLowerCase().indexOf('n') === 0 || !/^https?:\/\//i.test(String(f[3]))) return;
      if (Date.now() - inicio > VBUSCA.TEMPO_MAX_MS / 2) return;
      var linha = i + 2, tipo = String(f[1]).trim().toLowerCase(), nome = String(f[2] || f[3]), url = String(f[3]).trim();
      try {
        if (tipo === 'rss') {
          vbusca_lerFeed(url).forEach(function (u) { empurrar(u, nome); });
          abaFontes.getRange(linha, 5, 1, 1).setValue(new Date());
          abaFontes.getRange(linha, 8).setValue('');
        } else {
          var pg = vagas_abrir(url);
          abaFontes.getRange(linha, 5).setValue(new Date());
          if (!pg.ok || !pg.html) { abaFontes.getRange(linha, 8).setValue('não abriu (' + pg.codigo + ')'); return; }
          if (String(f[5]) === pg.assinatura) { abaFontes.getRange(linha, 8).setValue(''); return; } // nada mudou
          var itens = vbusca_listarPagina(url, pg.texto, pg.links);
          itens.forEach(function (u) { empurrar(u, nome); });
          abaFontes.getRange(linha, 6).setValue(pg.assinatura);
          abaFontes.getRange(linha, 8).setValue('');
        }
      } catch (e) {
        abaFontes.getRange(linha, 8).setValue(String(e.message || e).slice(0, 200));
      }
    });

    // 2. Avalia as candidatas
    var novas = [], vistasNovas = [];
    for (var c = 0; c < candidatas.length && c < VBUSCA.MAX_CANDIDATAS; c++) {
      if (Date.now() - inicio > VBUSCA.TEMPO_MAX_MS) break;
      var cand = candidatas[c];
      var res;
      try { res = vbusca_avaliar(cand.url, existentes, 0); } catch (e) { res = { ok: false, motivo: 'erro: ' + String(e.message || e).slice(0, 150) }; }
      if (res.erroTemporario) continue; // tenta de novo na próxima rodada
      vistasNovas.push([cand.url, new Date(), res.ok ? 'entrou no site' : 'descartada', res.ok ? res.vaga.titulo : res.motivo, cand.fonte]);
      if (res.via && res.via !== cand.url) vistasNovas.push([res.via, new Date(), res.ok ? 'entrou no site' : 'descartada', '(link oficial citado por ' + cand.url + ')', cand.fonte]);
      if (res.ok) {
        res.vaga.origem = 'busca automática · ' + cand.fonte;
        novas.push(res.vaga);
        existentes.links[vbusca_normUrl(res.vaga.link)] = 1;
        existentes.titulos.push(vbusca_tokens(res.vaga.titulo));
      }
    }

    // 3. Grava
    if (vistasNovas.length) abaVistas.getRange(abaVistas.getLastRow() + 1, 1, vistasNovas.length, VBUSCA.CABECALHO_VISTAS.length).setValues(vistasNovas);
    if (novas.length) {
      var largura = abaVagas.getLastColumn();
      var head = abaVagas.getRange(1, 1, 1, largura).getValues()[0].map(function (h) { return String(h).trim(); });
      var linhas = novas.map(function (o) { return head.map(function (h) { return o[h] === undefined || o[h] === null ? '' : o[h]; }); });
      abaVagas.getRange(abaVagas.getLastRow() + 1, 1, linhas.length, largura).setValues(linhas);
      vbusca_contarFontes(abaFontes, novas);
      if (VBUSCA.EMAIL) vbusca_email(ss, abaVagas, novas);
    }
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- uma candidata

/** Abre a página, pede ao Gemini os dados e confere tudo. Devolve {ok, vaga} ou {ok:false, motivo}. */
function vbusca_avaliar(url, existentes, nivel) {
  var pg = vagas_abrir(url);
  if (!pg.ok) {
    if (pg.codigo === -1 || pg.codigo >= 500 || pg.codigo === 429) return { ok: false, erroTemporario: true };
    return { ok: false, motivo: 'página não abriu (' + pg.codigo + ')' };
  }
  if (!pg.html) return { ok: false, motivo: 'PDF: a busca só lê páginas' };
  var g = vbusca_gemini(vbusca_promptVaga(url, pg.texto, pg.links));
  if (!g || g.eh_vaga !== true) return { ok: false, motivo: 'não é vaga: ' + String((g && g.motivo) || '').slice(0, 120) };

  // Notícia ou agregador: troca pelo link oficial que a página cita (uma vez só)
  if (g.pagina_oficial === false) {
    var oficial = String(g.link_oficial || '').trim();
    var citado = pg.links.some(function (l) { return l.url === oficial; });
    if (nivel === 0 && oficial && citado && !vbusca_bloqueado(oficial)) {
      var k = vbusca_normUrl(oficial);
      if (existentes.links[k]) return { ok: false, motivo: 'já está na planilha (link oficial)' };
      var r = vbusca_avaliar(oficial, existentes, 1);
      r.via = oficial;
      return r;
    }
    return { ok: false, motivo: 'notícia ou agregador sem link oficial' };
  }

  var v = vbusca_validar(g, pg.texto, url, existentes, true);
  if (!v.ok) return v;
  return { ok: true, vaga: v.vaga };
}

/** Releitura de uma vaga que já está no site (chamada pelo vagasConferir quando a página muda ou o prazo passa). */
function vbusca_reler(url, texto, links, atual) {
  var g = vbusca_gemini(vbusca_promptVaga(url, texto, links || []));
  if (!g || g.eh_vaga !== true) return { ok: false, motivo: 'a IA não achou a vaga na página' };
  var v = vbusca_validar(g, texto, url, null, false);
  if (!v.ok) return v;
  // Precisa ser a mesma vaga: boa parte das palavras do título atual aparece no título lido
  if (vbusca_semelhanca(vbusca_tokens(atual.titulo), vbusca_tokens(v.vaga.titulo)) < 0.4 &&
      !vbusca_apoiado(atual.titulo, texto, 0.6)) return { ok: false, motivo: 'a página parece ser de outra vaga' };
  return { ok: true, inscricao_abre: v.vaga.inscricao_abre, inscricao_fecha: v.vaga.inscricao_fecha, periodo_programa: v.vaga.periodo_programa };
}

/** Todas as checagens. Devolve {ok, vaga} pronta para a planilha ou {ok:false, motivo}. */
function vbusca_validar(g, texto, url, existentes, nova) {
  var limpa = function (s, max) { return vbusca_semContato(String(s == null ? '' : s).replace(/\s+/g, ' ').trim()).slice(0, max || 300); };
  var hoje = vagas_inicioDoDia(new Date());
  if (g.aceita_estudante_medicina !== true) return { ok: false, motivo: 'não é para estudante de medicina' };
  if (g.aceita_outras_faculdades !== true) return { ok: false, motivo: 'só para alunos da própria instituição' };
  var categoria = VAGAS.OPCOES.categoria.filter(function (c) { return vbusca_norm(c) === vbusca_norm(g.categoria); })[0];
  if (!categoria) return { ok: false, motivo: 'categoria fora do site: ' + limpa(g.categoria, 40) };
  var titulo = limpa(g.titulo, 120), instituicao = limpa(g.instituicao, 120);
  if (titulo.length < 6) return { ok: false, motivo: 'sem título' };
  if (!vbusca_apoiado(titulo, texto, 0.6)) return { ok: false, motivo: 'título não aparece na página' };
  if (!instituicao || !vbusca_apoiado(instituicao, texto + ' ' + url, 0.5)) return { ok: false, motivo: 'instituição não aparece na página' };

  var fecha = vagas_paraData(g.inscricao_fecha);
  var abre = vagas_paraData(g.inscricao_abre);
  if (!fecha) return { ok: false, motivo: 'sem prazo de inscrição' };
  if (fecha < hoje) return { ok: false, motivo: 'inscrições já fecharam (' + vagas_fmt(fecha) + ')' };
  if ((fecha - hoje) / 86400000 > VBUSCA.DIAS_MAX_PRAZO) return { ok: false, motivo: 'prazo longe demais (' + vagas_fmt(fecha) + ')' };
  if (!vbusca_dataNoTexto(fecha, texto)) return { ok: false, motivo: 'o prazo ' + vagas_fmt(fecha) + ' não está escrito na página' };
  if (g.trecho_prazo && !vbusca_apoiado(g.trecho_prazo, texto, 0.7)) return { ok: false, motivo: 'o trecho do prazo não está na página' };
  if (abre && (abre > fecha || !vbusca_dataNoTexto(abre, texto))) abre = '';

  if (nova && existentes) {
    if (existentes.links[vbusca_normUrl(url)]) return { ok: false, motivo: 'já está na planilha' };
    var tk = vbusca_tokens(titulo);
    if (existentes.titulos.some(function (t) { return vbusca_semelhanca(t, tk) >= 0.8; })) return { ok: false, motivo: 'vaga parecida já está na planilha' };
  }

  var escolhe = function (valor, lista, padrao) {
    return lista.filter(function (x) { return vbusca_norm(x) === vbusca_norm(valor); })[0] || padrao;
  };
  return { ok: true, vaga: {
    publicar: 'sim',
    titulo: titulo,
    instituicao: instituicao,
    categoria: categoria,
    pais: limpa(g.pais, 40) || 'Brasil',
    cidade: limpa(g.cidade, 60),
    modalidade: escolhe(g.modalidade, VAGAS.OPCOES.modalidade, 'Presencial'),
    publico: limpa(g.publico, 140),
    custo: escolhe(g.custo, VAGAS.OPCOES.custo, 'Varia'),
    inscricao_abre: abre || '',
    inscricao_fecha: fecha,
    periodo_programa: limpa(g.periodo_programa, 80),
    recorrente: g.recorrente === true ? 'sim' : 'não',
    ciclo: limpa(g.ciclo, 260),
    resumo: limpa(g.resumo, 320),
    link: url,
    conferido_em: new Date(),
  } };
}

// ---------------------------------------------------------------- Gemini

function vbusca_promptVaga(url, texto, links) {
  return [
    'Você lê páginas da web e decide se elas anunciam UMA oportunidade com inscrição para estudantes de graduação em medicina:',
    'curso de verão ou inverno, estágio ou observership, eletiva clínica, intercâmbio, programa de pesquisa, bolsa, evento científico ou congresso, ou curso online gratuito com certificado.',
    '',
    'Endereço da página: ' + url,
    'Texto da página (pode estar cortado):',
    '"""',
    String(texto || '').slice(0, VBUSCA.MAX_TEXTO),
    '"""',
    'Links da página (texto → endereço):',
    (links || []).slice(0, 120).map(function (l) { return '- ' + l.texto + ' → ' + l.url; }).join('\n'),
    '',
    'Responda só com JSON: {"eh_vaga": false, "motivo": "", "pagina_oficial": true, "link_oficial": "", "titulo": "", "instituicao": "",',
    ' "categoria": "", "pais": "", "cidade": "", "modalidade": "", "publico": "", "aceita_estudante_medicina": false,',
    ' "aceita_outras_faculdades": false, "custo": "", "inscricao_abre": "", "inscricao_fecha": "", "trecho_prazo": "",',
    ' "periodo_programa": "", "recorrente": false, "ciclo": "", "resumo": ""}',
    'Regras:',
    '- Nunca invente. Campo sem informação no texto fica "" (ou false).',
    '- eh_vaga = true só se a página é sobre UMA oportunidade específica com inscrição. É false para notícia sem inscrição, lista de vários programas, página institucional, evento que já aconteceu ou vaga de emprego para médico formado. Em "motivo", diga em poucas palavras por quê.',
    '- pagina_oficial = false se for notícia, blog ou agregador. Nesse caso, em link_oficial, copie EXATAMENTE da lista de links o endereço da página oficial da oportunidade (ou deixe "").',
    '- categoria: exatamente uma destas: Eletiva clínica, Curso de verão/inverno, Estágio/Observership, Pesquisa, Intercâmbio, Evento científico, Bolsa, Curso online.',
    '- aceita_estudante_medicina: true se estudantes de graduação em medicina podem se inscrever.',
    '- aceita_outras_faculdades: true se alunos de outras faculdades ou universidades podem se inscrever (não é só para alunos da casa).',
    '- inscricao_abre e inscricao_fecha: datas das INSCRIÇÕES (não do evento), no formato AAAA-MM-DD. Para congresso, o fim das inscrições ou do envio de trabalhos.',
    '- trecho_prazo: copie exatamente do texto o trecho, de até 200 caracteres, onde está o prazo de inscrição.',
    '- pais em português (Brasil, Estados Unidos, Reino Unido, Portugal...). cidade: "São Paulo/SP" no Brasil; fora, só o nome da cidade.',
    '- modalidade: Presencial, Online ou Híbrido. custo: Gratuito, Pago, Com bolsa/auxílio ou Varia.',
    '- titulo: o nome oficial da oportunidade, até 90 caracteres. instituicao: quem oferece, como escrito na página.',
    '- publico: para quem é, até 100 caracteres. periodo_programa: quando acontece, curto (ex.: "jan–fev 2027").',
    '- recorrente: true se a página indica que acontece todo ano ou em ciclos. ciclo: 1 frase sobre isso.',
    '- resumo: 1 a 2 frases em português, com suas palavras, sem e-mail, telefone, Pix ou nome de pessoa.',
  ].join('\n');
}

function vbusca_listarPagina(url, texto, links) {
  if (!links || !links.length) return [];
  var g = vbusca_gemini([
    'Esta página lista oportunidades (cursos, programas, estágios, eletivas, bolsas, congressos).',
    'Endereço: ' + url,
    'Texto (pode estar cortado):',
    '"""', String(texto || '').slice(0, VBUSCA.MAX_TEXTO), '"""',
    'Links da página (texto → endereço):',
    links.slice(0, 200).map(function (l) { return '- ' + l.texto + ' → ' + l.url; }).join('\n'),
    '',
    'Liste até ' + VBUSCA.MAX_ITENS_POR_PAGINA + ' oportunidades que estudantes de graduação em medicina possam fazer e que ainda',
    'não tenham acontecido. Para cada uma, copie EXATAMENTE da lista de links o endereço da página dela (nunca invente endereço,',
    'nunca use o endereço desta página).',
    'Responda só com JSON: {"itens": [{"titulo": "", "link": ""}]}',
  ].join('\n'));
  var validos = {};
  links.forEach(function (l) { validos[l.url] = 1; });
  return ((g && g.itens) || []).map(function (i) { return String(i.link || '').trim(); })
    .filter(function (u) { return validos[u] && vbusca_normUrl(u) !== vbusca_normUrl(url); })
    .slice(0, VBUSCA.MAX_ITENS_POR_PAGINA);
}

function vbusca_gemini(prompt) {
  var chave = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!chave) throw new Error('Sem GEMINI_API_KEY.');
  var modelo = (typeof PUB !== 'undefined' && PUB.GEMINI_MODELO) || 'gemini-3.5-flash-lite';
  var resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': chave },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0.1 },
    }),
  });
  if (resp.getResponseCode() !== 200) throw new Error('Gemini respondeu ' + resp.getResponseCode());
  var corpo = JSON.parse(resp.getContentText());
  var partes = (((corpo.candidates || [])[0] || {}).content || {}).parts || [];
  var txt = partes.map(function (p) { return p.text || ''; }).join('');
  var m = txt.match(/\{[\s\S]*\}/);
  return m ? JSON.parse(m[0]) : null;
}

// ---------------------------------------------------------------- feeds (alertas do Google)

/** Endereços dos itens de um feed RSS ou Atom. Links de redirecionamento do Google viram o endereço de verdade. */
function vbusca_lerFeed(url) {
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('feed respondeu ' + resp.getResponseCode());
  var raiz = XmlService.parse(resp.getContentText()).getRootElement();
  var links = [];
  var visitar = function (el) {
    var nome = el.getName();
    if (nome === 'entry' || nome === 'item') {
      var l = '';
      el.getChildren().forEach(function (f) {
        if (f.getName() !== 'link' || l) return;
        var href = f.getAttribute('href');
        l = href ? href.getValue() : f.getText();
      });
      if (l) links.push(vbusca_desredirecionar(l.trim()));
      return;
    }
    el.getChildren().forEach(visitar);
  };
  visitar(raiz);
  return links.filter(function (u) { return /^https?:\/\//i.test(u); });
}

function vbusca_desredirecionar(u) {
  var m = String(u).match(/^https?:\/\/(?:www\.)?google\.[^\/]+\/url\?(.*)$/i);
  if (!m) return u;
  var p = m[1].split('&').map(function (x) { return x.split('='); }).filter(function (x) { return x[0] === 'url' || x[0] === 'q'; })[0];
  return p ? decodeURIComponent(p[1]) : u;
}

// ---------------------------------------------------------------- checagens de texto

function vbusca_norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}
/** Pelo menos `minimo` (0 a 1) das palavras do valor aparecem no texto. */
function vbusca_apoiado(valor, texto, minimo) {
  var fonte = vbusca_norm(texto);
  var palavras = vbusca_norm(valor).split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 3 || /\d/.test(w); });
  if (!palavras.length) return false;
  var achadas = palavras.filter(function (w) { return fonte.indexOf(w) >= 0; }).length;
  return achadas / palavras.length >= minimo;
}
function vbusca_tokens(s) {
  var pare = { de: 1, da: 1, do: 1, das: 1, dos: 1, the: 1, and: 1, for: 1, em: 1, na: 1, no: 1, of: 1, 'in': 1, e: 1, a: 1, o: 1 };
  return vbusca_norm(s).split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 2 && !pare[w]; });
}
function vbusca_semelhanca(a, b) {
  if (!a.length || !b.length) return 0;
  var sb = {};
  b.forEach(function (w) { sb[w] = 1; });
  var comum = a.filter(function (w) { return sb[w]; }).length;
  return comum / Math.max(a.length, b.length);
}

var VBUSCA_MESES_PT = ['janeiro', 'fevereiro', 'marco', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
var VBUSCA_MESES_EN = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
var VBUSCA_MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

/** A data aparece escrita no texto em algum formato comum (25/10/2026, 25/10, 25 de outubro, October 25, 2026-10-25...). */
function vbusca_dataNoTexto(d, texto) {
  var t = ' ' + vbusca_norm(texto).replace(/(\d)(?:st|nd|rd|th|º|o)\b/g, '$1') + ' ';
  var dia = d.getDate(), mes = d.getMonth() + 1, ano = d.getFullYear();
  var dd = ('0' + dia).slice(-2), mm = ('0' + mes).slice(-2);
  var D = '(?:0?' + dia + ')', M = '(?:0?' + mes + ')';
  var padroes = [
    '\\b' + D + '[\\/.\\-]' + M + '(?:[\\/.\\-](?:' + ano + '|' + String(ano).slice(2) + '))?(?![\\d]|[\\/.\\-]\\d)',
    '\\b' + ano + '-' + mm + '-' + dd + '\\b',
  ];
  [VBUSCA_MESES_PT, VBUSCA_MESES_EN, VBUSCA_MESES_ES].forEach(function (lista) {
    var nome = lista[mes - 1], abrev = nome.slice(0, 3);
    var n = '(?:' + nome + '|' + abrev + '\\.?)';
    padroes.push('\\b' + D + '\\s*(?:de\\s+)?' + n + '\\b');
    padroes.push('\\b' + n + '\\s+' + D + '\\b');
  });
  return padroes.some(function (p) { return new RegExp(p, 'i').test(t); });
}

function vbusca_semContato(s) {
  return String(s).replace(/\S+@\S+\.\S+/g, '').replace(/\bpix\b[^.;]*/gi, '').replace(/\(?\b\d{2}\)?\s?9?\d{4}[-\s]?\d{4}\b/g, '').replace(/\s{2,}/g, ' ').trim();
}

// ---------------------------------------------------------------- planilha

function vbusca_normUrl(u) {
  var s = String(u || '').trim();
  if (!/^https?:\/\//i.test(s)) return '';
  s = s.replace(/#.*$/, '').replace(/([?&])(utm_[^=&]+|fbclid|gclid|ref)=[^&]*/gi, '$1').replace(/[?&]+$/, '');
  s = s.replace(/^https?:\/\/(www\.)?/i, '').replace(/\/+$/, '');
  return s.toLowerCase();
}
function vbusca_bloqueado(u) {
  var m = String(u).match(/^https?:\/\/([^\/?#]+)/i);
  return !m || VBUSCA.BLOQUEADOS.test(m[1].replace(/^www\./i, ''));
}
function vbusca_existentes(sh, ix) {
  var out = { links: {}, titulos: [] };
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function (r) {
    var k = vbusca_normUrl(r[ix.link]);
    if (k) out.links[k] = 1;
    if (r[ix.titulo]) out.titulos.push(vbusca_tokens(r[ix.titulo]));
  });
  return out;
}
function vbusca_contarFontes(abaFontes, novas) {
  if (abaFontes.getLastRow() < 2) return;
  var nomes = abaFontes.getRange(2, 3, abaFontes.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); });
  var cont = abaFontes.getRange(2, 7, nomes.length, 1).getValues();
  novas.forEach(function (v) {
    var i = nomes.indexOf(String(v.origem || '').replace(/^busca automática · /, ''));
    if (i >= 0) cont[i][0] = (Number(cont[i][0]) || 0) + 1;
  });
  abaFontes.getRange(2, 7, nomes.length, 1).setValues(cont);
}
function vbusca_email(ss, sh, novas) {
  var html = '<div style="font-family:Arial;color:#0E1738"><p style="font-size:15px">A busca automática colocou ' + novas.length +
    (novas.length === 1 ? ' vaga nova' : ' vagas novas') + ' no site. Nada aqui precisa de aprovação.</p><ul style="padding-left:18px;font-size:14px">' +
    novas.map(function (v) {
      return '<li style="margin:4px 0"><b>' + vagas_esc(v.titulo) + '</b> · ' + vagas_esc(v.instituicao) + ' · até ' + vagas_fmt(v.inscricao_fecha) +
        ' · <a href="' + vagas_esc(v.link) + '">página</a> <span style="color:#59628A">(' + vagas_esc(v.origem) + ')</span></li>';
    }).join('') + '</ul><p style="font-size:13px;color:#59628A">Para tirar uma do site, coloque publicar = não na aba ' + VAGAS.ABA + '.</p></div>';
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'Handoff Vagas · ' + novas.length + (novas.length === 1 ? ' vaga nova no site' : ' vagas novas no site'), htmlBody: html });
}

// ---------------------------------------------------------------- fontes iniciais (usadas só pelo vagasBuscaInstalar)

var VBUSCA_FONTES_INICIAIS = [
 {
  "tipo": "pagina",
  "nome": "FMUSP · Eventos",
  "url": "https://www.fm.usp.br/fmusp/institucional/eventos"
 },
 {
  "tipo": "pagina",
  "nome": "Agência FAPESP · Agenda de cursos e escolas",
  "url": "https://agencia.fapesp.br/agenda/"
 },
 {
  "tipo": "pagina",
  "nome": "INCA · Eventos e curso de verão",
  "url": "https://www.gov.br/inca/pt-br/assuntos/eventos"
 },
 {
  "tipo": "pagina",
  "nome": "Pequeno Príncipe · Eventos",
  "url": "https://pequenoprincipe.org.br/evento/"
 },
 {
  "tipo": "pagina",
  "nome": "CFM · Eventos",
  "url": "https://portal.cfm.org.br/eventos"
 },
 {
  "tipo": "pagina",
  "nome": "Utrecht Summer School · Life Sciences",
  "url": "https://utrechtsummerschool.nl/courses/life-sciences"
 },
 {
  "tipo": "pagina",
  "nome": "Leiden University · Summer schools",
  "url": "https://www.universiteitleiden.nl/en/education/study-programmes/summer-schools"
 },
 {
  "tipo": "pagina",
  "nome": "UCL Summer School · Módulos",
  "url": "https://www.ucl.ac.uk/prospective-students/summer-school/modules"
 },
 {
  "tipo": "pagina",
  "nome": "LSHTM · Short courses",
  "url": "https://www.lshtm.ac.uk/study/courses/short-courses/a-z"
 },
 {
  "tipo": "pagina",
  "nome": "Karolinska · Summer courses abroad",
  "url": "https://education.ki.se/student-at-ki/studies-and-traineeships-abroad/summer-courses-abroad"
 },
 {
  "tipo": "pagina",
  "nome": "Institut Pasteur · Cursos",
  "url": "https://www.pasteur.fr/en/about-us/our-missions/education/training-programs/pasteur-courses"
 },
 {
  "tipo": "pagina",
  "nome": "EMBL · Cursos e conferências",
  "url": "https://www.embl.org/events/"
 },
 {
  "tipo": "pagina",
  "nome": "Calendário de congressos estudantis (MU Varna)",
  "url": "https://bsysb.mu-varna.bg/conference-calender/"
 },
 {
  "tipo": "pagina",
  "nome": "Universidade de Copenhague · Summer courses de saúde global",
  "url": "https://globalhealth.ku.dk/studies/summer-courses/"
 }
];
