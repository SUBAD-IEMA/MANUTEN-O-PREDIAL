/* =========================================================================
   GESTÃO DE MANUTENÇÃO PREDIAL · IEMA - SUBAD
   Cole este arquivo inteiro no editor do Apps Script da planilha (Extensões > Apps Script).
   Depois execute a função "instalar" uma vez e implante como App da Web (veja o README).
   ========================================================================= */

/* =========================================================================
   GOOGLE APPS SCRIPT: a planilha Google funciona como banco de dados.
   - doPost: API usada pelo site publicado no GitHub Pages
   - instalar: cria as abas, perfis, o usuário master e o gatilho de backup
   - backupMensal: no 1º dia útil do mês envia XLSX, JSON e CSV por e-mail
   ========================================================================= */

/** Senha provisória do usuário "master" criada na instalação (troque no 1º acesso). */
var SENHA_INICIAL_MASTER = 'Iema@2026';

/** Feriados fixos considerados no cálculo do 1º dia útil (dia-mês). */
var FERIADOS_FIXOS = ['01-01', '21-04', '01-05', '23-05', '07-09', '12-10', '02-11', '15-11', '20-11', '25-12'];

var ABAS = ['Demandas', 'NotasFiscais', 'Empenhos', 'Lotes', 'Usuarios', 'Perfis', 'Historico'];
var NOME_ABA = { Demandas: 'Demandas', NotasFiscais: 'Notas Fiscais', Empenhos: 'Empenhos', Lotes: 'Lotes', Usuarios: 'Usuários', Perfis: 'Perfis', Historico: 'Histórico' };
var ABA_CONFIG = 'Configurações';

/* ---------------------------------------------------------------------
   Repositório sobre a planilha
   --------------------------------------------------------------------- */
function planilha_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function fuso_() { return planilha_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
function aba_(tabela) {
  var sh = planilha_().getSheetByName(NOME_ABA[tabela] || tabela);
  if (!sh) throw new Error('A aba "' + (NOME_ABA[tabela] || tabela) + '" não existe. Execute a função instalar.');
  return sh;
}
function lerCelula_(v, tipo) {
  if (v instanceof Date) {
    if (tipo === 'datahora') return Utilities.formatDate(v, fuso_(), "yyyy-MM-dd'T'HH:mm:ss");
    return Utilities.formatDate(v, fuso_(), 'yyyy-MM-dd');
  }
  return v;
}
function escreverCelula_(v, tipo) {
  if (tipo === 'lista') return (v || []).join('; ');
  if (tipo === 'bool') return !!v;
  if (tipo === 'num' || tipo === 'int') return Number(v) || 0;
  if ((tipo === 'data') && v) { var p = String(v).slice(0, 10).split('-'); return new Date(+p[0], +p[1] - 1, +p[2], 12); }
  if (tipo === 'datahora' && v) { var d = new Date(v); return isNaN(d) ? String(v) : d; }
  return v == null ? '' : String(v);
}
var cache_ = {};
var Repo = {
  listar: function (tabela) {
    if (cache_[tabela]) return cache_[tabela].map(function (x) { return JSON.parse(JSON.stringify(x)); });
    var sh = aba_(tabela), def = Logica.TABELAS[tabela];
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) { cache_[tabela] = []; return []; }
    var head = vals[0].map(String), idx = def.colunas.map(function (c) { return head.indexOf(c[1]); });
    var out = [];
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      if (row.join('') === '') continue;
      var o = {};
      def.colunas.forEach(function (c, i) { o[c[0]] = idx[i] >= 0 ? lerCelula_(row[idx[i]], c[2]) : ''; });
      out.push(Logica.normalizar(tabela, o));
    }
    cache_[tabela] = out;
    return out.map(function (x) { return JSON.parse(JSON.stringify(x)); });
  },
  linha_: function (tabela, o) {
    var sh = aba_(tabela), def = Logica.TABELAS[tabela];
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var row = head.map(function () { return ''; });
    def.colunas.forEach(function (c) { var i = head.indexOf(c[1]); if (i >= 0) row[i] = escreverCelula_(o[c[0]], c[2]); });
    return row;
  },
  salvar: function (tabela, o) {
    var sh = aba_(tabela), def = Logica.TABELAS[tabela];
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var ci = head.indexOf(def.colunas.filter(function (c) { return c[0] === def.chave; })[0][1]);
    var n = sh.getLastRow(), linha = -1;
    if (n > 1) {
      var chaves = sh.getRange(2, ci + 1, n - 1, 1).getValues();
      for (var i = 0; i < chaves.length; i++) if (String(chaves[i][0]) === String(o[def.chave])) { linha = i + 2; break; }
    }
    var row = this.linha_(tabela, o);
    if (linha < 0) linha = n + 1;
    formatarLinhas_(sh, tabela, linha, 1);
    sh.getRange(linha, 1, 1, row.length).setValues([row]);
    delete cache_[tabela];
  },
  remover: function (tabela, chave) {
    var sh = aba_(tabela), def = Logica.TABELAS[tabela];
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var ci = head.indexOf(def.colunas.filter(function (c) { return c[0] === def.chave; })[0][1]);
    var n = sh.getLastRow();
    if (n < 2) return;
    var chaves = sh.getRange(2, ci + 1, n - 1, 1).getValues();
    for (var i = chaves.length - 1; i >= 0; i--) if (String(chaves[i][0]) === String(chave)) sh.deleteRow(i + 2);
    delete cache_[tabela];
  },
  substituir: function (tabela, lista) {
    var sh = aba_(tabela), self = this;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    if (lista.length) {
      var rows = lista.map(function (o) { return self.linha_(tabela, o); });
      formatarLinhas_(sh, tabela, 2, rows.length);
      sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    delete cache_[tabela];
  },
  historico: function (regs) {
    if (!regs || !regs.length) return;
    var sh = aba_('Historico'), self = this;
    var rows = regs.map(function (o) { return self.linha_('Historico', o); });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    delete cache_.Historico;
  },
  config: function () {
    var sh = planilha_().getSheetByName(ABA_CONFIG);
    var cfg = JSON.parse(JSON.stringify(Logica.CONFIG_PADRAO));
    if (!sh) return cfg;
    sh.getDataRange().getValues().slice(1).forEach(function (r) {
      var k = String(r[0]); if (!k || !(k in cfg)) return;
      var t = Logica.CONFIG_TIPOS[k], v = r[1];
      if (t === 'num') v = Number(v) || 0;
      else if (t === 'numOpc') v = v === '' || v == null ? '' : Number(v) || 0;
      else if (t === 'bool') v = v === true || v === 'TRUE' || v === 'Sim';
      else if (t === 'lista') v = String(v).split(/\n|;/).map(function (x) { return x.trim(); }).filter(String);
      else v = lerCelula_(v, 'data') === v ? String(v) : lerCelula_(v, 'data');
      cfg[k] = v;
    });
    return cfg;
  },
  salvarConfig: function (cfg) {
    var sh = planilha_().getSheetByName(ABA_CONFIG) || planilha_().insertSheet(ABA_CONFIG);
    var rows = [['Chave', 'Valor', 'Descrição']].concat(Object.keys(Logica.CONFIG_PADRAO).map(function (k) {
      var v = cfg[k];
      if (Array.isArray(v)) v = v.join('\n');
      return [k, v, DESCR_CONFIG_[k] || ''];
    }));
    sh.clearContents();
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 420); sh.setColumnWidth(3, 420);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0B4F4A').setFontColor('#FFFFFF');
  },
  agora: function () { return Utilities.formatDate(new Date(), fuso_(), "yyyy-MM-dd'T'HH:mm:ss"); },
  uuid: function () { return Utilities.getUuid().replace(/-/g, ''); },
  tokenSalvar: function (t, login) { CacheService.getScriptCache().put('tk_' + t, login, 21600); },
  tokenLer: function (t) { var c = CacheService.getScriptCache(), v = c.get('tk_' + t); if (v) c.put('tk_' + t, v, 21600); return v; },
  tokenRemover: function (t) { CacheService.getScriptCache().remove('tk_' + t); },
  versao: function () { return PropertiesService.getScriptProperties().getProperty('versao') || '0'; },
  marcarVersao: function () { PropertiesService.getScriptProperties().setProperty('versao', String(Date.now())); },
  travar: function (fn) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { cache_ = {}; return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
  },
  enviarBackup: function (origem) { return enviarBackup_(origem); },
  aposGravar: function () { atualizarLooker_(); },
  looker: function () {
    atualizarLooker_();
    var ss = planilha_(), id = function (n) { var sh = ss.getSheetByName(n); return sh ? String(sh.getSheetId()) : ''; };
    return { spreadsheetId: ss.getId(), demandas: id(LOOKER_ABAS.demandas), notas: id(LOOKER_ABAS.notas), saldos: id(LOOKER_ABAS.saldos), nome: ss.getName() };
  }
};

/* ---------------------------------------------------------------------
   Base para o Google Looker Studio (somente leitura, refeita a cada gravação)
   --------------------------------------------------------------------- */
var LOOKER_ABAS = { demandas: 'Looker - Demandas', notas: 'Looker - Notas Fiscais', saldos: 'Looker - Saldos' };
function dataLk_(s) { if (!s) return ''; var p = String(s).slice(0, 10).split('-'); return new Date(+p[0], +p[1] - 1, +p[2], 12); }
function escreverAba_(nome, cab, linhas, formatos) {
  var ss = planilha_(), sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
  sh.clearContents();
  var rows = [cab].concat(linhas);
  sh.getRange(1, 1, rows.length, cab.length).setValues(rows);
  sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#0B4F4A').setFontColor('#FFFFFF');
  if (linhas.length && formatos) formatos.forEach(function (f, i) { if (f) sh.getRange(2, i + 1, linhas.length, 1).setNumberFormat(f); });
  sh.setFrozenRows(1);
}
function atualizarLooker_() {
  var dados = { config: Repo.config(), lotes: Repo.listar('Lotes'), empenhos: Repo.listar('Empenhos'), demandas: Repo.listar('Demandas'), nfs: Repo.listar('NotasFiscais') };
  var fin = Logica.financeiro(dados), hoje = Logica.hoje();
  var lote = {}; dados.lotes.forEach(function (l) { lote[l.id] = l; });
  var emp = {}; dados.empenhos.forEach(function (e) { emp[e.id] = e; });
  var dem = {}; dados.demandas.forEach(function (d) { dem[d.id] = d; });
  var D = 'dd/mm/yyyy', M = '"R$" #,##0.00';
  escreverAba_(LOOKER_ABAS.demandas,
    ['Nº', 'Título', 'Modalidade', 'Lote', 'Descrição do lote', 'Situação', 'Prazo', 'Em atraso', 'Dias de atraso', 'Unidade demandante', 'Setores atendidos', 'Local',
      'Sistema afetado', 'Criticidade', 'Data da solicitação', 'Mês da solicitação', 'Ano da solicitação', 'Início previsto', 'Término previsto',
      'Execução: início', 'Execução: término', 'Ateste: data', 'Pagamento: data', 'Valor estimado', 'Valor atestado', 'Valor das notas fiscais',
      'Nº da Ordem de Serviço', 'Nº do Atendimento (Help Desk)', 'Registrado por', 'Atualizado em'],
    dados.demandas.map(function (d) {
      var l = lote[d.loteId] || {}, pz = Logica.prazo(d, hoje);
      return [d.id, d.titulo, d.modalidade, l.numero ? 'Lote ' + l.numero : '', l.descricao || '', Logica.situacao(d).nome, pz.nome, pz.k === 'atrasada' ? 'Sim' : 'Não',
        pz.k === 'atrasada' || pz.k === 'concluida-atraso' ? pz.dias : 0, d.unidade, d.setores.join(', '), d.local, d.sistema, d.criticidade,
        dataLk_(d.dataSolicitacao), String(d.dataSolicitacao).slice(0, 7), String(d.dataSolicitacao).slice(0, 4), dataLk_(d.inicioPrevisto), dataLk_(d.fimPrevisto),
        dataLk_(d.execucaoInicio), d.execucaoOk ? dataLk_(d.execucaoFim) : '', d.atesteOk ? dataLk_(d.atesteData) : '', d.pagamentoOk ? dataLk_(d.pagamentoData) : '',
        d.valorEstimado, d.atesteOk ? d.atesteValor : 0, fin.nfPorDemanda[d.id] || 0, d.osNumero, d.atendimentoNumero, d.criadoPor, d.atualizadoEm];
    }), [0, 0, 0, 0, 0, 0, 0, 0, '0', 0, 0, 0, 0, 0, D, '@', '@', D, D, D, D, D, D, M, M, M, '@', '@', 0, 0]);
  escreverAba_(LOOKER_ABAS.notas,
    ['Data de emissão', 'Mês', 'Ano', 'Nº da Nota Fiscal', 'Valor do serviço', 'Demanda', 'Título', 'Modalidade', 'Lote', 'Unidade demandante', 'Local',
      'Empenho', 'Tipo do empenho', 'Empenho principal'],
    dados.nfs.map(function (n) {
      var d = dem[n.demandaId] || {}, l = lote[d.loteId] || {}, e = emp[n.empenhoId] || {}, pr = e.tipo === 'Reforço' ? (emp[e.vinculadoA] || {}) : e;
      return [dataLk_(n.data), String(n.data).slice(0, 7), String(n.data).slice(0, 4), n.numero, n.valor, n.demandaId, d.titulo || '', d.modalidade || '',
        l.numero ? 'Lote ' + l.numero : '', d.unidade || '', d.local || '', e.numero || '', e.tipo || '', pr.numero || ''];
    }), [D, '@', '@', '@', M, 0, 0, 0, 0, 0, 0, '@', 0, '@']);
  var linha = function (tipo, nome, o) { return [tipo, nome, o.total, o.usado, o.saldo, Math.round(o.pctUsado * 10) / 10, o.limite, o.alerta ? 'Sim' : 'Não', Repo.agora()]; };
  escreverAba_(LOOKER_ABAS.saldos, ['Tipo', 'Nome', 'Valor total', 'Executado (notas fiscais)', 'Saldo', '% executado', 'Limite de alerta (%)', 'Em alerta', 'Atualizado em'],
    [linha('Valor global', 'Valor global da vigência', fin.global)]
      .concat(fin.lotes.map(function (x) { return linha('Lote', 'Lote ' + x.lote.numero + ' - ' + x.lote.descricao, x); }))
      .concat(fin.empenhos.map(function (x) { return linha('Empenho', x.empenho.numero + (x.membros.length > 1 ? ' (com reforços)' : ''), x); })),
    [0, 0, M, M, M, '0.0', '0', 0, 0]);
}
var DESCR_CONFIG_ = {
  orgao: 'Nome exibido no sistema e nos relatórios', contratoGestao: 'Processo de gestão', contratoExecucao: 'Processo de execução',
  numeroContrato: 'Número do contrato', ataNumero: 'Número da ata', ataAno: 'Ano da ata (AAAA)', contratada: 'Empresa contratada',
  vigenciaInicio: 'Início da vigência (AAAA-MM-DD)', vigenciaFim: 'Fim da vigência (AAAA-MM-DD)',
  valorGlobal: 'Valor global da vigência (usado apenas se os lotes não tiverem valor). Alterar exige senha no sistema.',
  alertaGlobalPct: 'Alerta quando o valor global executado atingir este percentual (padrão fixo: 70). Alterar exige senha no sistema.',
  alertaEmpenhoPct: 'Alerta quando o empenho executado atingir este percentual (padrão fixo: 70). Alterar exige senha no sistema.',
  setores: 'Setores e unidades (um por linha)', locais: 'Não utilizado nesta versão', emailsBackup: 'E-mails que recebem o backup mensal (separados por vírgula)',
  backupAutomatico: 'TRUE para enviar o backup no 1º dia útil do mês', lookerRelatorio: 'Endereço do relatório salvo no Looker Studio', chavePainel: 'Chave do link do painel somente leitura (gerada pelo sistema; apague para revogar)', ultimoBackup: 'Preenchido automaticamente'
};

function formatarLinhas_(sh, tabela, linha, qtd) {
  var def = Logica.TABELAS[tabela];
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var fmt = head.map(function (h) {
    var c = def.colunas.filter(function (x) { return x[1] === h; })[0];
    if (!c) return '@';
    return { texto: '@', lista: '@', data: 'dd/mm/yyyy', datahora: 'dd/mm/yyyy hh:mm', num: '"R$" #,##0.00', int: '0', bool: '@' }[c[2]];
  });
  var formatos = [];
  for (var i = 0; i < qtd; i++) formatos.push(fmt.slice());
  var rng = sh.getRange(linha, 1, qtd, head.length);
  rng.setNumberFormats(formatos);
  def.colunas.forEach(function (c) { if (c[2] === 'bool') { var ix = head.indexOf(c[1]); if (ix >= 0) sh.getRange(linha, ix + 1, qtd, 1).setNumberFormat('General').insertCheckboxes(); } });
}

/* ---------------------------------------------------------------------
   API web (chamada pelo site)
   --------------------------------------------------------------------- */
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, codigo: 'REQ', mensagem: 'Requisição inválida.' }); }
  cache_ = {};
  return json_(Servico.executar(Repo, req));
}
function doGet() {
  return json_({ ok: true, sistema: 'Gestão de Manutenção Predial - IEMA', planilha: planilha_().getName(), versao: Repo.versao() });
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------------------------------------------------------------------
   Instalação (execute uma vez pelo editor ou pelo menu da planilha)
   --------------------------------------------------------------------- */
function instalar() {
  var ss = planilha_();
  ABAS.forEach(function (t) {
    var nome = NOME_ABA[t], sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    var cols = Logica.TABELAS[t].colunas.map(function (c) { return c[1]; });
    var atual = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    cols.forEach(function (c) { if (atual.indexOf(c) < 0) atual.push(c); });
    sh.getRange(1, 1, 1, atual.length).setValues([atual]).setFontWeight('bold').setBackground('#0B4F4A').setFontColor('#FFFFFF').setWrap(true);
    sh.setFrozenRows(1);
    if (t === 'Usuarios') { var ix = atual.indexOf('Hash da senha'); if (ix >= 0) sh.hideColumns(ix + 1); ix = atual.indexOf('Salt'); if (ix >= 0) sh.hideColumns(ix + 1); }
  });
  if (!ss.getSheetByName(ABA_CONFIG)) Repo.salvarConfig(Logica.CONFIG_PADRAO);
  else {
    var cfg = Repo.config();
    // Atualização da versão anterior: troca a lista antiga de setores pelas siglas e remove a lista de imóveis
    if (!cfg.setores.length || cfg.setores[0] === 'Diretoria-Geral') cfg.setores = Logica.CONFIG_PADRAO.setores.slice();
    cfg.locais = [];
    if (cfg.alertaGlobalPct === '' || !(Number(cfg.alertaGlobalPct) > 0)) cfg.alertaGlobalPct = 70;
    if (cfg.alertaEmpenhoPct === '' || !(Number(cfg.alertaEmpenhoPct) > 0)) cfg.alertaEmpenhoPct = 70;
    Repo.salvarConfig(cfg);
  }
  var perfis = Repo.listar('Perfis');
  Logica.PERFIS_PADRAO.forEach(function (p) { if (!perfis.some(function (x) { return x.nome === p.nome; })) Repo.salvar('Perfis', Logica.normalizar('Perfis', p)); });
  if (!Repo.listar('Usuarios').some(function (u) { return u.perfil === 'Gestor Master'; })) {
    var salt = Repo.uuid();
    Repo.salvar('Usuarios', Logica.normalizar('Usuarios', { login: 'master', nome: 'Gestor Master', email: '', perfil: 'Gestor Master', ativo: true, trocarSenha: true, salt: salt, hash: Logica.hashSenha(SENHA_INICIAL_MASTER, salt), criadoEm: Repo.agora() }));
  }
  var p1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (p1 && p1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p1);
  instalarGatilho_();
  atualizarLooker_();
  Repo.marcarVersao();
  Repo.historico([{ em: Repo.agora(), usuario: 'sistema', perfil: '', acao: 'Estrutura instalada ou atualizada', demandaId: '', campo: '', anterior: '', novo: '', motivo: '' }]);
  return 'Instalação concluída. Usuário: master · senha provisória: ' + SENHA_INICIAL_MASTER;
}
function instalarGatilho_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupMensal') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupMensal').timeBased().everyDays(1).atHour(7).create();
}
function redefinirSenhaMaster() {
  var u = Repo.listar('Usuarios').filter(function (x) { return x.login === 'master'; })[0];
  if (!u) return instalar();
  u.salt = Repo.uuid(); u.hash = Logica.hashSenha(SENHA_INICIAL_MASTER, u.salt); u.trocarSenha = true; u.ativo = true; u.perfil = 'Gestor Master';
  Repo.salvar('Usuarios', u);
  return 'Senha do usuário master redefinida para: ' + SENHA_INICIAL_MASTER;
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Manutenção Predial')
    .addItem('Instalar ou atualizar estrutura', 'menuInstalar')
    .addItem('Enviar backup agora', 'menuBackup')
    .addItem('Redefinir senha do usuário master', 'menuSenha')
    .addToUi();
}
function menuInstalar() { SpreadsheetApp.getUi().alert(instalar()); }
function menuBackup() { var r = enviarBackup_('manual pelo menu da planilha'); SpreadsheetApp.getUi().alert('Backup enviado para ' + r.destinatarios + '.'); }
function menuSenha() { SpreadsheetApp.getUi().alert(redefinirSenhaMaster()); }

/* ---------------------------------------------------------------------
   Backup mensal (1º dia útil) em XLSX, JSON e CSV
   --------------------------------------------------------------------- */
function pascoa_(ano) {
  var a = ano % 19, b = Math.floor(ano / 100), c = ano % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  var mes = Math.floor((h + l - 7 * m + 114) / 31), dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}
function ehDiaUtil_(d) {
  var dow = d.getDay();
  if (dow === 0 || dow === 6) return false;
  var dm = Utilities.formatDate(d, fuso_(), 'dd-MM');
  if (FERIADOS_FIXOS.indexOf(dm) >= 0) return false;
  var p = pascoa_(d.getFullYear()), moveis = [-48, -47, -2, 60].map(function (n) { var x = new Date(p); x.setDate(x.getDate() + n); return Utilities.formatDate(x, fuso_(), 'dd-MM'); });
  return moveis.indexOf(dm) < 0;
}
function ehPrimeiroDiaUtil_(hoje) {
  if (!ehDiaUtil_(hoje)) return false;
  for (var d = 1; d < hoje.getDate(); d++) if (ehDiaUtil_(new Date(hoje.getFullYear(), hoje.getMonth(), d))) return false;
  return true;
}
function backupMensal() {
  var cfg = Repo.config();
  if (!cfg.backupAutomatico) return;
  var hoje = new Date(), mes = Utilities.formatDate(hoje, fuso_(), 'yyyy-MM');
  var props = PropertiesService.getScriptProperties();
  if (!ehPrimeiroDiaUtil_(hoje) || props.getProperty('backupMes') === mes) return;
  enviarBackup_('automático do 1º dia útil');
  props.setProperty('backupMes', mes);
}
function csv_(tabela, lista) {
  var cols = Logica.TABELAS[tabela].colunas.filter(function (c) { return c[0] !== 'hash' && c[0] !== 'salt'; });
  var esc = function (v) { v = Logica.valorTexto(v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  return '\uFEFF' + [cols.map(function (c) { return esc(c[1]); }).join(';')].concat(lista.map(function (o) { return cols.map(function (c) { return esc(o[c[0]]); }).join(';'); })).join('\r\n');
}
function enviarBackup_(origem) {
  var ss = planilha_(), cfg = Repo.config(), carimbo = Utilities.formatDate(new Date(), fuso_(), 'yyyy-MM-dd_HHmm');
  var dados = { geradoEm: Repo.agora(), origem: origem, config: cfg };
  ['Demandas', 'NotasFiscais', 'Empenhos', 'Lotes', 'Historico'].forEach(function (t) { dados[{ Demandas: 'demandas', NotasFiscais: 'nfs', Empenhos: 'empenhos', Lotes: 'lotes', Historico: 'historico' }[t]] = Repo.listar(t); });
  var anexos = [];
  var xlsx = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  anexos.push(xlsx.getBlob().setName('manutencao-predial_' + carimbo + '.xlsx'));
  anexos.push(Utilities.newBlob(JSON.stringify(dados, null, 2), 'application/json', 'manutencao-predial_' + carimbo + '.json'));
  anexos.push(Utilities.newBlob(csv_('Demandas', dados.demandas), 'text/csv', 'demandas_' + carimbo + '.csv'));
  anexos.push(Utilities.newBlob(csv_('NotasFiscais', dados.nfs), 'text/csv', 'notas-fiscais_' + carimbo + '.csv'));
  anexos.push(Utilities.newBlob(csv_('Empenhos', dados.empenhos), 'text/csv', 'empenhos_' + carimbo + '.csv'));
  anexos.push(Utilities.newBlob(csv_('Lotes', dados.lotes), 'text/csv', 'lotes_' + carimbo + '.csv'));
  anexos.push(Utilities.newBlob(csv_('Historico', dados.historico), 'text/csv', 'historico_' + carimbo + '.csv'));
  var pastaNome = 'Backups - Manutenção Predial IEMA', it = DriveApp.getFoldersByName(pastaNome);
  var pasta = it.hasNext() ? it.next() : DriveApp.createFolder(pastaNome);
  anexos.forEach(function (b) { pasta.createFile(b.copyBlob()); });
  var fin = Logica.financeiro(dados);
  var dest = String(cfg.emailsBackup || '').split(/[,;\s]+/).filter(function (x) { return x.indexOf('@') > 0; }).join(',');
  if (dest) {
    var brl = function (v) { return 'R$ ' + (Number(v) || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); };
    MailApp.sendEmail({
      to: dest,
      subject: 'Backup da Manutenção Predial (IEMA) - ' + Utilities.formatDate(new Date(), fuso_(), 'dd/MM/yyyy'),
      htmlBody: '<p>Segue o backup ' + origem + ' do sistema de Gestão de Manutenção Predial.</p>' +
        '<ul><li>Demandas: ' + dados.demandas.length + '</li><li>Notas Fiscais: ' + dados.nfs.length + '</li><li>Empenhos: ' + dados.empenhos.length + '</li>' +
        '<li>Valor global da vigência: ' + brl(fin.global.total) + '</li><li>Saldo global: ' + brl(fin.global.saldo) + ' (' + fin.global.pctSaldo.toFixed(1).replace('.', ',') + '%)</li></ul>' +
        '<p>Anexos: planilha completa (Excel), base completa (JSON) e tabelas em CSV. Cópias também foram salvas na pasta "' + pastaNome + '" do Google Drive.</p>',
      attachments: anexos
    });
  }
  cfg.ultimoBackup = Repo.agora() + ' (' + origem + ')';
  Repo.salvarConfig(cfg);
  return { destinatarios: dest || 'nenhum e-mail configurado', arquivos: anexos.length };
}


/* =========================================================================
   LÓGICA COMPARTILHADA
   Este mesmo código roda no navegador (modo demonstração) e no Google Apps
   Script (modo Google Sheets), garantindo as mesmas regras nos dois lados.
   ========================================================================= */
var Logica = (function () {
  'use strict';
  var L = {};

  /* ---------------------------------------------------------------------
     Estrutura das planilhas (abas, rótulos de coluna e tipos)
     --------------------------------------------------------------------- */
  L.TABELAS = {
    Demandas: { chave: 'id', colunas: [
      ['id', 'Nº', 'texto'], ['titulo', 'Título', 'texto'], ['modalidade', 'Modalidade', 'texto'], ['loteId', 'Lote', 'texto'],
      ['unidade', 'Unidade demandante', 'texto'], ['solicitante', 'Solicitante', 'texto'], ['setores', 'Setores atendidos', 'lista'],
      ['local', 'Local', 'texto'], ['localInterno', 'Localização interna', 'texto'], ['sistema', 'Sistema afetado', 'texto'],
      ['criticidade', 'Criticidade', 'texto'], ['risco', 'Risco declarado', 'bool'], ['pessoas', 'Pessoas impactadas', 'int'],
      ['responsavelLocal', 'Responsável local', 'texto'], ['descricao', 'Descrição', 'texto'],
      ['processo', 'Nº do processo', 'texto'], ['osNumero', 'Nº da Ordem de Serviço', 'texto'], ['atendimentoNumero', 'Nº do Atendimento (Help Desk)', 'texto'],
      ['valorEstimado', 'Valor estimado', 'num'], ['dataSolicitacao', 'Data da solicitação', 'data'], ['inicioPrevisto', 'Início previsto', 'data'],
      ['prazoQtd', 'Prazo estimado', 'int'], ['prazoUnidade', 'Unidade do prazo', 'texto'], ['fimPrevisto', 'Término previsto', 'data'],
      ['inicialOk', 'Demanda inicial', 'bool'], ['inicialData', 'Demanda inicial: data', 'data'], ['inicialPor', 'Demanda inicial: por', 'texto'],
      ['aprovacaoOk', 'Aprovação', 'bool'], ['aprovacaoData', 'Aprovação: data', 'data'], ['aprovacaoObs', 'Aprovação: observação', 'texto'], ['aprovacaoPor', 'Aprovação: por', 'texto'],
      ['execucaoOk', 'Execução', 'bool'], ['execucaoInicio', 'Execução: início', 'data'], ['execucaoFim', 'Execução: término', 'data'], ['execucaoPor', 'Execução: por', 'texto'],
      ['atesteOk', 'Ateste', 'bool'], ['atesteData', 'Ateste: data', 'data'], ['atesteValor', 'Ateste: valor', 'num'], ['atestePor', 'Ateste: por', 'texto'],
      ['pagamentoOk', 'Pagamento', 'bool'], ['pagamentoData', 'Pagamento: data', 'data'], ['pagamentoPor', 'Pagamento: por', 'texto'],
      ['cancelada', 'Cancelada', 'bool'], ['motivoCancelamento', 'Motivo do cancelamento', 'texto'], ['obs', 'Observações', 'texto'],
      ['criadoPor', 'Registrado por', 'texto'], ['criadoEm', 'Registrado em', 'datahora'], ['atualizadoPor', 'Atualizado por', 'texto'], ['atualizadoEm', 'Atualizado em', 'datahora']
    ] },
    NotasFiscais: { chave: 'id', colunas: [
      ['id', 'Código interno', 'texto'], ['demandaId', 'Demanda', 'texto'], ['numero', 'Nº da Nota Fiscal', 'texto'], ['valor', 'Valor do serviço', 'num'],
      ['data', 'Data de emissão', 'data'], ['empenhoId', 'Empenho', 'texto'], ['obs', 'Observação', 'texto'],
      ['registradoPor', 'Registrado por', 'texto'], ['registradoEm', 'Registrado em', 'datahora']
    ] },
    Empenhos: { chave: 'id', colunas: [
      ['id', 'Código interno', 'texto'], ['numero', 'Nº do empenho', 'texto'], ['tipo', 'Tipo', 'texto'], ['vinculadoA', 'Reforço do empenho', 'texto'],
      ['valor', 'Valor', 'num'], ['data', 'Data', 'data'], ['descricao', 'Descrição', 'texto']
    ] },
    Lotes: { chave: 'id', colunas: [
      ['id', 'Código interno', 'texto'], ['numero', 'Nº do lote', 'texto'], ['descricao', 'Descrição', 'texto'], ['valor', 'Valor do lote', 'num']
    ] },
    Usuarios: { chave: 'login', colunas: [
      ['login', 'Login', 'texto'], ['nome', 'Nome', 'texto'], ['email', 'E-mail', 'texto'], ['perfil', 'Perfil', 'texto'], ['ativo', 'Ativo', 'bool'],
      ['trocarSenha', 'Trocar senha no próximo acesso', 'bool'], ['salt', 'Salt', 'texto'], ['hash', 'Hash da senha', 'texto'], ['criadoEm', 'Criado em', 'datahora']
    ] },
    Perfis: { chave: 'nome', colunas: [
      ['nome', 'Perfil', 'texto'], ['verTodas', 'Vê todas as demandas', 'bool'], ['criar', 'Registra demandas', 'bool'],
      ['editarOutras', 'Edita demandas de outros', 'bool'], ['campos', 'Grupos de campos editáveis', 'lista'], ['etapas', 'Marca etapas', 'bool'],
      ['financeiro', 'Lança notas, empenhos e lotes', 'bool'], ['areaGestor', 'Acessa a Área do Gestor', 'bool'],
      ['configuracoes', 'Altera configurações', 'bool'], ['usuarios', 'Gerencia usuários e perfis', 'bool'], ['excluir', 'Exclui e cancela demandas', 'bool'],
      ['protegido', 'Perfil do sistema', 'bool']
    ] },
    Historico: { chave: null, colunas: [
      ['em', 'Data e hora', 'datahora'], ['usuario', 'Usuário', 'texto'], ['perfil', 'Perfil', 'texto'], ['acao', 'Ação', 'texto'],
      ['demandaId', 'Demanda', 'texto'], ['campo', 'Campo', 'texto'], ['anterior', 'Valor anterior', 'texto'], ['novo', 'Valor novo', 'texto'], ['motivo', 'Motivo', 'texto']
    ] }
  };

  L.CONFIG_PADRAO = {
    orgao: 'Instituto Estadual de Meio Ambiente (IEMA) - SUBAD',
    contratoGestao: '', contratoExecucao: '', numeroContrato: '', ataNumero: '', ataAno: '', contratada: '', vigenciaInicio: '', vigenciaFim: '',
    valorGlobal: 0, alertaGlobalPct: 70, alertaEmpenhoPct: 70,
    setores: ['APACB', 'APAGO', 'APAGUA', 'APAPE', 'APAPM', 'APAS', 'ARIEMV', 'ASSCOM', 'ASSJUR', 'CONADM', 'DIAF', 'DILCAM', 'DISPN', 'GEA', 'GEAD',
      'GECAM', 'GEDOC', 'GEFIN', 'GEIC', 'GEOCIP', 'GEPREVINES', 'GEPRO', 'GEREF', 'GESIM', 'GFAU', 'GFI', 'GGEP', 'GGUC', 'GP', 'GQA', 'GTI',
      'MONACP', 'MONAFF', 'MONAST', 'PECF', 'PEFG', 'PEI', 'PEMF', 'PEPAZ', 'PEPCV', 'RDSCO', 'REBIODB', 'SUAP', 'SUBAD', 'SUBCON', 'SUBEC',
      'SUBED', 'SUBEI', 'SUBEUC', 'SUBFAU', 'SUBIO', 'SUBMIN', 'SUBOC', 'SUBOR', 'SUBPOI', 'SUBSAN', 'SUBSID', 'SUCOP', 'SUFAM', 'SUGEO', 'SUPAT'],
    locais: [],
    emailsBackup: 'cleilson.costa@iema.es.gov.br, subad@iema.es.gov.br',
    backupAutomatico: true,
    lookerRelatorio: '',
    chavePainel: '',
    ultimoBackup: ''
  };
  L.ALERTA_PADRAO = 70;
  L.ROTULOS_CONFIG = {
    contratoGestao: 'Processo de gestão', contratoExecucao: 'Processo de execução', numeroContrato: 'Número do contrato', ataNumero: 'Número da ata',
    ataAno: 'Ano da ata', contratada: 'Contratada', vigenciaInicio: 'Início da vigência', vigenciaFim: 'Fim da vigência', valorGlobal: 'Valor global da vigência',
    alertaGlobalPct: 'Alerta de execução do valor global (%)', alertaEmpenhoPct: 'Alerta de execução do empenho (%)', setores: 'Setores e unidades',
    emailsBackup: 'E-mails do backup', backupAutomatico: 'Backup automático', lookerRelatorio: 'Relatório salvo no Looker Studio', orgao: 'Órgão'
  };
  L.CONFIG_TIPOS = { valorGlobal: 'num', alertaGlobalPct: 'numOpc', alertaEmpenhoPct: 'numOpc', setores: 'lista', locais: 'lista', backupAutomatico: 'bool' };

  L.GRUPOS = [
    { id: 'identificacao', nome: 'Identificação e problema', campos: ['titulo', 'descricao', 'modalidade', 'unidade', 'solicitante', 'sistema', 'criticidade', 'risco', 'pessoas'] },
    { id: 'local', nome: 'Local, setores e lote', campos: ['local', 'localInterno', 'setores', 'responsavelLocal', 'loteId'] },
    { id: 'prazo', nome: 'Datas, prazo estimado e valor estimado', campos: ['dataSolicitacao', 'inicioPrevisto', 'prazoQtd', 'prazoUnidade', 'valorEstimado'] },
    { id: 'contrato', nome: 'Ordem de Serviço, atendimento e observações', campos: ['osNumero', 'atendimentoNumero', 'obs'] }
  ];
  L.ETAPAS = [
    { id: 'inicial', nome: 'Demanda inicial', campos: ['inicialOk', 'inicialData'], obrig: ['inicialData'] },
    { id: 'aprovacao', nome: 'Aprovação', campos: ['aprovacaoOk', 'aprovacaoData', 'aprovacaoObs'], obrig: ['aprovacaoData'], soProgramado: true, opcional: true },
    { id: 'execucao', nome: 'Execução', campos: ['execucaoOk', 'execucaoInicio', 'execucaoFim'], obrig: ['execucaoInicio', 'execucaoFim'] },
    { id: 'ateste', nome: 'Ateste', campos: ['atesteOk', 'atesteData', 'atesteValor'], obrig: ['atesteData', 'atesteValor'] },
    { id: 'pagamento', nome: 'Pagamento', campos: ['pagamentoOk', 'pagamentoData'], obrig: ['pagamentoData'] }
  ];
  L.PERFIS_PADRAO = [
    { nome: 'Gestor Master', verTodas: true, criar: true, editarOutras: true, campos: ['identificacao', 'local', 'prazo', 'contrato'], etapas: true, financeiro: true, areaGestor: true, configuracoes: true, usuarios: true, excluir: true, protegido: true },
    { nome: 'Gestor', verTodas: true, criar: true, editarOutras: true, campos: ['identificacao', 'local', 'prazo', 'contrato'], etapas: true, financeiro: true, areaGestor: true, configuracoes: true, usuarios: false, excluir: true, protegido: true },
    { nome: 'Solicitante', verTodas: false, criar: true, editarOutras: false, campos: ['identificacao', 'local', 'prazo'], etapas: false, financeiro: false, areaGestor: false, configuracoes: false, usuarios: false, excluir: false, protegido: false },
    { nome: 'Contratada (consulta)', verTodas: true, criar: false, editarOutras: false, campos: [], etapas: false, financeiro: false, areaGestor: false, configuracoes: false, usuarios: false, excluir: false, protegido: false },
    { nome: 'Somente leitura', verTodas: true, criar: false, editarOutras: false, campos: [], etapas: false, financeiro: false, areaGestor: true, configuracoes: false, usuarios: false, excluir: false, protegido: false }
  ];
  L.SISTEMAS = ['Hidrossanitário', 'Elétrico', 'Cobertura', 'Alvenaria e revestimentos', 'Esquadrias', 'Climatização',
    'Combate a incêndio', 'Acessibilidade', 'Pavimentação', 'Urbanização'];
  L.CRITICIDADES = ['Baixa', 'Média', 'Alta', 'Crítica'];
  L.MODALIDADES = ['Programado', 'Contínuo'];
  L.UNIDADES_PRAZO = ['dias', 'semanas', 'meses'];
  L.ROTULOS = {};
  Object.keys(L.TABELAS).forEach(function (t) { L.TABELAS[t].colunas.forEach(function (c) { if (t === 'Demandas') L.ROTULOS[c[0]] = c[1]; }); });

  /* ---------------------------------------------------------------------
     Datas
     --------------------------------------------------------------------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  L.pad = pad;
  L.iso = function (d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };
  L.hoje = function () { return L.iso(new Date()); };
  L.parse = function (s) { if (!s) return null; var p = String(s).slice(0, 10).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); };
  L.diff = function (a, b) { return Math.round((L.parse(b) - L.parse(a)) / 86400000); };
  L.addDias = function (s, n) { var d = L.parse(s); d.setDate(d.getDate() + n); return L.iso(d); };
  L.addMeses = function (s, n) {
    var d = L.parse(s), dia = d.getDate();
    d.setDate(1); d.setMonth(d.getMonth() + n);
    var ult = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    d.setDate(Math.min(dia, ult));
    return L.iso(d);
  };
  L.fimPrevisto = function (inicio, qtd, unidade) {
    qtd = Number(qtd) || 0;
    if (!inicio || !qtd) return '';
    if (unidade === 'semanas') return L.addDias(inicio, qtd * 7);
    if (unidade === 'quinzenas') return L.addDias(inicio, qtd * 15);
    if (unidade === 'meses') return L.addMeses(inicio, qtd);
    return L.addDias(inicio, qtd);
  };
  L.brl = function (v) {
    var n = Math.round((Number(v) || 0) * 100) / 100, neg = n < 0, p = Math.abs(n).toFixed(2).split('.');
    return (neg ? '-' : '') + 'R$ ' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1];
  };
  L.prazoTexto = function (qtd, unidade) {
    qtd = Number(qtd) || 0;
    if (!qtd) return '—';
    var sing = { dias: 'dia', semanas: 'semana', quinzenas: 'quinzena', meses: 'mês' };
    return qtd + ' ' + (qtd === 1 ? (sing[unidade] || unidade) : unidade);
  };

  /* ---------------------------------------------------------------------
     SHA-256 (senhas nunca são gravadas em texto)
     --------------------------------------------------------------------- */
  L.sha256 = function (str) {
    str = unescape(encodeURIComponent(str));
    var rr = function (v, a) { return (v >>> a) | (v << (32 - a)); };
    var mp = Math.pow, maxWord = mp(2, 32), result = '', words = [], bitLen = str.length * 8;
    var hash = [], k = [], primeCounter = 0, isComposite = {}, i, j;
    for (var candidate = 2; primeCounter < 64; candidate++) {
      if (!isComposite[candidate]) {
        for (i = 0; i < 313; i += candidate) isComposite[i] = candidate;
        hash[primeCounter] = (mp(candidate, 0.5) * maxWord) | 0;
        k[primeCounter++] = (mp(candidate, 1 / 3) * maxWord) | 0;
      }
    }
    hash = hash.slice(0, 8);
    str += '\x80';
    while (str.length % 64 - 56) str += '\x00';
    for (i = 0; i < str.length; i++) { j = str.charCodeAt(i); words[i >> 2] |= j << ((3 - i) % 4) * 8; }
    words[words.length] = ((bitLen / maxWord) | 0);
    words[words.length] = bitLen;
    for (j = 0; j < words.length;) {
      var w = words.slice(j, j += 16), oldHash = hash;
      hash = hash.slice(0, 8);
      for (i = 0; i < 64; i++) {
        var w15 = w[i - 15], w2 = w[i - 2], a = hash[0], e = hash[4];
        var temp1 = hash[7] + (rr(e, 6) ^ rr(e, 11) ^ rr(e, 25)) + ((e & hash[5]) ^ ((~e) & hash[6])) + k[i] +
          (w[i] = (i < 16) ? w[i] : (w[i - 16] + (rr(w15, 7) ^ rr(w15, 18) ^ (w15 >>> 3)) + w[i - 7] + (rr(w2, 17) ^ rr(w2, 19) ^ (w2 >>> 10))) | 0);
        var temp2 = (rr(a, 2) ^ rr(a, 13) ^ rr(a, 22)) + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
        hash = [(temp1 + temp2) | 0].concat(hash);
        hash[4] = (hash[4] + temp1) | 0;
      }
      for (i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
    }
    for (i = 0; i < 8; i++) for (j = 3; j + 1; j--) { var b = (hash[i] >> (j * 8)) & 255; result += ((b < 16) ? 0 : '') + b.toString(16); }
    return result;
  };
  L.hashSenha = function (senha, salt) { return L.sha256(salt + '|' + senha); };

  /* ---------------------------------------------------------------------
     Normalização de registros
     --------------------------------------------------------------------- */
  L.normalizar = function (tabela, o) {
    var out = {};
    L.TABELAS[tabela].colunas.forEach(function (c) {
      var k = c[0], t = c[2], v = o ? o[k] : undefined;
      if (t === 'num') out[k] = Math.round((Number(v) || 0) * 100) / 100;
      else if (t === 'int') out[k] = parseInt(v, 10) || 0;
      else if (t === 'bool') out[k] = v === true || v === 'TRUE' || v === 'true' || v === 1 || v === 'Sim';
      else if (t === 'lista') out[k] = Array.isArray(v) ? v.map(String).filter(function (x) { return x; }) : String(v || '').split(';').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
      else out[k] = v == null ? '' : String(v).trim();
    });
    return out;
  };
  L.valorTexto = function (v) {
    if (Array.isArray(v)) return v.join('; ');
    if (v === true) return 'Sim';
    if (v === false) return 'Não';
    return v == null ? '' : String(v);
  };

  /* ---------------------------------------------------------------------
     Situação e prazo da demanda
     --------------------------------------------------------------------- */
  L.SITUACOES = [
    { id: 'analise', nome: 'Aguardando análise', cor: '#64748B' },
    { id: 'aprovacao', nome: 'Aguardando aprovação', cor: '#7C3AED' },
    { id: 'aguardando', nome: 'Aguardando execução', cor: '#2563EB' },
    { id: 'execucao', nome: 'Em execução', cor: '#0E7490' },
    { id: 'ateste', nome: 'Aguardando ateste', cor: '#B45309' },
    { id: 'pagamento', nome: 'Aguardando pagamento', cor: '#C2410C' },
    { id: 'paga', nome: 'Paga', cor: '#15803D' },
    { id: 'cancelada', nome: 'Cancelada', cor: '#94A3B8' }
  ];
  L.situacao = function (d) {
    var id;
    if (d.cancelada) id = 'cancelada';
    else if (d.pagamentoOk) id = 'paga';
    else if (d.atesteOk) id = 'pagamento';
    else if (d.execucaoOk) id = 'ateste';
    else if (d.execucaoInicio && d.inicialOk) id = 'execucao';
    else if (!d.inicialOk) id = 'analise';
    else if (d.modalidade === 'Programado' && !d.aprovacaoOk) id = 'aprovacao';
    else id = 'aguardando';
    for (var i = 0; i < L.SITUACOES.length; i++) if (L.SITUACOES[i].id === id) return L.SITUACOES[i];
  };
  L.prazo = function (d, hoje) {
    hoje = hoje || L.hoje();
    if (d.cancelada) return { k: 'neutro', nome: 'Cancelada', dias: 0 };
    if (d.execucaoOk) {
      var atrasou = d.fimPrevisto && d.execucaoFim && d.execucaoFim > d.fimPrevisto;
      return { k: atrasou ? 'concluida-atraso' : 'concluida', nome: atrasou ? 'Executada com atraso' : 'Executada no prazo', dias: atrasou ? L.diff(d.fimPrevisto, d.execucaoFim) : 0 };
    }
    if (!d.fimPrevisto) return { k: 'neutro', nome: 'Sem prazo', dias: 0 };
    var n = L.diff(hoje, d.fimPrevisto);
    if (n < 0) return { k: 'atrasada', nome: 'Em atraso', dias: -n };
    if (n <= 7) return { k: 'atencao', nome: n === 0 ? 'Vence hoje' : 'Vence em ' + n + (n === 1 ? ' dia' : ' dias'), dias: n };
    return { k: 'ok', nome: 'No prazo', dias: n };
  };

  /* ---------------------------------------------------------------------
     Financeiro: valor global, lotes, empenhos e notas fiscais
     --------------------------------------------------------------------- */
  L.financeiro = function (dados) {
    var cfg = dados.config || {}, lotes = dados.lotes || [], emp = dados.empenhos || [], nfs = dados.nfs || [], dem = dados.demandas || [];
    var somaLotes = lotes.reduce(function (a, l) { return a + (Number(l.valor) || 0); }, 0);
    var global = somaLotes > 0 ? somaLotes : (Number(cfg.valorGlobal) || 0);
    var nfTotal = nfs.reduce(function (a, n) { return a + (Number(n.valor) || 0); }, 0);
    var demPorId = {};
    dem.forEach(function (d) { demPorId[d.id] = d; });
    var nfPorDemanda = {};
    nfs.forEach(function (n) { nfPorDemanda[n.demandaId] = (nfPorDemanda[n.demandaId] || 0) + (Number(n.valor) || 0); });
    var limG = Number(cfg.alertaGlobalPct) > 0 ? Number(cfg.alertaGlobalPct) : L.ALERTA_PADRAO;
    var limE = Number(cfg.alertaEmpenhoPct) > 0 ? Number(cfg.alertaEmpenhoPct) : L.ALERTA_PADRAO;
    var item = function (o, lim) { o.saldo = o.total - o.usado; o.pctUsado = o.total ? o.usado / o.total * 100 : 0; o.pctSaldo = o.total ? o.saldo / o.total * 100 : 0; o.limite = lim; o.alerta = o.total > 0 && o.pctUsado >= lim - 1e-9; return o; };

    var porLote = lotes.map(function (l) {
      var usado = nfs.reduce(function (a, n) { var d = demPorId[n.demandaId]; return a + (d && d.loteId === l.id ? (Number(n.valor) || 0) : 0); }, 0);
      return item({ lote: l, total: Number(l.valor) || 0, usado: usado }, limG);
    });
    var novos = emp.filter(function (e) { return e.tipo !== 'Reforço' || !emp.some(function (x) { return x.id === e.vinculadoA && x.tipo !== 'Reforço'; }); });
    var familias = novos.map(function (n) {
      var membros = [n].concat(emp.filter(function (e) { return e.tipo === 'Reforço' && e.vinculadoA === n.id && e.id !== n.id; }));
      var ids = membros.map(function (m) { return m.id; });
      var usado = nfs.reduce(function (a, x) { return a + (ids.indexOf(x.empenhoId) >= 0 ? (Number(x.valor) || 0) : 0); }, 0);
      return item({ empenho: n, membros: membros, ids: ids, total: membros.reduce(function (a, m) { return a + (Number(m.valor) || 0); }, 0), usado: usado }, limE);
    });
    var comprometido = dem.reduce(function (a, d) {
      if (d.cancelada || d.pagamentoOk) return a;
      var base = d.atesteOk && d.atesteValor ? d.atesteValor : d.valorEstimado;
      return a + Math.max(0, (Number(base) || 0) - (nfPorDemanda[d.id] || 0));
    }, 0);
    var g = item({ total: global, usado: nfTotal }, limG);
    return {
      global: g, origemGlobal: somaLotes > 0 ? 'lotes' : 'manual', somaLotes: somaLotes, lotes: porLote, empenhos: familias,
      totalEmpenhado: emp.reduce(function (a, e) { return a + (Number(e.valor) || 0); }, 0),
      comprometido: comprometido, saldoProjetado: g.saldo - comprometido, nfPorDemanda: nfPorDemanda, limiteGlobal: limG, limiteEmpenho: limE,
      alertas: [g.alerta ? { tipo: 'global', nome: 'Valor global da vigência', item: g } : null]
        .concat(porLote.filter(function (x) { return x.alerta; }).map(function (x) { return { tipo: 'lote', nome: 'Lote ' + x.lote.numero, item: x }; }))
        .concat(familias.filter(function (x) { return x.alerta; }).map(function (x) { return { tipo: 'empenho', nome: 'Empenho ' + x.empenho.numero, item: x }; }))
        .filter(function (x) { return x; })
    };
  };
  L.familiaDoEmpenho = function (fin, empenhoId) {
    for (var i = 0; i < fin.empenhos.length; i++) if (fin.empenhos[i].ids.indexOf(empenhoId) >= 0) return fin.empenhos[i];
    return null;
  };

  /* ---------------------------------------------------------------------
     Permissões
     --------------------------------------------------------------------- */
  L.grupoDoCampo = function (campo) {
    for (var i = 0; i < L.GRUPOS.length; i++) if (L.GRUPOS[i].campos.indexOf(campo) >= 0) return L.GRUPOS[i].id;
    for (i = 0; i < L.ETAPAS.length; i++) if (L.ETAPAS[i].campos.indexOf(campo) >= 0) return 'etapa:' + L.ETAPAS[i].id;
    if (campo === 'cancelada' || campo === 'motivoCancelamento') return 'cancelar';
    return null;
  };
  L.etapaDoCampo = function (campo) {
    for (var i = 0; i < L.ETAPAS.length; i++) if (L.ETAPAS[i].campos.indexOf(campo) >= 0) return L.ETAPAS[i];
    return null;
  };
  L.podeEditarCampo = function (perfil, campo, demanda, login) {
    var g = L.grupoDoCampo(campo);
    if (!g) return false;
    if (g.indexOf('etapa:') === 0) return !!perfil.etapas;
    if (g === 'cancelar') return !!perfil.excluir;
    if ((perfil.campos || []).indexOf(g) < 0) return false;
    if (!demanda) return true;
    if (perfil.editarOutras) return true;
    return demanda.criadoPor === login && !demanda.inicialOk;
  };
  L.CAMPOS_EDITAVEIS = [].concat.apply([], L.GRUPOS.map(function (g) { return g.campos; }))
    .concat([].concat.apply([], L.ETAPAS.map(function (e) { return e.campos; }))).concat(['cancelada', 'motivoCancelamento']);

  /* Compara antes/depois e indica o que mudou e se exige senha */
  L.analisarAlteracao = function (antes, depois) {
    var mudancas = [], exigeSenha = false;
    L.CAMPOS_EDITAVEIS.forEach(function (k) {
      var a = L.valorTexto(antes[k]), b = L.valorTexto(depois[k]);
      if (a === b) return;
      mudancas.push({ campo: k, anterior: a, novo: b });
      var et = L.etapaDoCampo(k);
      if (et && antes[et.id + 'Ok']) exigeSenha = true;
      if (antes.pagamentoOk || antes.cancelada) exigeSenha = true;
    });
    return { mudancas: mudancas, exigeSenha: exigeSenha };
  };

  /* Regras de sequência e preenchimento das etapas */
  L.validarEtapas = function (d, qtdNotas) {
    var erros = [];
    var programado = d.modalidade === 'Programado';
    if (!programado && d.aprovacaoOk) erros.push('A etapa de aprovação só se aplica a demandas programadas.');
    if (d.aprovacaoOk && !d.inicialOk) erros.push('Marque a demanda inicial antes da aprovação.');
    if (d.execucaoOk && !d.inicialOk) erros.push('Marque a demanda inicial antes da execução.');
    if (d.atesteOk && !d.execucaoOk) erros.push('Marque a execução antes do ateste.');
    if (d.pagamentoOk && !d.atesteOk) erros.push('Marque o ateste antes do pagamento.');
    if (d.pagamentoOk && !qtdNotas) erros.push('Registre ao menos uma Nota Fiscal antes de marcar o pagamento.');
    L.ETAPAS.forEach(function (e) {
      if (!d[e.id + 'Ok']) return;
      e.obrig.forEach(function (c) { if (!d[c] || (c === 'atesteValor' && !(Number(d[c]) > 0))) erros.push(e.nome + ': preencha ' + L.ROTULOS[c].replace(e.nome + ': ', '') + '.'); });
    });
    if (d.execucaoInicio && d.execucaoFim && d.execucaoFim < d.execucaoInicio) erros.push('O término da execução deve ser igual ou posterior ao início.');
    if (d.cancelada && !d.motivoCancelamento) erros.push('Informe o motivo do cancelamento.');
    return erros;
  };
  L.validarObrigatorios = function (d) {
    var faltam = [];
    [['titulo', 'título'], ['modalidade', 'modalidade'], ['loteId', 'lote'], ['unidade', 'unidade demandante'], ['local', 'local'], ['sistema', 'sistema afetado'],
      ['descricao', 'descrição'], ['responsavelLocal', 'responsável local'], ['dataSolicitacao', 'data da solicitação'], ['prazoUnidade', 'unidade do prazo']]
      .forEach(function (x) { if (!d[x[0]]) faltam.push(x[1]); });
    if (!(Number(d.prazoQtd) > 0)) faltam.push('prazo estimado');
    if (d.modalidade === 'Programado' && !(Number(d.valorEstimado) > 0)) faltam.push('valor estimado (obrigatório para demandas programadas)');
    return faltam;
  };

  return L;
})();


/* =========================================================================
   SERVIÇO (ações da API)
   Recebe um "repositório" (Google Sheets ou navegador) e aplica as regras.
   Toda gravação registra o histórico e devolve os dados atualizados.
   ========================================================================= */
var Servico = (function (L) {
  'use strict';
  var S = {};

  function Erro(codigo, mensagem) { this.codigo = codigo; this.mensagem = mensagem; }
  function falha(codigo, mensagem) { throw new Erro(codigo, mensagem); }

  /* ---------------- Sessão ---------------- */
  function perfilDe(repo, nome) {
    var p = repo.listar('Perfis').filter(function (x) { return x.nome === nome; })[0];
    return p || null;
  }
  function sessao(repo, token) {
    if (!token) falha('SESSAO', 'Sessão expirada. Entre novamente.');
    var login = repo.tokenLer(token);
    if (!login) falha('SESSAO', 'Sessão expirada. Entre novamente.');
    var u = repo.listar('Usuarios').filter(function (x) { return x.login === login; })[0];
    if (!u || !u.ativo) falha('SESSAO', 'Usuário inativo ou removido.');
    var p = perfilDe(repo, u.perfil);
    if (!p) falha('PERFIL', 'O perfil "' + u.perfil + '" não existe mais. Procure o gestor master.');
    return { login: u.login, nome: u.nome, usuario: u, perfil: p };
  }
  function exigir(cond, msg) { if (!cond) falha('PERMISSAO', msg || 'Seu perfil não permite esta ação.'); }
  function conferirSenha(ses, senha) {
    if (!senha) falha('SENHA', 'Confirme a operação com a sua senha.');
    if (L.hashSenha(senha, ses.usuario.salt) !== ses.usuario.hash) falha('SENHA', 'Senha incorreta. A alteração não foi gravada.');
  }
  function hist(repo, ses, acao, demandaId, campo, anterior, novo, motivo) {
    return { em: repo.agora(), usuario: ses.login, perfil: ses.perfil.nome, acao: acao, demandaId: demandaId || '', campo: campo || '', anterior: anterior == null ? '' : String(anterior), novo: novo == null ? '' : String(novo), motivo: motivo || '' };
  }
  function publico(u) { return { login: u.login, nome: u.nome, email: u.email, perfil: u.perfil, ativo: u.ativo, trocarSenha: u.trocarSenha, criadoEm: u.criadoEm }; }

  /* ---------------- Dados visíveis para a sessão ---------------- */
  function montarDados(repo, ses) {
    var p = ses.perfil;
    var cfg = repo.config();
    var demandas = repo.listar('Demandas');
    if (!p.verTodas) demandas = demandas.filter(function (d) { return d.criadoPor === ses.login; });
    var ids = {};
    demandas.forEach(function (d) { ids[d.id] = 1; });
    var nfs = repo.listar('NotasFiscais').filter(function (n) { return ids[n.demandaId]; });
    var verFin = p.financeiro || p.areaGestor;
    var usuarios = repo.listar('Usuarios');
    var nomes = {};
    usuarios.forEach(function (u) { nomes[u.login] = u.nome; });
    var historico = repo.listar('Historico');
    if (!p.areaGestor) historico = historico.filter(function (h) { return ids[h.demandaId]; });
    if (historico.length > 4000) historico = historico.slice(historico.length - 4000);
    if (!p.configuracoes) delete cfg.emailsBackup;
    if (!p.areaGestor) delete cfg.chavePainel;
    return {
      versao: repo.versao(), config: cfg, lotes: repo.listar('Lotes'), empenhos: verFin ? repo.listar('Empenhos') : [],
      demandas: demandas, nfs: nfs, perfis: repo.listar('Perfis'), nomes: nomes,
      usuarios: p.usuarios ? usuarios.map(publico) : [], historico: historico.reverse(),
      sessao: { login: ses.login, nome: ses.nome, perfil: p, trocarSenha: ses.usuario.trocarSenha }
    };
  }
  function dadosCompletos(repo) {
    return { config: repo.config(), lotes: repo.listar('Lotes'), empenhos: repo.listar('Empenhos'), demandas: repo.listar('Demandas'), nfs: repo.listar('NotasFiscais') };
  }

  /* =====================================================================
     AÇÕES
     ===================================================================== */
  var A = {};

  A.versao = function (repo) { return { versao: repo.versao() }; };

  A.login = function (repo, req) {
    var login = String(req.login || '').trim().toLowerCase();
    var u = repo.listar('Usuarios').filter(function (x) { return x.login.toLowerCase() === login; })[0];
    if (!u || L.hashSenha(String(req.senha || ''), u.salt) !== u.hash) falha('LOGIN', 'Login ou senha incorretos.');
    if (!u.ativo) falha('LOGIN', 'Usuário desativado. Procure o gestor master.');
    var token = repo.uuid() + repo.uuid();
    repo.tokenSalvar(token, u.login);
    var ses = sessao(repo, token);
    repo.historico([hist(repo, ses, 'Entrada no sistema')]);
    return { token: token, dados: montarDados(repo, ses) };
  };

  A.sair = function (repo, req) { if (req.token) repo.tokenRemover(req.token); return { ok: true }; };

  A.dados = function (repo, req) { return { dados: montarDados(repo, sessao(repo, req.token)) }; };

  A.trocarSenha = function (repo, req) {
    var ses = sessao(repo, req.token);
    conferirSenha(ses, req.senhaAtual);
    var nova = String(req.novaSenha || '');
    if (nova.length < 6) falha('VALIDACAO', 'A nova senha precisa ter ao menos 6 caracteres.');
    var u = ses.usuario;
    u.salt = repo.uuid();
    u.hash = L.hashSenha(nova, u.salt);
    u.trocarSenha = false;
    repo.salvar('Usuarios', u);
    repo.historico([hist(repo, ses, 'Senha alterada pelo próprio usuário')]);
    return { dados: montarDados(repo, ses) };
  };

  /* ---------------- Demandas ---------------- */
  A.salvarDemanda = function (repo, req) {
    var ses = sessao(repo, req.token), p = ses.perfil;
    var entrada = req.demanda || {};
    var todas = repo.listar('Demandas');
    var antes = entrada.id ? todas.filter(function (x) { return x.id === entrada.id; })[0] : null;
    var registros = [];
    var nfsDem = function (id) { return repo.listar('NotasFiscais').filter(function (n) { return n.demandaId === id; }).length; };
    var lotes = repo.listar('Lotes');
    var d;

    if (!antes) {
      exigir(p.criar, 'Seu perfil não permite registrar demandas.');
      d = L.normalizar('Demandas', {});
      L.CAMPOS_EDITAVEIS.forEach(function (k) { if (entrada[k] !== undefined && L.podeEditarCampo(p, k, null, ses.login)) d[k] = entrada[k]; });
      d = L.normalizar('Demandas', d);
      if (!p.campos || p.campos.indexOf('identificacao') < 0) falha('PERMISSAO', 'Seu perfil não permite preencher os dados da demanda.');
      var faltam = L.validarObrigatorios(d);
      if (faltam.length) falha('VALIDACAO', 'Preencha os campos obrigatórios: ' + faltam.join(', ') + '.');
      if (!lotes.some(function (l) { return l.id === d.loteId; })) falha('VALIDACAO', 'Selecione um lote cadastrado. Toda demanda precisa estar vinculada a um lote.');
      var ano = String(d.dataSolicitacao).slice(0, 4);
      var pre = (d.modalidade === 'Contínuo' ? 'MC' : 'MP') + '-' + ano + '-';
      var nums = todas.filter(function (x) { return x.id.indexOf(pre) === 0; }).map(function (x) { return parseInt(x.id.slice(pre.length), 10) || 0; });
      d.id = pre + ('00' + ((nums.length ? Math.max.apply(null, nums) : 0) + 1)).slice(-3);
      d.criadoPor = ses.login;
      d.criadoEm = repo.agora();
      registros.push(hist(repo, ses, 'Demanda registrada', d.id, '', '', d.titulo, ''));
    } else {
      var visivel = p.verTodas || antes.criadoPor === ses.login;
      exigir(visivel, 'Você não tem acesso a esta demanda.');
      d = L.normalizar('Demandas', antes);
      L.CAMPOS_EDITAVEIS.forEach(function (k) { if (entrada[k] !== undefined) d[k] = entrada[k]; });
      d = L.normalizar('Demandas', d);
      var an = L.analisarAlteracao(antes, d);
      if (!an.mudancas.length) return { semAlteracao: true, dados: montarDados(repo, ses) };
      an.mudancas.forEach(function (m) {
        if (!L.podeEditarCampo(p, m.campo, antes, ses.login)) falha('PERMISSAO', 'Seu perfil não permite alterar o campo "' + (L.ROTULOS[m.campo] || m.campo) + '"' + (!p.editarOutras && antes.inicialOk ? ', pois a demanda já foi recebida pelo gestor' : '') + '.');
      });
      if (an.exigeSenha) {
        conferirSenha(ses, req.senha);
        if (!String(req.motivo || '').trim()) falha('VALIDACAO', 'Informe o motivo da alteração.');
      }
      var falt = L.validarObrigatorios(d);
      if (falt.length) falha('VALIDACAO', 'Preencha os campos obrigatórios: ' + falt.join(', ') + '.');
      if (!lotes.some(function (l) { return l.id === d.loteId; })) falha('VALIDACAO', 'Selecione um lote cadastrado.');
      var etapasMudadas = {};
      an.mudancas.forEach(function (m) {
        var et = L.etapaDoCampo(m.campo);
        if (et && m.campo === et.id + 'Ok') {
          etapasMudadas[et.id] = 1;
          registros.push(hist(repo, ses, (d[m.campo] ? 'Etapa marcada: ' : 'Etapa desmarcada: ') + et.nome, d.id, L.ROTULOS[m.campo], m.anterior, m.novo, req.motivo));
        } else {
          registros.push(hist(repo, ses, an.exigeSenha ? 'Alteração com senha' : 'Demanda alterada', d.id, L.ROTULOS[m.campo] || m.campo, m.anterior, m.novo, req.motivo));
        }
      });
      L.ETAPAS.forEach(function (e) { if (etapasMudadas[e.id]) d[e.id + 'Por'] = d[e.id + 'Ok'] ? ses.login : ''; });
    }
    if (d.inicialOk && !d.inicialData) d.inicialData = L.hoje();
    if (!d.inicioPrevisto) d.inicioPrevisto = d.dataSolicitacao;
    d.fimPrevisto = L.fimPrevisto(d.inicioPrevisto, d.prazoQtd, d.prazoUnidade);
    var erros = L.validarEtapas(d, antes ? nfsDem(d.id) : 0);
    if (erros.length) falha('VALIDACAO', erros.join(' '));
    if (!antes && L.ETAPAS.some(function (e) { return d[e.id + 'Ok']; })) L.ETAPAS.forEach(function (e) { if (d[e.id + 'Ok']) d[e.id + 'Por'] = ses.login; });
    d.atualizadoPor = ses.login;
    d.atualizadoEm = repo.agora();
    repo.salvar('Demandas', d);
    repo.historico(registros);
    return { id: d.id, dados: montarDados(repo, ses) };
  };

  A.excluirDemanda = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.excluir, 'Seu perfil não permite excluir demandas.');
    conferirSenha(ses, req.senha);
    if (!String(req.motivo || '').trim()) falha('VALIDACAO', 'Informe o motivo da exclusão.');
    var d = repo.listar('Demandas').filter(function (x) { return x.id === req.id; })[0];
    if (!d) falha('VALIDACAO', 'Demanda não encontrada.');
    var notas = repo.listar('NotasFiscais').filter(function (n) { return n.demandaId === d.id; });
    var registros = [];
    notas.forEach(function (n) { repo.remover('NotasFiscais', n.id); registros.push(hist(repo, ses, 'Nota Fiscal excluída com a demanda', d.id, 'NF ' + n.numero, n.valor, '', req.motivo)); });
    repo.remover('Demandas', d.id);
    registros.push(hist(repo, ses, 'Demanda excluída', d.id, d.titulo, L.situacao(d).nome + (notas.length ? ' · ' + notas.length + ' nota(s) fiscal(is)' : ''), '', req.motivo));
    repo.historico(registros);
    return { notasExcluidas: notas.length, dados: montarDados(repo, ses) };
  };

  /* ---------------- Notas fiscais ---------------- */
  A.salvarNF = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro, 'Seu perfil não permite lançar Notas Fiscais.');
    var nf = L.normalizar('NotasFiscais', req.nf || {});
    var todas = repo.listar('NotasFiscais');
    var antes = nf.id ? todas.filter(function (x) { return x.id === nf.id; })[0] : null;
    var d = repo.listar('Demandas').filter(function (x) { return x.id === nf.demandaId; })[0];
    if (!d) falha('VALIDACAO', 'Demanda não encontrada.');
    if (!d.atesteOk) falha('VALIDACAO', 'A Nota Fiscal só pode ser registrada depois do ateste.');
    if (!nf.numero) falha('VALIDACAO', 'Informe o número da Nota Fiscal.');
    if (!(nf.valor > 0)) falha('VALIDACAO', 'Informe o valor do serviço da Nota Fiscal.');
    if (!nf.data) falha('VALIDACAO', 'Informe a data de emissão.');
    if (!nf.empenhoId) falha('VALIDACAO', 'Vincule a Nota Fiscal a um empenho.');
    var base = dadosCompletos(repo);
    if (antes) base.nfs = base.nfs.filter(function (x) { return x.id !== antes.id; });
    var fin = L.financeiro(base);
    var fam = L.familiaDoEmpenho(fin, nf.empenhoId);
    if (!fam) falha('VALIDACAO', 'Empenho não encontrado.');
    if (nf.valor > fam.saldo + 0.004) falha('SALDO', 'O empenho ' + fam.empenho.numero + ' tem saldo de ' + L.brl(fam.saldo) + ', insuficiente para esta nota. Cadastre um reforço de empenho.');
    if (antes || d.pagamentoOk) {
      conferirSenha(ses, req.senha);
      if (!String(req.motivo || '').trim()) falha('VALIDACAO', 'Informe o motivo da alteração.');
    }
    var registros = [];
    if (antes) {
      ['numero', 'valor', 'data', 'empenhoId', 'obs'].forEach(function (k) {
        if (String(antes[k]) !== String(nf[k])) registros.push(hist(repo, ses, 'Nota Fiscal alterada', d.id, 'NF ' + antes.numero + ': ' + L.TABELAS.NotasFiscais.colunas.filter(function (c) { return c[0] === k; })[0][1], antes[k], nf[k], req.motivo));
      });
      nf.registradoPor = antes.registradoPor; nf.registradoEm = antes.registradoEm;
    } else {
      nf.id = 'NF' + repo.uuid().slice(0, 8).toUpperCase();
      nf.registradoPor = ses.login; nf.registradoEm = repo.agora();
      registros.push(hist(repo, ses, 'Nota Fiscal registrada', d.id, 'NF ' + nf.numero, '', nf.valor, req.motivo));
    }
    repo.salvar('NotasFiscais', nf);
    repo.historico(registros);
    return { dados: montarDados(repo, ses) };
  };
  A.excluirNF = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro, 'Seu perfil não permite excluir Notas Fiscais.');
    conferirSenha(ses, req.senha);
    if (!String(req.motivo || '').trim()) falha('VALIDACAO', 'Informe o motivo da exclusão.');
    var nfs = repo.listar('NotasFiscais'), nf = nfs.filter(function (x) { return x.id === req.id; })[0];
    if (!nf) falha('VALIDACAO', 'Nota Fiscal não encontrada.');
    var d = repo.listar('Demandas').filter(function (x) { return x.id === nf.demandaId; })[0];
    if (d && d.pagamentoOk && nfs.filter(function (x) { return x.demandaId === d.id; }).length === 1) falha('VALIDACAO', 'Esta é a única nota de uma demanda paga. Desmarque o pagamento antes de excluí-la.');
    repo.remover('NotasFiscais', nf.id);
    repo.historico([hist(repo, ses, 'Nota Fiscal excluída', nf.demandaId, 'NF ' + nf.numero, nf.valor, '', req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };

  /* ---------------- Empenhos e lotes ---------------- */
  A.salvarEmpenho = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro, 'Seu perfil não permite cadastrar empenhos.');
    var e = L.normalizar('Empenhos', req.empenho || {});
    var todos = repo.listar('Empenhos');
    var antes = e.id ? todos.filter(function (x) { return x.id === e.id; })[0] : null;
    if (!e.numero) falha('VALIDACAO', 'Informe o número do empenho.');
    if (!(e.valor > 0)) falha('VALIDACAO', 'Informe o valor do empenho.');
    if (e.tipo !== 'Novo' && e.tipo !== 'Reforço') falha('VALIDACAO', 'Indique se o empenho é novo ou reforço.');
    if (e.tipo === 'Reforço') {
      var orig = todos.filter(function (x) { return x.id === e.vinculadoA && x.tipo === 'Novo'; })[0];
      if (!orig) falha('VALIDACAO', 'Selecione o empenho novo ao qual este reforço está vinculado.');
    } else e.vinculadoA = '';
    if (todos.some(function (x) { return x.numero === e.numero && x.id !== e.id; })) falha('VALIDACAO', 'Já existe um empenho com o número ' + e.numero + '.');
    if (antes) {
      conferirSenha(ses, req.senha);
      var base = dadosCompletos(repo);
      base.empenhos = base.empenhos.map(function (x) { return x.id === e.id ? e : x; });
      var fam = L.familiaDoEmpenho(L.financeiro(base), e.id);
      if (fam && fam.saldo < -0.004) falha('SALDO', 'Com esse valor, o empenho ficaria com saldo negativo diante das notas já lançadas.');
      if (antes.tipo === 'Novo' && e.tipo === 'Reforço' && todos.some(function (x) { return x.vinculadoA === e.id; })) falha('VALIDACAO', 'Este empenho tem reforços vinculados e não pode virar reforço.');
    } else e.id = 'EMP' + repo.uuid().slice(0, 8).toUpperCase();
    repo.salvar('Empenhos', e);
    repo.historico([hist(repo, ses, antes ? 'Empenho alterado' : 'Empenho cadastrado (' + e.tipo + ')', '', e.numero, antes ? antes.valor : '', e.valor, req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };
  A.excluirEmpenho = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro, 'Seu perfil não permite excluir empenhos.');
    conferirSenha(ses, req.senha);
    var e = repo.listar('Empenhos').filter(function (x) { return x.id === req.id; })[0];
    if (!e) falha('VALIDACAO', 'Empenho não encontrado.');
    if (repo.listar('NotasFiscais').some(function (n) { return n.empenhoId === e.id; })) falha('VALIDACAO', 'O empenho tem Notas Fiscais vinculadas e não pode ser excluído.');
    if (repo.listar('Empenhos').some(function (x) { return x.vinculadoA === e.id; })) falha('VALIDACAO', 'Exclua antes os reforços vinculados a este empenho.');
    repo.remover('Empenhos', e.id);
    repo.historico([hist(repo, ses, 'Empenho excluído', '', e.numero, e.valor, '', req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };
  A.salvarLote = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro || ses.perfil.configuracoes, 'Seu perfil não permite cadastrar lotes.');
    var l = L.normalizar('Lotes', req.lote || {});
    var todos = repo.listar('Lotes');
    var antes = l.id ? todos.filter(function (x) { return x.id === l.id; })[0] : null;
    if (!l.numero) falha('VALIDACAO', 'Informe o número do lote.');
    if (!l.descricao) falha('VALIDACAO', 'Informe a descrição do lote.');
    if (todos.some(function (x) { return x.numero === l.numero && x.id !== l.id; })) falha('VALIDACAO', 'Já existe o lote ' + l.numero + '.');
    if (antes) {
      conferirSenha(ses, req.senha);
      var base = dadosCompletos(repo);
      base.lotes = base.lotes.map(function (x) { return x.id === l.id ? l : x; });
      var it = L.financeiro(base).lotes.filter(function (x) { return x.lote.id === l.id; })[0];
      if (l.valor > 0 && it && it.saldo < -0.004) falha('SALDO', 'O valor informado é menor que as notas já lançadas neste lote.');
    } else l.id = 'LOTE' + repo.uuid().slice(0, 6).toUpperCase();
    repo.salvar('Lotes', l);
    repo.historico([hist(repo, ses, antes ? 'Lote alterado' : 'Lote cadastrado', '', 'Lote ' + l.numero, antes ? antes.descricao + ' / ' + antes.valor : '', l.descricao + ' / ' + l.valor, req.motivo)]);
    return { id: l.id, dados: montarDados(repo, ses) };
  };
  A.excluirLote = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.financeiro || ses.perfil.configuracoes, 'Seu perfil não permite excluir lotes.');
    conferirSenha(ses, req.senha);
    var l = repo.listar('Lotes').filter(function (x) { return x.id === req.id; })[0];
    if (!l) falha('VALIDACAO', 'Lote não encontrado.');
    if (repo.listar('Demandas').some(function (d) { return d.loteId === l.id; })) falha('VALIDACAO', 'Há demandas vinculadas a este lote. Mude o lote dessas demandas antes de excluí-lo.');
    repo.remover('Lotes', l.id);
    repo.historico([hist(repo, ses, 'Lote excluído', '', 'Lote ' + l.numero, l.descricao, '', req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };

  /* ---------------- Configurações ---------------- */
  A.salvarConfig = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.configuracoes, 'Seu perfil não permite alterar as configurações.');
    var atual = repo.config(), novo = req.config || {}, registros = [];
    var permitidas = ['orgao', 'contratoGestao', 'contratoExecucao', 'numeroContrato', 'ataNumero', 'ataAno', 'contratada', 'vigenciaInicio', 'vigenciaFim',
      'valorGlobal', 'alertaGlobalPct', 'alertaEmpenhoPct', 'setores', 'emailsBackup', 'backupAutomatico', 'lookerRelatorio'];
    var muda = permitidas.filter(function (k) { return novo[k] !== undefined && L.valorTexto(novo[k]) !== L.valorTexto(atual[k]); });
    if (!muda.length) return { semAlteracao: true, dados: montarDados(repo, ses) };
    var protegidos = ['valorGlobal', 'alertaGlobalPct', 'alertaEmpenhoPct'].filter(function (k) { return muda.indexOf(k) >= 0 && atual[k] !== '' && Number(atual[k]) > 0; });
    if (protegidos.length) {
      conferirSenha(ses, req.senha);
      if (!String(req.motivo || '').trim()) falha('VALIDACAO', 'Informe o motivo da alteração.');
    }
    muda.forEach(function (k) {
      var v = novo[k];
      if (L.CONFIG_TIPOS[k] === 'num') v = Number(v) || 0;
      if (L.CONFIG_TIPOS[k] === 'numOpc') v = v === '' || v == null ? '' : Number(v) || 0;
      if (L.CONFIG_TIPOS[k] === 'bool') v = v === true || v === 'true';
      if (L.CONFIG_TIPOS[k] === 'lista') v = (Array.isArray(v) ? v : String(v).split(/\n|;/)).map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
      registros.push(hist(repo, ses, 'Configuração alterada', '', L.ROTULOS_CONFIG[k] || k, L.valorTexto(atual[k]), L.valorTexto(v), req.motivo));
      atual[k] = v;
    });
    ['alertaGlobalPct', 'alertaEmpenhoPct'].forEach(function (k) { if (atual[k] !== '' && (Number(atual[k]) < 1 || Number(atual[k]) > 100)) falha('VALIDACAO', 'O percentual de alerta deve ficar entre 1 e 100.'); });
    if (String(atual.ataAno || '') && !/^\d{4}$/.test(String(atual.ataAno))) falha('VALIDACAO', 'Informe o ano da ata com quatro dígitos.');
    repo.salvarConfig(atual);
    repo.historico(registros);
    return { dados: montarDados(repo, ses) };
  };

  /* ---------------- Usuários e perfis ---------------- */
  A.salvarUsuario = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.usuarios, 'Somente o gestor master gerencia usuários.');
    var e = req.usuario || {};
    var login = String(e.login || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,30}$/.test(login)) falha('VALIDACAO', 'Login inválido: use de 3 a 30 letras minúsculas, números, ponto, hífen ou sublinhado.');
    var todos = repo.listar('Usuarios'), antes = todos.filter(function (x) { return x.login === login; })[0];
    if (req.novo && antes) falha('VALIDACAO', 'Já existe o login ' + login + '.');
    if (!perfilDe(repo, e.perfil)) falha('VALIDACAO', 'Selecione um perfil válido.');
    if (!String(e.nome || '').trim()) falha('VALIDACAO', 'Informe o nome do usuário.');
    var u = antes ? L.normalizar('Usuarios', antes) : L.normalizar('Usuarios', { login: login, ativo: true, criadoEm: repo.agora() });
    u.nome = String(e.nome).trim(); u.email = String(e.email || '').trim(); u.perfil = e.perfil; u.ativo = e.ativo !== false;
    if (antes && login === ses.login && (!u.ativo || u.perfil !== antes.perfil)) falha('VALIDACAO', 'Você não pode desativar nem mudar o perfil do próprio usuário.');
    if (antes && antes.perfil === 'Gestor Master' && (u.perfil !== 'Gestor Master' || !u.ativo)) {
      var masters = todos.filter(function (x) { return x.perfil === 'Gestor Master' && x.ativo; });
      if (masters.length <= 1) falha('VALIDACAO', 'É preciso manter ao menos um Gestor Master ativo.');
    }
    if (!antes || e.senha) {
      if (String(e.senha || '').length < 6) falha('VALIDACAO', 'Defina uma senha provisória com ao menos 6 caracteres.');
      u.salt = repo.uuid(); u.hash = L.hashSenha(String(e.senha), u.salt); u.trocarSenha = true;
    }
    repo.salvar('Usuarios', u);
    repo.historico([hist(repo, ses, antes ? (e.senha ? 'Usuário alterado e senha redefinida' : 'Usuário alterado') : 'Usuário criado', '', login, antes ? antes.perfil + (antes.ativo ? '' : ' (inativo)') : '', u.perfil + (u.ativo ? '' : ' (inativo)'), '')]);
    return { dados: montarDados(repo, ses) };
  };
  A.salvarPerfil = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.usuarios, 'Somente o gestor master gerencia perfis.');
    var p = L.normalizar('Perfis', req.perfil || {});
    var nomeAnt = String(req.nomeAnterior || '');
    if (!p.nome) falha('VALIDACAO', 'Informe o nome do perfil.');
    var todos = repo.listar('Perfis'), antes = nomeAnt ? todos.filter(function (x) { return x.nome === nomeAnt; })[0] : null;
    if (antes && antes.nome === 'Gestor Master') falha('VALIDACAO', 'O perfil Gestor Master não pode ser alterado.');
    if (todos.some(function (x) { return x.nome === p.nome && (!antes || x.nome !== antes.nome); })) falha('VALIDACAO', 'Já existe o perfil ' + p.nome + '.');
    p.protegido = antes ? antes.protegido : false;
    if (antes && antes.protegido && p.nome !== antes.nome) falha('VALIDACAO', 'Perfis do sistema não podem ser renomeados.');
    if (antes && p.nome !== antes.nome) {
      repo.remover('Perfis', antes.nome);
      repo.listar('Usuarios').filter(function (u) { return u.perfil === antes.nome; }).forEach(function (u) { u.perfil = p.nome; repo.salvar('Usuarios', u); });
    }
    repo.salvar('Perfis', p);
    repo.historico([hist(repo, ses, antes ? 'Perfil alterado' : 'Perfil criado', '', p.nome, '', JSON.stringify(p), '')]);
    return { dados: montarDados(repo, ses) };
  };
  A.excluirPerfil = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.usuarios, 'Somente o gestor master gerencia perfis.');
    var p = perfilDe(repo, req.nome);
    if (!p) falha('VALIDACAO', 'Perfil não encontrado.');
    if (p.protegido) falha('VALIDACAO', 'Perfis do sistema não podem ser excluídos.');
    if (repo.listar('Usuarios').some(function (u) { return u.perfil === p.nome; })) falha('VALIDACAO', 'Há usuários com este perfil. Mude o perfil deles antes.');
    repo.remover('Perfis', p.nome);
    repo.historico([hist(repo, ses, 'Perfil excluído', '', p.nome, '', '', '')]);
    return { dados: montarDados(repo, ses) };
  };

  /* ---------------- Backup e restauração ---------------- */
  A.exportar = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.areaGestor || ses.perfil.configuracoes, 'Seu perfil não permite exportar a base.');
    var b = dadosCompletos(repo);
    b.historico = repo.listar('Historico');
    b.geradoEm = repo.agora();
    return { backup: b };
  };
  A.importar = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.usuarios, 'Somente o gestor master restaura backups.');
    conferirSenha(ses, req.senha);
    var b = req.backup || {};
    if (!Array.isArray(b.demandas)) falha('VALIDACAO', 'Arquivo inválido: falta a lista de demandas.');
    repo.substituir('Lotes', (b.lotes || []).map(function (x) { return L.normalizar('Lotes', x); }));
    repo.substituir('Empenhos', (b.empenhos || []).map(function (x) { return L.normalizar('Empenhos', x); }));
    repo.substituir('Demandas', b.demandas.map(function (x) { return L.normalizar('Demandas', x); }));
    repo.substituir('NotasFiscais', (b.nfs || []).map(function (x) { return L.normalizar('NotasFiscais', x); }));
    if (b.config) { var c = repo.config(); Object.keys(L.CONFIG_PADRAO).forEach(function (k) { if (b.config[k] !== undefined && k !== 'ultimoBackup' && k !== 'chavePainel') c[k] = b.config[k]; }); repo.salvarConfig(c); }
    repo.historico([hist(repo, ses, 'Backup restaurado', '', '', '', b.demandas.length + ' demandas', req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };
  A.backupAgora = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.configuracoes, 'Seu perfil não permite enviar backups.');
    if (!repo.enviarBackup) falha('INDISPONIVEL', 'O envio por e-mail funciona apenas com a planilha Google conectada.');
    var r = repo.enviarBackup('manual (' + ses.login + ')');
    repo.historico([hist(repo, ses, 'Backup enviado por e-mail', '', '', '', r.destinatarios, '')]);
    return { mensagem: 'Backup enviado para ' + r.destinatarios + '.', dados: montarDados(repo, ses) };
  };

  /* ---------------- Link do painel (somente leitura) para o Looker Studio ---------------- */
  var PERFIL_PAINEL = { nome: 'Painel somente leitura', verTodas: true, criar: false, editarOutras: false, campos: [], etapas: false, financeiro: false, areaGestor: true, configuracoes: false, usuarios: false, excluir: false, protegido: true };
  A.gerarChavePainel = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.areaGestor && (ses.perfil.configuracoes || ses.perfil.usuarios), 'Somente gestores geram o link do painel.');
    conferirSenha(ses, req.senha);
    var c = repo.config(), havia = !!c.chavePainel;
    c.chavePainel = repo.uuid() + repo.uuid().slice(0, 8);
    repo.salvarConfig(c);
    repo.historico([hist(repo, ses, havia ? 'Link do painel refeito (o anterior foi revogado)' : 'Link do painel para o Looker Studio gerado', '', '', '', '', req.motivo)]);
    return { chave: c.chavePainel, dados: montarDados(repo, ses) };
  };
  A.revogarChavePainel = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.areaGestor && (ses.perfil.configuracoes || ses.perfil.usuarios), 'Somente gestores revogam o link do painel.');
    conferirSenha(ses, req.senha);
    var c = repo.config(); c.chavePainel = ''; repo.salvarConfig(c);
    repo.historico([hist(repo, ses, 'Link do painel revogado', '', '', '', '', req.motivo)]);
    return { dados: montarDados(repo, ses) };
  };
  A.painelPublico = function (repo, req) {
    var c = repo.config();
    if (!c.chavePainel || String(req.chave || '') !== c.chavePainel) falha('PAINEL', 'Link do painel inválido ou revogado. Peça um novo link ao gestor.');
    var pub = {};
    ['orgao', 'contratoGestao', 'contratoExecucao', 'numeroContrato', 'ataNumero', 'ataAno', 'contratada', 'vigenciaInicio', 'vigenciaFim', 'valorGlobal', 'alertaGlobalPct', 'alertaEmpenhoPct', 'setores', 'locais'].forEach(function (k) { pub[k] = c[k]; });
    return { dados: {
      versao: repo.versao(), config: pub, lotes: repo.listar('Lotes'), empenhos: repo.listar('Empenhos'),
      demandas: repo.listar('Demandas').map(function (d) { d.criadoPor = ''; d.atualizadoPor = ''; d.inicialPor = d.aprovacaoPor = d.execucaoPor = d.atestePor = d.pagamentoPor = ''; return d; }),
      nfs: repo.listar('NotasFiscais').map(function (n) { n.registradoPor = ''; return n; }), perfis: [], nomes: {}, usuarios: [], historico: [],
      sessao: { login: 'painel', nome: 'Painel somente leitura', perfil: PERFIL_PAINEL, trocarSenha: false }
    } };
  };

  A.looker = function (repo, req) {
    var ses = sessao(repo, req.token);
    exigir(ses.perfil.areaGestor, 'Seu perfil não acessa os dashboards.');
    if (!repo.looker) falha('INDISPONIVEL', 'O Looker Studio usa a planilha Google: conecte o sistema à planilha para gerar o link.');
    return { looker: repo.looker(), relatorio: repo.config().lookerRelatorio || '' };
  };

  /* =====================================================================
     Ponto de entrada
     ===================================================================== */
  var GRAVAM = { gerarChavePainel: 1, revogarChavePainel: 1, trocarSenha: 1, salvarDemanda: 1, excluirDemanda: 1, salvarNF: 1, excluirNF: 1, salvarEmpenho: 1, excluirEmpenho: 1, salvarLote: 1, excluirLote: 1, salvarConfig: 1, salvarUsuario: 1, salvarPerfil: 1, excluirPerfil: 1, importar: 1, backupAgora: 1 };
  S.executar = function (repo, req) {
    try {
      var fn = A[req && req.acao];
      if (!fn) falha('ACAO', 'Ação desconhecida.');
      var exec = function () {
        var r = fn(repo, req);
        if (GRAVAM[req.acao]) {
          repo.marcarVersao();
          if (r.dados) r.dados.versao = repo.versao();
          if (repo.aposGravar && !r.semAlteracao) { try { repo.aposGravar(); } catch (e) { /* a base do Looker é refeita na próxima gravação */ } }
        }
        return r;
      };
      var r = GRAVAM[req.acao] && repo.travar ? repo.travar(exec) : exec();
      r.ok = true;
      return r;
    } catch (e) {
      if (e instanceof Erro) return { ok: false, codigo: e.codigo, mensagem: e.mensagem };
      return { ok: false, codigo: 'INTERNO', mensagem: 'Erro interno: ' + (e && e.message ? e.message : String(e)) };
    }
  };
  return S;
})(Logica);
