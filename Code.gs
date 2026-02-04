/**************** CONFIG ****************/
const SHEET_ADMIN = 'admin';
const SHEET_DADOS = 'dados';

const REPORT_TEMPLATE_ID = '1vRa8lccjBdlpUELeXqbF3DV4FzKiwGl2kN8ed-nZhuQ';
const REPORT_FOLDER_ID   = '1URzRxSVUsM449WM30lqfwpn13Y9br3To';

/**
 * Aba dados (A..S):
 * A Congregação
 * B Área
 * C Data Inicial
 * D Data Final
 * E Mês/Ano
 * F Dízimo
 * G Ofertas
 * H Ofertas EBD
 * I Oferta Especial
 * J Total Entrada
 * K Saldo Anterior
 * L Percentual Despesa (25% de J)
 * M Despesa Congregação
 * N Supervisor (25% de J)
 * O Transporte Dirigente
 * P Despesa Total (M+O+N)
 * Q Entrada Pix
 * R Saldo Atual ((M+O) - L + K)  // conforme você pediu
 * S PDF URL
 */
const COL = {
  CONG: 1,
  AREA: 2,
  DATA_INI: 3,
  DATA_FIM: 4,
  MES_ANO: 5,
  DIZIMO: 6,
  OFERTAS: 7,
  OFERTAS_EBD: 8,
  OFERTA_ESP: 9,
  TOTAL_ENTRADA: 10,
  SALDO_ANT: 11,
  PERC_DESP: 12,
  DESP_CONG: 13,
  SUPERVISOR: 14,
  TRANSP_DIR: 15,
  DESP_TOTAL: 16,
  PIX: 17,
  SALDO_ATUAL: 18,
  PDF_URL: 19,
};

const MONEY_COLS = [6,7,8,9,10,11,12,13,14,15,16,17,18];

/**************** WEB APP ****************/
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Prestação de Contas - Igreja')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**************** SAFE WRAPPER ****************/
function safe_(fn) {
  try {
    return fn();
  } catch (e) {
    return {
      ok: false,
      message: String(e && e.message ? e.message : e),
      stack: String(e && e.stack ? e.stack : ''),
    };
  }
}

/**************** HELPERS ****************/
function norm_(v){ return String(v || '').trim(); }
function num_(v){ const n = Number(v); return isNaN(n) ? 0 : n; }

function toDate_(iso) {
  const s = String(iso || '').trim();
  const parts = s.split('-');
  if (parts.length !== 3) throw new Error('Data inválida.');
  const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
  const dt = new Date(y, m - 1, d);
  if (isNaN(dt.getTime())) throw new Error('Data inválida.');
  return dt;
}

function pad2_(n){ return (n < 10 ? '0' : '') + n; }

function parseMesAnoKey_(mesAno) {
  const s = String(mesAno || '').trim();
  const parts = s.split('/');
  if (parts.length !== 2) return null;
  const mm = Number(parts[0]);
  const yy = Number(parts[1]);
  if (!mm || !yy || mm < 1 || mm > 12) return null;
  return yy * 100 + mm;
}

function formatMoneyColumns_(sh) {
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  MONEY_COLS.forEach(c => sh.getRange(2, c, rows, 1).setNumberFormat('R$ #,##0.00'));
}

/**************** AUTH (CACHE TOKEN) ****************/
function api_login(payload) {
  return safe_(() => {
    ensureSheets_();
    const login = norm_(payload?.login);
    const senha = norm_(payload?.senha);
    if (!login || !senha) return { ok:false, message:'Informe login e senha.' };

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ADMIN);
    if (!sh) return { ok:false, message:'Aba admin não encontrada.' };

    const lastRow = sh.getLastRow();
    if (lastRow < 1) return { ok:false, message:'Aba admin está vazia.' };

    // admin: A login, B senha, C congregações (CSV), D área
    const values = sh.getRange(1, 1, lastRow, 4).getValues();
    let row = null;

    for (const r of values) {
      if (norm_(r[0]) === login && norm_(r[1]) === senha) { row = r; break; }
    }
    if (!row) return { ok:false, message:'Login ou senha inválidos.' };

    const rawCong = norm_(row[2]);
    const rawArea = norm_(row[3]);
    if (!rawCong) return { ok:false, message:'Usuário sem congregação na coluna C.' };
    if (!rawArea) return { ok:false, message:'Usuário sem área na coluna D.' };

    const congregacoes = rawCong.split(',').map(s => s.trim()).filter(Boolean);
    const profile = { login, area: rawArea, congregacoes };

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('auth:' + token, JSON.stringify(profile), 6 * 60 * 60);

    return { ok:true, token };
  });
}

function checkAuth_(token) {
  token = norm_(token);
  if (!token) return { ok:false };
  const raw = CacheService.getScriptCache().get('auth:' + token);
  if (!raw) return { ok:false };

  let profile;
  try { profile = JSON.parse(raw); } catch(e) { return { ok:false }; }
  if (!profile || !profile.login || !profile.area || !Array.isArray(profile.congregacoes)) return { ok:false };
  return { ok:true, profile };
}

function api_getProfile(payload) {
  return safe_(() => {
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };
    return { ok:true, profile: auth.profile };
  });
}

/**************** VALIDATION + CALC ****************/
function validatePayload_(p, profile) {
  if (!p) throw new Error('Dados do formulário vazios.');

  const required = [
    ['congregacao','Congregação'],
    ['area','Área'],
    ['dataInicial','Data Inicial'],
    ['dataFinal','Data Final'],
    ['mesAno','Mês/Ano'],
    ['dizimo','Dízimo'],
    ['ofertas','Ofertas'],
    ['ofertasEbd','Ofertas EBD'],
    ['saldoAnterior','Saldo anterior'],
    ['despesaCongregacao','Despesa da congregação'],
    ['entradaPix','Entrada em Pix'],
  ];
  for (const [k, label] of required) {
    const v = p[k];
    if (v === null || v === undefined || String(v).trim() === '') {
      throw new Error(`Preencha o campo obrigatório: ${label}.`);
    }
  }

  p.congregacao = norm_(p.congregacao);
  p.area = norm_(p.area);
  p.mesAno = norm_(p.mesAno);

  if (!/^(0[1-9]|1[0-2])\/\d{4}$/.test(p.mesAno)) {
    throw new Error('Mês/Ano inválido. Use o formato MM/AAAA.');
  }

  const allowedCongLower = profile.congregacoes.map(c => c.toLowerCase());
  if (!allowedCongLower.includes(p.congregacao.toLowerCase())) {
    throw new Error('Congregação inválida para este usuário.');
  }
  if (p.area !== String(profile.area).trim()) {
    throw new Error('Área inválida para este usuário.');
  }

  const dataInicial = toDate_(p.dataInicial);
  const dataFinal = toDate_(p.dataFinal);
  if (dataFinal.getTime() < dataInicial.getTime()) {
    throw new Error('Data Final deve ser igual ou posterior à Data Inicial.');
  }

  const moneyKeys = [
    ['dizimo','Dízimo'],
    ['ofertas','Ofertas'],
    ['ofertasEbd','Ofertas EBD'],
    ['saldoAnterior','Saldo anterior'],
    ['despesaCongregacao','Despesa da congregação'],
    ['entradaPix','Entrada em Pix'],
  ];
  for (const [k,label] of moneyKeys) {
    if (isNaN(Number(p[k]))) throw new Error(`Valor monetário inválido em: ${label}.`);
  }

  if (p.ofertaEspecial !== undefined && String(p.ofertaEspecial).trim() !== '' && isNaN(Number(p.ofertaEspecial))) {
    throw new Error('Valor monetário inválido em: Oferta especial.');
  }
  if (p.transporteDirigente !== undefined && String(p.transporteDirigente).trim() !== '' && isNaN(Number(p.transporteDirigente))) {
    throw new Error('Valor monetário inválido em: Transporte do dirigente.');
  }
}

function calcFields_(payload) {
  const dizimo = Number(payload.dizimo);
  const ofertas = Number(payload.ofertas);
  const ofertasEbd = Number(payload.ofertasEbd);
  const ofertaEspecial = Number(payload.ofertaEspecial || 0);

  const saldoAnterior = Number(payload.saldoAnterior);
  const despesaCong = Number(payload.despesaCongregacao);
  const transpDir = Number(payload.transporteDirigente || 0);
  const pix = Number(payload.entradaPix);

  const totalEntrada = dizimo + ofertas + ofertasEbd + ofertaEspecial; // J
  const percentualDespesa = totalEntrada * 0.25;                      // L
  const supervisor = totalEntrada * 0.25;                             // N
  const despesaTotal = despesaCong + transpDir + supervisor;          // P

  // conforme sua regra literal:
  const saldoAtual = (despesaCong + transpDir) - percentualDespesa + saldoAnterior; // R

  // para o PDF (repasse templo central) - inferência operacional:
  const repasseTemploCentral = totalEntrada - percentualDespesa - supervisor;

  return { dizimo, ofertas, ofertasEbd, ofertaEspecial, saldoAnterior, despesaCong, transpDir, pix,
           totalEntrada, percentualDespesa, supervisor, despesaTotal, saldoAtual, repasseTemploCentral };
}

/**************** CRUD ****************/
function api_saveLancamento(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const form = payload?.form;
    validatePayload_(form, auth.profile);

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const c = calcFields_(form);

    const row = new Array(19).fill('');

    row[COL.CONG-1] = form.congregacao;
    row[COL.AREA-1] = form.area;
    row[COL.DATA_INI-1] = toDate_(form.dataInicial);
    row[COL.DATA_FIM-1] = toDate_(form.dataFinal);
    row[COL.MES_ANO-1] = form.mesAno;

    row[COL.DIZIMO-1] = c.dizimo;
    row[COL.OFERTAS-1] = c.ofertas;
    row[COL.OFERTAS_EBD-1] = c.ofertasEbd;
    row[COL.OFERTA_ESP-1] = c.ofertaEspecial;

    row[COL.TOTAL_ENTRADA-1] = c.totalEntrada;
    row[COL.SALDO_ANT-1] = c.saldoAnterior;
    row[COL.PERC_DESP-1] = c.percentualDespesa;
    row[COL.DESP_CONG-1] = c.despesaCong;
    row[COL.SUPERVISOR-1] = c.supervisor;
    row[COL.TRANSP_DIR-1] = c.transpDir;
    row[COL.DESP_TOTAL-1] = c.despesaTotal;
    row[COL.PIX-1] = c.pix;
    row[COL.SALDO_ATUAL-1] = c.saldoAtual;
    row[COL.PDF_URL-1] = '';

    sh.appendRow(row);
    formatMoneyColumns_(sh);

    return { ok:true, message:'Lançamento salvo com sucesso.' };
  });
}

function api_listLancamentos(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const filterCong = norm_(payload?.congregacaoFilter);
    if (filterCong && !allowed.includes(filterCong.toLowerCase())) {
      return { ok:false, message:'Filtro de congregação inválido.' };
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, items: [] };

    const lim = Math.min(Math.max(Number(payload?.limit || 60), 1), 200);

    const data = sh.getRange(2, 1, lastRow - 1, 19).getValues();
    const items = [];

    for (let i=0; i<data.length; i++) {
      const r = data[i];
      const rowIndex = i + 2;

      const cong = norm_(r[COL.CONG-1]);
      const rowArea = norm_(r[COL.AREA-1]);
      if (!cong || !rowArea) continue;
      if (!allowed.includes(cong.toLowerCase())) continue;
      if (rowArea !== area) continue;
      if (filterCong && cong.toLowerCase() !== filterCong.toLowerCase()) continue;

      const dtIni = (r[COL.DATA_INI-1] instanceof Date) ? r[COL.DATA_INI-1] : null;
      const dtFim = (r[COL.DATA_FIM-1] instanceof Date) ? r[COL.DATA_FIM-1] : null;

      items.push({
        rowIndex,
        congregacao: cong,
        area: rowArea,
        dataInicial: dtIni ? dtIni.getTime() : null,
        dataFinal: dtFim ? dtFim.getTime() : null,
        mesAno: norm_(r[COL.MES_ANO-1]),
        totalEntrada: num_(r[COL.TOTAL_ENTRADA-1]),
        pdfUrl: norm_(r[COL.PDF_URL-1]),
      });
    }

    items.sort((a,b) => (b.dataFinal || 0) - (a.dataFinal || 0));
    return { ok:true, items: items.slice(0, lim) };
  });
}

function api_getLancamento(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const row = Number(payload?.rowIndex);
    if (!row || row < 2) return { ok:false, message:'Linha inválida.' };

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();
    if (row > lastRow) return { ok:false, message:'Linha inválida.' };
    const v = sh.getRange(row, 1, 1, 19).getValues()[0];

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const cong = norm_(v[COL.CONG-1]);
    const rowArea = norm_(v[COL.AREA-1]);
    if (!allowed.includes(cong.toLowerCase()) || rowArea !== area) {
      return { ok:false, message:'Sem permissão para acessar este lançamento.' };
    }

    const dtIni = v[COL.DATA_INI-1] instanceof Date ? v[COL.DATA_INI-1] : null;
    const dtFim = v[COL.DATA_FIM-1] instanceof Date ? v[COL.DATA_FIM-1] : null;

    return {
      ok:true,
      item: {
        rowIndex: row,
        congregacao: cong,
        area: rowArea,
        dataInicial: dtIni ? Utilities.formatDate(dtIni, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        dataFinal: dtFim ? Utilities.formatDate(dtFim, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        mesAno: norm_(v[COL.MES_ANO-1]),
        dizimo: num_(v[COL.DIZIMO-1]),
        ofertas: num_(v[COL.OFERTAS-1]),
        ofertasEbd: num_(v[COL.OFERTAS_EBD-1]),
        ofertaEspecial: num_(v[COL.OFERTA_ESP-1]),
        saldoAnterior: num_(v[COL.SALDO_ANT-1]),
        despesaCongregacao: num_(v[COL.DESP_CONG-1]),
        transporteDirigente: num_(v[COL.TRANSP_DIR-1]),
        entradaPix: num_(v[COL.PIX-1]),
        pdfUrl: norm_(v[COL.PDF_URL-1]),
      }
    };
  });
}

function api_updateLancamento(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const row = Number(payload?.rowIndex);
    if (!row || row < 2) return { ok:false, message:'Linha inválida.' };

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();
    if (row > lastRow) return { ok:false, message:'Linha inválida.' };
    const curr = sh.getRange(row, 1, 1, 19).getValues()[0];

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const currCong = norm_(curr[COL.CONG-1]);
    const currArea = norm_(curr[COL.AREA-1]);
    if (!allowed.includes(currCong.toLowerCase()) || currArea !== area) {
      return { ok:false, message:'Sem permissão para editar este lançamento.' };
    }

    const form = payload?.form;
    validatePayload_(form, profile);
    const c = calcFields_(form);

    const keepPdf = norm_(curr[COL.PDF_URL-1]);

    const rowVals = new Array(19).fill('');
    rowVals[COL.CONG-1] = form.congregacao;
    rowVals[COL.AREA-1] = form.area;
    rowVals[COL.DATA_INI-1] = toDate_(form.dataInicial);
    rowVals[COL.DATA_FIM-1] = toDate_(form.dataFinal);
    rowVals[COL.MES_ANO-1] = form.mesAno;

    rowVals[COL.DIZIMO-1] = c.dizimo;
    rowVals[COL.OFERTAS-1] = c.ofertas;
    rowVals[COL.OFERTAS_EBD-1] = c.ofertasEbd;
    rowVals[COL.OFERTA_ESP-1] = c.ofertaEspecial;

    rowVals[COL.TOTAL_ENTRADA-1] = c.totalEntrada;
    rowVals[COL.SALDO_ANT-1] = c.saldoAnterior;
    rowVals[COL.PERC_DESP-1] = c.percentualDespesa;
    rowVals[COL.DESP_CONG-1] = c.despesaCong;
    rowVals[COL.SUPERVISOR-1] = c.supervisor;
    rowVals[COL.TRANSP_DIR-1] = c.transpDir;
    rowVals[COL.DESP_TOTAL-1] = c.despesaTotal;
    rowVals[COL.PIX-1] = c.pix;
    rowVals[COL.SALDO_ATUAL-1] = c.saldoAtual;
    rowVals[COL.PDF_URL-1] = keepPdf;

    sh.getRange(row, 1, 1, 19).setValues([rowVals]);
    formatMoneyColumns_(sh);

    return { ok:true, message:'Lançamento atualizado com sucesso.' };
  });
}

function api_deleteLancamento(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const row = Number(payload?.rowIndex);
    if (!row || row < 2) return { ok:false, message:'Linha inválida.' };

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();
    if (row > lastRow) return { ok:false, message:'Linha inválida.' };
    const curr = sh.getRange(row, 1, 1, 19).getValues()[0];

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const cong = norm_(curr[COL.CONG-1]);
    const rowArea = norm_(curr[COL.AREA-1]);
    if (!allowed.includes(cong.toLowerCase()) || rowArea !== area) {
      return { ok:false, message:'Sem permissão para excluir este lançamento.' };
    }

    // tenta jogar o PDF na lixeira (se for arquivo drive)
    const pdfUrl = norm_(curr[COL.PDF_URL-1]);
    try {
      const id = extractDriveFileId_(pdfUrl);
      if (id) DriveApp.getFileById(id).setTrashed(true);
    } catch(_) {}

    sh.deleteRow(row);
    return { ok:true, message:'Lançamento excluído.' };
  });
}

/**************** DASHBOARD + CHART ****************/
function api_getDashboardBundle(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const filter = norm_(payload?.congregacaoFilter);
    if (filter && !allowed.includes(filter.toLowerCase())) {
      return { ok:false, message:'Filtro de congregação inválido.' };
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();

    if (lastRow < 2) {
      return {
        ok:true,
        cards: { saldoAtual:0, dizimosSemana:0, entradas7d:0, saidas7d:0, entradasMes:0, saidasMes:0 },
        chart: { labels:[], values:[] }
      };
    }

    const data = sh.getRange(2, 1, lastRow - 1, 19).getValues();

    const today = new Date();
    const start7 = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 6);
    const currentMesAno = pad2_(today.getMonth()+1) + '/' + today.getFullYear();

    let lastSaldoAtual = null;
    let lastSaldoDate = null;

    let dizimosSemana = 0;
    let entradas7d = 0; // J
    let saidas7d = 0;   // (M+O)
    let entradasMes = 0;
    let saidasMes = 0;

    const mapMes = new Map(); // mesAno -> soma total entrada

    for (const r of data) {
      const cong = norm_(r[COL.CONG-1]);
      const rowArea = norm_(r[COL.AREA-1]);
      if (!cong || !rowArea) continue;

      if (!allowed.includes(cong.toLowerCase())) continue;
      if (rowArea !== area) continue;
      if (filter && cong.toLowerCase() !== filter.toLowerCase()) continue;

      const dtFinal = (r[COL.DATA_FIM-1] instanceof Date) ? r[COL.DATA_FIM-1] : null;

      const dizimo = num_(r[COL.DIZIMO-1]);
      const totalEntrada = num_(r[COL.TOTAL_ENTRADA-1]);
      const despesaCong = num_(r[COL.DESP_CONG-1]);
      const transpDir = num_(r[COL.TRANSP_DIR-1]);
      const saldoAtual = num_(r[COL.SALDO_ATUAL-1]);
      const mesAno = norm_(r[COL.MES_ANO-1]);

      if (dtFinal) {
        if (!lastSaldoDate || dtFinal.getTime() > lastSaldoDate.getTime()) {
          lastSaldoDate = dtFinal;
          lastSaldoAtual = saldoAtual;
        }
      }

      if (dtFinal && dtFinal >= start7 && dtFinal <= today) {
        dizimosSemana += dizimo;
        entradas7d += totalEntrada;
        saidas7d += (despesaCong + transpDir);
      }

      if (mesAno === currentMesAno) {
        entradasMes += totalEntrada;
        saidasMes += (despesaCong + transpDir);
      }

      if (mesAno) mapMes.set(mesAno, (mapMes.get(mesAno) || 0) + totalEntrada);
    }

    if (lastSaldoAtual === null) lastSaldoAtual = 0;

    const items = Array.from(mapMes.entries())
      .map(([k,v]) => ({ mesAno:k, total:v, sortKey: parseMesAnoKey_(k) }))
      .filter(x => x.sortKey !== null)
      .sort((a,b) => a.sortKey - b.sortKey);

    const last12 = items.slice(Math.max(items.length - 12, 0));

    return {
      ok:true,
      cards: { saldoAtual:lastSaldoAtual, dizimosSemana, entradas7d, saidas7d, entradasMes, saidasMes },
      chart: { labels: last12.map(x=>x.mesAno), values: last12.map(x=>x.total) }
    };
  });
}

/**************** PDF GENERATION ****************/
function api_generatePdfForLancamento(payload) {
  return safe_(() => {
    ensureSheets_();
    const auth = checkAuth_(payload?.token);
    if (!auth.ok) return { ok:false, message:'Sessão expirada. Faça login novamente.' };

    const row = Number(payload?.rowIndex);
    if (!row || row < 2) return { ok:false, message:'Linha inválida.' };

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DADOS);
    const lastRow = sh.getLastRow();
    if (row > lastRow) return { ok:false, message:'Linha inválida.' };
    const r = sh.getRange(row, 1, 1, 19).getValues()[0];

    const profile = auth.profile;
    const area = String(profile.area).trim();
    const allowed = profile.congregacoes.map(c => c.toLowerCase());

    const cong = norm_(r[COL.CONG-1]);
    const rowArea = norm_(r[COL.AREA-1]);
    if (!allowed.includes(cong.toLowerCase()) || rowArea !== area) {
      return { ok:false, message:'Sem permissão para gerar PDF deste lançamento.' };
    }

    const dtIni = r[COL.DATA_INI-1] instanceof Date ? r[COL.DATA_INI-1] : null;
    const dtFim = r[COL.DATA_FIM-1] instanceof Date ? r[COL.DATA_FIM-1] : null;

    const d = {
      congregacao: cong,
      area: rowArea,
      dataInicial: dtIni ? formatDateBR_(dtIni) : '',
      dataFinal: dtFim ? formatDateBR_(dtFim) : '',
      mesAno: norm_(r[COL.MES_ANO-1]),

      dizimo: num_(r[COL.DIZIMO-1]),
      ofertas: num_(r[COL.OFERTAS-1]),
      ofertasEbd: num_(r[COL.OFERTAS_EBD-1]),
      ofertaEspecial: num_(r[COL.OFERTA_ESP-1]),

      totalEntrada: num_(r[COL.TOTAL_ENTRADA-1]),
      saldoAnterior: num_(r[COL.SALDO_ANT-1]),
      percentualDespesa: num_(r[COL.PERC_DESP-1]),
      despesaCongregacao: num_(r[COL.DESP_CONG-1]),
      supervisor: num_(r[COL.SUPERVISOR-1]),
      transporteDirigente: num_(r[COL.TRANSP_DIR-1]),
      despesaTotal: num_(r[COL.DESP_TOTAL-1]),
      saldoAtual: num_(r[COL.SALDO_ATUAL-1]),
    };

    const repasseTemploCentral = d.totalEntrada - d.percentualDespesa - d.supervisor;

    const folder = DriveApp.getFolderById(REPORT_FOLDER_ID);
    const safeName = `Relatorio_${d.congregacao}_${d.mesAno}_L${row}`.replace(/[\\\/:*?"<>|]/g, '-');

    const copyFile = DriveApp.getFileById(REPORT_TEMPLATE_ID).makeCopy(safeName, folder);
    const presId = copyFile.getId();

    const pres = SlidesApp.openById(presId);
    const repl = buildReplacements_(d, repasseTemploCentral);
    applyReplacements_(pres, repl);
    pres.saveAndClose();

    const pdfBlob = DriveApp.getFileById(presId).getBlob().getAs(MimeType.PDF).setName(safeName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);

    try { DriveApp.getFileById(presId).setTrashed(true); } catch(_) {}

    const url = pdfFile.getUrl();
    sh.getRange(row, COL.PDF_URL).setValue(url);

    return { ok:true, pdfUrl:url, message:'PDF gerado e salvo com sucesso.' };
  });
}

function formatDateBR_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatMoneyBR_(n) {
  const v = Number(n || 0);
  return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function buildReplacements_(d, repasseTemploCentral) {
  return {
    'B': d.dataInicial,
    'C': d.dataFinal,
    'D': d.mesAno,
    'E': formatMoneyBR_(d.dizimo),
    'F': formatMoneyBR_(d.ofertas),
    'G': formatMoneyBR_(d.ofertasEbd),
    'Q': formatMoneyBR_(d.ofertaEspecial),
    'H': formatMoneyBR_(d.totalEntrada),
    'I': formatMoneyBR_(d.percentualDespesa),
    'J': formatMoneyBR_(d.saldoAnterior),
    'K': formatMoneyBR_(d.despesaCongregacao),
    'L': formatMoneyBR_(d.supervisor),
    'M': formatMoneyBR_(d.despesaTotal),
    'N': formatMoneyBR_(repasseTemploCentral),
    'P': formatMoneyBR_(d.saldoAtual),
    'U': formatMoneyBR_(d.transporteDirigente),
    'CONGREGACAO': d.congregacao,
    'ÁREA': d.area,
    'AREA': d.area,
  };
}

function applyReplacements_(presentation, dict) {
  const slides = presentation.getSlides();
  slides.forEach(slide => {
    Object.keys(dict).forEach(key => {
      const val = String(dict[key] ?? '');
      slide.replaceAllText('{{' + key + '}}', val);
      slide.replaceAllText('[[' + key + ']]', val);
      slide.replaceAllText('<<' + key + '>>', val);
    });
  });
}

function extractDriveFileId_(url) {
  if (!url) return null;
  const s = String(url);
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m2 && m2[1]) return m2[1];
  return null;
}

/**************** SETUP SHEETS ****************/
function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();

  let admin = ss.getSheetByName(SHEET_ADMIN);
  if (!admin) admin = ss.insertSheet(SHEET_ADMIN);

  let dados = ss.getSheetByName(SHEET_DADOS);
  if (!dados) dados = ss.insertSheet(SHEET_DADOS);

  if (dados.getLastRow() === 0) {
    dados.getRange(1, 1, 1, 19).setValues([[
      'Congregação','Área','Data Inicial','Data Final','Mês/Ano',
      'Dízimo','Ofertas','Ofertas EBD','Oferta Especial',
      'Total Entrada','Saldo Anterior','Percentual Despesa',
      'Despesa Congregação','Supervisor','Transporte Dirigente',
      'Despesa Total','Entrada Pix','Saldo Atual','PDF URL'
    ]]);
    dados.setFrozenRows(1);
  }

  formatMoneyColumns_(dados);
}
