/**
 * API de funis da Pati — lê a planilha de estratégia AO VIVO e devolve JSON (JSONP).
 * A planilha NÃO é compartilhada: o script roda como o dono (eleuterio.corp@gmail,
 * que já tem acesso) e expõe só os números processados.
 *
 * Deploy (IMPORTANTE — fazer na conta PESSOAL @gmail, não na do Workspace):
 *  1. script.google.com -> Novo projeto -> cole este arquivo.
 *  2. Implantar -> Nova implantação -> tipo "App da Web".
 *     - Executar como: Eu (eleuterio.corp@gmail.com)
 *     - Quem tem acesso: QUALQUER PESSOA
 *  3. Copie a URL que termina em /exec e me mande.
 *  Teste: abra <URL>/exec?callback=cb no navegador — tem que vir cb({...}).
 */
var SHEET_ID = '1WpJOUA4bgC235ZmFwS7WPSgfA-XBoLMRqzwJA1zoBzU';
var TAB = 'Total'; // a aba que o José preenche

function doGet(e) {
  var out;
  try { out = buildData(); }
  catch (err) { out = { erro: String(err) }; }
  var json = JSON.stringify(out);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function buildData() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
  var last = sh.getLastRow();
  var vals = sh.getRange(1, 1, last, 16).getValues(); // A..P
  var funis = [];
  var cur = null;
  for (var i = 0; i < vals.length; i++) {
    var b = String(vals[i][1] || '').trim(); // coluna B = rótulo
    if (b.indexOf('●') === 0) {           // "●" = nova origem/funil
      cur = {
        nome: b.replace('●', '').trim(),
        inv: zeros12(), fat: zeros12(), leads: zeros12()
      };
      funis.push(cur);
    } else if (cur) {
      var meses = []; // C..N = 12 meses (índices 2..13)
      for (var m = 0; m < 12; m++) meses.push(num(vals[i][2 + m]));
      if (b === 'Investimento') cur.inv = meses;
      else if (b.indexOf('Faturamento') === 0) for (var k = 0; k < 12; k++) cur.fat[k] += meses[k]; // soma origem + tag
      else if (b === 'Leads') cur.leads = meses;
    }
  }
  return { atualizado: new Date().toISOString(), tab: TAB, funis: funis };
}

function zeros12() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }
function num(x) {
  if (typeof x === 'number') return x;
  var s = String(x || '').replace(/R\$/gi, '').replace(/\s/g, '');
  if (s === '' || s === '-') return 0;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
