/* Painel de Vendas: conversa com servidor.py, que grava tudo em vendas.xlsx. */

let CFG = null;            // preços, frete e estados (vêm do servidor.py)
let sales = [];            // vendas lidas da planilha
let loaded = false;        // já tentou carregar?
let online = false;        // servidor respondeu?
let confirmId = null;      // pedido aguardando confirmação de remoção

const $ = id => document.getElementById(id);
const brl = v => v.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
const pct = v => v.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + '%';
const nint = v => v.toLocaleString('pt-BR');
const esc = s => String(s).replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const dateBR = iso => { const p = String(iso).split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : esc(iso); };
const todayISO = () => { const d = new Date(), z = n => String(n).padStart(2, '0'); return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()); };

/* ---------- comunicação com o servidor ---------- */
async function api(metodo, caminho, corpo) {
  let resp;
  try {
    resp = await fetch('/api' + caminho, {
      method: metodo,
      headers: corpo ? {'Content-Type': 'application/json'} : undefined,
      body: corpo ? JSON.stringify(corpo) : undefined
    });
  } catch (e) {
    throw new Error('Servidor desligado. Rode "python servidor.py" e abra http://localhost:8000');
  }
  let dados = null;
  try { dados = await resp.json(); } catch (e) { /* sem corpo */ }
  if (!resp.ok) throw new Error((dados && dados.erro) || 'Erro ' + resp.status);
  return dados;
}

function setStatus(tipo, texto) {
  const el = $('status');
  el.className = 'chip ' + tipo;
  el.innerHTML = '<i></i><span>' + esc(texto) + '</span>';
}

async function carregarVendas() {
  try {
    sales = await api('GET', '/vendas');
    online = true;
    setStatus('ok', 'Ligado à planilha vendas.xlsx');
  } catch (e) {
    online = false;
    setStatus('err', e.message);
  }
  loaded = true;
  render();
}

/* ---------- formulário ---------- */
function buildForm() {
  $('f-produto').innerHTML = '<option value="">Selecione o produto</option>' +
    Object.keys(CFG.precos).map(p => '<option value="' + esc(p) + '">' + esc(p) + ' · ' + brl(CFG.precos[p]) + '</option>').join('');
  $('f-estado').innerHTML = '<option value="">UF</option>' + CFG.regioes.map(r =>
    '<optgroup label="' + r + '">' + Object.keys(CFG.estados[r]).map(u =>
      '<option value="' + u + '">' + u + ' · ' + esc(CFG.estados[r][u]) + '</option>').join('') + '</optgroup>').join('');
  $('flt-reg').innerHTML = '<option value="">Todas</option>' + CFG.regioes.map(r => '<option>' + r + '</option>').join('');
  $('flt-prod').innerHTML = '<option value="">Todos</option>' + Object.keys(CFG.precos).map(p => '<option value="' + esc(p) + '">' + esc(p) + '</option>').join('');
}

function regiaoDe(uf) {
  for (const r of CFG.regioes) if (CFG.estados[r][uf]) return r;
  return '';
}
function nextPedido() { return sales.reduce((m, s) => Math.max(m, s.pedido), 1000) + 1; }

function current() {
  if (!CFG) return {produto: '', uf: '', regiao: '', qtd: 0, valorUnit: 0, frete: 0};
  const produto = $('f-produto').value, uf = $('f-estado').value;
  const qtd = parseInt($('f-qtd').value, 10);
  const regiao = uf ? regiaoDe(uf) : '';
  return {produto, uf, regiao, qtd: qtd > 0 ? qtd : 0,
    valorUnit: produto ? CFG.precos[produto] : 0, frete: regiao ? CFG.frete[regiao] : 0};
}

function renderCalc() {
  const c = current(), dash = '—';
  const prods = c.produto && c.qtd ? c.qtd * c.valorUnit : null;
  const fr = c.regiao && c.qtd ? c.qtd * c.frete : null;
  $('calc').innerHTML =
    '<div><dt>Região</dt><dd>' + (c.regiao || dash) + '</dd></div>' +
    '<div><dt>Valor unitário</dt><dd>' + (c.produto ? brl(c.valorUnit) : dash) + '</dd></div>' +
    '<div><dt>Produtos' + (prods !== null ? ' (' + c.qtd + ' un)' : '') + '</dt><dd>' + (prods !== null ? brl(prods) : dash) + '</dd></div>' +
    '<div><dt>Frete' + (fr !== null ? ' (' + c.qtd + ' × ' + brl(c.frete) + ')' : '') + '</dt><dd>' + (fr !== null ? brl(fr) : dash) + '</dd></div>' +
    '<div class="total"><dt>Total da venda</dt><dd>' + (prods !== null && fr !== null ? brl(prods + fr) : dash) + '</dd></div>';
  $('f-pedido').value = nextPedido();
}

function say(texto, tipo) {
  const m = $('msg');
  m.textContent = texto;
  m.className = 'msg ' + (tipo || '');
}

async function submitSale(ev) {
  ev.preventDefault();
  const c = current(), data = $('f-data').value;
  if (!online) return say('Sem conexão com o servidor. Rode "python servidor.py" e recarregue a página.', 'err');
  if (!data) return say('Informe a data da venda.', 'err');
  if (!c.produto) return say('Escolha o produto.', 'err');
  if (!c.uf) return say('Escolha o estado.', 'err');
  if (!c.qtd) return say('A quantidade deve ser 1 ou mais.', 'err');

  const btn = $('f-ok');
  btn.disabled = true;
  say('Gravando na planilha…');
  try {
    const venda = await api('POST', '/vendas', {data, estado: c.uf, produto: c.produto, quantidade: c.qtd});
    sales.push(venda);
    say('Venda ' + venda.pedido + ' gravada em vendas.xlsx: ' + brl(venda.total) + '.', 'ok');
    $('f-qtd').value = 1;
    render();
  } catch (e) {
    say(e.message, 'err');
  } finally {
    btn.disabled = false;
  }
}

/* ---------- dashboard ---------- */
function barHTML(nome, ratio, texto, destaque, tag) {
  return '<div class="bar' + (destaque ? ' top' : '') + '"><div class="name">' + esc(nome) + (tag ? '<span class="tag">' + tag + '</span>' : '') +
    '</div><div class="track"><div class="fill" style="width:' + Math.max(ratio * 100, 0) + '%"></div></div><div class="val">' + texto + '</div></div>';
}

function renderDash() {
  const total = sales.reduce((s, x) => s + x.total, 0);
  const units = sales.reduce((s, x) => s + x.quantidade, 0);
  const freight = sales.reduce((s, x) => s + x.quantidade * x.frete, 0);
  $('kpis').innerHTML = [
    ['Faturamento total', brl(total)], ['Pedidos', nint(sales.length)],
    ['Unidades vendidas', nint(units)], ['Frete cobrado', brl(freight)]
  ].map(k => '<div class="kpi"><span>' + k[0] + '</span><strong>' + k[1] + '</strong></div>').join('');

  if (!sales.length) {
    const t = !loaded ? 'Carregando vendas…' : (online ? 'Nenhuma venda cadastrada ainda. Preencha o formulário para registrar a primeira.' : 'Sem conexão com o servidor.');
    $('reg').innerHTML = '<div class="empty">' + t + '</div>';
    $('prod').innerHTML = '<div class="empty">' + t + '</div>';
    return;
  }

  const porRegiao = CFG.regioes.map(r => ({r, t: sales.filter(s => s.regiao === r).reduce((a, s) => a + s.total, 0)})).sort((a, b) => b.t - a.t);
  const maxR = porRegiao[0].t || 1;
  $('reg').innerHTML = porRegiao.map((x, i) => barHTML(x.r, x.t / maxR,
    pct(total ? x.t / total * 100 : 0) + ' · ' + brl(x.t), i === 0,
    i === 0 ? 'maior' : (i === porRegiao.length - 1 ? 'menor' : ''))).join('');

  const mapa = {};
  sales.forEach(s => { const m = mapa[s.produto] || (mapa[s.produto] = {q: 0, t: 0}); m.q += s.quantidade; m.t += s.quantidade * s.valorUnit; });
  const porProduto = Object.keys(mapa).map(p => ({p, q: mapa[p].q, t: mapa[p].t})).sort((a, b) => b.q - a.q);
  const maxP = porProduto[0].q || 1;
  $('prod').innerHTML = porProduto.map((x, i) => barHTML(x.p, x.q / maxP,
    nint(x.q) + ' un · ' + brl(x.t) + ' em produtos', i === 0, '')).join('');
}

function renderTable() {
  const fr = $('flt-reg').value, fp = $('flt-prod').value;
  const lista = sales.filter(s => (!fr || s.regiao === fr) && (!fp || s.produto === fp)).sort((a, b) => b.pedido - a.pedido);
  $('count').textContent = lista.length + (lista.length === 1 ? ' venda' : ' vendas');
  if (!lista.length) {
    const t = sales.length ? 'Nenhuma venda com esses filtros.' : (!loaded ? 'Carregando vendas…' : (online ? 'Nenhuma venda cadastrada ainda.' : 'Sem conexão com o servidor.'));
    $('rows').innerHTML = '<tr><td colspan="9" class="empty">' + t + '</td></tr>';
    return;
  }
  $('rows').innerHTML = lista.map(s => {
    const acao = confirmId === s.pedido
      ? '<button class="link crit" data-act="yes" data-id="' + s.pedido + '">Confirmar</button> · <button class="link" data-act="no">Cancelar</button>'
      : '<button class="link" data-act="del" data-id="' + s.pedido + '">Remover</button>';
    return '<tr><td class="num">' + s.pedido + '</td><td>' + dateBR(s.data) + '</td><td>' + esc(s.regiao) + '<span class="uf">' + esc(s.estado) + '</span></td><td>' + esc(s.produto) +
      '</td><td class="num">' + s.quantidade + '</td><td class="num">' + brl(s.valorUnit) + '</td><td class="num">' + brl(s.quantidade * s.frete) +
      '</td><td class="num strong">' + brl(s.total) + '</td><td>' + acao + '</td></tr>';
  }).join('');
}

function render() { renderDash(); renderTable(); renderCalc(); }

async function onRowClick(ev) {
  const b = ev.target.closest('button');
  if (!b) return;
  const acao = b.dataset.act, id = Number(b.dataset.id);
  if (acao === 'del') { confirmId = id; renderTable(); }
  else if (acao === 'no') { confirmId = null; renderTable(); }
  else if (acao === 'yes') {
    confirmId = null;
    try {
      await api('DELETE', '/vendas/' + id);
      await carregarVendas();
    } catch (e) {
      say(e.message, 'err');
      renderTable();
    }
  }
}

/* ---------- início ---------- */
$('f-data').value = todayISO();
$('f').addEventListener('submit', submitSale);
['f-produto', 'f-estado', 'f-qtd'].forEach(id => $(id).addEventListener('input', renderCalc));
$('flt-reg').addEventListener('change', renderTable);
$('flt-prod').addEventListener('change', renderTable);
$('rows').addEventListener('click', onRowClick);
render();

(async () => {
  try {
    CFG = await api('GET', '/config');
  } catch (e) {
    loaded = true;
    setStatus('err', e.message);
    render();
    return;
  }
  buildForm();
  await carregarVendas();
  window.addEventListener('focus', carregarVendas);   // se você editar a planilha, a página atualiza ao voltar para ela
})();
