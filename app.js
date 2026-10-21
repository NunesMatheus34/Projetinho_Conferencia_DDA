/* ===================================================================
   0) NAVEGAÇÃO PELO MENU LATERAL
   =================================================================== */

// MODO_TESTE (só no testes.html): a página de testes não tem a tela; aí um elemento que
// falta vira um <div> solto, para o app.js carregar e os testes usarem as funções.
// No index.html todos os elementos existem e nada muda.
const $ = id => document.getElementById(id) || (window.MODO_TESTE ? document.createElement('div') : null);

// seções que existem; qualquer outro #endereço (inclusive o antigo #dda) abre a primeira
const PAINEIS = ['santander', 'itau', 'codigo-barras', 'boletos', 'ajuda'];

function mostraPainel(nome){
  document.querySelectorAll('.painel').forEach(p => p.classList.add('oculto'));
  document.querySelectorAll('.menu-item[data-painel]').forEach(b => {
    b.classList.remove('ativo');
    b.removeAttribute('aria-current');
  });
  const painel = $('painel-' + nome);
  const btn = document.querySelector(`.menu-item[data-painel="${nome}"]`);
  if(painel) painel.classList.remove('oculto');
  if(btn){ btn.classList.add('ativo'); btn.setAttribute('aria-current', 'page'); }
}
function painelDoHash(){
  const h = location.hash.replace('#','');
  return PAINEIS.includes(h) ? h : 'santander';
}
function sincronizaPainel(){
  if(window.MODO_TESTE) return;
  const nome = painelDoHash();
  // endereço sem # ou com # desconhecido: corrige pra #santander sem criar
  // entrada nova no histórico (o Voltar continua funcionando normal)
  if(location.hash !== '#' + nome){ location.replace('#' + nome); return; }
  mostraPainel(nome);
  fechaMenuCelular();
}

// botão ☰: no computador recolhe/expande o menu; no celular abre/fecha por cima
const celular = window.matchMedia('(max-width:760px)');

function atualizaToggle(){
  const corpo = document.body.classList;
  const aberto = celular.matches ? corpo.contains('menu-aberto') : !corpo.contains('menu-recolhido');
  $('menuToggle').setAttribute('aria-expanded', String(aberto));
  $('menuToggle').setAttribute('aria-label', aberto ? 'Fechar menu' : 'Abrir menu');
}
function fechaMenuCelular(){
  document.body.classList.remove('menu-aberto');
  atualizaToggle();
}

// lembra se o menu estava recolhido (só neste navegador; se não der, ignora)
try{ if(localStorage.getItem('menuRecolhido') === '1') document.body.classList.add('menu-recolhido'); }catch(e){}

$('menuToggle').addEventListener('click', () => {
  if(celular.matches){
    document.body.classList.toggle('menu-aberto');
  }else{
    const recolhido = document.body.classList.toggle('menu-recolhido');
    try{ localStorage.setItem('menuRecolhido', recolhido ? '1' : '0'); }catch(e){}
  }
  atualizaToggle();
});
$('menuFundo').addEventListener('click', fechaMenuCelular);
$('menu').addEventListener('click', e => { if(e.target.closest('a.menu-item')) fechaMenuCelular(); });
document.addEventListener('keydown', e => { if(e.key === 'Escape') fechaMenuCelular(); });
celular.addEventListener('change', fechaMenuCelular);

window.addEventListener('hashchange', sincronizaPainel);
sincronizaPainel();


/* ===================================================================
   1) FUNÇÕES COMPARTILHADAS PELAS DUAS FERRAMENTAS
   =================================================================== */

// formata um número como dinheiro brasileiro: 1234.5 -> "R$ 1.234,50"
const brl = n => (n===null || n===undefined || isNaN(n)) ? '—' : n.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});

// transforma qualquer valor (texto ou número) num número limpo com 2 casas.
// resolve o caso do valor vir como texto tipo "1.234,56"
function paraNumero(v){
  if(v===null || v===undefined || v==='') return null;
  if(typeof v==='number') return Math.round(v*100)/100;
  let s = String(v).replace(/[R$\s]/g,'')
                   .replace(/\./g,'')
                   .replace(',','.');
  let n = parseFloat(s);
  return isNaN(n) ? null : Math.round(n*100)/100;
}

// formata uma data como dd/mm/aaaa, venha ela como texto ou como data
function formatData(v){
  if(v===null || v===undefined || v==='') return '';
  if(v instanceof Date){
    const d = String(v.getDate()).padStart(2,'0');
    const m = String(v.getMonth()+1).padStart(2,'0');
    return `${d}/${m}/${v.getFullYear()}`;
  }
  return String(v).trim();
}

// segurança: evita que caracteres do nome quebrem o HTML
function escapa(s){
  return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
}

// abre o Excel e devolve a primeira aba (objeto do SheetJS)
function lerPrimeiraAba(file){
  return new Promise((res,rej)=>{
    const fr = new FileReader();
    fr.onload = e => {
      try{
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
        res(wb.Sheets[wb.SheetNames[0]]);
      }catch(err){ rej(err); }
    };
    fr.onerror = () => rej(new Error('Não consegui ler o arquivo.'));
    fr.readAsArrayBuffer(file);
  });
}

const matrizDaAba = ws => XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});

// abre um Excel e devolve uma matriz: lista de linhas, cada linha uma lista de células
async function lerMatriz(file){
  return matrizDaAba(await lerPrimeiraAba(file));
}

// Igual ao lerMatriz, mas também dá o TEXTO de cada célula, para ler código de barras.
// Na matriz, uma célula numérica vira Number, e o JavaScript arredonda depois do
// 15º dígito sem dar erro. textoDe(i, j) usa o texto da célula (o valor original
// quando já é texto; senão o texto formatado, cell.w), nunca Number.
// i e j são os mesmos índices da matriz (linha e coluna).
// eraTexto = false quando o Excel não gravou a célula como texto (número, data...):
// número no Excel só guarda 15 dígitos, então um código de barras ali já pode estar alterado.
// Célula vazia conta como texto (vazio).
// linhaExcel(i) = número da linha no Excel (1, 2, 3...), para mostrar ao usuário.
async function lerMatrizComTexto(file){
  const ws = await lerPrimeiraAba(file);
  const matriz = matrizDaAba(ws);
  const ini = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const textoDe = (i, j) => {
    const cel = ws[XLSX.utils.encode_cell({ r: ini.s.r + i, c: ini.s.c + j })];
    if(!cel || cel.v === null || cel.v === undefined) return { texto:'', eraTexto:true };
    if(typeof cel.v === 'string') return { texto:cel.v, eraTexto:true };
    return { texto: cel.w !== undefined ? String(cel.w) : '', eraTexto:false };
  };
  return { matriz, textoDe, linhaExcel: i => ini.s.r + i + 1 };
}

// código de barras de uma célula lida por lerMatrizComTexto: só os dígitos, e
// "suspeito" quando a célula não era texto (num número o Excel pode ter trocado os
// últimos dígitos por zero ou gravado em notação científica, ex.: 2,38E+43)
function codigoBarrasDaCelula(leitura, i, j){
  const { texto, eraTexto } = leitura.textoDe(i, j);
  return { digitos: apenasDigitos(texto), suspeito: !eraTexto };
}

// prepara uma caixa de upload (clicar OU arrastar arquivo)
function ligaDrop(dropId, inputId, txtId, onFile){
  const drop = $(dropId);
  const inp  = $(inputId);
  const txt  = $(txtId);

  inp.addEventListener('change', () => { if(inp.files[0]) aceita(inp.files[0]); });

  ['dragover','dragenter'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave','drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if(f){ inp.files = e.dataTransfer.files; aceita(f); }
  });

  function aceita(f){
    drop.classList.add('ok');
    txt.textContent = '✓ ' + f.name;
    onFile(f);
  }
}

function mostraErro(id, msg){ const e=$(id); e.textContent='⚠ '+msg; e.style.display='block'; }
function limpaErro(id){ $(id).style.display='none'; }

const BORDA_FINA = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

// data só pra mostrar na tela, sempre dd/mm/aaaa (não mexe no dado usado na conferência)
function dataBR(v){
  const s = formatData(v);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return `${m[3]}/${m[2]}/${m[1]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if(m) return `${m[1].padStart(2,'0')}/${m[2].padStart(2,'0')}/${m[3].length===2 ? '20'+m[3] : m[3]}`;
  return s;
}

const centavos = v => (v===null || v===undefined || isNaN(v)) ? '' : String(Math.round(Math.abs(v)*100));

// nome de cada grupo do resultado (etiqueta da tabela)
const NOMES_GRUPO = {
  ok:'Conferido', data:'Data divergente', nome:'Nome divergente', banco:'Só no banco', sis:'Só no sistema',
  valor:'Divergência de valor', dup:'Duplicado',
};
const etiqueta = (grupo, texto) => `<span class="status g-${grupo}">${texto || NOMES_GRUPO[grupo]}</span>`;

// texto curto que aparece na tabela quando o grupo escolhido não tem nenhum título
const MSG_VAZIO = {
  nome:  'Nenhum boleto com nome diferente. Tudo que bateu no valor também bateu no nome do fornecedor.',
  data:  'Nenhum vencimento diferente. As datas batem em tudo que foi encontrado.',
  banco: 'Nenhum boleto do banco ficou de fora: todos foram achados no sistema.',
  sis:   'Nenhum título do sistema ficou de fora: todos foram achados no banco.',
  ok:    'Nenhum título conferido. Veja os outros grupos.',
  valor: 'Nenhuma diferença: o valor do código de barras bate com o saldo lançado em todos os títulos.',
  dup:   'Nenhum código de barras lançado mais de uma vez.',
};

// nome dos arquivos escolhidos, no alto da tela
function mostraArquivos(pref, arquivos){
  const nomes = arquivos.filter(Boolean).map(f => f.name);
  $(pref + 'Arquivos').textContent = nomes.length ? nomes.join(' · ') : 'Nenhum arquivo enviado';
}

// libera os botões "Baixar resultado (Excel)" da tela. O do alto só aciona o de baixo,
// que é quem gera o Excel.
function habilitaBaixar(pref){
  document.querySelectorAll(`[data-baixar="${pref}"]`).forEach(b => b.disabled = false);
}
document.querySelectorAll('[data-baixar]').forEach(b =>
  b.addEventListener('click', () => $(b.dataset.baixar + 'BtnXlsx').click()));


/* -------------------------------------------------------------------
   TABELA DE RESULTADO (igual nas três conferências)
   Cada tela entrega a lista de linhas já com o grupo de cada uma. A busca
   e o clique nos cartões de resumo só escolhem quais linhas aparecem;
   nada disso muda o resultado da conferência nem o Excel.
   Cada linha: { grupos:[...], valor, texto, digitos, ... }
     grupos  = em quais cartões a linha conta (ex: ['ok','dup'])
     valor   = o que entra no total do cartão e no rodapé
     texto / digitos = onde a busca procura
   ------------------------------------------------------------------- */
const TELAS = {};

function registraTela(pref, { colunas, linhaHtml, rodapeHtml }){
  TELAS[pref] = { linhas:[], grupo:null, busca:'', colunas, linhaHtml, rodapeHtml };
  $(pref + 'Busca').addEventListener('input', e => { TELAS[pref].busca = e.target.value; desenhaTabela(pref); });
  document.querySelectorAll(`#${pref}Resumo .stat`).forEach(card => {
    card.querySelector('.stat-filtro').addEventListener('click', () => {
      const t = TELAS[pref];
      t.grupo = t.grupo === card.dataset.grupo ? null : card.dataset.grupo;
      desenhaTabela(pref);
    });
  });
  $(pref + 'Filtro').addEventListener('click', e => {
    if(e.target.closest('.limpa-filtro')){ TELAS[pref].grupo = null; desenhaTabela(pref); }
  });
}

const somaValores = linhas => Math.round(linhas.reduce((s, l) => s + (l.valor || 0), 0) * 100) / 100;

// mostra um resultado novo: preenche os cartões e desenha a tabela
function mostraLinhas(pref, linhas){
  const t = TELAS[pref];
  t.linhas = linhas;
  document.querySelectorAll(`#${pref}Resumo .stat`).forEach(card => {
    const doGrupo = linhas.filter(l => l.grupos.includes(card.dataset.grupo));
    card.querySelector('.stat-qtd').textContent = doGrupo.length;
    card.querySelector('.stat-total').textContent = brl(somaValores(doGrupo));
  });
  desenhaTabela(pref);
}

// a busca procura no texto (fornecedor, NF, CNPJ...) e, se for número, nos valores e documentos
function buscaCasa(linha, busca){
  const q = normaliza(busca).replace(/\s+/g,' ');
  if(!q) return true;
  if(linha.texto.includes(q)) return true;
  if(/^[R$\s\d.,\/-]+$/.test(q)){
    const d = q.replace(/\D/g,'');
    return d.length>0 && linha.digitos.includes(d);
  }
  return false;
}

function desenhaTabela(pref){
  const t = TELAS[pref];
  const visiveis = t.linhas.filter(l => (!t.grupo || l.grupos.includes(t.grupo)) && buscaCasa(l, t.busca));

  let vazio = 'Nenhum título para mostrar.';
  if(t.busca.trim()) vazio = `Nada encontrado para “${escapa(t.busca.trim())}”${t.grupo ? ' neste grupo' : ''}.`;
  else if(t.grupo) vazio = MSG_VAZIO[t.grupo];
  $(pref + 'Corpo').innerHTML = visiveis.length
    ? visiveis.map(t.linhaHtml).join('')
    : `<tr><td class="vazio" colspan="${t.colunas}">${vazio}</td></tr>`;
  $(pref + 'Rodape').innerHTML = t.rodapeHtml(visiveis);

  document.querySelectorAll(`#${pref}Resumo .stat`).forEach(card => {
    const ativo = card.dataset.grupo === t.grupo;
    card.classList.toggle('ativo', ativo);
    card.querySelector('.stat-filtro').setAttribute('aria-pressed', String(ativo));
  });
  const filtro = $(pref + 'Filtro');
  filtro.hidden = !t.grupo;
  if(t.grupo){
    const nome = document.querySelector(`#${pref}Resumo .stat[data-grupo="${t.grupo}"] .stat-nome`).textContent;
    filtro.innerHTML = `Mostrando só: <b>${escapa(nome)}</b> <button type="button" class="limpa-filtro">Ver todos</button>`;
  }
}

const qtdTitulos = n => n === 1 ? '1 título' : `${n} títulos`;


/* ===================================================================
   2) LEITORES POR BANCO
   Cada banco exporta o DDA num layout diferente. O leitor de cada banco
   recebe a matriz do Excel (de lerMatriz) e devolve uma lista no formato
   comum, que é o que a conferência usa:
     { venc, beneficiario, cnpj, doc, valor, valorAPagar, codigoBarras,
       pagador, banco, comDesconto }
   "valor" é o que é comparado com o saldo do sistema. "banco" é o banco
   emissor do boleto (código, ex: 237). O que o banco não tiver fica vazio
   ('' nos textos, null nos valores, false no comDesconto).
   Pra incluir um banco novo, é só escrever o leitor dele aqui e ensinar
   o detectaBanco a reconhecer o arquivo.
   (normaliza e apenasDigitos ficam na seção de código de barras)
   =================================================================== */

const NOMES_BANCO = { santander:'Santander', itau:'Itaú' };

// descobre de qual banco é o DDA pelo conteúdo do arquivo (não pelo nome).
// devolve 'santander', 'itau' ou null se não reconhecer.
function detectaBanco(m){
  if(achaCabecalhoItau(m) >= 0) return 'itau';
  for(let i=0; i<Math.min(m.length, 30); i++){
    const cels = (m[i] || []).map(normaliza);
    if(cels.includes('NOSSO NUMERO') && cels.includes('BENEFICIARIO')) return 'santander';
  }
  return null;
}

function avisoBancoErrado(achado, tela){
  return `Este arquivo é um DDA do ${NOMES_BANCO[achado]}, mas foi enviado na Conferência ${NOMES_BANCO[tela]}. ` +
         `Abra a Conferência ${NOMES_BANCO[achado]} no menu e envie o arquivo lá.`;
}

// SANTANDER: cabeçalho na linha 5, dados a partir da linha 6 (índice 5).
// Colunas: A=Vencimento(0) B=Beneficiário(1) C=Nosso Número(2) D=Pagador(3) E=Valor(4)
// Não traz CNPJ, valor a pagar, banco emissor nem código de barras.
function leitorSantander(m){
  const out = [];
  for(let i=5; i<m.length; i++){
    const linha = m[i] || [];
    const venc  = linha[0];
    const nome  = linha[1];
    const doc   = linha[2];
    const pagador = linha[3];
    const valor = paraNumero(linha[4]);
    if(valor===null) continue;
    out.push({
      venc: formatData(venc),
      beneficiario: nome ? String(nome).trim() : '',
      cnpj: '',
      doc: doc!=null ? String(doc).trim() : '',
      valor,
      valorAPagar: null,
      codigoBarras: '',
      pagador: pagador!=null ? String(pagador).trim() : '',
      banco: '',
      comDesconto: false,
    });
  }
  return out;
}

// ITAÚ ("Consulta Boletos Recebidos" do internet banking): as primeiras linhas
// têm os dados da conta; o cabeçalho é a linha com "Beneficiário" e
// "Código de barras" (achada pelo conteúdo, não pelo número da linha).
function achaCabecalhoItau(m){
  for(let i=0; i<m.length; i++){
    const cels = (m[i] || []).map(normaliza);
    if(cels.some(c => c.includes('BENEFICIARIO')) &&
       cels.some(c => c.includes('CODIGO') && c.includes('BARRA'))) return i;
  }
  return -1;
}

// acha cada coluna pelo texto do cabeçalho; se não achar, usa a posição padrão:
// A=Pagador/Agregado B=Beneficiário C=CPF/CNPJ D=Venc. E=Nº Doc. F=Até Venc.
// G=A Pagar H=Tipo de Boleto I=Banco J=Observações K=Código de barras
function colunasItau(linhaCab){
  const cab = (linhaCab || []).map(normaliza);
  const acha = (teste, padrao) => { const i = cab.findIndex(teste); return i>=0 ? i : padrao; };
  return {
    pagador: acha(c => c.startsWith('PAGADOR'), 0),
    benef:   acha(c => c.includes('BENEFICIARIO'), 1),
    cnpj:    acha(c => c.includes('CNPJ'), 2),
    venc:    acha(c => c.startsWith('VENC'), 3),
    doc:     acha(c => c.includes('DOC'), 4),
    ateVenc: acha(c => c.startsWith('ATE VENC'), 5),
    aPagar:  acha(c => c === 'A PAGAR', 6),
    banco:   acha(c => c === 'BANCO', 8),
    codigo:  acha(c => c.includes('CODIGO') && c.includes('BARRA'), 10),
  };
}

// "Até Venc." é o valor comparado com o sistema; "A Pagar" fica em valorAPagar.
// Quando os dois são diferentes, o boleto é marcado como "com desconto".
// leitura (opcional, de lerMatrizComTexto): o código de barras vem do TEXTO da
// célula, nunca de Number, e cada boleto ganha codigoTexto e codigoSuspeito
// (a célula não estava gravada como texto). Sem leitura, nada muda.
function leitorItau(m, leitura){
  const cab = achaCabecalhoItau(m);
  if(cab<0) return [];
  const c = colunasItau(m[cab]);
  const texto = v => v!=null ? String(v).trim() : '';
  const out = [];
  for(let i=cab+1; i<m.length; i++){
    const linha = m[i] || [];
    const valor = paraNumero(linha[c.ateVenc]);
    if(valor===null) continue;
    const valorAPagar = paraNumero(linha[c.aPagar]);
    const cod = leitura ? codigoBarrasDaCelula(leitura, i, c.codigo) : null;
    out.push({
      venc: formatData(linha[c.venc]),
      beneficiario: texto(linha[c.benef]),
      cnpj: texto(linha[c.cnpj]),
      doc: texto(linha[c.doc]),
      valor,
      valorAPagar,
      codigoBarras: cod ? cod.digitos : apenasDigitos(linha[c.codigo]),
      pagador: texto(linha[c.pagador]),
      banco: texto(linha[c.banco]),
      comDesconto: valorAPagar!==null && valorAPagar!==valor,
      ...(cod ? { codigoTexto: leitura.textoDe(i, c.codigo).texto, codigoSuspeito: cod.suspeito,
                  linhaExcel: leitura.linhaExcel(i) } : {}),
    });
  }
  return out;
}

// data e hora da consulta, que o Itaú escreve antes do cabeçalho:
// "Data/Hora: 15/10/2026 às 10:39:21" -> "15/10/2026 às 10:39:21"
function dataHoraConsultaItau(m){
  const cab = achaCabecalhoItau(m);
  const fim = cab>=0 ? cab : m.length;
  for(let i=0; i<fim; i++){
    for(const cel of (m[i] || [])){
      const achou = cel!=null && String(cel).match(/data\s*\/\s*hora\s*:?\s*(.+)/i);
      if(achou) return achou[1].trim();
    }
  }
  return '';
}


/* ===================================================================
   3) CONFERÊNCIA DE DDA
   =================================================================== */

// deixa um nome "padronizado" pra comparar: maiúsculo, sem espaço extra, sem pontuação
function limpaNome(s){
  if(s===null || s===undefined) return '';
  s = String(s).trim().toUpperCase().replace(/\s+/g,' ');
  s = s.replace(/[.,/\-&]/g,'');
  return s.replace(/\s+/g,' ').trim();
}

// decide se dois nomes são "a mesma empresa", mesmo que escritos diferente
function nomesBatem(a,b){
  a = limpaNome(a); b = limpaNome(b);
  if(!a || !b) return false;
  const [curto,longo] = a.length<=b.length ? [a,b] : [b,a];
  if(longo.startsWith(curto.slice(0, Math.min(curto.length,10)))) return true;
  return similaridade(a,b) >= 0.6;
}

// calcula o quanto dois textos são parecidos (0 = nada, 1 = idênticos)
function similaridade(a,b){
  if(a===b) return 1;
  if(a.length<2 || b.length<2) return 0;
  const bigramas = s => {
    const m = new Map();
    for(let i=0;i<s.length-1;i++){
      const g = s.substr(i,2);
      m.set(g,(m.get(g)||0)+1);
    }
    return m;
  };
  const A = bigramas(a), B = bigramas(b);
  let inter = 0;
  A.forEach((c,g)=>{ if(B.has(g)) inter += Math.min(c,B.get(g)); });
  return (2*inter) / ((a.length-1)+(b.length-1));
}

// SISTEMA: dados começam na linha 7 (índice 6). Nome=coluna C(2), Saldo=coluna P(15)
// Título (NF/duplicata)=coluna D(3): só aparece na tabela e entra na busca; não é usado na conferência.
function extraiSistema(m){
  const out = [];
  for(let i=6; i<m.length; i++){
    const linha = m[i] || [];
    const venc  = linha[0];
    const nome  = linha[2];
    if(nome===null || String(nome).trim()==='') continue;
    const saldo = paraNumero(linha[15]);
    if(saldo===null) continue;
    const doc = linha[3];
    out.push({ venc: formatData(venc), nome: String(nome).trim(), valor: saldo, doc: doc!=null ? String(doc).trim() : '' });
  }
  return out;
}

// diferença de até 1 centavo ainda é considerada "mesmo valor" (evita falso
// "não bateu" por arredondamento entre o boleto e o lançamento no sistema)
const TOLERANCIA_VALOR = 0.01;
function valoresBatem(a, b){
  return Math.abs(a - b) <= TOLERANCIA_VALOR + 1e-9;
}

// agrupa os índices de "sis" por valor em centavos, pra achar candidatos a
// parear sem varrer a lista inteira a cada título do banco (O(1) em vez de O(m))
function indexaPorValor(sis){
  const idx = new Map();
  sis.forEach((s,i) => {
    const c = Math.round(s.valor*100);
    if(!idx.has(c)) idx.set(c, []);
    idx.get(c).push(i);
  });
  return idx;
}

// devolve os índices de "sis" cujo valor está dentro da TOLERANCIA_VALOR,
// na mesma ordem em que apareceriam numa varredura completa (índice crescente)
function candidatosPorValor(idx, valor){
  const centro = Math.round(valor*100);
  const tol = Math.round(TOLERANCIA_VALOR*100);
  let out = [];
  for(let c=centro-tol; c<=centro+tol; c++){
    const lista = idx.get(c);
    if(lista) out = out.concat(lista);
  }
  return tol>0 ? out.sort((a,b)=>a-b) : out;
}

// "banco" é a lista no formato comum (vem de um leitor de banco, ver seção 2).
// O resto da conferência compara pelo campo "nome", igual ao sistema, então o
// beneficiário é copiado pra "nome" aqui.
function conferir(banco, sistema){
  const sis = sistema.map(s => ({ ...s, usado:false }));
  const ban = banco.map(b => ({ ...b, nome:b.beneficiario, usado:false }));
  const idxValor = indexaPorValor(sis);

  const ok        = [];
  const datasDiff = [];
  const div       = [];

  // acha, entre os candidatos com valor compatível, o primeiro ainda não
  // usado que também passa no teste extra da camada (nome, data, etc.)
  function achaPar(b, testeExtra){
    for(const i of candidatosPorValor(idxValor, b.valor)){
      const s = sis[i];
      if(!s.usado && valoresBatem(s.valor,b.valor) && testeExtra(s)) return i;
    }
    return -1;
  }

  // Camada 1: valor + nome + vencimento (tudo bate) -> CONFERIDO
  ban.forEach(b => {
    const i = achaPar(b, s => nomesBatem(b.nome,s.nome) && mesmaData(b.venc,s.venc));
    if(i>=0){ sis[i].usado = true; b.usado = true; ok.push({ banco:b, sistema:sis[i] }); }
  });

  // Camada 2: valor + nome batem, mas a DATA é diferente -> DATA DIVERGENTE
  ban.forEach(b => {
    if(b.usado) return;
    const i = achaPar(b, s => nomesBatem(b.nome,s.nome));
    if(i>=0){ sis[i].usado = true; b.usado = true; datasDiff.push({ banco:b, sistema:sis[i] }); }
  });

  // Camada 3: só o valor bate (nome diferente) -> NOME DIVERGENTE
  ban.forEach(b => {
    if(b.usado) return;
    const i = achaPar(b, () => true);
    if(i>=0){ sis[i].usado = true; b.usado = true; div.push({ banco:b, sistema:sis[i] }); }
  });

  const soBanco = ban.filter(b => !b.usado);
  const soSis   = sis.filter(s => !s.usado);

  return { ok, div, soBanco, soSis, datasDiff };
}

function mesmaData(a, b){
  const so = s => (s ? String(s).match(/\d+/g) || [] : []).join('-');
  const A = so(a), B = so(b);
  if(!A || !B) return true;
  return A === B;
}

// transforma o resultado da conferência (ok, div, datasDiff, soBanco, soSis) nas
// linhas da tabela: primeiro o que precisa de atenção, os conferidos por último.
// O valor da linha é o do boleto; no "Só no sistema", o saldo lançado.
function linhasDda(r){
  const linhas = [];
  const add = (grupo, b, s) => {
    const valor = b ? b.valor : s.valor;
    const textos = [b && b.nome, s && s.nome, b && b.doc, s && s.doc, b && b.cnpj, b && b.banco,
                    b && dataBR(b.venc), s && dataBR(s.venc), brl(valor)];
    const digitos = [centavos(valor), b && centavos(b.valor), s && centavos(s.valor),
                     b && apenasDigitos(b.doc), s && apenasDigitos(s.doc), b && apenasDigitos(b.cnpj)];
    linhas.push({ grupo, grupos:[grupo], banco:b, sistema:s, valor,
                  texto: normaliza(textos.filter(Boolean).join(' | ')),
                  digitos: digitos.filter(Boolean).join('|') });
  };
  r.div.forEach(d => add('nome', d.banco, d.sistema));
  r.datasDiff.forEach(d => add('data', d.banco, d.sistema));
  r.soBanco.forEach(b => add('banco', b, null));
  r.soSis.forEach(s => add('sis', null, s));
  r.ok.forEach(d => add('ok', d.banco, d.sistema));
  return linhas;
}

// uma linha da tabela da conferência de DDA. "extras" = colunas do Itaú
// (CNPJ, banco emissor, código de barras) e a marca "com desconto".
function linhaDdaHtml(l, extras){
  const b = l.banco, s = l.sistema;
  const venc = escapa(dataBR(b ? b.venc : s.venc));
  const vencSistema = l.grupo==='data' ? `<span class="sub">Sistema: ${escapa(dataBR(s.venc))}</span>` : '';
  const nome = escapa(b ? b.nome : s.nome);
  const nomeSistema = l.grupo==='nome' ? `<span class="sub">No sistema: ${escapa(s.nome)}</span>` : '';

  // Santander: NF do sistema na frente; o número do boleto no banco embaixo.
  // Itaú: só a NF aqui; o Nº Doc. do boleto vai na coluna própria.
  const nf = s && s.doc ? s.doc : '';
  const docBanco = b && b.doc ? b.doc : '';
  let doc;
  if(extras){
    doc = nf ? `<span class="mono">${escapa(nf)}</span>` : '—';
  }else{
    doc = nf ? `<span class="mono">${escapa(nf)}</span>` : (docBanco ? `<span class="mono">${escapa(docBanco)}</span>` : '—');
    if(nf && docBanco) doc += `<span class="sub mono">Banco: ${escapa(docBanco)}</span>`;
  }
  const colunaDocItau = !extras ? '' : `<td class="mono">${docBanco ? escapa(docBanco) : '—'}</td>`;

  let valor = brl(l.valor);
  if(b && s){
    const diff = Math.round((b.valor - s.valor)*100)/100;
    if(diff!==0) valor += ` <span class="tol" title="Sistema: ${brl(s.valor)} (dif. ${diff>0?'+':''}${brl(diff)})">±</span>`;
  }
  if(extras && b && b.comDesconto){
    valor += `<span class="sub"><span class="status g-ok status-mini" title="A pagar: ${brl(b.valorAPagar)}">com desconto</span></span>`;
  }

  const colunasExtras = !extras ? '' :
    `<td class="mono">${b ? escapa(b.cnpj) : '—'}</td><td class="mono">${b ? escapa(b.banco) : '—'}</td>` +
    `<td class="mono cod">${b && b.codigoBarras ? escapa(b.codigoBarras) : '—'}</td>`;

  return `<tr><td>${etiqueta(l.grupo)}</td><td class="mono">${venc}${vencSistema}</td>` +
         `<td class="forn">${nome}${nomeSistema}</td><td>${doc}</td>${colunaDocItau}${colunasExtras}` +
         `<td class="num mono">${valor}</td></tr>`;
}

function rodapeDda(colunas){
  return linhas => `<tr><td colspan="${colunas-1}">Total · ${qtdTitulos(linhas.length)}</td>` +
                   `<td class="num mono">${brl(somaValores(linhas))}</td></tr>`;
}

let arqBanco = null, arqSis = null, ddaUltimo = null;

registraTela('dda', { colunas:5, linhaHtml: l => linhaDdaHtml(l, false), rodapeHtml: rodapeDda(5) });

function ddaCheca(){ $('ddaBtnGo').disabled = !(arqBanco && arqSis); mostraArquivos('dda', [arqBanco, arqSis]); }
ligaDrop('dropBanco','fileBanco','txtBanco', f => { arqBanco = f; ddaCheca(); });
ligaDrop('dropSis','fileSis','txtSis', f => { arqSis = f; ddaCheca(); });

$('ddaBtnGo').addEventListener('click', async () => {
  limpaErro('ddaErro');
  try{
    const [mB, mS] = await Promise.all([lerMatriz(arqBanco), lerMatriz(arqSis)]);
    const bancoDoArquivo = detectaBanco(mB);
    if(bancoDoArquivo && bancoDoArquivo!=='santander') throw new Error(avisoBancoErrado(bancoDoArquivo, 'santander'));
    const banco   = leitorSantander(mB);
    const sistema = extraiSistema(mS);

    if(!banco.length)   throw new Error('Não encontrei valores no arquivo do banco. Confira se o cabeçalho está na linha 5 e os valores na coluna E.');
    if(!sistema.length) throw new Error('Não encontrei saldos no arquivo do sistema. Confira se os nomes estão na coluna C e o saldo na coluna P.');

    const r = conferir(banco, sistema);
    ddaUltimo = r;

    mostraLinhas('dda', linhasDda(r));

    $('ddaResultado').style.display = 'block';
    habilitaBaixar('dda');
    $('ddaResultado').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    mostraErro('ddaErro', err.message || 'Erro ao processar os arquivos.');
  }
});

$('ddaBtnReset').addEventListener('click', () => location.reload());

const COLS_PADRAO = [
  { cab:'Vencimento',        w:14 },
  { cab:'Nome/Beneficiário', w:46 },
  { cab:'Para NF',           w:16 },
  { cab:'Pedidos',           w:16 },
  { cab:'Valor',             w:16, fmt:'R$ #,##0.00' },
];
const COLS_DATAS = [
  { cab:'Nome/Beneficiário',            w:46 },
  { cab:'Valor',                        w:16, fmt:'R$ #,##0.00' },
  { cab:'Data no DDA (boleto)',         w:20 },
  { cab:'Data no sistema (lançada)',    w:24 },
];
const COLS_DIV = [
  { cab:'Vencimento',      w:14 },
  { cab:'Nome no banco',   w:40 },
  { cab:'Nome no sistema', w:40 },
  { cab:'Para NF',         w:16 },
  { cab:'Pedidos',         w:16 },
  { cab:'Valor',           w:16, fmt:'R$ #,##0.00' },
];

function montaAbaDda(wb, nomeAba, titulo, colunas, linhas){
  const ws = wb.addWorksheet(nomeAba);
  ws.columns = colunas.map(c => ({ width:c.w }));
  const ncols = colunas.length;
  const ultimaCol = String.fromCharCode(64 + ncols);

  ws.mergeCells(`A1:${ultimaCol}1`);
  const t = ws.getCell('A1');
  t.value = titulo;
  t.font = { bold:true, size:13, color:{argb:'FFFFFFFF'} };
  t.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF12283A'} };
  t.alignment = { vertical:'middle', horizontal:'left', indent:1 };
  ws.getRow(1).height = 22;

  const cab = ws.getRow(2);
  cab.values = colunas.map(c => c.cab);
  cab.eachCell(c => {
    c.font = { bold:true, color:{argb:'FF12283A'} };
    c.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFEDE7DA'} };
    c.alignment = { vertical:'middle', horizontal:'center' };
    c.border = BORDA_FINA;
  });
  cab.height = 18;

  linhas.forEach(vals => {
    const r = ws.addRow(vals);
    r.eachCell({ includeEmpty:true }, c => { c.border = BORDA_FINA; });
    colunas.forEach((c, i) => {
      if(c.fmt){
        r.getCell(i+1).numFmt = c.fmt;
        if(c.fmt !== '@') r.getCell(i+1).alignment = { horizontal:'right' };
      }
    });
  });

  if(linhas.length === 0){
    ws.mergeCells(`A3:${ultimaCol}3`);
    ws.getCell('A3').value = 'Nenhum registro nesta categoria.';
    ws.getCell('A3').font = { italic:true, color:{argb:'FF5C6B78'} };
  }
  return ws;
}

$('ddaBtnXlsx').addEventListener('click', async () => {
  if(!ddaUltimo) return;

  const wb = new ExcelJS.Workbook();

  const linhaPadrao = (venc, nome, valor) => [ venc || '', nome || '', '', '', valor ];
  const linhaDiv = (venc, nomeBanco, nomeSistema, valor) =>
    [ venc || '', nomeBanco || '', nomeSistema || '', '', '', valor ];

  montaAbaDda(wb, 'Nome divergente', 'Nome divergente (conferir · possível factoring)', COLS_DIV,
           ddaUltimo.div.map(d => linhaDiv(d.sistema.venc, d.banco.nome, d.sistema.nome, d.sistema.valor)));

  montaAbaDda(wb, 'Datas divergentes', 'Datas divergentes (mesmo título, vencimento diferente)', COLS_DATAS,
           ddaUltimo.datasDiff.map(d => [ d.sistema.nome, d.sistema.valor, d.banco.venc, d.sistema.venc ]));

  montaAbaDda(wb, 'So no banco', 'Só no banco (não achado no sistema)', COLS_PADRAO,
           ddaUltimo.soBanco.map(d => linhaPadrao(d.venc, d.nome, d.valor)));
  montaAbaDda(wb, 'So no sistema', 'Só no sistema (não achado no banco)', COLS_PADRAO,
           ddaUltimo.soSis.map(d => linhaPadrao(d.venc, d.nome, d.valor)));
  montaAbaDda(wb, 'Conferidos', 'Conferidos (valor e nome batem)', COLS_PADRAO,
           ddaUltimo.ok.map(d => linhaPadrao(d.sistema.venc, d.sistema.nome, d.sistema.valor)));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'conferencia-dda.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});


/* ===================================================================
   4) CONFERÊNCIA ITAÚ
   Mesma regra da Conferência Santander (função conferir), com o DDA lido
   pelo leitorItau, filtro por unidade e as colunas extras do Itaú.
   =================================================================== */

// o Pagador vem escrito de vários jeitos (CRV INDUSTRIAL LTDA, C R V ..., CRV INDL ...):
// quem tem "CAPINOPOLIS" (com ou sem acento/espaço) é a Unidade Capinópolis; o resto é Matriz
const NOMES_UNIDADE = { matriz:'Matriz', capinopolis:'Unidade Capinópolis' };
function unidadeDoPagador(p){
  return normaliza(p).replace(/\s+/g,'').includes('CAPINOPOLIS') ? 'capinopolis' : 'matriz';
}

let itauArqBanco = null, itauArqSis = null;
let itauBanco = null, itauSistema = null, itauConsulta = '', itauUltimo = null;

registraTela('itau', { colunas:9, linhaHtml: l => linhaDdaHtml(l, true), rodapeHtml: rodapeDda(9) });

function itauCheca(){ $('itauBtnGo').disabled = !(itauArqBanco && itauArqSis); mostraArquivos('itau', [itauArqBanco, itauArqSis]); }
ligaDrop('itauDropBanco','itauFileBanco','itauTxtBanco', f => { itauArqBanco = f; itauCheca(); });
ligaDrop('itauDropSis','itauFileSis','itauTxtSis', f => { itauArqSis = f; itauCheca(); });

function itauUnidadeEscolhida(){ return $('itauUnidade').value; }

// roda a conferência com os boletos da unidade escolhida e mostra na tela
function itauMostra(){
  const unidade = itauUnidadeEscolhida();
  const banco = unidade==='todas' ? itauBanco : itauBanco.filter(b => unidadeDoPagador(b.pagador)===unidade);
  const r = conferir(banco, itauSistema);
  itauUltimo = r;

  const porUnidade = { matriz:0, capinopolis:0 };
  itauBanco.forEach(b => porUnidade[unidadeDoPagador(b.pagador)]++);
  const comDesconto = banco.filter(b => b.comDesconto).length;
  $('itauInfo').innerHTML =
    `Consulta no Itaú: <b class="mono">${escapa(itauConsulta || 'data/hora não encontrada no arquivo')}</b> · ` +
    `${itauBanco.length} boletos no arquivo (Matriz: ${porUnidade.matriz} · Unidade Capinópolis: ${porUnidade.capinopolis}) · ` +
    `conferindo <b>${unidade==='todas' ? 'todas as unidades' : NOMES_UNIDADE[unidade]}</b>: ${banco.length} boletos` +
    (comDesconto ? `, ${comDesconto} com desconto` : '');

  mostraLinhas('itau', linhasDda(r));
}

$('itauBtnGo').addEventListener('click', async () => {
  limpaErro('itauErro');
  try{
    const [mB, mS] = await Promise.all([lerMatriz(itauArqBanco), lerMatriz(itauArqSis)]);
    const bancoDoArquivo = detectaBanco(mB);
    if(bancoDoArquivo && bancoDoArquivo!=='itau') throw new Error(avisoBancoErrado(bancoDoArquivo, 'itau'));
    if(bancoDoArquivo!=='itau') throw new Error('Não reconheci o arquivo do Itaú: falta a linha de cabeçalho com "Beneficiário" e "Código de barras". Confira se é o Excel "Consulta Boletos Recebidos" do internet banking.');

    const banco   = leitorItau(mB);
    const sistema = extraiSistema(mS);

    if(!banco.length)   throw new Error('Não encontrei boletos no arquivo do Itaú. Confira se os valores estão na coluna "Até Venc.".');
    if(!sistema.length) throw new Error('Não encontrei saldos no arquivo do sistema. Confira se os nomes estão na coluna C e o saldo na coluna P.');

    itauBanco = banco;
    itauSistema = sistema;
    itauConsulta = dataHoraConsultaItau(mB);
    itauMostra();

    $('itauResultado').style.display = 'block';
    habilitaBaixar('itau');
    $('itauResultado').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    mostraErro('itauErro', err.message || 'Erro ao processar os arquivos.');
  }
});

// trocar a unidade depois de conferir refaz a conferência na hora
$('itauUnidade').addEventListener('change', () => { if(itauBanco) itauMostra(); });

$('itauBtnReset').addEventListener('click', () => location.reload());

const FMT_REAL = 'R$ #,##0.00';
// colunas que o Itaú acrescenta no Excel, antes e depois das que já existem
const COLS_ITAU_MEIO = [
  { cab:'CNPJ',             w:20 },
  { cab:'Banco',            w:8 },
  { cab:'Código de barras', w:50 },
  { cab:'Nº Doc. (Itaú)',   w:18, fmt:'@' },   // texto, para não perder zeros à esquerda
];
const COLS_ITAU_FIM = [
  { cab:'A pagar',  w:16, fmt:FMT_REAL },
  { cab:'Desconto', w:14 },
  { cab:'Unidade',  w:22 },
];
const COLS_ITAU_PADRAO = [
  { cab:'Vencimento',        w:14 },
  { cab:'Nome/Beneficiário', w:46 },
  ...COLS_ITAU_MEIO,
  { cab:'Para NF',           w:16 },
  { cab:'Pedidos',           w:16 },
  { cab:'Valor (até venc.)', w:18, fmt:FMT_REAL },
  ...COLS_ITAU_FIM,
];
const COLS_ITAU_DIV = [
  { cab:'Vencimento',        w:14 },
  { cab:'Nome no banco',     w:40 },
  { cab:'Nome no sistema',   w:40 },
  ...COLS_ITAU_MEIO,
  { cab:'Para NF',           w:16 },
  { cab:'Pedidos',           w:16 },
  { cab:'Valor (até venc.)', w:18, fmt:FMT_REAL },
  ...COLS_ITAU_FIM,
];
const COLS_ITAU_DATAS = [
  { cab:'Nome/Beneficiário',         w:46 },
  { cab:'Valor (até venc.)',         w:18, fmt:FMT_REAL },
  { cab:'Data no DDA (boleto)',      w:20 },
  { cab:'Data no sistema (lançada)', w:24 },
  ...COLS_ITAU_MEIO,
  ...COLS_ITAU_FIM,
];

$('itauBtnXlsx').addEventListener('click', async () => {
  if(!itauUltimo) return;
  const r = itauUltimo;
  const unidade = itauUnidadeEscolhida();
  const sufixo = unidade==='todas' ? '' : ` — ${NOMES_UNIDADE[unidade]}`;

  const wb = new ExcelJS.Workbook();

  const meio = b => [ b.cnpj, b.banco, b.codigoBarras, b.doc ? String(b.doc) : '' ];
  const fim  = b => [ b.valorAPagar, b.comDesconto ? 'com desconto' : '', NOMES_UNIDADE[unidadeDoPagador(b.pagador)] ];
  const linhaPadrao = (venc, nome, valor) => [ venc || '', nome || '', '', '', valor ];

  montaAbaDda(wb, 'Nome divergente', 'Nome divergente (conferir · possível factoring)' + sufixo, COLS_ITAU_DIV,
           r.div.map(d => [ d.sistema.venc || '', d.banco.nome || '', d.sistema.nome || '', ...meio(d.banco), '', '', d.sistema.valor, ...fim(d.banco) ]));

  montaAbaDda(wb, 'Datas divergentes', 'Datas divergentes (mesmo título, vencimento diferente)' + sufixo, COLS_ITAU_DATAS,
           r.datasDiff.map(d => [ d.sistema.nome, d.sistema.valor, d.banco.venc, d.sistema.venc, ...meio(d.banco), ...fim(d.banco) ]));

  montaAbaDda(wb, 'So no banco', 'Só no banco (não achado no sistema)' + sufixo, COLS_ITAU_PADRAO,
           r.soBanco.map(d => [ d.venc || '', d.nome || '', ...meio(d), '', '', d.valor, ...fim(d) ]));
  montaAbaDda(wb, 'So no sistema', 'Só no sistema (não achado no banco)' + sufixo, COLS_PADRAO,
           r.soSis.map(d => linhaPadrao(d.venc, d.nome, d.valor)));
  montaAbaDda(wb, 'Conferidos', 'Conferidos (valor e nome batem)' + sufixo, COLS_ITAU_PADRAO,
           r.ok.map(d => [ d.sistema.venc || '', d.sistema.nome || '', ...meio(d.banco), '', '', d.sistema.valor, ...fim(d.banco) ]));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'conferencia-itau.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});


/* ===================================================================
   5) CONFERÊNCIA DE CÓDIGO DE BARRAS
   =================================================================== */

// maiúsculo, sem acento, sem espaço nas pontas — pra comparar cabeçalhos
function normaliza(s){
  if(s===null || s===undefined) return '';
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim();
}

// deixa só os dígitos de um código de barras (tira ponto, espaço, etc.)
function apenasDigitos(v){
  if(v===null || v===undefined) return '';
  return String(v).replace(/\D+/g,'');
}

// linha digitável de boleto: o valor são os ÚLTIMOS 10 dígitos, em centavos
function valorDoCodigoBarras(digitos){
  const ult10 = digitos.slice(-10);
  const n = parseInt(ult10,10);
  return isNaN(n) ? null : Math.round(n)/100;
}

// procura a linha que tem, ao mesmo tempo, uma célula com "CÓDIGO DE BARRAS"
// (o subtítulo do relatório só tem "Cód. Barras" abreviado, então exigir a
// palavra completa "CODIGO" evita pegar essa linha errada) e uma com "SALDO".
function acharCabecalho(m){
  for(let i=0;i<m.length;i++){
    const linha = m[i] || [];
    let temCodBarras = false, temSaldo = false;
    linha.forEach(cel=>{
      const n = normaliza(cel);
      if(n.includes('CODIGO') && n.includes('BARRA')) temCodBarras = true;
      if(n.includes('SALDO')) temSaldo = true;
    });
    if(temCodBarras && temSaldo) return i;
  }
  return -1;
}

// mapeia as colunas pelo texto do cabeçalho; usa fallback fixo se não achar
function mapeiaColunas(linhaCab){
  const cols = { venc:-1, cliente:-1, nome:-1, duplicata:-1, barra:-1, saldo:-1 };
  (linhaCab || []).forEach((cel,i)=>{
    const n = normaliza(cel);
    if(cols.barra<0      && n.includes('BARRA'))     cols.barra = i;
    if(cols.saldo<0      && n.includes('SALDO'))      cols.saldo = i;
    if(cols.venc<0        && n.includes('VENC'))        cols.venc = i;
    if(cols.duplicata<0 && n.includes('DUPLICATA')) cols.duplicata = i;
    if(cols.cliente<0    && n.includes('CLIENTE'))    cols.cliente = i;
  });
  cols.nome = cols.cliente>=0 ? cols.cliente+1 : -1;

  // fallback fixo: A=Vencto, B=Cliente, C=Nome, D=Duplicata, K=Código de barras, L=Saldo
  if(cols.venc<0)       cols.venc = 0;
  if(cols.nome<0)       cols.nome = 2;
  if(cols.duplicata<0)  cols.duplicata = 3;
  if(cols.barra<0)      cols.barra = 10;
  if(cols.saldo<0)       cols.saldo = 11;
  return cols;
}

// extrai os títulos com código de barras válido (>=40 dígitos) e saldo numérico
function extraiTitulos(m, cols, headerIdx){
  const out = [];
  for(let i=headerIdx+1; i<m.length; i++){
    const linha = m[i] || [];
    const saldo = paraNumero(linha[cols.saldo]);
    if(saldo===null) continue;
    const digitos = apenasDigitos(linha[cols.barra]);
    if(digitos.length<40) continue;
    out.push({
      venc: formatData(linha[cols.venc]),
      nome: linha[cols.nome]!=null ? String(linha[cols.nome]).trim() : '',
      duplicata: linha[cols.duplicata]!=null ? String(linha[cols.duplicata]).trim() : '',
      codigo: digitos,
      valorCodigo: valorDoCodigoBarras(digitos),
      saldo,
    });
  }
  return out;
}

function achaDivergencias(titulos){
  return titulos
    .map(t => ({ ...t, diff: Math.round((t.valorCodigo - t.saldo)*100)/100 }))
    .filter(t => Math.abs(t.diff) >= 0.005)
    .sort((a,b) => Math.abs(b.diff) - Math.abs(a.diff));
}

function achaDuplicados(titulos){
  const porCodigo = new Map();
  titulos.forEach(t => {
    if(!porCodigo.has(t.codigo)) porCodigo.set(t.codigo, []);
    porCodigo.get(t.codigo).push(t);
  });
  const grupos = [];
  porCodigo.forEach(lista => { if(lista.length>1) grupos.push(lista); });
  return grupos;
}

// linhas da tabela do código de barras. Cada título aparece uma vez:
// primeiro as divergências (maior diferença primeiro), depois os duplicados
// (lado a lado, pelo número do grupo) e por fim os demais.
// Um título pode estar em dois grupos ao mesmo tempo (divergente e duplicado).
function linhasCb(titulos, divergencias, duplicados){
  const divDoTitulo = new Map(divergencias.map(d => [d.idx, d]));
  const grupoDup = new Map();
  duplicados.forEach((lista, gi) => lista.forEach(t => grupoDup.set(t.idx, gi+1)));

  const ordem = [
    ...divergencias.map(d => titulos[d.idx]),
    ...duplicados.flat().filter(t => !divDoTitulo.has(t.idx)),
    ...titulos.filter(t => !divDoTitulo.has(t.idx) && !grupoDup.has(t.idx)),
  ];
  return ordem.map(t => {
    const div = divDoTitulo.get(t.idx);
    const dup = grupoDup.get(t.idx);
    const grupos = [div ? 'valor' : 'ok'];
    if(dup) grupos.push('dup');
    const textos = [t.nome, t.duplicata, dataBR(t.venc), t.codigo, brl(t.saldo), brl(t.valorCodigo)];
    const digitos = [centavos(t.saldo), centavos(t.valorCodigo), apenasDigitos(t.duplicata)];
    return { titulo:t, diff: div ? div.diff : null, dup, grupos, valor:t.saldo,
             texto: normaliza(textos.join(' | ')), digitos: digitos.filter(Boolean).join('|') };
  });
}

function linhaCbHtml(l){
  const t = l.titulo;
  let status = etiqueta(l.grupos[0]);
  if(l.dup) status += `<br>${etiqueta('dup', 'Duplicado · grupo ' + l.dup)}`;
  const diff = l.diff===null ? '—' : `${l.diff>=0 ? '+' : ''}${brl(l.diff)}`;
  return `<tr><td>${status}</td><td class="mono">${escapa(dataBR(t.venc))}</td><td class="forn">${escapa(t.nome)}</td>` +
         `<td class="mono">${escapa(t.duplicata) || '—'}</td><td class="mono cod">${escapa(t.codigo)}</td>` +
         `<td class="num mono">${brl(t.valorCodigo)}</td><td class="num mono">${brl(t.saldo)}</td>` +
         `<td class="num mono${l.diff===null ? '' : ' diff'}">${diff}</td></tr>`;
}

function rodapeCb(linhas){
  const soma = campo => Math.round(linhas.reduce((s, l) => s + (l.titulo[campo] || 0), 0) * 100) / 100;
  const somaDiff = Math.round(linhas.reduce((s, l) => s + (l.diff || 0), 0) * 100) / 100;
  return `<tr><td colspan="5">Total · ${qtdTitulos(linhas.length)}</td>` +
         `<td class="num mono">${brl(soma('valorCodigo'))}</td><td class="num mono">${brl(soma('saldo'))}</td>` +
         `<td class="num mono">${somaDiff ? (somaDiff>0 ? '+' : '') + brl(somaDiff) : '—'}</td></tr>`;
}

let cbArquivo = null, cbUltimo = null;

registraTela('cb', { colunas:8, linhaHtml: linhaCbHtml, rodapeHtml: rodapeCb });

ligaDrop('cbDrop','cbFile','cbTxt', f => { cbArquivo = f; $('cbBtnGo').disabled = false; mostraArquivos('cb', [f]); });

$('cbBtnGo').addEventListener('click', async () => {
  limpaErro('cbErro');
  try{
    const m = await lerMatriz(cbArquivo);

    const headerIdx = acharCabecalho(m);
    if(headerIdx<0) throw new Error('Não encontrei a linha de cabeçalho (preciso de uma linha com "Código de Barras" e "Saldo" juntos).');

    const cols = mapeiaColunas(m[headerIdx]);
    const titulos = extraiTitulos(m, cols, headerIdx);
    if(!titulos.length) throw new Error('Não encontrei títulos com código de barras válido (40+ dígitos) e saldo numérico.');
    // número de cada título, só pra tabela da tela saber quem é quem (o Excel não usa)
    titulos.forEach((t, i) => { t.idx = i; });

    const divergencias = achaDivergencias(titulos);
    const duplicados = achaDuplicados(titulos);

    cbUltimo = { titulos, divergencias, duplicados };

    mostraLinhas('cb', linhasCb(titulos, divergencias, duplicados));

    $('cbResultado').style.display = 'block';
    habilitaBaixar('cb');
    $('cbResultado').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    mostraErro('cbErro', err.message || 'Erro ao processar o arquivo.');
  }
});

$('cbBtnReset').addEventListener('click', () => location.reload());

const AZUL_CAB = 'FF1F4E8C';
const BRANCO = 'FFFFFFFF';

function montaAbaCb(wb, nome, colunas, linhas){
  const ws = wb.addWorksheet(nome);
  ws.columns = colunas;
  ws.views = [{ state:'frozen', ySplit:1 }];

  const cab = ws.getRow(1);
  cab.eachCell(c => {
    c.font = { bold:true, color:{argb:BRANCO} };
    c.fill = { type:'pattern', pattern:'solid', fgColor:{argb:AZUL_CAB} };
    c.alignment = { vertical:'middle', horizontal:'center' };
    c.border = BORDA_FINA;
  });
  cab.height = 20;

  if(!linhas.length){
    ws.mergeCells(2,1,2,colunas.length);
    ws.getCell(2,1).value = 'Nenhum registro nesta categoria.';
    ws.getCell(2,1).font = { italic:true, color:{argb:'FF5C6B78'} };
    return ws;
  }

  linhas.forEach(obj => {
    const r = ws.addRow(obj);
    r.eachCell({ includeEmpty:true }, c => { c.border = BORDA_FINA; });
  });

  colunas.forEach(c => {
    if(c.fmt){
      const col = ws.getColumn(c.key);
      col.numFmt = c.fmt;
      col.alignment = { horizontal:'right' };
    }
  });

  return ws;
}

$('cbBtnXlsx').addEventListener('click', async () => {
  if(!cbUltimo) return;

  const wb = new ExcelJS.Workbook();

  const COLS_DIV_CB = [
    { header:'Vencto.',           key:'venc',        width:14 },
    { header:'Nome',              key:'nome',         width:40 },
    { header:'Duplicata',         key:'duplicata',   width:16 },
    { header:'Código de barras',  key:'codigo',        width:50 },
    { header:'Valor no código',   key:'valorCodigo', width:18, fmt:'R$ #,##0.00' },
    { header:'Saldo lançado',     key:'saldo',        width:18, fmt:'R$ #,##0.00' },
    { header:'Diferença',         key:'diff',          width:16, fmt:'R$ #,##0.00' },
  ];
  const COLS_DUP_CB = [
    { header:'Grupo',             key:'grupo',        width:10 },
    { header:'Vencto.',           key:'venc',           width:14 },
    { header:'Nome',              key:'nome',            width:40 },
    { header:'Duplicata',         key:'duplicata',      width:16 },
    { header:'Código de barras',  key:'codigo',           width:50 },
    { header:'Saldo lançado',     key:'saldo',           width:18, fmt:'R$ #,##0.00' },
  ];

  montaAbaCb(wb, 'Divergências de Valor', COLS_DIV_CB, cbUltimo.divergencias);

  const linhasDup = cbUltimo.duplicados.flatMap((lista,gi) =>
    lista.map(t => ({ ...t, grupo: gi+1 })));
  montaAbaCb(wb, 'Códigos Duplicados', COLS_DUP_CB, linhasDup);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'Conferencia_Codigo_Barras.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});


/* ===================================================================
   6) AJUDA NOS CARTÕES DE RESUMO
   Cada cartão com data-ajuda="..." no index.html ganha um "?" que explica
   o grupo. Abre ao passar o mouse, ao tocar no "?" ou pelo teclado (Tab).
   =================================================================== */

const AJUDA_RESUMO = {
  conferidos: 'Boletos que bateram em tudo com o sistema: valor, nome do fornecedor e vencimento. Não precisa fazer nada.',
  nomeDiv:    'O valor bateu, mas o nome do fornecedor é diferente. Pode ser factoring (outra empresa cobrando no lugar do fornecedor). Confirme com o fornecedor antes de pagar.',
  dataDiv:    'É o mesmo boleto (valor e nome batem), mas o vencimento no sistema está diferente do boleto. Acerte a data no sistema.',
  soBanco:    'Boletos que estão no banco, mas não foram achados no sistema. Veja se falta lançar a nota. Se não reconhecer o boleto, não pague.',
  soSis:      'Títulos lançados no sistema que não apareceram no banco. Pode ser conta de consumo, pagamento por Pix ou transferência, ou boleto que ainda não foi emitido.',
  cbDiv:      'O valor que está no código de barras é diferente do saldo lançado. Corrija o lançamento ou peça o boleto certo ao fornecedor.',
  cbDup:      'O mesmo código de barras foi lançado em mais de um título. Risco de pagar duas vezes: apague o lançamento repetido.',
  cbOk:       'Títulos em que o valor do código de barras bate com o saldo lançado.',
  bolOk:      'Títulos achados no DDA um para um, sem nenhuma pendência, com o código conferido (tamanho, dígitos, banco e valor). Pode copiar a linha digitável e pagar.',
  bolVerif:   'Precisa de atenção antes de pagar: boletos agrupados (azul), casos com algo a conferir (amarelo) e códigos bloqueados (vermelho). A coluna "O que verificar" diz o quê.',
  bolSem:     'Títulos do sistema sem boleto no DDA (pode ser Pix, transferência ou boleto ainda não emitido) e contas de consumo, marcadas como "conta de consumo".',
};

function fechaDicas(exceto){
  document.querySelectorAll('.stat.dica-aberta').forEach(card => {
    if(card===exceto) return;
    card.classList.remove('dica-aberta');
    card.querySelector('.ajuda-btn').setAttribute('aria-expanded', 'false');
  });
}

document.querySelectorAll('.stat[data-ajuda]').forEach((card, i) => {
  const texto = AJUDA_RESUMO[card.dataset.ajuda];
  if(!texto) return;
  const nome = card.querySelector('.stat-nome').textContent;
  card.insertAdjacentHTML('beforeend',
    `<button type="button" class="ajuda-btn" aria-label="O que é ${escapa(nome)}?" aria-expanded="false" aria-describedby="dica-${i}">?</button>` +
    `<span class="dica" role="tooltip" id="dica-${i}">${texto}</span>`);
  const btn = card.querySelector('.ajuda-btn');
  btn.addEventListener('click', e => {
    e.stopPropagation();
    const abrir = !card.classList.contains('dica-aberta');
    fechaDicas(card);
    card.classList.toggle('dica-aberta', abrir);
    btn.setAttribute('aria-expanded', String(abrir));
  });
});
document.addEventListener('click', () => fechaDicas());
document.addEventListener('keydown', e => { if(e.key === 'Escape') fechaDicas(); });


/* ===================================================================
   7) BOLETO (linha digitável e código de barras)
   Funções puras, sem tela: conferem se um código de boleto bancário está
   íntegro antes de alguém copiar e pagar. Nada aqui muda as outras telas.
   Índices abaixo começam em 0, como no JavaScript (cb[0..3] = 4 primeiros).
   Código de barras (44): banco(0..2) moeda(3) DV geral(4) fator(5..8)
     valor(9..18) campo livre(19..43).
   Linha digitável (47): campo1(0..9) campo2(10..20) campo3(21..31)
     DV geral(32) fator+valor(33..46).
   =================================================================== */

// só os dígitos: tira espaços, pontos e traços. Devolve null se sobrar outra
// coisa (letra, vírgula...), por exemplo notação científica "3,4191E+43"
function soDigitos(texto){
  if(texto===null || texto===undefined) return null;
  const limpo = String(texto).replace(/[\s.\-]+/g, '');
  return /^\d*$/.test(limpo) ? limpo : null;
}
const ehNotacaoCientifica = texto => /\d[.,]?\d*\s*E\s*[+-]?\s*\d+/i.test(String(texto || ''));

// DV de um campo da linha digitável: da direita para a esquerda, pesos 2,1,2,1...
// soma os dígitos de cada produto (12 -> 1+2)
function dvMod10(campo){
  let soma = 0;
  for(let i = campo.length - 1, peso = 2; i >= 0; i--, peso = peso===2 ? 1 : 2){
    const p = Number(campo[i]) * peso;
    soma += p > 9 ? p - 9 : p;
  }
  return (10 - soma % 10) % 10;
}

// DV geral do código de barras: os 43 dígitos sem a posição 4 (o próprio DV),
// pesos 2 a 9 da direita para a esquerda, repetindo. 0, 10 ou 11 viram 1.
function dvMod11Geral(cb44){
  const base = cb44.slice(0, 4) + cb44.slice(5);
  let soma = 0;
  for(let i = base.length - 1, peso = 2; i >= 0; i--, peso = peso===9 ? 2 : peso + 1){
    soma += Number(base[i]) * peso;
  }
  const dv = 11 - (soma % 11);
  return (dv===0 || dv===10 || dv===11) ? 1 : dv;
}

// código de barras (44) -> linha digitável (47), só números
function barrasParaLinha(cb44){
  const c1 = cb44.slice(0, 4) + cb44.slice(19, 24);
  const c2 = cb44.slice(24, 34);
  const c3 = cb44.slice(34, 44);
  return c1 + dvMod10(c1) + c2 + dvMod10(c2) + c3 + dvMod10(c3) + cb44[4] + cb44.slice(5, 19);
}

// linha digitável (47) -> código de barras (44)
function linhaParaBarras(l47){
  return l47.slice(0, 4) + l47[32] + l47.slice(33, 47) + l47.slice(4, 9) + l47.slice(10, 20) + l47.slice(21, 31);
}

// o que está gravado no código: banco, vencimento (dd/mm/aaaa) e valor.
// Fator 0000 = boleto sem vencimento (venc = '').
// Desde 22/02/2025 o fator recomeçou em 1000: venc = 22/02/2025 + (fator - 1000) dias.
function decodifica(cb44){
  const fator = Number(cb44.slice(5, 9));
  let venc = '';
  if(fator > 0){
    const d = new Date(Date.UTC(2025, 1, 22) + (fator - 1000) * 86400000);
    venc = `${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCMonth()+1).padStart(2,'0')}/${d.getUTCFullYear()}`;
  }
  return { banco: cb44.slice(0, 3), fator, venc, valor: Number(cb44.slice(9, 19)) / 100 };
}

// confere só a estrutura do código (regras a, b e c): tamanho, DVs e ida e volta.
// Devolve { ok, tipo, linha, cb, motivo }. Não compara com nada do DDA.
// tipo: 'boleto' | 'consumo' (conta de consumo/convênio, começa com 8) | 'invalido'
function validaEstrutura(texto){
  const bloqueia = (motivo, tipo) => ({ ok:false, tipo: tipo || 'invalido', linha:'', cb:'', motivo });

  if(ehNotacaoCientifica(texto)) return bloqueia('código ilegível: veio em notação científica (o Excel gravou como número e perdeu dígitos)');
  const d = soDigitos(texto);
  if(d===null) return bloqueia('o código tem caracteres que não são números');
  if(!d) return bloqueia('sem código de barras');

  // a) tamanho: 44 (código de barras) ou 47 (linha digitável); 48 ou começo 8 = conta de consumo
  if(d[0]==='8' && (d.length===48 || d.length===44)) return bloqueia('conta de consumo (não é boleto bancário)', 'consumo');
  let cb;
  if(d.length===44) cb = d;
  else if(d.length===47){
    // b) os 3 DVs de campo que vieram na linha digitável
    const campos = [[0, 9], [10, 20], [21, 31]];
    for(let k=0; k<3; k++){
      const [ini, posDv] = campos[k];
      if(dvMod10(d.slice(ini, posDv)) !== Number(d[posDv])) return bloqueia(`dígito verificador do campo ${k+1} da linha digitável não confere`);
    }
    cb = linhaParaBarras(d);
  }
  else return bloqueia(`tem ${d.length} dígitos (o boleto tem 44 no código de barras ou 47 na linha digitável)`);

  // b) DV geral
  if(dvMod11Geral(cb) !== Number(cb[4])) return bloqueia('dígito verificador geral (módulo 11) não confere');

  // c) ida e volta
  const linha = barrasParaLinha(cb);
  if(linhaParaBarras(linha) !== cb) return bloqueia('conversão código de barras ↔ linha digitável não fecha');
  if(d.length===47 && linha !== d) return bloqueia('a linha digitável não bate com o código de barras dela');

  return { ok:true, tipo:'boleto', linha, cb, motivo:'' };
}

// confere um código de boleto (linha digitável ou código de barras) contra o DDA.
// esperado = { banco, valor, venc } do DDA (coluna Banco, "Até Venc." e "Venc.").
// Devolve { ok, tipo, linha, cb, motivo, avisos }. Se não for ok, linha e cb vêm
// vazios: o boleto fica BLOQUEADO (nada para copiar) e motivo diz por quê.
function validaBoleto(texto, esperado){
  const esp = esperado || {};
  const avisos = [];
  const bloqueia = (motivo, tipo) => ({ ok:false, tipo: tipo || 'invalido', linha:'', cb:'', motivo, avisos });

  const e = validaEstrutura(texto);
  if(!e.ok) return bloqueia(e.motivo, e.tipo);
  const { linha, cb } = e;

  // d) bate com o DDA
  const dec = decodifica(cb);
  const bancoDda = apenasDigitos(esp.banco);
  if(!bancoDda) return bloqueia('o DDA não tem o banco para comparar');
  if(bancoDda.padStart(3, '0') !== dec.banco) return bloqueia(`banco do código (${dec.banco}) diferente do banco no DDA (${esp.banco})`);
  const valorDda = paraNumero(esp.valor);
  if(valorDda===null) return bloqueia('o DDA não tem o valor para comparar');
  if(Math.round(valorDda * 100) !== Math.round(dec.valor * 100)) return bloqueia(`valor do código (${brl(dec.valor)}) diferente do "Até Venc." no DDA (${brl(valorDda)})`);

  // vencimento diferente só avisa
  if(formatData(esp.venc) !== dec.venc) avisos.push('vencimento do código diferente do DDA (boleto prorrogado?)');

  return { ok:true, tipo:'boleto', linha, cb, motivo:'', avisos };
}

// linha digitável no formato impresso no boleto, só para mostrar:
// 00190.50095 40144.816069 06809.350314 3 37370000000100
function formataLinha(l47){
  if(!l47 || l47.length!==47) return l47 || '';
  return `${l47.slice(0,5)}.${l47.slice(5,10)} ${l47.slice(10,15)}.${l47.slice(15,21)} ` +
         `${l47.slice(21,26)}.${l47.slice(26,32)} ${l47[32]} ${l47.slice(33)}`;
}


/* ===================================================================
   8) CONFERÊNCIA DE BOLETOS
   DDA do Itaú x "Duplicatas em Aberto Modelo A" do sistema. Para cada título
   do sistema SEM código de barras, acha o boleto no DDA (regras em
   PAREAMENTO, abaixo: nome, NF e valor) e entrega a linha digitável,
   mas só se o código do DDA passar no validaBoleto (seção 7).
   Títulos que JÁ têm código no sistema ficam fora da conferência; o código
   deles só é conferido. Nada aqui corrige o código do sistema.
   Grupos:
     bolOk    Com boleto     título achado no DDA (nome + valor, ou boleto agrupado), código válido
     bolVerif Verificação    código bloqueado, nome ou valor diferente, ou código do sistema errado
     bolJa    Já tem código  código no sistema válido e igual ao do DDA (ou conta de consumo)
     bolSem   Sem boleto     título sem código que não está no DDA
     bolDda   Só no DDA      boleto do DDA sem título no sistema
   =================================================================== */

// SISTEMA ("Duplicatas em Aberto Modelo A"): o cabeçalho é a linha com
// "Código de barras" e "Saldo" (acharCabecalho); as colunas são achadas pelo
// texto (mapeiaColunas) + Prf. Entram só as linhas com código do fornecedor
// (coluna Cliente) numérico: subtotais e totais ficam de fora.
// leitura = lerMatrizComTexto(arquivo). O código de barras vem do texto da célula.
function leitorSistemaBoletos(leitura){
  const m = leitura.matriz;
  const cab = acharCabecalho(m);
  if(cab<0) throw new Error('Não reconheci o relatório do sistema: falta a linha de cabeçalho com "Código de barras" e "Saldo". Confira se é o "Duplicatas em Aberto Modelo A".');
  const nomes = (m[cab] || []).map(normaliza);
  const cols = mapeiaColunas(m[cab]);
  cols.prf = nomes.findIndex(c => c.startsWith('PRF'));
  const falta = [['Cliente', cols.cliente], ['Vencto.', nomes.findIndex(c => c.includes('VENC'))],
                 ['Duplicata', nomes.findIndex(c => c.includes('DUPLICATA'))]].filter(([, i]) => i<0).map(([n]) => n);
  if(falta.length) throw new Error(`No relatório do sistema não achei a coluna ${falta.join(', ')}. Confira se é o "Duplicatas em Aberto Modelo A".`);

  const texto = v => v!=null ? String(v).trim() : '';
  const out = [];
  for(let i=cab+1; i<m.length; i++){
    const linha = m[i] || [];
    if(!/^\d+$/.test(texto(linha[cols.cliente]))) continue;
    const saldo = paraNumero(linha[cols.saldo]);
    if(saldo===null) continue;
    const cel = leitura.textoDe(i, cols.barra);
    out.push({
      venc: formatData(linha[cols.venc]),
      fornecedor: texto(linha[cols.cliente]),
      nome: texto(linha[cols.nome]),
      doc: texto(linha[cols.duplicata]),
      prf: cols.prf>=0 ? texto(linha[cols.prf]) : '',
      valor: saldo,
      codigoTexto: cel.texto.trim(),
      codigoSuspeito: !cel.eraTexto,
      temCodigo: /\d/.test(cel.texto),
      linhaExcel: leitura.linhaExcel(i),
    });
  }
  return out;
}

/* -------------------------------------------------------------------
   CHECAGENS DOS ARQUIVOS (avisos no topo, antes do resultado)
   Conferem se os dois arquivos combinam: mesma data, mesma empresa e
   período. Não mudam o resultado da conferência, só avisam.
   ------------------------------------------------------------------- */

// datas dd/mm/aaaa (ou dd/mm/aa) que aparecem num texto ou numa célula de data,
// na ordem em que aparecem. Cada uma: { txt:'dd/mm/aaaa', t: milissegundos UTC }
function datasDoTexto(v){
  const s = formatData(v);
  const out = [];
  for(const m of s.matchAll(/(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})(?!\d)/g)){
    const d = Number(m[1]), mes = Number(m[2]), a = m[3].length===2 ? 2000 + Number(m[3]) : Number(m[3]);
    const t = Date.UTC(a, mes - 1, d);
    const dt = new Date(t);
    if(dt.getUTCDate()!==d || dt.getUTCMonth()!==mes-1) continue;   // 31/02 etc.
    out.push({ txt:`${String(d).padStart(2,'0')}/${String(mes).padStart(2,'0')}/${a}`, t });
  }
  return out;
}
const DIA_MS = 86400000;
const NOMES_DIA = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
const txtData = t => { const d = new Date(t); return `${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCMonth()+1).padStart(2,'0')}/${d.getUTCFullYear()}`; };

// valor de um rótulo nas primeiras linhas (antes do cabeçalho): "Rótulo: valor" na
// mesma célula, ou o rótulo numa célula e o valor nas seguintes da mesma linha
function valorDoRotulo(m, fim, teste){
  for(let i=0; i<fim; i++){
    const linha = m[i] || [];
    for(let j=0; j<linha.length; j++){
      const n = normaliza(linha[j]);
      if(!n || !teste(n)) continue;
      const s = formatData(linha[j]);
      const depois = (s.includes(':') ? s.slice(s.indexOf(':') + 1) : s.replace(/^\S+/, '')).trim();
      if(depois) return depois;
      return linha.slice(j+1).filter(c => c!==null && c!=='').map(formatData).join(' ').trim();
    }
  }
  return '';
}

// topo do "Duplicatas em Aberto Modelo A": código da empresa (A1), data/hora
// do relatório (linha 1) e período (célula "Período")
function topoSistemaBoletos(m){
  const cab = acharCabecalho(m);
  const fim = cab>=0 ? cab : Math.min(m.length, 15);
  const a1 = formatData((m[0] || [])[0]);
  const empresa = (a1.match(/\d+/) || [''])[0];
  // data/hora: a primeira data da linha 1 (e a hora, se vier junto)
  let dataHora = '', data = null;
  for(const cel of (m[0] || [])){
    if(cel===null || cel==='') continue;
    const ds = datasDoTexto(cel);
    if(!ds.length) continue;
    data = ds[0];
    const s = formatData(cel);
    const hora = s.match(/\d{1,2}:\d{2}(:\d{2})?/) ||
                 formatData((m[0] || [])[(m[0] || []).indexOf(cel) + 1]).match(/^\d{1,2}:\d{2}(:\d{2})?$/);
    dataHora = data.txt + (hora ? ' ' + hora[0] : '');
    break;
  }
  const periodoTxt = valorDoRotulo(m, fim, n => n.startsWith('PERIODO'));
  const ds = datasDoTexto(periodoTxt);
  const periodo = ds.length ? { ini: ds[0], fim: ds[ds.length-1], txt: periodoTxt } : null;
  return { empresa, a1, dataHora, data, periodo };
}

// nome da empresa e CNPJ do DDA do Itaú (linhas antes do cabeçalho)
function contaDoDdaItau(m){
  const cab = achaCabecalhoItau(m);
  const fim = cab>=0 ? cab : Math.min(m.length, 15);
  return {
    nome: valorDoRotulo(m, fim, n => n.startsWith('NOME DA EMPRESA')),
    cnpj: valorDoRotulo(m, fim, n => n.startsWith('CPF/CNPJ') || n.startsWith('CNPJ')),
  };
}

// período efetivo: se começa numa segunda-feira, inclui o sábado e o domingo
// anteriores (boleto que vence no fim de semana é pago na segunda)
function periodoEfetivo(periodo){
  if(!periodo) return null;
  const segunda = new Date(periodo.ini.t).getUTCDay() === 1;
  return { ini: segunda ? periodo.ini.t - 2*DIA_MS : periodo.ini.t, fim: periodo.fim.t, segunda };
}

// devolve a lista de checagens: { tipo:'ok'|'aviso'|'alerta', titulo, texto, itens:[...] }
// e marca nos boletos de fim de semana o aviso com a data original.
// linhas = resultado do conferirBoletos; dda = boletos lidos do DDA.
function checagensBoletos({ topo, consultaDda, conta, dda, linhas }){
  const out = [];

  // 1) mesma data
  const dDda = datasDoTexto(consultaDda)[0];
  if(!dDda || !topo.data){
    out.push({ tipo:'aviso', titulo:'Data dos arquivos',
      texto:`Não consegui comparar as datas. DDA: ${consultaDda || 'data/hora não encontrada'} · Relatório do sistema: ${topo.dataHora || 'data/hora não encontrada na linha 1'}.` });
  }else if(dDda.t !== topo.data.t){
    out.push({ tipo:'aviso', titulo:'Arquivos de dias diferentes',
      texto:`O DDA foi consultado em ${consultaDda} e o relatório do sistema é de ${topo.dataHora}. Gere os dois no mesmo dia para não faltar boleto ou título.` });
  }else{
    out.push({ tipo:'ok', titulo:'Mesma data', texto:`DDA consultado em ${consultaDda} · relatório do sistema de ${topo.dataHora}.` });
  }

  // 2) período: boletos do DDA que vencem fora dele
  const per = periodoEfetivo(topo.periodo);
  const dentro = [];
  if(!per){
    out.push({ tipo:'aviso', titulo:'Período do relatório',
      texto:'Não achei a célula "Período" no relatório do sistema, então não contei os boletos do DDA que vencem fora dele.' });
    dentro.push(...dda);
  }else{
    const fora = [], fimDeSemana = [];
    dda.forEach(b => {
      const v = datasDoTexto(b.venc)[0];
      if(!v || v.t < per.ini || v.t > per.fim){ fora.push(b); return; }
      dentro.push(b);
      if(per.segunda && v.t < topo.periodo.ini.t) fimDeSemana.push({ b, v });
    });
    const p = topo.periodo;
    const txtPer = p.ini.t===p.fim.t ? p.ini.txt : `${p.ini.txt} a ${p.fim.txt}`;
    const itens = [];
    if(fimDeSemana.length){
      itens.push(`O período começa numa segunda-feira: ${fimDeSemana.length} ${fimDeSemana.length===1 ? 'boleto que vence' : 'boletos que vencem'} no sábado ou domingo anterior ${fimDeSemana.length===1 ? 'entrou' : 'entraram'} no período (é pago na segunda):`);
      fimDeSemana.forEach(({ b, v }) => itens.push(`${NOMES_DIA[new Date(v.t).getUTCDay()]} ${v.txt} · ${b.beneficiario} · ${brl(b.valor)}`));
    }
    if(fora.length){
      itens.push(`${fora.length} ${fora.length===1 ? 'boleto vence' : 'boletos vencem'} fora do período:`);
      fora.forEach(b => itens.push(`${dataBR(b.venc) || 'sem vencimento'} · ${b.beneficiario} · ${brl(b.valor)}`));
    }
    out.push({ tipo: fora.length ? 'aviso' : 'ok', titulo:`Período do relatório: ${txtPer}`,
      texto: fora.length ? `${fora.length} de ${dda.length} boletos do DDA vencem fora do período do relatório; eles aparecem na Verificação como "boleto no DDA sem título no sistema".`
                         : `Todos os ${dda.length} boletos do DDA vencem dentro do período${per.segunda ? ' (contando o fim de semana anterior)' : ''}.`,
      itens });
    // aviso na própria linha do boleto de fim de semana, com a data original
    fimDeSemana.forEach(({ b, v }) => {
      linhas.filter(l => l.boleto && l.boleto.idx===dda.indexOf(b)).forEach(l =>
        l.avisos.push(`vence ${NOMES_DIA[new Date(v.t).getUTCDay()]} ${v.txt}; pago na segunda ${p.ini.txt}`));
    });
  }

  // 3) mesma empresa: quantos boletos do DDA (dentro do período) casaram com o sistema
  const casados = new Set(linhas.filter(l => l.boleto && l.titulo).map(l => l.boleto.idx));
  const base = dentro.length ? dentro : dda;
  const nCasados = base.filter(b => casados.has(dda.indexOf(b))).length;
  const pct = base.length ? nCasados / base.length : 0;
  const quem = `Relatório do sistema: empresa ${topo.empresa || '(não achei o código na célula A1)'} · ` +
               `DDA: ${conta.nome || 'nome não encontrado'}${conta.cnpj ? ' · CNPJ ' + conta.cnpj : ''}`;
  const conta_ = `${nCasados} de ${base.length} boletos do DDA${dentro.length && per ? ' (dentro do período)' : ''} casaram com títulos do sistema (${Math.round(pct*100)}%, contando os códigos já lançados).`;
  if(pct < 0.10){
    out.unshift({ tipo:'alerta', titulo:'Os arquivos podem ser de empresas diferentes. Confira a empresa do relatório do sistema',
      texto: `${quem}. ${conta_}` });
  }else{
    out.push({ tipo:'ok', titulo:'Mesma empresa', texto:`${quem}. ${conta_}` });
  }
  return out;
}

function checagensHtml(lista){
  const icone = { ok:'✓', aviso:'!', alerta:'!' };
  return lista.map(c =>
    `<div class="checagem checagem-${c.tipo}"${c.tipo==='alerta' ? ' role="alert"' : ''}>` +
      `<span class="checagem-ico" aria-hidden="true">${icone[c.tipo]}</span>` +
      `<div><p class="checagem-titulo">${escapa(c.titulo)}</p><p class="checagem-texto">${escapa(c.texto)}</p>` +
      (c.itens && c.itens.length ? `<details class="checagem-lista"><summary>Ver a lista</summary><ul>${c.itens.map(i => `<li>${escapa(i)}</li>`).join('')}</ul></details>` : '') +
    `</div></div>`).join('');
}

// código de um boleto do DDA: bloqueia se a célula não era texto (pode ter perdido dígitos)
function validaBoletoDda(b){
  if(b.codigoSuspeito && b.codigoTexto)
    return { ok:false, tipo:'invalido', linha:'', cb:'', avisos:[],
             motivo:'código ilegível: no Excel do DDA o código de barras não está gravado como texto (pode ter perdido dígitos)' };
  return validaBoleto(b.codigoTexto, { banco:b.banco, valor:b.valor, venc:b.venc });
}

/* -------------------------------------------------------------------
   PAREAMENTO DA CONFERÊNCIA DE BOLETOS
   1) Um para um: título x boleto precisa de PELO MENOS 2 sinais entre
      nome, NF e valor. Só o valor nunca casa.
   2) Boleto agrupado: um boleto que paga 2 ou mais títulos do mesmo
      fornecedor, com soma exata ao centavo (subset-sum em centavos).
   3) O resto fica sem boleto.
   ------------------------------------------------------------------- */

// palavras que não ajudam a dizer quem é o fornecedor. Comparadas pelas 5 primeiras
// letras, então pegam também singular/plural (ALIMENTO/ALIMENTOS, TRANSPORTE/TRANSPORTES).
// As da primeira linha são as da regra original; as outras são genéricas de ramo,
// que apareciam em fornecedores diferentes e faziam casar o nome errado.
const PALAVRAS_IGNORADAS = new Set(['LTDA', 'S/A', 'SA', 'DE', 'DO', 'DA', 'E', 'ME', 'EPP', 'EIRELI', 'IND', 'COM',
  'COMERCIO', 'INDUSTRIA', 'INDUSTRIAL', 'BRASIL', 'PECAS', 'DISTRIBUIDORA', 'LIMITADA', 'SERVICOS', 'AUTO',
  'MAQUINAS', 'EQUIPAMENTOS', 'COMERCIAL', 'PRODUTOS',
  'ALIMENTOS', 'TRANSPORTES', 'TRANSPORTADORA', 'LOGISTICA', 'MOTORES', 'TINTAS', 'TECIDOS', 'QUIMICA',
  'METAIS', 'METALURGICA', 'EMBALAGENS', 'FERRAGENS', 'FERRAMENTAS', 'MATERIAIS', 'CONSTRUCAO', 'CONSTRUTORA',
  'ENGENHARIA', 'TECNOLOGIA', 'SOLUCOES', 'SISTEMAS', 'ATACADO', 'ATACADISTA', 'IMPORTADORA', 'EXPORTADORA',
  'IMPORTACAO', 'EXPORTACAO', 'NACIONAL', 'GRUPO', 'EMPRESA', 'EMPREENDIMENTOS', 'PARTICIPACOES',
  'REPRESENTACOES', 'ELETRICA', 'ELETRICOS', 'HIDRAULICA', 'AGRICOLA', 'AGROPECUARIA', 'VEICULOS',
  'ROLAMENTOS', 'PNEUS', 'COMBUSTIVEIS', 'INSUMOS', 'SUPRIMENTOS', 'MINERACAO', 'CENTRO', 'MINAS', 'GERAIS',
  'FORNECEDOR', 'FILIAL', 'MATRIZ', 'UNIDADE', 'BANCO', 'FOMENTO', 'MERCANTIL', 'SECURITIZADORA']
  .map(p => p.slice(0, 5)));

// as 5 primeiras letras de cada palavra relevante (4 letras ou mais, fora as
// ignoradas e as "genéricas" desta conferência — ver palavrasGenericas)
function palavrasDoNome(nome, genericas){
  const out = new Set();
  for(const p of normaliza(nome).match(/[A-Z]+/g) || []){
    const ini = p.slice(0, 5);
    if(p.length >= 4 && !PALAVRAS_IGNORADAS.has(ini) && !(genericas && genericas.has(ini))) out.add(ini);
  }
  return out;
}

// palavra que aparece no nome de 3 ou mais fornecedores DIFERENTES do mesmo arquivo
// (ex.: o nome da cidade) não diz quem é o fornecedor: nesta conferência ela não conta.
// Cada arquivo é contado separado, porque o mesmo fornecedor costuma vir escrito de
// um jeito no DDA e de outro no sistema.
const MIN_FORNECEDORES_GENERICA = 3;
function palavrasGenericas(...listasDeNomes){
  const genericas = new Set();
  listasDeNomes.forEach(nomes => {
    const porPalavra = new Map();
    new Set(nomes.map(n => limpaNome(normaliza(n)))).forEach(n => {
      palavrasDoNome(n).forEach(p => porPalavra.set(p, (porPalavra.get(p) || 0) + 1));
    });
    porPalavra.forEach((qtd, p) => { if(qtd >= MIN_FORNECEDORES_GENERICA) genericas.add(p); });
  });
  return genericas;
}
// as 5 primeiras letras do nome (só letras), para juntar títulos do mesmo fornecedor
const inicioDoNome = nome => (normaliza(nome).match(/[A-Z]+/g) || []).join('').slice(0, 5);

// NF sem zeros à esquerda; com menos de 3 dígitos não serve de sinal
function nfSinal(doc){
  const d = apenasDigitos(doc).replace(/^0+/, '');
  return d.length >= 3 ? d : '';
}

const emCentavos = v => (v===null || v===undefined || isNaN(v)) ? null : Math.round(v * 100);

// sinais entre um título do sistema e um boleto do DDA
function sinaisPar(t, b){
  const nome = [...t.palavras].some(p => b.palavras.has(p));
  const nf = !!t.nf && apenasDigitos(b.doc).includes(t.nf);
  const c = emCentavos(t.valor);
  const valorAteVenc = c === emCentavos(b.valor);
  const valorAPagar = !valorAteVenc && b.valorAPagar!==null && c === emCentavos(b.valorAPagar);
  const valor = valorAteVenc || valorAPagar;
  return { nome, nf, valor, valorAPagar, n: nome + nf + valor };
}

// distância em dias entre dois vencimentos (desempate: o mais perto ganha)
function distanciaDias(a, b){
  const da = datasDoTexto(a)[0], db = datasDoTexto(b)[0];
  return da && db ? Math.abs(da.t - db.t) / DIA_MS : 9999;
}

// 1) um para um. Monta todos os pares com 2 sinais ou mais e escolhe do mais forte
// para o mais fraco (3 sinais antes de 2; empate: vencimento mais perto, depois a
// ordem dos arquivos). Cada título e cada boleto entram em um par só.
function parearUmParaUm(titulos, boletos){
  const cand = [];
  titulos.forEach(t => boletos.forEach(b => {
    const s = sinaisPar(t, b);
    if(s.n >= 2) cand.push({ t, b, s, dist: distanciaDias(t.venc, b.venc) });
  }));
  cand.sort((x, y) => y.s.n - x.s.n || x.dist - y.dist || x.t.ref - y.t.ref || x.b.idx - y.b.idx);
  const tUsado = new Set(), bUsado = new Set(), pares = [];
  for(const c of cand){
    if(tUsado.has(c.t.ref) || bUsado.has(c.b.idx)) continue;
    tUsado.add(c.t.ref); bUsado.add(c.b.idx); pares.push(c);
  }
  return pares;
}

// subset-sum em centavos: acha itens cuja soma dá EXATAMENTE "alvo" e diz se essa
// combinação é a ÚNICA. Programação dinâmica: para cada soma alcançada guarda de
// onde ela veio (a soma anterior e o item somado) e de quantas formas dá para
// chegar nela (contando até 2: "uma" ou "mais de uma"); no fim refaz o caminho.
// itens: [{ c: centavos, ... }], todos positivos.
// Devolve { itens: [...] | null, unica: true/false, estourou: true/false }.
// estourou = passou de "limite" somas guardadas (desiste para não travar o navegador).
const LIMITE_SOMAS = 2000000;
function somaExata(itens, alvo, limite = LIMITE_SOMAS){
  const uteis = itens.filter(it => it.c > 0 && it.c <= alvo);
  if(alvo <= 0) return { itens:null, unica:false, estourou:false };
  // somas alcançadas e, para cada uma (mesma posição): a soma anterior, o item somado
  // e de quantas formas se chega nela (1 ou 2 = "mais de uma")
  const somas = [0], anterior = [-1], item = [-1], formas = [1];
  const pos = new Map([[0, 0]]);
  const caminho = () => {
    const usados = [];
    for(let k = pos.get(alvo); k > 0; k = pos.get(anterior[k])) usados.push(uteis[item[k]]);
    return usados.reverse();
  };
  for(let i = 0; i < uteis.length; i++){
    const v = uteis[i].c;
    const n = somas.length;
    // as formas são lidas de uma cópia de ANTES deste item, para ele não ser contado duas vezes
    const formasAntes = formas.slice(0, n);
    for(let j = 0; j < n; j++){
      const nova = somas[j] + v;
      if(nova > alvo) continue;
      const k = pos.get(nova);
      if(k !== undefined){ formas[k] = Math.min(2, formas[k] + formasAntes[j]); continue; }
      pos.set(nova, somas.length);
      somas.push(nova); anterior.push(somas[j]); item.push(i); formas.push(formasAntes[j]);
    }
    const kAlvo = pos.get(alvo);
    if(kAlvo !== undefined && formas[kAlvo] >= 2) return { itens: caminho(), unica:false, estourou:false };  // já é ambíguo
    if(somas.length > limite) return { itens:null, unica:false, estourou:true };
  }
  if(!pos.has(alvo)) return { itens:null, unica:false, estourou:false };
  return { itens: caminho(), unica: formas[pos.get(alvo)] === 1, estourou:false };
}

// combinação de 2 ou mais títulos para um boleto, com preferência para as notas
// cuja NF aparece no Nº Doc. do boleto. Devolve { titulos, unica, estourou } ou null.
function combinacaoDoBoleto(b, itens){
  const alvo = emCentavos(b.valor);
  // um título sozinho igual ao boleto não é agrupado (e não pode estar numa soma de 2+)
  const uteis = itens.filter(it => it.c < alvo);
  const docB = apenasDigitos(b.doc);
  const comNf = uteis.filter(it => it.t.nf && docB.includes(it.t.nf));
  if(comNf.length){
    const somaNf = comNf.reduce((s, it) => s + it.c, 0);
    if(somaNf === alvo && comNf.length >= 2) return { titulos: comNf.map(it => it.t), unica:true, estourou:false };
    if(somaNf < alvo){
      const r = somaExata(uteis.filter(it => !comNf.includes(it)), alvo - somaNf);
      if(r.itens) return { titulos: [...comNf, ...r.itens].map(it => it.t), unica: r.unica, estourou:false };
    }
    // as NFs do Nº Doc. não fecham a soma: tenta sem elas como preferência
  }
  const r = somaExata(uteis, alvo);
  if(r.estourou) return { titulos:[], unica:false, estourou:true };
  return r.itens ? { titulos: r.itens.map(it => it.t), unica: r.unica, estourou:false } : null;
}

// 2) boleto agrupado. Para cada boleto livre (do maior valor para o menor), junta
// os títulos sem par e os provisórios (bateram NF e nome, mas não o valor) com as
// mesmas 5 primeiras letras no nome e procura uma soma exata.
// Só agrupa quando a combinação é a ÚNICA; se houver mais de uma (ou notas demais
// para conferir), o boleto vai para "ambiguos" e nenhuma nota é marcada por ele.
function agruparBoletos(boletosLivres, titulosCandidatos){
  const grupos = [], ambiguos = [];
  const usado = new Set();
  [...boletosLivres].sort((a, b) => b.valor - a.valor || a.idx - b.idx).forEach(b => {
    const ini = inicioDoNome(b.beneficiario);
    if(!ini) return;
    const itens = titulosCandidatos
      .filter(t => !usado.has(t.ref) && inicioDoNome(t.nome) === ini)
      .map(t => ({ c: emCentavos(t.valor), t }));
    if(itens.length < 2) return;
    const achou = combinacaoDoBoleto(b, itens);
    if(!achou) return;
    if(!achou.unica){
      ambiguos.push({ b, exemplo: achou.titulos, candidatos: itens.map(it => it.t), estourou: achou.estourou });
      return;
    }
    achou.titulos.forEach(t => usado.add(t.ref));
    grupos.push({ b, titulos: achou.titulos });
  });
  return { grupos, ambiguos };
}

/* -------------------------------------------------------------------
   MOTIVOS DE VERIFICAÇÃO (além do boleto agrupado ambíguo)
   Um par título x boleto vai para a Verificação, com o motivo escrito, se:
   - o boleto está em nome de outra empresa (3 primeiras letras do primeiro
     nome diferentes, ou o nome não bateu no pareamento): "título cedido a
     factoring/fundo" se o beneficiário for FOMENTO, FIDC, FUNDO, FACTORING,
     CRED..., COBRAN... ou SECURITIZ...; senão "confirmar se é o mesmo fornecedor";
   - "A Pagar" diferente de "Até Venc.": "Desconto/abatimento: pagar R$ X";
   - bateram NF e nome mas não o valor: "Valor diferente: boleto R$ X x sistema R$ Y";
   - código do DDA inválido, código do sistema errado/fora do DDA, ou o aviso de
     vencimento do validaBoleto (boleto prorrogado?);
   - possível pagamento em duplicidade (ver marcaDuplicidade).
   "Já tem código" nunca muda de grupo (fica fora da conferência): só ganha o aviso.
   ------------------------------------------------------------------- */

// primeiro nome (primeira palavra só de letras), para comparar as 3 primeiras letras
const primeiroNome = nome => ((normaliza(nome).match(/[A-Z]+/g) || [''])[0]);
function outraEmpresa(nomeSistema, beneficiario){
  const a = primeiroNome(nomeSistema).slice(0, 3), b = primeiroNome(beneficiario).slice(0, 3);
  return !!a && !!b && a !== b;
}
const PALAVRAS_FACTORING = ['FOMENTO', 'FIDC', 'FUNDO', 'FACTORING', 'CRED', 'COBRAN', 'SECURITIZ'];
function cedidoAFactoring(beneficiario){
  const palavras = normaliza(beneficiario).match(/[A-Z]+/g) || [];
  return palavras.some(p => PALAVRAS_FACTORING.some(f => p.startsWith(f)));
}

// possível pagamento em duplicidade, entre as linhas que têm título e boleto, e os
// boletos do DDA sem título:
//  a) mesmo fornecedor (código do sistema) e mesmo valor em NFs diferentes;
//  b) a mesma linha digitável em mais de uma linha (as notas de um mesmo boleto
//     agrupado contam como uma só).
// Quem estava em "Com boleto" ou "Só no DDA" vai para a Verificação; "Já tem código"
// só ganha o aviso.
function marcaDuplicidade(linhas){
  const marca = (l, texto) => {
    if(l.grupo === 'bolJa'){ l.avisos.push(texto); return; }
    l.motivo = l.motivo ? `${l.motivo} · ${texto}` : texto;
    l.grupo = 'bolVerif';
  };
  // a) por fornecedor + valor
  const porFornValor = new Map();
  linhas.filter(l => l.titulo && l.boleto).forEach(l => {
    const k = `${l.titulo.fornecedor}|${emCentavos(l.titulo.valor)}`;
    porFornValor.set(k, [...(porFornValor.get(k) || []), l]);
  });
  porFornValor.forEach(ls => {
    const nfs = [...new Set(ls.map(l => l.titulo.doc))];
    const boletos = new Set(ls.map(l => l.boleto.idx));
    if(nfs.length < 2 || boletos.size < 2) return;   // notas do mesmo agrupado não contam
    const texto = `possível pagamento em duplicidade: mesmo fornecedor e mesmo valor (${brl(ls[0].titulo.valor)}) nas NFs ${nfs.join(', ')}`;
    ls.forEach(l => marca(l, texto));
  });
  // b) mesma linha digitável
  const porLinha = new Map();
  linhas.filter(l => l.linha).forEach(l => porLinha.set(l.linha, [...(porLinha.get(l.linha) || []), l]));
  porLinha.forEach(ls => {
    const unidades = new Set(ls.map(l => l.agrupado ? 'g' + l.agrupado.idx : 'l' + linhas.indexOf(l)));
    if(unidades.size < 2) return;
    const nfs = ls.map(l => l.titulo ? l.titulo.doc : 'só no DDA').join(', ');
    const texto = `possível pagamento em duplicidade: a mesma linha digitável aparece em ${unidades.size} lugares (${nfs})`;
    ls.forEach(l => marca(l, texto));
  });
}

// a conferência: devolve as linhas da tabela, cada uma com
// { grupo, titulo, boleto, linha, bloqueio, motivo, avisos }
//   linha    = linha digitável (47) só quando dá para usar; '' quando bloqueado
//   bloqueio = por que o código do DDA foi bloqueado
//   motivo   = o que verificar (grupo Verificação)
//   sinais   = o que bateu no um para um: 'nome', 'nf', 'valor'
//   agrupado = { n, k, idx } quando o boleto paga n títulos (este é o k-ésimo)
function conferirBoletos(dda, titulos){
  const boletos = dda.map((b, i) => ({ ...b, idx:i, val: validaBoletoDda(b) }));
  const porCb = new Map();
  boletos.forEach(b => { if(b.val.ok && !porCb.has(b.val.cb)) porCb.set(b.val.cb, b); });
  const usados = new Set();
  const linhas = [];
  const add = (grupo, titulo, boleto, extra) =>
    linhas.push({ grupo, titulo, boleto, linha:'', bloqueio:'', motivo:'', avisos:[], ...extra });

  // 1) títulos que já têm código no sistema: só confere o código
  const semCodigo = [], problemas = [];
  titulos.forEach((t, ref) => {
    const tt = { ...t, ref };
    if(!t.temCodigo){ semCodigo.push(tt); return; }
    const onde = `linha ${t.linhaExcel} do relatório`;
    if(t.codigoSuspeito){
      problemas.push({ t:tt, invalido:true, motivo:`código digitado no sistema (${onde}) ilegível: não está gravado como texto e pode ter perdido dígitos; use o do DDA` });
      return;
    }
    const e = validaEstrutura(t.codigoTexto);
    if(e.tipo==='consumo'){ add('bolJa', tt, null, { avisos:['conta de consumo (não é conferida como boleto)'] }); return; }
    if(!e.ok){ problemas.push({ t:tt, invalido:true, motivo:`código digitado no sistema (${onde}) é inválido: ${e.motivo}; use o do DDA` }); return; }
    const b = porCb.get(e.cb);
    if(!b){ problemas.push({ t:tt, motivo:`código digitado no sistema (${onde}) não está no DDA; use o do DDA` }); return; }
    if(usados.has(b.idx)){ problemas.push({ t:tt, motivo:`o mesmo código do sistema (${onde}) está em outro título; use o do DDA` }); return; }
    const v = validaBoleto(t.codigoTexto, { banco:b.banco, valor:b.valor, venc:b.venc });
    if(!v.ok){ problemas.push({ t:tt, invalido:true, motivo:`código digitado no sistema (${onde}) é inválido: ${v.motivo}; use o do DDA` }); return; }
    usados.add(b.idx);
    // fora da conferência: os motivos de Verificação viram só aviso
    const avisos = [...v.avisos];
    if(outraEmpresa(t.nome, b.beneficiario))
      avisos.push(cedidoAFactoring(b.beneficiario) ? 'título cedido a factoring/fundo' : 'confirmar se é o mesmo fornecedor');
    if(b.comDesconto && b.valorAPagar!==null) avisos.push(`Desconto/abatimento: pagar ${brl(b.valorAPagar)}`);
    add('bolJa', tt, b, { linha:v.linha, avisos });
  });

  // 2) os títulos sem código e os de código errado procuram o boleto no DDA
  //    (regras em PAREAMENTO, acima). Os de código errado só entram no um para um.
  const motivoDe = new Map(problemas.map(p => [p.t.ref, p.motivo]));
  const invalidoDe = new Set(problemas.filter(p => p.invalido).map(p => p.t.ref));
  const genericas = palavrasGenericas(titulos.map(t => t.nome), dda.map(b => b.beneficiario));
  const tits = [...semCodigo, ...problemas.map(p => p.t)]
    .map(t => ({ ...t, palavras: palavrasDoNome(t.nome, genericas), nf: nfSinal(t.doc) }));
  const livres = boletos.filter(b => !usados.has(b.idx))
    .map(b => ({ ...b, palavras: palavrasDoNome(b.beneficiario, genericas) }));
  const titulo = s => ({ ...titulos[s.ref], ref:s.ref });
  const boleto = b => boletos[b.idx];

  // 2.1) um para um. Par sem o sinal de valor (bateram NF e nome) é provisório:
  //      o título pode ser uma das notas de um boleto agrupado
  const pares = parearUmParaUm(tits, livres);
  const firmes = pares.filter(p => p.s.valor);
  const provisorios = pares.filter(p => !p.s.valor);
  const tPareado = new Set(pares.map(p => p.t.ref));
  const bFirme = new Set(firmes.map(p => p.b.idx));

  // 2.2) boleto agrupado: boletos sem par firme x títulos sem par + provisórios
  const candidatos = [...tits.filter(t => !tPareado.has(t.ref)), ...provisorios.map(p => p.t)]
    .filter(t => !motivoDe.has(t.ref));
  const { grupos, ambiguos } = agruparBoletos(livres.filter(b => !bFirme.has(b.idx)), candidatos);
  const tNoGrupo = new Set(grupos.flatMap(g => g.titulos.map(t => t.ref)));
  // boleto ambíguo também fica "ocupado": vai sozinho para a Verificação
  const bNoGrupo = new Set([...grupos, ...ambiguos].map(g => g.b.idx));
  // o provisório vale se nem o título nem o boleto dele foram para um grupo
  const finais = [...firmes, ...provisorios.filter(p => !tNoGrupo.has(p.t.ref) && !bNoGrupo.has(p.b.idx))];

  // cada par (um para um ou nota de um agrupado) vai para "Com boleto" só se não
  // tiver nenhum motivo de Verificação (regras em MOTIVOS DE VERIFICAÇÃO, acima)
  const poePar = (t, bo, s, extra) => {
    const v = bo.val;
    const avisos = [...(extra.avisos || [])];
    const motivos = [...(extra.motivos || [])];
    if(!v.ok) motivos.push('código do boleto no DDA bloqueado');
    if(s && !s.valor) motivos.push(`Valor diferente: boleto ${brl(bo.valor)} x sistema ${brl(t.valor)}`);
    const outra = outraEmpresa(t.nome, bo.beneficiario) || (s && !s.nome);
    if(outra) motivos.push(cedidoAFactoring(bo.beneficiario) ? 'título cedido a factoring/fundo' : 'confirmar se é o mesmo fornecedor');
    if(bo.comDesconto && bo.valorAPagar!==null) motivos.push(`Desconto/abatimento: pagar ${brl(bo.valorAPagar)}`);
    (v.avisos || []).forEach(a => (a.startsWith('vencimento do código') ? motivos : avisos).push(a));
    if(!mesmaData(t.venc, bo.venc)) avisos.push('vencimento no sistema diferente do DDA');
    add(motivos.length ? 'bolVerif' : 'bolOk', t, bo, { linha: v.ok ? v.linha : '', bloqueio: v.ok ? '' : v.motivo,
      avisos, motivo: motivos.join(' · '), ...(s ? { sinais: ['nome', 'nf', 'valor'].filter(k => s[k]) } : {}),
      ...(extra.agrupado ? { agrupado: extra.agrupado } : {}) });
  };

  finais.forEach(({ t: tp, b: bp, s }) => {
    const bo = boleto(bp);
    const motivos = motivoDe.has(tp.ref) ? [motivoDe.get(tp.ref)] : [];
    poePar(titulo(tp), bo, s, { motivos });
  });

  grupos.forEach(g => {
    const bo = boleto(g.b), n = g.titulos.length;
    const nfs = g.titulos.map(t => t.doc || '?').join(', ');
    g.titulos.forEach((tg, k) => {
      const info = `boleto agrupado (${k + 1} de ${n}): um boleto de ${brl(bo.valor)} paga ${n} títulos (NF ${nfs}). Pague uma vez só.`;
      poePar(titulo(tg), bo, null, { avisos:[info], agrupado:{ n, k: k + 1, idx: bo.idx } });
    });
  });

  // boleto que pode pagar mais de uma combinação de notas: nenhuma nota é marcada;
  // o boleto vai para a Verificação com as notas candidatas
  ambiguos.forEach(a => {
    const bo = boleto(a.b), v = bo.val;
    const lista = ts => ts.map(t => `${t.doc || '?'} (${brl(t.valor)})`).join(', ');
    const motivo = a.estourou
      ? `boleto agrupado: há notas demais de ${bo.beneficiario} para conferir a soma (${a.candidatos.length} notas). Confira no boleto quais notas ele paga.`
      : `boleto agrupado ambíguo: há mais de uma combinação de notas que soma ${brl(bo.valor)}. Exemplo: NF ${lista(a.exemplo)}. ` +
        `Confira no boleto quais notas ele paga. Notas candidatas: NF ${lista(a.candidatos)}.`;
    add('bolVerif', null, bo, { linha: v.ok ? v.linha : '', bloqueio: v.ok ? '' : v.motivo, avisos:[...(v.avisos || [])], motivo, ambiguo:true });
  });

  // 3) o resto: título sem boleto e boleto sem título
  const tFinal = new Set([...finais.map(p => p.t.ref), ...tNoGrupo]);
  const bFinal = new Set([...finais.map(p => p.b.idx), ...bNoGrupo]);
  tits.filter(t => !tFinal.has(t.ref)).forEach(s => {
    if(motivoDe.has(s.ref)) add('bolVerif', titulo(s), null, { motivo: motivoDe.get(s.ref) + ' (boleto não encontrado no DDA)',
                                                               codigoSistemaInvalido: invalidoDe.has(s.ref) });
    else add('bolSem', titulo(s), null, {});
  });
  livres.filter(b => !bFinal.has(b.idx)).forEach(b => {
    const bo = boleto(b);
    add('bolDda', null, bo, { linha: bo.val.ok ? bo.val.linha : '', bloqueio: bo.val.ok ? '' : bo.val.motivo, avisos:[...(bo.val.avisos || [])] });
  });

  marcaDuplicidade(linhas);

  // ordem: o que precisa de atenção primeiro
  const ORDEM = ['bolVerif', 'bolOk', 'bolDda', 'bolSem', 'bolJa'];
  linhas.sort((a, b) => ORDEM.indexOf(a.grupo) - ORDEM.indexOf(b.grupo));
  linhas.forEach(l => {
    const t = l.titulo, b = l.boleto;
    l.grupos = [l.grupo];
    // boleto agrupado: cada título entra com o saldo dele (a soma dá o valor do boleto)
    l.valor = l.agrupado ? t.valor : (b ? b.valor : t.valor);
    const textos = [t && t.nome, b && b.beneficiario, t && t.doc, t && t.prf, t && t.fornecedor, b && b.doc, b && b.cnpj,
                    b && b.banco, dataBR(t ? t.venc : b.venc), brl(l.valor), l.motivo, l.bloqueio];
    const digitos = [centavos(l.valor), t && centavos(t.valor), t && apenasDigitos(t.doc), t && t.fornecedor,
                     b && apenasDigitos(b.doc), b && apenasDigitos(b.cnpj), l.linha, b && b.codigoBarras, t && apenasDigitos(t.codigoTexto)];
    l.texto = normaliza(textos.filter(Boolean).join(' | '));
    l.digitos = digitos.filter(Boolean).join('|');
  });
  return linhas;
}

/* -------------------------------------------------------------------
   TELA E EXCEL DA CONFERÊNCIA DE BOLETOS
   O resultado do conferirBoletos é mostrado em 3 abas, iguais na tela e no
   Excel: Com boleto, Verificação e Sem boleto. Cada linha tem um status; o
   nome e a cor de cada status ficam só em STATUS_BOL (tela, Excel e Ajuda).
   As cores são as mesmas do style.css (variáveis --conferido-fundo etc.).
   ------------------------------------------------------------------- */
const STATUS_BOL = {
  ok:      { nome:'Com boleto',       fundo:'FFE2F2E8', texto:'FF17653B',
             ajuda:'Título achado no DDA, um para um, sem nenhuma pendência. Pode copiar a linha digitável e pagar.' },
  grupo:   { nome:'Boleto agrupado',  fundo:'FFDCE8F7', texto:'FF1F4E8C',
             ajuda:'Um boleto só que paga várias NFs: a soma delas dá exatamente o valor do boleto. Pague uma vez só.' },
  compoe:  { nome:'Compõe o boleto',  fundo:'FFF1F6FC', texto:'FF1F4E8C',
             ajuda:'NF que faz parte do boleto agrupado logo acima. Não tem linha digitável própria.' },
  verif:   { nome:'Verificar',        fundo:'FFFBEFD3', texto:'FF7A4F00',
             ajuda:'Precisa de conferência antes de pagar. A coluna "O que verificar" diz o quê.' },
  bloq:    { nome:'Bloqueado',        fundo:'FFFBE4E1', texto:'FF9A2A1C',
             ajuda:'O código de barras não passou na conferência (tamanho, dígitos, banco ou valor). Não há linha digitável para copiar: peça o boleto certo.' },
  sem:     { nome:'Sem boleto',       fundo:'FFECEFEE', texto:'FF44534E',
             ajuda:'Título do sistema sem boleto no DDA. Pode ser Pix, transferência ou boleto ainda não emitido.' },
  consumo: { nome:'Conta de consumo', fundo:'FFECEFEE', texto:'FF44534E',
             ajuda:'Conta de consumo (energia, água, telefone...). Não é boleto bancário e não é conferida aqui.' },
};
const ABAS_BOL = [
  { id:'ok',    nome:'Com boleto',  planilha:'Com boleto' },
  { id:'verif', nome:'Verificação', planilha:'Verificacao' },
  { id:'sem',   nome:'Sem boleto',  planilha:'Sem boleto' },
];
// colunas de cada aba (iguais na tela e no Excel)
function colunasAbaBol(aba){
  const cols = [
    { id:'codForn',    cab:'Cód. Fornecedor', w:16, texto:true },
    { id:'fornecedor', cab:'Fornecedor',      w:44 },
    { id:'nf',         cab:'NF',              w:14, texto:true },
    { id:'valor',      cab:'Valor',           w:16 },
  ];
  if(aba !== 'sem') cols.push({ id:'cb',    cab:'Código de Barras', w:50, texto:true },
                              { id:'linha', cab:'Linha Digitável',  w:52, texto:true });
  if(aba === 'verif') cols.push({ id:'verificar', cab:'O que verificar', w:70 });
  return cols;
}

// ordem: fornecedor e NF (NF como número: 900 antes de 1000)
const ordemFornNf = (a, b) => normaliza(a.fornecedor).localeCompare(normaliza(b.fornecedor)) ||
                              String(a.nf).localeCompare(String(b.nf), 'pt-BR', { numeric:true });

// uma linha de aba. conta = entra na quantidade e no total do cartão (as NFs que
// compõem um boleto agrupado não contam: o boleto já conta uma vez)
function linhaAbaBol(status, { t, b, linha, valor, verificar, nf, obs, conta = true, bloco }){
  const r = {
    status, bloco, conta,
    codForn: t ? String(t.fornecedor || '') : '',
    fornecedor: t ? t.nome : (b ? b.beneficiario : ''),
    nf: nf !== undefined ? nf : (t ? String(t.doc || '') : ''),
    valor, linha: linha || '', cb: linha ? linhaParaBarras(linha) : '',
    verificar: verificar || '', obs: obs || [],
  };
  r.dec = r.cb ? decodifica(r.cb) : null;
  r.texto = normaliza([r.codForn, r.fornecedor, b && b.beneficiario, r.nf, b && b.doc, brl(valor), r.verificar].filter(Boolean).join(' | '));
  r.digitos = [centavos(valor), r.codForn, apenasDigitos(r.nf), b && apenasDigitos(b.doc), r.linha, r.cb].filter(Boolean).join('|');
  return r;
}

// distribui as linhas do conferirBoletos nas 3 abas
function montaAbasBol(linhas){
  const ok = [], grupos = new Map(), amarelos = [], vermelhos = [], sem = [];
  const juntaVerificar = l => [l.bloqueio && 'Bloqueado: ' + l.bloqueio, l.motivo, ...l.avisos].filter(Boolean).join(' · ');
  let bloco = 0;
  linhas.forEach(l => {
    const t = l.titulo, b = l.boleto;
    if(l.agrupado){ grupos.set(l.agrupado.idx, [...(grupos.get(l.agrupado.idx) || []), l]); return; }
    bloco++;
    if(l.grupo === 'bolOk'){
      ok.push(linhaAbaBol('ok', { t, b, linha:l.linha, valor:b.valor, obs:l.avisos, bloco }));
    }else if(l.grupo === 'bolJa' && !b){
      sem.push(linhaAbaBol('consumo', { t, valor:t.valor, bloco }));
    }else if(l.grupo === 'bolJa'){
      // código já lançado no sistema, igual ao do DDA: sem aviso é "Com boleto"
      if(l.avisos.length) amarelos.push(linhaAbaBol('verif', { t, b, linha:l.linha, valor:b.valor, bloco,
        verificar:'Código já lançado no sistema. ' + l.avisos.join(' · ') }));
      else ok.push(linhaAbaBol('ok', { t, b, linha:l.linha, valor:b.valor, obs:['código já lançado no sistema'], bloco }));
    }else if(l.grupo === 'bolSem'){
      sem.push(linhaAbaBol('sem', { t, valor:t.valor, bloco }));
    }else if(l.grupo === 'bolDda'){
      const texto = `Boleto no DDA sem título no sistema${b.doc ? ` (Nº Doc. ${b.doc})` : ''}: veja se falta lançar a nota. Se ninguém reconhecer o boleto, não pague.`;
      const verificar = [l.bloqueio && 'Bloqueado: ' + l.bloqueio, l.motivo || texto, ...l.avisos].filter(Boolean).join(' · ');
      (l.bloqueio ? vermelhos : amarelos).push(linhaAbaBol(l.bloqueio ? 'bloq' : 'verif', { t:null, b, linha:l.linha, valor:b.valor, verificar, bloco }));
    }else{
      // bolVerif: vermelho quando o código não pode ser usado (do DDA bloqueado ou do sistema inválido sem boleto no DDA)
      const bloq = !!l.bloqueio || (!b && l.codigoSistemaInvalido);
      (bloq ? vermelhos : amarelos).push(linhaAbaBol(bloq ? 'bloq' : 'verif', { t, b, linha:l.linha,
        valor: b ? b.valor : t.valor, verificar: juntaVerificar(l), bloco }));
    }
  });

  // boletos agrupados: uma linha azul para o boleto e, abaixo, as NFs (sem repetir a linha digitável)
  const blocosGrupo = [...grupos.values()].map(ls => {
    bloco++;
    ls.sort((x, y) => String(x.titulo.doc).localeCompare(String(y.titulo.doc), 'pt-BR', { numeric:true }));
    const b = ls[0].boleto, n = ls.length;
    const bloqueado = !!ls[0].bloqueio;
    const codigos = [...new Set(ls.map(l => l.titulo.fornecedor))];
    const extras = [...new Set(ls.flatMap(l => (l.motivo || '').split(' · ').filter(Boolean)))];
    const verificar = [bloqueado && 'Bloqueado: ' + ls[0].bloqueio,
      `Boleto único: a soma das ${n} NFs abaixo dá ${brl(b.valor)}. Pagar uma vez só.`, ...extras].filter(Boolean).join(' · ');
    const cab = linhaAbaBol(bloqueado ? 'bloq' : 'grupo', { t:null, b, linha: bloqueado ? '' : ls[0].linha, valor:b.valor,
      nf:`${n} NFs`, verificar, bloco });
    cab.codForn = codigos.length === 1 ? String(codigos[0]) : '';
    cab.fornecedor = ls[0].titulo.nome;
    cab.texto = normaliza([cab.texto, ...ls.map(l => l.titulo.doc)].join(' | '));
    const filhos = ls.map(l => linhaAbaBol('compoe', { t:l.titulo, b, valor:l.titulo.valor, verificar:'compõe o boleto acima', conta:false, bloco }));
    return [cab, ...filhos];
  }).sort((x, y) => ordemFornNf(x[0], y[0]));

  return {
    ok: ok.sort(ordemFornNf),
    verif: [...blocosGrupo.flat(), ...amarelos.sort(ordemFornNf), ...vermelhos.sort(ordemFornNf)],
    sem: sem.sort(ordemFornNf),
  };
}

const resumoAba = linhas => {
  const conta = linhas.filter(r => r.conta);
  return { qtd: conta.length, total: somaValores(conta) };
};

/* ---------- tela ---------- */
let bolAbas = null, bolAbaAtiva = 'ok', bolBuscaTxt = '';

function celulaLinhaDigitavel(r){
  if(!r.linha) return r.status === 'compoe' ? '<span class="sub">compõe o boleto acima</span>' : '—';
  const d = r.dec;
  return `<span class="mono linha-dig">${escapa(formataLinha(r.linha))}</span>` +
         `<span class="sub mono">Banco ${escapa(d.banco)} · ${brl(d.valor)} · venc. ${escapa(d.venc || 'sem vencimento')}</span>` +
         `<button type="button" class="btn-copiar" data-linha="${escapa(r.linha)}">Copiar</button>`;
}

function linhaAbaHtml(r, aba){
  const st = STATUS_BOL[r.status];
  const marca = (aba !== 'ok' && !(aba === 'sem' && r.status === 'sem')) ? `<span class="status st-${r.status}">${escapa(st.nome)}</span>` : '';
  const obs = r.obs.map(o => `<span class="sub">${escapa(o)}</span>`).join('');
  let html = `<td class="mono">${escapa(r.codForn) || '—'}</td><td class="forn">${marca}<span class="forn-nome">${escapa(r.fornecedor)}</span>${obs}</td>` +
             `<td class="mono">${escapa(r.nf) || '—'}</td><td class="num mono">${brl(r.valor)}</td>`;
  if(aba !== 'sem') html += `<td class="mono cod">${r.cb ? escapa(r.cb) : '—'}</td><td class="cod-bol">${celulaLinhaDigitavel(r)}</td>`;
  if(aba === 'verif') html += `<td class="verificar">${escapa(r.verificar)}</td>`;
  return `<tr class="st-${r.status}">${html}</tr>`;
}

function desenhaAbaBol(){
  if(!bolAbas) return;
  const aba = bolAbaAtiva, cols = colunasAbaBol(aba);
  const todas = bolAbas[aba];
  // a busca mostra o bloco inteiro (boleto agrupado + suas NFs) quando qualquer linha dele casa
  const blocos = new Set(todas.filter(r => buscaCasa(r, bolBuscaTxt)).map(r => r.bloco));
  const visiveis = todas.filter(r => blocos.has(r.bloco));
  $('bolCabecalho').innerHTML = '<tr>' + cols.map(c => `<th${c.id === 'valor' ? ' class="num"' : ''}>${c.cab}</th>`).join('') + '</tr>';
  const vazio = bolBuscaTxt.trim() ? `Nada encontrado para “${escapa(bolBuscaTxt.trim())}” nesta aba.` : MSG_VAZIO_BOL[aba];
  $('bolCorpo').innerHTML = visiveis.length ? visiveis.map(r => linhaAbaHtml(r, aba)).join('')
                                            : `<tr><td class="vazio" colspan="${cols.length}">${vazio}</td></tr>`;
  const res = resumoAba(visiveis);
  $('bolRodape').innerHTML = `<tr><td colspan="3">Total · ${res.qtd === 1 ? '1 item' : res.qtd + ' itens'}</td>` +
                             `<td class="num mono">${brl(res.total)}</td>${cols.length > 4 ? `<td colspan="${cols.length - 4}"></td>` : ''}</tr>`;
  document.querySelectorAll('#bolResumo .stat').forEach(card => {
    const ativo = card.dataset.aba === aba;
    card.classList.toggle('ativo', ativo);
    card.querySelector('.stat-filtro').setAttribute('aria-pressed', String(ativo));
  });
}
const MSG_VAZIO_BOL = {
  ok:    'Nenhum título com boleto sem pendência. Veja a Verificação.',
  verif: 'Nada para verificar.',
  sem:   'Todo título do sistema tem boleto ou código.',
};

function mostraAbasBol(abas){
  bolAbas = abas;
  document.querySelectorAll('#bolResumo .stat').forEach(card => {
    const res = resumoAba(abas[card.dataset.aba]);
    card.querySelector('.stat-qtd').textContent = res.qtd;
    card.querySelector('.stat-total').textContent = brl(res.total);
  });
  desenhaAbaBol();
}

document.querySelectorAll('#bolResumo .stat').forEach(card =>
  card.querySelector('.stat-filtro').addEventListener('click', () => { bolAbaAtiva = card.dataset.aba; desenhaAbaBol(); }));
$('bolBusca').addEventListener('input', e => { bolBuscaTxt = e.target.value; desenhaAbaBol(); });

// Copiar: só pela área de transferência do navegador (nada de seleção manual)
$('bolCorpo').addEventListener('click', async e => {
  const btn = e.target.closest('.btn-copiar');
  if(!btn) return;
  const volta = texto => { btn.textContent = texto; setTimeout(() => { btn.textContent = 'Copiar'; btn.classList.remove('copiado', 'falhou'); }, 2000); };
  try{
    if(!navigator.clipboard || !navigator.clipboard.writeText) throw new Error('sem acesso');
    await navigator.clipboard.writeText(btn.dataset.linha);
    btn.classList.add('copiado'); volta('Copiado ✓');
  }catch(err){
    btn.classList.add('falhou'); volta('Não copiou');
  }
});

let bolArqBanco = null, bolArqSis = null, bolUltimo = null;

function bolCheca(){ $('bolBtnGo').disabled = !(bolArqBanco && bolArqSis); mostraArquivos('bol', [bolArqBanco, bolArqSis]); }
ligaDrop('bolDropBanco','bolFileBanco','bolTxtBanco', f => { bolArqBanco = f; bolCheca(); });
ligaDrop('bolDropSis','bolFileSis','bolTxtSis', f => { bolArqSis = f; bolCheca(); });

$('bolBtnGo').addEventListener('click', async () => {
  limpaErro('bolErro');
  try{
    const [lB, lS] = await Promise.all([lerMatrizComTexto(bolArqBanco), lerMatrizComTexto(bolArqSis)]);
    const bancoDoArquivo = detectaBanco(lB.matriz);
    if(bancoDoArquivo==='santander') throw new Error('Este arquivo é um DDA do Santander. A Conferência de Boletos usa o DDA do Itaú ("Consulta Boletos Recebidos").');
    if(bancoDoArquivo!=='itau') throw new Error('Não reconheci o arquivo do Itaú: falta a linha de cabeçalho com "Beneficiário" e "Código de barras". Confira se é o Excel "Consulta Boletos Recebidos" do internet banking.');

    const dda = leitorItau(lB.matriz, lB);
    const titulos = leitorSistemaBoletos(lS);
    if(!dda.length) throw new Error('Não encontrei boletos no arquivo do Itaú. Confira se os valores estão na coluna "Até Venc.".');
    if(!titulos.length) throw new Error('Não encontrei títulos no relatório do sistema (linhas com código do fornecedor numérico e saldo).');

    const linhas = conferirBoletos(dda, titulos);
    bolUltimo = { linhas, consulta: dataHoraConsultaItau(lB.matriz) };

    // checagens dos arquivos (mesma data, mesma empresa, período), antes do resultado
    const checagens = checagensBoletos({ topo: topoSistemaBoletos(lS.matriz), consultaDda: bolUltimo.consulta,
                                         conta: contaDoDdaItau(lB.matriz), dda, linhas });
    $('bolChecagens').innerHTML = checagensHtml(checagens);

    const comCodigo = titulos.filter(t => t.temCodigo).length;
    $('bolInfo').innerHTML =
      `Consulta no Itaú: <b class="mono">${escapa(bolUltimo.consulta || 'data/hora não encontrada no arquivo')}</b> · ` +
      `${dda.length} boletos no DDA · ${titulos.length} títulos no sistema (${comCodigo} já com código de barras)`;
    bolUltimo.abas = montaAbasBol(linhas);
    bolAbaAtiva = 'ok';
    mostraAbasBol(bolUltimo.abas);

    $('bolResultado').style.display = 'block';
    habilitaBaixar('bol');
    $('bolResultado').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    mostraErro('bolErro', err.message || 'Erro ao processar os arquivos.');
  }
});

$('bolBtnReset').addEventListener('click', () => location.reload());

/* ---------- Excel ---------- */
const ARIAL = { name:'Arial', size:10 };

// as 3 abas, com as mesmas colunas e cores da tela. NF, código de barras e linha
// digitável vão como TEXTO (string + numFmt '@'), para não perder zero nem dígito.
function planilhaBoletos(abas){
  const wb = new ExcelJS.Workbook();
  ABAS_BOL.forEach(a => {
    const cols = colunasAbaBol(a.id);
    const ws = wb.addWorksheet(a.planilha, { views:[{ state:'frozen', ySplit:1 }] });
    ws.columns = cols.map(c => ({ width:c.w }));
    const cab = ws.getRow(1);
    cab.values = cols.map(c => c.cab);
    cab.eachCell(c => {
      c.font = { ...ARIAL, bold:true, color:{ argb:'FFFFFFFF' } };
      c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0F4C47' } };
      c.alignment = { vertical:'middle', horizontal:'center' };
      c.border = BORDA_FINA;
    });
    cab.height = 20;
    abas[a.id].forEach(r => {
      const valores = cols.map(c => {
        if(c.id === 'valor') return r.valor;
        if(c.id === 'fornecedor') return r.status === 'consumo' ? `${r.fornecedor} — conta de consumo` : r.fornecedor;
        return String(r[c.id] || '');
      });
      const row = ws.addRow(valores);
      const st = STATUS_BOL[r.status];
      cols.forEach((c, i) => {
        const cel = row.getCell(i + 1);
        cel.font = { ...ARIAL, bold: r.status === 'grupo', color:{ argb: st.texto } };
        cel.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: st.fundo } };
        cel.border = BORDA_FINA;
        cel.alignment = { vertical:'top', wrapText: c.id === 'verificar' || c.id === 'fornecedor' };
        if(c.texto) cel.numFmt = '@';
        if(c.id === 'valor'){ cel.numFmt = FMT_REAL; cel.alignment = { vertical:'top', horizontal:'right' }; }
      });
    });
    const res = resumoAba(abas[a.id]);
    const tot = ws.addRow(cols.map(c => c.id === 'codForn' ? `Total · ${res.qtd} ${res.qtd === 1 ? 'item' : 'itens'}` : c.id === 'valor' ? res.total : ''));
    tot.eachCell({ includeEmpty:true }, (cel, i) => {
      cel.font = { ...ARIAL, bold:true };
      cel.border = BORDA_FINA;
      if(cols[i - 1] && cols[i - 1].id === 'valor') cel.numFmt = FMT_REAL;
    });
  });
  return wb;
}

// lê de volta o Excel gerado e compara, aba por aba e linha por linha, a linha
// digitável e o código de barras com os da tela. Devolve a lista de diferenças.
async function confereExcelBoletos(buffer, abas){
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const erros = [];
  ABAS_BOL.forEach(a => {
    const ws = wb.getWorksheet(a.planilha);
    if(!ws){ erros.push(`a aba "${a.planilha}" não está no arquivo`); return; }
    const cols = colunasAbaBol(a.id);
    const iLinha = cols.findIndex(c => c.id === 'linha') + 1, iCb = cols.findIndex(c => c.id === 'cb') + 1;
    abas[a.id].forEach((r, k) => {
      const row = ws.getRow(k + 2);
      const lido = i => { const v = row.getCell(i).value; return v === null || v === undefined ? '' : v; };
      if(iLinha && (lido(iLinha) !== r.linha || (r.linha && typeof lido(iLinha) !== 'string')))
        erros.push(`${a.nome}, linha ${k + 2}: linha digitável no Excel "${lido(iLinha)}" ≠ tela "${r.linha}"`);
      if(iCb && (lido(iCb) !== r.cb || (r.cb && typeof lido(iCb) !== 'string')))
        erros.push(`${a.nome}, linha ${k + 2}: código de barras no Excel ≠ tela`);
    });
    if(ws.rowCount !== abas[a.id].length + 2) erros.push(`${a.nome}: o Excel tem ${ws.rowCount - 2} linhas e a tela ${abas[a.id].length}`);
  });
  return erros;
}

$('bolBtnXlsx').addEventListener('click', async () => {
  if(!bolUltimo) return;
  limpaErro('bolErro');
  try{
    const buffer = await planilhaBoletos(bolUltimo.abas).xlsx.writeBuffer();
    const erros = await confereExcelBoletos(buffer, bolUltimo.abas);
    if(erros.length){
      mostraErro('bolErro', `O Excel gerado não confere com a tela, então o download foi cancelado. ${erros.slice(0, 3).join(' · ')}` +
                            (erros.length > 3 ? ` (e mais ${erros.length - 3})` : ''));
      return;
    }
    const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'conferencia-boletos.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  }catch(err){
    mostraErro('bolErro', 'Não consegui gerar o Excel: ' + (err.message || err));
  }
});

// legenda de status na Ajuda (vem de STATUS_BOL: mesmo nome e cor da tela e do Excel)
const legendaBol = $('ajudaBolStatus');
if(legendaBol) legendaBol.innerHTML = Object.entries(STATUS_BOL).map(([id, st]) =>
  `<li><span class="status st-${id}">${escapa(st.nome)}</span> ${escapa(st.ajuda)}</li>`).join('');
