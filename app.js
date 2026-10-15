/* ===================================================================
   0) NAVEGAÇÃO PELO MENU LATERAL
   =================================================================== */

const $ = id => document.getElementById(id);

// seções que existem; qualquer outro #endereço (inclusive o antigo #dda) abre a primeira
const PAINEIS = ['santander', 'itau', 'codigo-barras', 'ajuda'];

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

// abre um Excel e devolve uma matriz: lista de linhas, cada linha uma lista de células
function lerMatriz(file){
  return new Promise((res,rej)=>{
    const fr = new FileReader();
    fr.onload = e => {
      try{
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
        const ws = wb.Sheets[wb.SheetNames[0]];
        res(XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null}));
      }catch(err){ rej(err); }
    };
    fr.onerror = () => rej(new Error('Não consegui ler o arquivo.'));
    fr.readAsArrayBuffer(file);
  });
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
function leitorItau(m){
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
    out.push({
      venc: formatData(linha[c.venc]),
      beneficiario: texto(linha[c.benef]),
      cnpj: texto(linha[c.cnpj]),
      doc: texto(linha[c.doc]),
      valor,
      valorAPagar,
      codigoBarras: apenasDigitos(linha[c.codigo]),
      pagador: texto(linha[c.pagador]),
      banco: texto(linha[c.banco]),
      comDesconto: valorAPagar!==null && valorAPagar!==valor,
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

  // NF do sistema na frente; o número do boleto no banco embaixo
  const nf = s && s.doc ? s.doc : '';
  const docBanco = b && b.doc ? b.doc : '';
  let doc = nf ? `<span class="mono">${escapa(nf)}</span>` : (docBanco ? `<span class="mono">${escapa(docBanco)}</span>` : '—');
  if(nf && docBanco) doc += `<span class="sub mono">Banco: ${escapa(docBanco)}</span>`;

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
         `<td class="forn">${nome}${nomeSistema}</td><td>${doc}</td>${colunasExtras}` +
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
      if(c.fmt){ r.getCell(i+1).numFmt = c.fmt; r.getCell(i+1).alignment = { horizontal:'right' }; }
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

registraTela('itau', { colunas:8, linhaHtml: l => linhaDdaHtml(l, true), rodapeHtml: rodapeDda(8) });

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

  const meio = b => [ b.cnpj, b.banco, b.codigoBarras ];
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
