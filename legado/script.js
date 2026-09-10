/* ===================================================================
   LÓGICA DA CONFERÊNCIA (o cérebro da página)
   Este arquivo faz: ler os Excel, limpar, comparar e mostrar o resultado.
   =================================================================== */


/* -------------------------------------------------------------------
   1) FUNÇÕES AUXILIARES (pequenas ajudantes usadas em vários lugares)
   ------------------------------------------------------------------- */

// formata um número como dinheiro brasileiro: 1234.5 -> "R$ 1.234,50"
const brl = n => n.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});

// deixa um nome "padronizado" pra comparar: maiúsculo, sem espaço extra, sem pontuação
function limpaNome(s){
  if(s===null || s===undefined) return '';
  s = String(s).trim().toUpperCase().replace(/\s+/g,' '); // tira espaços duplicados
  s = s.replace(/[.,/\-&]/g,'');                          // tira . , / - &
  return s.replace(/\s+/g,' ').trim();
}

// transforma qualquer valor (texto ou número) num número limpo com 2 casas.
// resolve o caso do saldo vir como texto tipo "1.234,56"
function paraNumero(v){
  if(v===null || v===undefined || v==='') return null;
  if(typeof v==='number') return Math.round(v*100)/100;   // já é número
  let s = String(v).replace(/[R$\s]/g,'')  // tira "R$" e espaços
                   .replace(/\./g,'')       // tira o ponto de milhar
                   .replace(',','.');       // vírgula decimal vira ponto
  let n = parseFloat(s);
  return isNaN(n) ? null : Math.round(n*100)/100;
}

// formata a data de vencimento como dd/mm/aaaa, venha ela como texto ou como data
function formatData(v){
  if(v===null || v===undefined || v==='') return '';
  if(v instanceof Date){                     // veio como data de verdade (arquivo do sistema)
    const d = String(v.getDate()).padStart(2,'0');
    const m = String(v.getMonth()+1).padStart(2,'0');
    return `${d}/${m}/${v.getFullYear()}`;
  }
  return String(v).trim();                    // já veio como texto (arquivo do banco)
}

// decide se dois nomes são "a mesma empresa", mesmo que escritos diferente
function nomesBatem(a,b){
  a = limpaNome(a); b = limpaNome(b);
  if(!a || !b) return false;
  // o banco costuma CORTAR o nome, então checo se um começa igual ao outro
  const [curto,longo] = a.length<=b.length ? [a,b] : [b,a];
  if(longo.startsWith(curto.slice(0, Math.min(curto.length,10)))) return true;
  // senão, checo se são "muito parecidos" (60% ou mais)
  return similaridade(a,b) >= 0.6;
}

// calcula o quanto dois textos são parecidos (0 = nada, 1 = idênticos).
// compara os "pedaços de 2 letras" que os dois têm em comum.
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


/* -------------------------------------------------------------------
   2) LER OS ARQUIVOS EXCEL
   ------------------------------------------------------------------- */

// abre o arquivo e devolve uma matriz: lista de linhas, cada linha uma lista de células
function lerMatriz(file){
  return new Promise((res,rej)=>{
    const fr = new FileReader();
    fr.onload = e => {
      try{
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
        const ws = wb.Sheets[wb.SheetNames[0]];              // primeira aba
        res(XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null}));
      }catch(err){ rej(err); }
    };
    fr.onerror = () => rej(new Error('Não consegui ler o arquivo.'));
    fr.readAsArrayBuffer(file);
  });
}

// IMPORTANTE: as colunas contam a partir do ZERO. A=0, B=1, C=2, D=3, E=4 ... P=15

// BANCO: dados começam na linha 6 (índice 5). Beneficiário=coluna B(1), Valor=coluna E(4)
function extraiBanco(m){
  const out = [];
  for(let i=5; i<m.length; i++){
    const linha = m[i] || [];
    const venc  = linha[0];               // coluna A (Vencimento)
    const nome  = linha[1];               // coluna B
    const valor = paraNumero(linha[4]);   // coluna E
    if(valor===null) continue;            // pula linha sem valor
    out.push({ venc: formatData(venc), nome: nome ? String(nome).trim() : '', valor });
  }
  return out;
}

// SISTEMA: dados começam na linha 7 (índice 6). Nome=coluna C(2), Saldo=coluna P(15)
// A "limpeza" que você pediu acontece aqui: uso SÓ o nome e o saldo,
// as colunas E até O são simplesmente ignoradas.
function extraiSistema(m){
  const out = [];
  for(let i=6; i<m.length; i++){
    const linha = m[i] || [];
    const venc  = linha[0];               // coluna A (Vencimento)
    const nome  = linha[2];               // coluna C
    if(nome===null || String(nome).trim()==='') continue;  // pula linhas de total/despesas
    const saldo = paraNumero(linha[15]);  // coluna P
    if(saldo===null) continue;
    out.push({ venc: formatData(venc), nome: String(nome).trim(), valor: saldo });
  }
  return out;
}


/* -------------------------------------------------------------------
   3) A CONFERÊNCIA (o coração de tudo)
   ------------------------------------------------------------------- */

function conferir(banco, sistema){
  // cópia dos dois lados com marca "usado" pra não parear o mesmo título 2x
  const sis = sistema.map(s => ({ ...s, usado:false }));
  const ban = banco.map(b => ({ ...b, usado:false }));

  const ok        = [];  // valor + nome + VENCIMENTO batem (conferido de verdade)
  const datasDiff = [];  // valor + nome batem, mas o VENCIMENTO é diferente
  const div       = [];  // só o valor bate, nome NÃO (possível factoring)

  // O pareamento é feito em CAMADAS, do mais rigoroso ao mais frouxo.
  // Isso é importante: os casamentos perfeitos "pegam" primeiro os títulos do
  // sistema, evitando que um boleto seja pareado com o lançamento errado quando
  // existem vários com o mesmo valor.

  // Camada 1: valor + nome + vencimento (tudo bate) -> CONFERIDO
  ban.forEach(b => {
    const i = sis.findIndex(s => !s.usado && s.valor===b.valor && nomesBatem(b.nome,s.nome) && mesmaData(b.venc,s.venc));
    if(i>=0){ sis[i].usado = true; b.usado = true; ok.push({ banco:b, sistema:sis[i] }); }
  });

  // Camada 2: valor + nome batem, mas a DATA é diferente -> DATA DIVERGENTE
  ban.forEach(b => {
    if(b.usado) return;
    const i = sis.findIndex(s => !s.usado && s.valor===b.valor && nomesBatem(b.nome,s.nome));
    if(i>=0){ sis[i].usado = true; b.usado = true; datasDiff.push({ banco:b, sistema:sis[i] }); }
  });

  // Camada 3: só o valor bate (nome diferente) -> NOME DIVERGENTE
  ban.forEach(b => {
    if(b.usado) return;
    const i = sis.findIndex(s => !s.usado && s.valor===b.valor);
    if(i>=0){ sis[i].usado = true; b.usado = true; div.push({ banco:b, sistema:sis[i] }); }
  });

  // o que não achou par:
  const soBanco = ban.filter(b => !b.usado);  // existe no banco, não no sistema
  const soSis   = sis.filter(s => !s.usado);  // existe no sistema, não no banco

  return { ok, div, soBanco, soSis, datasDiff };
}

// compara duas datas ignorando diferença de formato (pega só dia/mês/ano)
function mesmaData(a, b){
  const so = s => (s ? String(s).match(/\d+/g) || [] : []).join('-'); // "10/08/2026" -> "10-08-2026"
  const A = so(a), B = so(b);
  if(!A || !B) return true;   // se falta data em algum lado, não trato como divergência
  return A === B;
}


/* -------------------------------------------------------------------
   4) MOSTRAR O RESULTADO NA TELA (montar as tabelas em HTML)
   ------------------------------------------------------------------- */

// tabela de pares (banco x sistema), usada em "divergente" e "conferido"
function tabelaPares(dados, destaque){
  if(!dados.length) return '<p class="vazio">Nenhum registro nesta categoria.</p>';
  let h = '<table><thead><tr><th>Valor</th><th>Nome no banco</th><th></th><th>Nome no sistema</th></tr></thead><tbody>';
  dados.forEach(d => {
    h += `<tr class="${destaque?'row-div':''}">
      <td class="val">${brl(d.banco.valor)}</td>
      <td>${escapa(d.banco.nome)}</td>
      <td class="arrow">≠</td>
      <td>${escapa(d.sistema.nome)}</td></tr>`;
  });
  return h + '</tbody></table>';
}

// tabela simples (só valor e nome), usada em "só no banco" e "só no sistema"
function tabelaSimples(dados){
  if(!dados.length) return '<p class="vazio">Nenhum registro nesta categoria — tudo certo aqui.</p>';
  let h = '<table><thead><tr><th>Valor</th><th>Nome</th></tr></thead><tbody>';
  dados.forEach(d => { h += `<tr><td class="val">${brl(d.valor)}</td><td>${escapa(d.nome)}</td></tr>`; });
  return h + '</tbody></table>';
}

// tabela de datas divergentes: mostra as duas datas lado a lado
function tabelaDatas(dados){
  if(!dados.length) return '<p class="vazio">Nenhuma data divergente — as datas batem em tudo que casou.</p>';
  let h = '<table><thead><tr><th>Valor</th><th>Nome</th><th>Data no DDA (boleto)</th><th>Data no sistema (lançada)</th></tr></thead><tbody>';
  dados.forEach(d => {
    h += `<tr class="row-data">
      <td class="val">${brl(d.sistema.valor)}</td>
      <td>${escapa(d.sistema.nome)}</td>
      <td class="val">${escapa(d.banco.venc)}</td>
      <td class="val">${escapa(d.sistema.venc)}</td></tr>`;
  });
  return h + '</tbody></table>';
}

// segurança: evita que caracteres do nome quebrem o HTML
function escapa(s){
  return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
}


/* -------------------------------------------------------------------
   5) LIGAR TUDO: eventos de upload, botões, etc.
   ------------------------------------------------------------------- */

let arqBanco = null, arqSis = null, ultimo = null;

// prepara uma caixa de upload (clicar OU arrastar arquivo)
function ligaDrop(dropId, inputId, txtId, onFile){
  const drop = document.getElementById(dropId);
  const inp  = document.getElementById(inputId);
  const txt  = document.getElementById(txtId);

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
    checa();
  }
}
ligaDrop('dropBanco','fileBanco','txtBanco', f => arqBanco = f);
ligaDrop('dropSis','fileSis','txtSis', f => arqSis = f);

// libera o botão "Conferir" só quando os dois arquivos estiverem prontos
function checa(){ document.getElementById('btnGo').disabled = !(arqBanco && arqSis); }

function mostraErro(msg){ const e=document.getElementById('erro'); e.textContent='⚠ '+msg; e.style.display='block'; }
function limpaErro(){ document.getElementById('erro').style.display='none'; }

// clique em CONFERIR: faz tudo acontecer na ordem
document.getElementById('btnGo').addEventListener('click', async () => {
  limpaErro();
  try{
    const [mB, mS] = await Promise.all([lerMatriz(arqBanco), lerMatriz(arqSis)]);
    const banco   = extraiBanco(mB);
    const sistema = extraiSistema(mS);

    if(!banco.length)   throw new Error('Não encontrei valores no arquivo do banco. Confira se o cabeçalho está na linha 5 e os valores na coluna E.');
    if(!sistema.length) throw new Error('Não encontrei saldos no arquivo do sistema. Confira se os nomes estão na coluna C e o saldo na coluna P.');

    const r = conferir(banco, sistema);
    ultimo = r;  // guarda pro botão de baixar Excel

    // preenche os números do resumo
    document.getElementById('nOk').textContent    = r.ok.length;
    document.getElementById('nDiv').textContent   = r.div.length;
    document.getElementById('nData').textContent  = r.datasDiff.length;
    document.getElementById('nBanco').textContent = r.soBanco.length;
    document.getElementById('nSis').textContent   = r.soSis.length;

    // preenche as tabelas
    document.getElementById('tblDiv').innerHTML   = tabelaPares(r.div, true);
    document.getElementById('tblData').innerHTML  = tabelaDatas(r.datasDiff);
    document.getElementById('tblBanco').innerHTML = tabelaSimples(r.soBanco);
    document.getElementById('tblSis').innerHTML   = tabelaSimples(r.soSis);
    document.getElementById('tblOk').innerHTML    = tabelaPares(r.ok, false);

    // mostra a área de resultado e os botões extras
    document.getElementById('resultado').style.display = 'block';
    document.getElementById('btnXlsx').classList.remove('oculto');
    document.getElementById('btnReset').classList.remove('oculto');
    document.getElementById('resultado').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(err){
    mostraErro(err.message || 'Erro ao processar os arquivos.');
  }
});

// botão NOVA CONFERÊNCIA: recarrega a página pra começar do zero
document.getElementById('btnReset').addEventListener('click', () => location.reload());

// botão BAIXAR RESULTADO: gera um Excel FORMATADO (com o ExcelJS).
// Cada aba tem 5 colunas: Vencimento | Nome/Beneficiário | Para NF | Pedidos | Valor
// As colunas "Para NF" e "Pedidos" saem em branco, pra você preencher à mão depois.

const BORDA_FINA = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

// definição das colunas de cada tipo de aba: { cab: título, w: largura, fmt: formato opcional }
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

// monta uma aba já formatada. "colunas" define os títulos/larguras; "linhas" são listas de valores.
function montaAba(wb, nomeAba, titulo, colunas, linhas){
  const ws = wb.addWorksheet(nomeAba);
  ws.columns = colunas.map(c => ({ width:c.w }));
  const ncols = colunas.length;
  const ultimaCol = String.fromCharCode(64 + ncols);   // A=65; 5 colunas -> "E", 4 colunas -> "D"

  // Linha 1: título da categoria (faixa escura ocupando todas as colunas)
  ws.mergeCells(`A1:${ultimaCol}1`);
  const t = ws.getCell('A1');
  t.value = titulo;
  t.font = { bold:true, size:13, color:{argb:'FFFFFFFF'} };
  t.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF12283A'} };
  t.alignment = { vertical:'middle', horizontal:'left', indent:1 };
  ws.getRow(1).height = 22;

  // Linha 2: cabeçalho das colunas
  const cab = ws.getRow(2);
  cab.values = colunas.map(c => c.cab);
  cab.eachCell(c => {
    c.font = { bold:true, color:{argb:'FF12283A'} };
    c.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFEDE7DA'} };
    c.alignment = { vertical:'middle', horizontal:'center' };
    c.border = BORDA_FINA;
  });
  cab.height = 18;

  // Linhas de dados
  linhas.forEach(vals => {
    const r = ws.addRow(vals);
    r.eachCell({ includeEmpty:true }, c => { c.border = BORDA_FINA; });
    colunas.forEach((c, i) => {
      if(c.fmt){ r.getCell(i+1).numFmt = c.fmt; r.getCell(i+1).alignment = { horizontal:'right' }; }
    });
  });

  // se não houver nenhuma linha, deixa um aviso
  if(linhas.length === 0){
    ws.mergeCells(`A3:${ultimaCol}3`);
    ws.getCell('A3').value = 'Nenhum registro nesta categoria.';
    ws.getCell('A3').font = { italic:true, color:{argb:'FF5C6B78'} };
  }
  return ws;
}

document.getElementById('btnXlsx').addEventListener('click', async () => {
  if(!ultimo) return;

  const wb = new ExcelJS.Workbook();

  // aba com as linhas no padrão de 5 colunas: [venc, nome, NF(vazio), pedido(vazio), valor]
  const linhaPadrao = (venc, nome, valor) => [ venc || '', nome || '', '', '', valor ];

  // no "Nome divergente" mostro os DOIS nomes lado a lado (banco e sistema),
  // pra ficar fácil ver a diferença sem precisar abrir os dois arquivos originais.
  const linhaDiv = (venc, nomeBanco, nomeSistema, valor) =>
    [ venc || '', nomeBanco || '', nomeSistema || '', '', '', valor ];

  montaAba(wb, 'Nome divergente', 'Nome divergente (conferir · possível factoring)', COLS_DIV,
           ultimo.div.map(d => linhaDiv(d.sistema.venc, d.banco.nome, d.sistema.nome, d.sistema.valor)));

  // NOVO: aba de datas divergentes — mostra a data do boleto (DDA) e a data lançada (sistema)
  montaAba(wb, 'Datas divergentes', 'Datas divergentes (mesmo título, vencimento diferente)', COLS_DATAS,
           ultimo.datasDiff.map(d => [ d.sistema.nome, d.sistema.valor, d.banco.venc, d.sistema.venc ]));

  montaAba(wb, 'So no banco', 'Só no banco (não achado no sistema)', COLS_PADRAO,
           ultimo.soBanco.map(d => linhaPadrao(d.venc, d.nome, d.valor)));
  montaAba(wb, 'So no sistema', 'Só no sistema (não achado no banco)', COLS_PADRAO,
           ultimo.soSis.map(d => linhaPadrao(d.venc, d.nome, d.valor)));
  montaAba(wb, 'Conferidos', 'Conferidos (valor e nome batem)', COLS_PADRAO,
           ultimo.ok.map(d => linhaPadrao(d.sistema.venc, d.sistema.nome, d.sistema.valor)));

  // gera o arquivo e dispara o download no navegador
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'conferencia-dda.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});
