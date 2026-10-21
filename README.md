# Conferência DDA

Aplicação web simples para o contas a pagar. Roda 100% no navegador, sem servidor, sem build e sem frameworks. Tem **quatro conferências** e uma página de **Ajuda**, num menu lateral:

- **Conferência Santander**: compara o DDA (Débito Direto Autorizado) do Santander com os títulos a pagar do sistema.
- **Conferência Itaú**: a mesma conferência com o DDA do Itaú, com CNPJ, banco emissor, código de barras, boletos com desconto e filtro por unidade.
- **Código de barras**: compara o valor gravado no código de barras de cada título com o saldo lançado no sistema e aponta códigos lançados mais de uma vez.
- **Conferência de Boletos**: acha no DDA do Itaú o boleto de cada título em aberto do sistema e entrega a linha digitável pronta para copiar, só depois de conferir o código. Junta notas pagas por um boleto só e separa o que precisa ser verificado antes de pagar.
- **Ajuda**: passo a passo de cada conferência, o que fazer com cada grupo ou caso, tabela de status e perguntas frequentes.

## Como usar

Abra o `index.html` no navegador. O duplo clique já funciona: não precisa instalar nada nem ter servidor. Os arquivos `style.css` e `app.js` precisam ficar na mesma pasta.

Cada seção tem seu endereço, então dá pra abrir direto numa delas, e o botão Voltar do navegador funciona:

- `index.html#santander`: Conferência Santander (o antigo `#dda` também abre aqui)
- `index.html#itau`: Conferência Itaú
- `index.html#codigo-barras`: Código de barras
- `index.html#boletos`: Conferência de Boletos
- `index.html#ajuda`: Ajuda

As conferências seguem a mesma ordem na tela (a Conferência de Boletos usa 3 abas no lugar dos grupos; veja a seção dela):

1. Envie os arquivos Excel (clicando na caixa ou arrastando o arquivo). O nome do arquivo aparece na caixa e no alto da tela.
2. Clique em **Conferir**.
3. Veja os **cartões de resumo**: um por grupo, com a quantidade e o valor total. O **?** de cada cartão explica o grupo.
4. Use a **tabela de resultado**: clique num cartão para ver só aquele grupo, busque por fornecedor, NF ou valor, e veja o total no rodapé (soma do que está aparecendo).
5. Clique em **Baixar resultado (Excel)** para gerar a planilha, com uma aba por grupo.

A página de Ajuda (no rodapé do menu) explica tudo isso em linguagem simples, para quem vai usar no dia a dia.

---

## Conferência de DDA (Santander e Itaú)

Você envia o DDA do banco e o relatório de títulos a pagar do sistema, e a página separa os títulos em cinco grupos:

| Grupo | Cor | O que significa |
| --- | --- | --- |
| Conferido | verde | valor, nome e vencimento batem |
| Data divergente | âmbar | valor e nome batem, mas o vencimento é diferente |
| Nome divergente | roxo | só o valor bate (possível factoring/cessão de crédito) |
| Só no banco | vermelho | boleto no DDA sem título no sistema |
| Só no sistema | cinza | título no sistema sem boleto no DDA |

### Arquivos esperados

| Arquivo | Como é lido |
| --- | --- |
| DDA do Santander | Excel do Pagador Eletrônico. Cabeçalho na linha 5; A = Vencimento, B = Beneficiário, C = Nosso Número, D = Pagador, E = Valor. |
| DDA do Itaú | Excel "Consulta Boletos Recebidos". As primeiras linhas têm os dados da conta (inclusive "Data/Hora" da consulta, mostrada na tela). O cabeçalho é achado pelo conteúdo: a linha com "Beneficiário" e "Código de barras". Colunas: Pagador/Agregado, Beneficiário, CPF/CNPJ, Venc., Nº Doc., Até Venc., A Pagar, Tipo de Boleto, Banco, Observações, Código de barras. O "Nº Doc." é guardado como texto (sem perder zeros à esquerda) e aparece na coluna própria "Nº Doc. (Itaú)" da tabela e do Excel. |
| Títulos a pagar do sistema | Dados a partir da linha 7; C = Nome, D = Título (NF), P = Saldo. O título (NF) só aparece na tabela e entra na busca; não é usado na conferência. |

O banco do arquivo é descoberto pelo **conteúdo**, não pelo nome. Se o DDA do Itaú for enviado na tela do Santander (ou o contrário), a página avisa e não confere.

### Particularidades do Itaú

- O valor comparado é o **"Até Venc."**. O **"A Pagar"** fica guardado à parte; quando os dois são diferentes, o boleto é marcado **com desconto**.
- **Filtro por unidade**: quem tem "CAPINOPOLIS" no Pagador (com ou sem acento e espaços) é **Unidade Capinópolis**; o resto é **Matriz**. O relatório do sistema não é separado por unidade, então, ao filtrar uma unidade, os títulos da outra caem em "Só no sistema".
- A tabela traz a coluna **Nº Doc. (Itaú)** (número do boleto no arquivo do Itaú; "—" em "Só no sistema"), CNPJ, banco emissor (código, ex.: 237) e código de barras. A coluna "NF / documento" mostra só a NF do sistema ("—" em "Só no banco").

#### Colunas do Excel do Itaú

| Aba | Colunas |
| --- | --- |
| Nome divergente | Vencimento, Nome no banco, Nome no sistema, CNPJ, Banco, Código de barras, Nº Doc. (Itaú), Para NF, Pedidos, Valor (até venc.), A pagar, Desconto, Unidade |
| Datas divergentes | Nome/Beneficiário, Valor (até venc.), Data no DDA (boleto), Data no sistema (lançada), CNPJ, Banco, Código de barras, Nº Doc. (Itaú), A pagar, Desconto, Unidade |
| So no banco / Conferidos | Vencimento, Nome/Beneficiário, CNPJ, Banco, Código de barras, Nº Doc. (Itaú), Para NF, Pedidos, Valor (até venc.), A pagar, Desconto, Unidade |
| So no sistema | Vencimento, Nome/Beneficiário, Para NF, Pedidos, Valor (sem colunas do boleto) |

"Nº Doc. (Itaú)" é gravado como texto (formato `@`), então números como `000040992001000` mantêm os zeros. "Para NF" e "Pedidos" vêm em branco para preencher.

### Como a conferência funciona

1. Para cada boleto do banco, procura no sistema um título com o **mesmo valor**, **nome parecido** e **mesmo vencimento** → **Conferido**.
2. Se valor e nome baterem mas o **vencimento for diferente** → **Data divergente**.
3. Se só o **valor** bater → **Nome divergente**.
4. O que sobrar de cada lado vai para **Só no banco** ou **Só no sistema**.

O pareamento é feito em camadas, da mais rigorosa (1) à mais frouxa (3): um título já pareado não entra mais nas camadas seguintes. Diferença de até **1 centavo** conta como mesmo valor (aparece um ± ao lado do valor na tela). A comparação de nomes tolera nomes cortados pelo banco e pequenas variações de escrita (prefixo e semelhança de texto).

### Leitores por banco

Cada banco tem um "leitor" no `app.js` (seção **LEITORES POR BANCO**) que transforma o Excel do banco numa lista no formato comum:

```
{ venc, beneficiario, cnpj, doc, valor, valorAPagar, codigoBarras, pagador, banco, comDesconto }
```

A conferência só trabalha com esse formato. Para incluir um banco novo, basta escrever o leitor dele e ensinar a função `detectaBanco` a reconhecer o arquivo.

---

## Código de barras

Você envia **um** relatório de títulos a pagar com as colunas de código de barras e saldo.

- O cabeçalho é encontrado sozinho: a linha que tem uma célula com **"Código de Barras"** e outra com **"Saldo"** (isso evita confundir com o subtítulo abreviado "Cód. Barras").
- As colunas são achadas pelo texto do cabeçalho (BARRA, SALDO, VENC, DUPLICATA e o nome logo à direita de "Cliente"). Se não achar, usa A = Vencto, B = Cliente, C = Nome, D = Duplicata, K = Código de barras, L = Saldo.
- Só entram títulos com saldo numérico e código com **40 ou mais dígitos**.
- O valor do boleto é lido dos **últimos 10 dígitos** do código, em centavos.

Grupos:

- **Conferido** (verde): o valor do código bate com o saldo.
- **Divergência de valor** (âmbar): diferença de 0,005 ou mais; a tabela começa pelas maiores diferenças.
- **Duplicado** (vermelho): o mesmo código em mais de um título; os repetidos têm o mesmo número de grupo. Um título pode ser divergente e duplicado ao mesmo tempo.

---

## Conferência de Boletos

Você envia o **DDA do Itaú** e o relatório **"Duplicatas em Aberto Modelo A"** do sistema, da **mesma empresa** e do **mesmo dia**. Para cada título em aberto, a página acha o boleto no DDA e entrega a **linha digitável pronta para copiar**, só depois de conferir o código. O resultado vem em **3 abas**, iguais na tela e no Excel.

### Arquivos esperados

| Arquivo | Como é lido |
| --- | --- |
| DDA do Itaú | O mesmo "Consulta Boletos Recebidos" da Conferência Itaú (leitor `leitorItau`). O código de barras é lido do **texto** da célula, nunca como número. Das primeiras linhas saem a Data/Hora da consulta, o nome da empresa e o CNPJ. |
| Duplicatas em Aberto Modelo A | Cabeçalho achado pelo conteúdo (a linha com "Código de barras" e "Saldo", como na tela Código de barras). Colunas: Vencto., Cliente (código do fornecedor), nome (a coluna logo à direita de Cliente), Duplicata (NF), Prf, Código de barras e Saldo. Só entram linhas com código do fornecedor **numérico** (subtotais e totais ficam de fora). Do topo saem o código da empresa (célula A1), a Data/Hora do relatório (linha 1) e o "Período". |

### Checagens dos arquivos (no topo, antes do resultado)

- **Mesma data**: a Data/Hora da consulta do DDA e a do relatório têm que ser do mesmo dia; se não forem, aparece um aviso.
- **Mesma empresa**: mostra o código da empresa do relatório e o nome/CNPJ do DDA. Se menos de **10%** dos boletos do DDA (dentro do período) casarem com o sistema, contando os códigos já lançados, aparece o alerta grande "Os arquivos podem ser de empresas diferentes".
- **Período**: mostra o período do relatório e quantos boletos do DDA vencem fora dele. Se o período começa numa **segunda-feira**, os boletos de sábado e domingo anteriores entram também, com a data original.

As checagens só avisam: não mudam o resultado.

### Títulos que já têm código no sistema

Ficam fora do pareamento. O código deles passa pelas mesmas conferências da linha digitável (abaixo) e é procurado no DDA:

- válido e igual a um boleto do DDA: vai para **Com boleto**, marcado "código já lançado no sistema" (ou para a **Verificação**, se tiver algum aviso);
- inválido, ilegível, repetido em dois títulos ou fora do DDA: vai para a **Verificação**, com "use o do DDA" e a linha correta do DDA quando houver. O código do sistema **nunca** é corrigido automaticamente;
- conta de consumo (começa com 8): vai para **Sem boleto**, marcada "Conta de consumo".

### Regras de pareamento

**1. Um para um.** Um título e um boleto casam com **pelo menos 2 destes 3 sinais**:

- **Nome**: alguma palavra relevante em comum, comparando as **5 primeiras letras** de palavras com 4 letras ou mais. Não contam:
  - as palavras da lista fixa `PALAVRAS_IGNORADAS` (LTDA, SA, ME, EPP, EIRELI, COMERCIO, INDUSTRIA, BRASIL, PECAS, SERVICOS e palavras de ramo como ALIMENTOS, TRANSPORTES, MOTORES, TINTAS…), também comparadas pelas 5 primeiras letras, o que pega singular e plural;
  - as palavras **genéricas desta conferência**: uma palavra que aparece no nome de **3 ou mais fornecedores diferentes** do mesmo arquivo (por exemplo, o nome de uma cidade).
- **NF**: o número da NF, sem zeros à esquerda e com **3 dígitos ou mais**, aparece nos dígitos do Nº Doc. do boleto. NF com 1 ou 2 dígitos não conta.
- **Valor**: igual **ao centavo** ao "Até Venc." ou ao "A Pagar".

Só o valor batendo **nunca** casa. Os pares são escolhidos do mais forte para o mais fraco (3 sinais antes de 2; no empate, o vencimento mais próximo). Cada título e cada boleto entram em um par só.

**2. Boleto agrupado.** Para cada boleto livre, do maior valor para o menor, a página junta os títulos ainda sem par cujo nome começa com as **mesmas 5 letras** do beneficiário. Também entram os títulos que bateram só NF e nome, com valor diferente. Ela então procura **2 ou mais** títulos cuja soma dê **exatamente** o valor do boleto, ao centavo.

- A soma usa programação dinâmica em centavos (`somaExata`): para cada soma, guarda de onde ela veio e quantas combinações chegam nela.
- Dá preferência às notas cuja NF aparece no Nº Doc. do boleto.
- Só agrupa quando a combinação é **a única**. Se houver mais de uma, ou notas demais para conferir, o boleto vai **sozinho** para a Verificação ("boleto agrupado ambíguo"), com as notas candidatas, e nenhuma nota é marcada como paga por ele.

**3. O resto** é "Sem boleto" (título sem par) ou "Boleto no DDA sem título no sistema" (na Verificação).

### O que vai para a Verificação

Um par só fica em **Com boleto** se não tiver nenhum destes motivos. A coluna "O que verificar" diz qual é:

| Motivo | Quando |
| --- | --- |
| Boleto único: a soma das N NFs abaixo dá R$ X. Pagar uma vez só. | boleto agrupado (linha azul, com as NFs logo abaixo) |
| boleto agrupado ambíguo | mais de uma combinação de notas dá o valor do boleto |
| título cedido a factoring/fundo | boleto em nome de outra empresa (3 primeiras letras do primeiro nome diferentes, ou o nome não bateu) com FOMENTO, FIDC, FUNDO, FACTORING, CRED…, COBRAN… ou SECURITIZ… |
| confirmar se é o mesmo fornecedor | boleto em nome de outra empresa, sem essas palavras |
| Desconto/abatimento: pagar R$ X | "A Pagar" diferente de "Até Venc." |
| Valor diferente: boleto R$ X x sistema R$ Y | bateram NF e nome, mas não o valor, e não virou grupo |
| vencimento do código diferente do DDA (boleto prorrogado?) | o vencimento gravado no código não é o "Venc." do DDA |
| possível pagamento em duplicidade | mesmo fornecedor e mesmo valor em NFs diferentes, ou a mesma linha digitável em mais de um lugar (as notas de um mesmo boleto agrupado não contam) |
| Boleto no DDA sem título no sistema | boleto que não casou com nenhum título |
| código digitado no sistema … use o do DDA | código do sistema inválido, ilegível ou fora do DDA |
| Bloqueado | código do DDA reprovado nas travas abaixo |

### Travas de segurança da linha digitável

1. **Leitura em texto.** O Excel é lido com `lerMatrizComTexto`: para o código de barras usa o texto da célula, nunca `Number`, porque o JavaScript arredonda depois do 15º dígito sem dar erro. Célula de código que não está gravada como texto (número, data) e código em notação científica (ex.: `3,4191E+43`) são **bloqueados como ilegíveis**.
2. **`validaBoleto`** (seção **BOLETO** do `app.js`). Para ser aceito, o código passa em **tudo**:
   - a) exatamente 44 dígitos (código de barras) ou 47 (linha digitável, convertida para 44);
   - b) DV geral módulo 11 e os 3 DVs de campo módulo 10 certos;
   - c) ida e volta: `linhaParaBarras(barrasParaLinha(cb)) === cb`;
   - d) banco do código = coluna Banco do DDA, e valor do código = "Até Venc." ao centavo.

   Vencimento diferente não bloqueia, mas vai para a Verificação. O vencimento é calculado pelo fator a partir de 22/02/2025 (fator 1000). Código que começa com 8 é conta de consumo e não é validado como boleto.
3. **Bloqueado = sem linha.** Código reprovado não tem linha digitável, não tem botão Copiar e nunca vai para "Com boleto"; aparece em vermelho com o motivo.
4. **Conferência visual.** Ao lado de cada linha digitável aparecem o banco, o valor e o vencimento **lidos do próprio código**.
5. **Copiar** usa só `navigator.clipboard.writeText` (os 47 dígitos), sem seleção manual. Se o navegador não deixar, o botão mostra "Não copiou".
6. **Excel conferido.** Código de barras, linha digitável e NF vão como texto (`'@'`). Antes de liberar o download, a página relê o Excel gerado e compara cada linha digitável e código de barras com os da tela. Se algum for diferente, o download é cancelado e o erro aparece na tela.
7. **O código do sistema nunca é alterado**: quando está errado, a página só avisa e mostra o do DDA.

Mesmo assim: **na hora de pagar, confira no banco se o beneficiário e o valor batem com a planilha.**

### Abas, status e Excel

| Aba | Colunas | O que entra |
| --- | --- | --- |
| Com boleto | Cód. Fornecedor, Fornecedor, NF, Valor, Código de Barras, Linha Digitável | pares um para um sem nenhuma pendência, por fornecedor e NF, com total |
| Verificação | as mesmas + O que verificar | agrupados (azul, com as NFs em azul claro logo abaixo, sem repetir a linha digitável), depois os casos a conferir (amarelo) e os bloqueados (vermelho) |
| Sem boleto | Cód. Fornecedor, Fornecedor, NF, Valor | títulos sem boleto e contas de consumo, com total |

Os cartões de resumo mostram a quantidade e o valor de cada aba e servem para trocar de aba. Cada status tem um nome e uma cor, definidos só em `STATUS_BOL` no `app.js` e usados pela tela, pelo Excel (fonte Arial, uma aba por aba da tela) e pela legenda da Ajuda: Com boleto (verde), Boleto agrupado (azul), Compõe o boleto (azul claro), Verificar (amarelo), Bloqueado (vermelho), Sem boleto e Conta de consumo (cinza).

---

## Testes

Abra o **`testes.html`** no navegador: os testes rodam sozinhos e mostram **verde** (passou) ou **vermelho** (falhou), com um placar no topo. Usam **só dados fictícios** e as funções do próprio `app.js`, que é carregado em modo teste (`window.MODO_TESTE`: sem a tela, os elementos que faltam viram `<div>` soltos). Precisam de internet, como a página principal, por causa das bibliotecas.

O que é testado:

- código válido → linha digitável, banco, valor e vencimento certos, e a volta para o mesmo código;
- um dígito trocado na linha → bloqueado;
- código vindo como número (notação científica, ou célula numérica num Excel fictício) → bloqueado como ilegível;
- 3 títulos que somam exatamente o valor de um boleto → 1 grupo com as 3 NFs;
- só o valor igual (nome e NF diferentes) → não casa;
- NF de 2 dígitos dentro do Nº Doc. → não conta como sinal de NF;
- beneficiário que é um fundo de investimento, com valor e NF iguais → Verificação, "título cedido";
- Excel gerado e lido de volta → linha digitável igual à da tela.

As telas Santander, Itaú e Código de barras são conferidas à parte com os arquivos da pasta `EXEMPLOS/`, que têm dados reais e por isso ficam só no computador (ignorados pelo `.gitignore`).

---

## Arquivos do projeto

- `index.html`: estrutura da página (menu, as quatro conferências e a Ajuda).
- `style.css`: aparência. Todas as cores, fontes e medidas ficam como variáveis no início do arquivo.
- `app.js`: toda a lógica, em seções: navegação, funções compartilhadas e tabela de resultado, leitores por banco, Conferência de DDA (Santander), Conferência Itaú, Código de barras, ajuda dos cartões, **BOLETO** (validação da linha digitável e do código de barras) e Conferência de Boletos (leitura do relatório, checagens, pareamento, motivos de Verificação, tela e Excel).
- `testes.html`: testes automáticos com dados fictícios (veja **Testes**).
- `legado/`: versões antigas, mantidas só como referência (não são usadas pelo app).
- `.gitignore`: impede o envio das pastas `EXEMPLOS/` e `backup/` e de qualquer `.xlsx`, `.xls` ou `.pdf` (arquivos com dados reais).

## Tecnologias

- HTML, CSS e JavaScript puro, sem frameworks e sem build.
- [SheetJS (xlsx.js) 0.18.5](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js), via CDN, para **ler** os Excel enviados.
- [ExcelJS 4.4.0](https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js), via CDN, para **gerar** as planilhas de resultado formatadas.
- Fontes IBM Plex Sans e IBM Plex Mono, do Google Fonts. Sem internet, o navegador usa as fontes do sistema.

As bibliotecas e as fontes vêm da internet; sem conexão, a página abre mas não consegue ler nem gerar Excel.

## Limitações conhecidas

- **Código de barras de 44 dígitos**: a leitura pelos últimos 10 dígitos está certa para a linha digitável (47 dígitos). No código de barras de 44 dígitos o valor fica no meio (do 10º ao 19º dígito); se o relatório do sistema trouxer esse formato, os títulos vão aparecer como divergentes.
- **Caminhos de exportação**: os caminhos no internet banking e no sistema descritos na Ajuda foram deduzidos dos arquivos de exemplo e podem precisar de ajuste.
- **Relatório "Duplicatas em Aberto Modelo A"**: a leitura do topo (empresa na A1, Data/Hora na linha 1 e "Período") é tolerante a formatos, mas foi testada só com planilhas fictícias. A tela sempre mostra o que leu, para conferir.
- **DV geral do boleto**: quando o DV geral do código é 1, a regra do padrão bancário (0, 10 e 11 viram 1) deixa passar alguns erros de um dígito no código de barras de 44 dígitos. Na linha digitável, os DVs de campo pegam esses erros (os que sobram, no vencimento, vão para a Verificação). Por isso, confira sempre o beneficiário e o valor no banco.
- **Nome do fornecedor**: fornecedor cujo nome só tem palavras ignoradas ou curtas não casa pelo nome (só por valor + NF, e aí vai para a Verificação).
- **Copiar**: depende de o navegador liberar a área de transferência; se aparecer "Não copiou", abra a página por um endereço `http://localhost` ou use outro navegador.

## Privacidade

Todo o processamento acontece no navegador. Nenhum dado das planilhas é enviado para servidor algum.
