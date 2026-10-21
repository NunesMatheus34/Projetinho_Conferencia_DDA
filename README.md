# Conferência DDA

Aplicação web simples para o contas a pagar. Roda 100% no navegador, sem servidor, sem build e sem frameworks. Tem **três conferências** e uma página de **Ajuda**, num menu lateral:

- **Conferência Santander**: compara o DDA (Débito Direto Autorizado) do Santander com os títulos a pagar do sistema.
- **Conferência Itaú**: a mesma conferência com o DDA do Itaú, com CNPJ, banco emissor, código de barras, boletos com desconto e filtro por unidade.
- **Código de barras**: compara o valor gravado no código de barras de cada título com o saldo lançado no sistema e aponta códigos lançados mais de uma vez.
- **Ajuda**: passo a passo de cada conferência, o que fazer com cada grupo, tabela de status e perguntas frequentes.

## Como usar

Abra o `index.html` no navegador. O duplo clique já funciona: não precisa instalar nada nem ter servidor. Os arquivos `style.css` e `app.js` precisam ficar na mesma pasta.

Cada seção tem seu endereço, então dá pra abrir direto numa delas, e o botão Voltar do navegador funciona:

- `index.html#santander`: Conferência Santander (o antigo `#dda` também abre aqui)
- `index.html#itau`: Conferência Itaú
- `index.html#codigo-barras`: Código de barras
- `index.html#ajuda`: Ajuda

Todas as conferências seguem a mesma ordem na tela:

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

## Arquivos do projeto

- `index.html`: estrutura da página (menu, as três conferências e a Ajuda).
- `style.css`: aparência. Todas as cores, fontes e medidas ficam como variáveis no início do arquivo.
- `app.js`: toda a lógica, em seções: navegação, funções compartilhadas e tabela de resultado, leitores por banco, Conferência de DDA (Santander), Conferência Itaú, Código de barras e ajuda dos cartões.
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

## Privacidade

Todo o processamento acontece no navegador. Nenhum dado das planilhas é enviado para servidor algum.
