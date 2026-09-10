# Conciliação Financeira

Aplicação web simples (roda 100% no navegador, num único arquivo HTML — sem servidor, sem build, sem frameworks) com **duas ferramentas de conciliação**, organizadas em abas:

- **Conferência de DDA** — compara o débito automático do banco com os títulos a pagar do sistema.
- **Conferência de Código de Barras** — compara o valor lido no código de barras (linha digitável) de cada boleto com o saldo lançado no sistema, e aponta códigos duplicados.

## Como usar

Basta abrir o `index.html` no navegador (duplo clique já funciona, não precisa de instalação nem servidor). As duas ferramentas ficam nas abas no topo da página.

Para abrir direto numa ferramenta específica, use o link com a âncora:

- `index.html#dda` → Conferência de DDA
- `index.html#codigo-barras` → Conferência de Código de Barras

### Conferência de DDA

1. Suba o arquivo do **DDA do banco** na primeira caixa.
2. Suba o arquivo do **sistema (títulos a pagar)** na segunda caixa.
3. Clique em **Conferir**.
4. Revise os cinco blocos de resultado na tela.
5. Se quiser, clique em **Baixar resultado (Excel)** para gerar uma planilha com uma aba por categoria.

### Conferência de Código de Barras

1. Suba o relatório de **títulos a pagar** do sistema.
2. Clique em **Conferir**.
3. Revise os blocos de **divergências de valor** e **códigos duplicados**.
4. Se quiser, clique em **Baixar resultado (Excel)** para gerar uma planilha com uma aba por categoria.

---

## Conferência de DDA

Você sobe as duas planilhas Excel (DDA do banco + títulos a pagar do sistema), e a página aponta:

- o que **bate** certinho (valor, nome e vencimento conferem);
- o que tem **valor e nome iguais mas vencimento diferente** (mesmo título, data divergente);
- o que tem **valor igual mas nome diferente** (possível factoring/cessão de crédito);
- o que está **só no banco** (não encontrado no sistema);
- o que está **só no sistema** (não encontrado no banco).

### Formato esperado das planilhas

| Arquivo                   | Cabeçalho na linha | Coluna do nome | Coluna do valor |
| ------------------------- | ------------------ | -------------- | --------------- |
| DDA do Banco               | 5                  | B              | E               |
| Sistema (Títulos a Pagar)  | 6                  | C              | P (saldo)       |

As colunas E até O do arquivo do sistema são ignoradas automaticamente — só o nome (coluna C) e o saldo (coluna P) são usados na conferência.

### Como a conferência funciona

1. Para cada título do banco, procura no sistema um título com o **mesmo valor**, **nome parecido** e **mesmo vencimento** → vai para **Conferidos**.
2. Se valor e nome baterem mas o **vencimento for diferente**, considera **Data divergente** (mesmo título, lançado com outra data).
3. Se não achar nome parecido mas o **valor bater**, considera **Nome divergente** (sinal de possível factoring, já que o beneficiário pode ter mudado por causa da cessão do título).
4. O que sobrar de cada lado, sem par nenhum, cai em **Só no banco** ou **Só no sistema**.

O pareamento é feito em camadas, do mais rigoroso (passo 1) ao mais frouxo (passo 3): um título já casado numa camada não entra mais na disputa das camadas seguintes.

A comparação de nomes é tolerante a diferenças: nomes truncados pelo banco e pequenas variações de escrita ainda são considerados o "mesmo" beneficiário (usando prefixo e similaridade de texto).

---

## Conferência de Código de Barras

Você sobe **um** relatório de títulos a pagar do sistema, e a página compara o valor lido em cada código de barras com o saldo lançado.

### Detecção do arquivo

- O cabeçalho é encontrado automaticamente: é a linha que tem, ao mesmo tempo, uma célula com **"Código de Barras"** e uma com **"Saldo"** (isso evita confundir com o subtítulo abreviado "Cód. Barras" que aparece antes do cabeçalho real).
- As colunas são mapeadas pelo texto do cabeçalho: código de barras (contém "BARRA"), saldo (contém "SALDO"), vencimento (contém "VENC"), duplicata (contém "DUPLICATA"), e nome (coluna imediatamente à direita de "Cliente").
- Se não conseguir mapear pelo texto, usa um fallback fixo: A=Vencto, B=Cliente, C=Nome, D=Duplicata, K=Código de barras, L=Saldo.
- Só entram na conferência linhas com saldo numérico e código de barras com **40 ou mais dígitos**.

### Como o valor é lido do código de barras

O valor de um boleto está nos **últimos 10 dígitos** da linha digitável, em centavos. A página remove pontos/espaços do código, pega os últimos 10 dígitos e divide por 100.

### As duas conferências

- **Divergência de valor**: quando a diferença entre o valor lido no código e o saldo lançado é de 0,005 ou mais (qualquer diferença, inclusive de centavos). A tabela é ordenada da maior diferença para a menor.
- **Códigos duplicados**: mesmo código de barras aparecendo em mais de um título, agrupados visualmente na tabela.

---

## Tecnologias

- HTML, CSS e JavaScript puro, num único arquivo autocontido (sem frameworks, sem build).
- [SheetJS (xlsx.js)](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js) via CDN, para **ler** os arquivos Excel enviados.
- [ExcelJS](https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js) via CDN, para **gerar** as planilhas de resultado já formatadas (cores, bordas, larguras de coluna, cabeçalho congelado).

## Arquivos do projeto

- `index.html` — página única e autocontida com as duas ferramentas: estrutura (HTML), aparência (CSS) e toda a lógica (JavaScript) num só arquivo.
- `legado/` — versão anterior, antes da fusão em abas: `index.html`/`style.css`/`script.js` (só a Conferência de DDA) e `conferencia-codigo-barras.html` (ferramenta avulsa). Mantidos apenas como referência; não são usados pela aplicação atual.

## Privacidade

Todo o processamento acontece localmente, no navegador. Nenhum dado das planilhas é enviado para servidor algum.
