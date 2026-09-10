# Documentação — Sistema de Compras (Marfim)

Sistema interno de gestão de pedidos de compra, recebimento de NF e cadastros,
construído como um Google Apps Script (backend) + página HTML única (frontend),
usando uma Google Sheet como banco de dados.

## Arquitetura

| Arquivo | Papel |
|---|---|
| `codemarco.gs` | Backend (Google Apps Script). Toda a lógica de negócio, leitura/escrita na planilha, envio de email. |
| `indexmarco.html` | Frontend (SPA de página única). HTML + CSS + JS puro, sem framework. Chama o backend via `google.script.run`. |

- **Banco de dados**: uma única Google Sheet (`SHEET_ID` no topo de `codemarco.gs`). Não há banco de dados relacional nem ORM — cada "tabela" é uma aba da planilha, lida via `sheetToArray()` (linhas → objetos, usando o cabeçalho como chave).
- **Não há ambiente de testes automatizado, nem CI.** Não existem testes unitários. A única verificação possível fora da planilha real é checagem de sintaxe (`node --check`) e revisão manual de código.
- **Implantação é manual**: alterações em `codemarco.gs`/`indexmarco.html` só entram em vigor depois de publicadas no Apps Script (Implantar → Gerenciar implantações). Mudar o arquivo no Git **não atualiza automaticamente o app em produção**.

## Modelo de dados (abas da planilha)

Definido em `setupPlanilha()` (`codemarco.gs`), constante `ABAS`:

| Aba | Colunas |
|---|---|
| `USUARIOS` | COD, NOME, USUARIO, SENHA, EMAIL, PERFIL, FILIAIS_LIBERADAS |
| `FORNECEDORES` | COD, NOME, CNPJ, EMAIL, CONTATO, COND_PAGAMENTO, CEP, BAIRRO, ENDERECO, CIDADE, ESTADO |
| `MATERIAS_PRIMAS` | COD, DESCRICAO, UNIDADE, CATEGORIA |
| `TRANSPORTADORAS` | COD, NOME, CNPJ, CONTATO, PRAZO, OBSERVACAO, CEP, BAIRRO, ENDERECO, CIDADE, ESTADO |
| `FILIAIS` | COD, NOME, CNPJ, CEP, BAIRRO, ENDERECO, CIDADE, ESTADO, EMAIL_RESPONSAVEL, COD_TRANSPORTADORA |
| `PEDIDOS` | ID_PEDIDO, DATA, COD_FILIAL, NOME_FILIAL, COD_FORNECEDOR, NOME_FORNECEDOR, FRETE, COD_TRANSPORTADORA, NOME_TRANSPORTADORA, PRAZO_ENTREGA, COND_PAGAMENTO, OBSERVACAO, USUARIO, VALOR_TOTAL, STATUS |
| `ITENS_PEDIDO` | ID_PEDIDO, COD_MP, DESCRICAO, QUANTIDADE, UNIDADE, PRECO_UNIT, SUBTOTAL |
| `PRECO_FORNECEDOR` | COD_FORNECEDOR, COD_MP, PRECO |
| `TRANSP_FORN_FILIAL` | COD_FORNECEDOR, COD_FILIAL, COD_TRANSPORTADORA |
| `RECEBIMENTOS` | ID_RECEBIMENTO, ID_PEDIDO, NF_NUMERO, DATA_NF, DATA_RECEBIMENTO, COD_FILIAL, NOME_FILIAL, COD_FORNECEDOR, NOME_FORNECEDOR, USUARIO, VALOR_TOTAL_PEDIDO, VALOR_TOTAL_RECEBIDO, DIVERGENCIA_QTD, DIVERGENCIA_PRECO, DIVERGENCIA_PAGTO, PAGTO_PRAZO, PAGTO_ESPERADO, PAGTO_NF, OBSERVACAO |
| `ITENS_RECEBIMENTO` | ID_RECEBIMENTO, ID_PEDIDO, COD_MP, DESCRICAO, QTD_PEDIDA, QTD_RECEBIDA, PRECO_PEDIDO, PRECO_RECEBIDO, SUBTOTAL_PEDIDO, SUBTOTAL_RECEBIDO, DIV_QTD, DIV_PRECO |
| `LOG_ERROS` | DATA, MENSAGEM |
| `LOG_NF` | DATA_HORA, ACAO, ID_RECEBIMENTO, ID_PEDIDO, NF_NUMERO, COD_FILIAL, NOME_FILIAL, USUARIO, DETALHE |

`LOG_NF` é a trilha de auditoria de um pedido. Ações registradas (`ACAO`): `LANÇAMENTO`,
`EXCLUSÃO`, `CANCELAMENTO`, `RETIFICAÇÃO` (e `EDIÇÃO`, valor legado de pedidos corrigidos antes da
fusão do fluxo de edição/retificação — não é mais gravado). Visível na tela de Histórico → detalhe
do pedido.

## Perfis de usuário e permissões

Coluna `PERFIL` em `USUARIOS`. Quatro valores possíveis:

| Perfil | Acesso |
|---|---|
| `ADMIN` | Tudo (Pedido, Histórico, NF, Cadastros). |
| `COMPRAS` | Pedido + Histórico + NF. Sem Cadastros. Pode retificar/cancelar pedidos e excluir NF. |
| `USUARIO` | Pedido + Histórico. Sem NF, sem Cadastros. **Não pode** retificar/cancelar pedidos já enviados. |
| `FILIAL` | NF + Histórico filtrado pelas filiais autorizadas (`FILIAIS_LIBERADAS`). Não acessa "Novo Pedido". |

Regras de gating ficam tanto no frontend (esconder botão/menu) quanto **revalidadas no backend**
via `_getPerfilReal(login)` — nunca confiar apenas no perfil que o frontend diz que o usuário tem.

**Importante**: `_getPerfilReal` espera o **login** (coluna `USUARIO`), não o nome de exibição
(`NOME`). O objeto de sessão no frontend (`sessao`) tem os dois campos (`sessao.usuario` e
`sessao.nome`) — são diferentes e não podem ser usados um pelo outro.

## Funcionalidades principais

### 1. Novo Pedido (`panel-pedido`)
Monta um pedido (filial, fornecedor, frete/transportadora, prazo, condição de pagamento, itens)
e envia por email ao fornecedor (`salvarPedido`). Sem checagem de perfil no backend — qualquer
usuário logado que não seja `FILIAL` pode criar e enviar pedidos.

### 2. Histórico (`panel-historico`)
Lista pedidos e mostra detalhe: itens, NF recebida, log de auditoria, e ações disponíveis
conforme perfil e status do pedido:

- **📧 Retificar e Reenviar** — abre o pedido no mesmo formulário de "Novo Pedido", com todos os
  campos (dados do cabeçalho e itens) editáveis, inclusive edição inline de um item já lançado
  (sem precisar excluir e relançar). Ao confirmar, salva as correções **e** reenvia ao fornecedor
  numa única ação, marcado como retificação (assunto e corpo do email avisam que o pedido foi
  atualizado e substitui a versão anterior). Se nenhum campo foi alterado desde a abertura do
  formulário, o modal de confirmação avisa disso antes de enviar — mas não bloqueia o reenvio.
  Bloqueado se já houver NF lançada para o pedido. Ação sempre explícita do usuário (botão
  dedicado) — nunca dispara como efeito colateral de outra tela.
- **Cancelar Pedido** — marca STATUS como CANCELADO e avisa o fornecedor por email. Bloqueado se
  já houver NF lançada.

Retificar/Cancelar são restritos a `ADMIN`/`COMPRAS` (frontend e backend).

### 3. Cadastros (`panel-cadastros`)
CRUD de Fornecedores, Matérias-Primas, Transportadoras, Filiais e Usuários. Só `ADMIN`.
Busca de CNPJ com auto-fill (BrasilAPI, com fallback para ReceitaWS se houver rate-limit).

### 4. Lançar NF (`panel-nf`)
Registra o recebimento de uma NF contra um pedido, compara quantidade/preço/condição de
pagamento esperados vs. recebidos, grava divergências e dispara email de alerta se houver.
Permite excluir uma NF lançada (libera o pedido para novo lançamento, ou para ser editado).

## Regras de negócio importantes

1. **Retificar é sempre uma ação explícita do usuário** — o botão "Retificar e Reenviar" é o único
   caminho para corrigir um pedido já enviado, e sempre reenvia ao fornecedor ao ser confirmado.
2. **Retificar sem alteração real não é bloqueado, apenas avisado** — se nada mudou desde que o
   formulário foi aberto, o modal de confirmação alerta o usuário, mas permite reenviar mesmo
   assim (ex.: fornecedor perdeu o email original).
3. **Pedido com NF já lançada não pode ser retificado nem cancelado** — evitaria divergência entre
   o que foi enviado ao fornecedor e o que já foi fisicamente recebido.
4. **Pedido cancelado não pode ser retificado, nem cancelado de novo.**
5. Verificação de perfil no backend sempre usa o **login** (`USUARIO`), nunca o nome de exibição.

## Histórico de alterações recentes (changelog)

- **Funde edição e retificação num único fluxo** — remove o botão "Editar Pedido" (correção
  silenciosa sem email) e o backend `editarPedido`/`reenviarPedidoRetificado` separados. Agora só
  existe "Retificar e Reenviar", que abre o formulário completo do pedido (cabeçalho + itens,
  com edição inline de item sem precisar excluir e relançar) e salva + reenvia ao fornecedor numa
  única chamada (`retificarPedido`). Se nada mudou, o modal de confirmação avisa antes de enviar.
- **Corrige verificação de permissão usando login em vez do nome de exibição** — funções de
  pedido e `excluirRecebimento` recebem o login (`sessao.usuario`) num parâmetro dedicado só para
  a checagem de permissão; nome de exibição continua sendo usado no LOG_NF e no email.
- **Adiciona fallback ReceitaWS** — busca de CNPJ não falha mais quando a BrasilAPI aplica
  rate-limit (HTTP 429).
- Histórico completo de commits: `git log --oneline`.

## Processo de verificação e testes (IMPORTANTE)

Este projeto **não tem suíte de testes automatizada nem ambiente de staging**. A única
verificação possível antes de uma mudança chegar à planilha real e aos usuários reais é:

1. Revisão de código e checagem de sintaxe (`node --check`) — feita a cada alteração, mas isso
   **não comprova comportamento correto**, só ausência de erro de sintaxe.
2. **Teste manual no ambiente real é obrigatório** antes de confiar em qualquer mudança em
   produção: publicar a nova versão no Apps Script e testar o fluxo afetado com um usuário de
   cada perfil relevante (ex.: ADMIN, COMPRAS, USUARIO) e com dados reais da planilha.

**Toda entrega de código (deste ou de qualquer assistente) deve terminar pedindo explicitamente
para o usuário verificar/testar a mudança antes de considerá-la concluída** — não há como
confirmar comportamento real sem acesso à implantação publicada e à planilha em produção.
Se encontrar um bug durante esse teste, reporte com o quê foi feito, o que era esperado e o que
aconteceu, para que a correção seja direcionada.
