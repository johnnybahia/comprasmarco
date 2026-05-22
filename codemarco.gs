// ============================================================
// CONFIGURAÇÃO — altere apenas esta seção
// ============================================================
const SHEET_ID = '1jJkc5LWZnTK4am6oAg_gPHdDJ-NqwT4fYSbY3-r2S_E';

const ABAS = {
  USUARIOS:      'USUARIOS',
  FORNECEDORES:  'FORNECEDORES',
  MATERIAS:      'MATERIAS_PRIMAS',
  TRANSPORTADORAS:'TRANSPORTADORAS',
  FILIAIS:       'FILIAIS',
  PEDIDOS:       'PEDIDOS',
  ITENS_PEDIDO:  'ITENS_PEDIDO',
  LOG:           'LOG_ERROS'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema de Compras')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// HELPERS
// ============================================================
function getSheet(nome) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(nome);
}

function sheetToArray(nome) {
  const sh = getSheet(nome);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function logErro(msg) {
  try {
    getSheet(ABAS.LOG).appendRow([new Date(), msg]);
  } catch(e) {}
}

// ============================================================
// AUTH
// ============================================================
function validarLogin(usuario, senha) {
  try {
    const rows = sheetToArray(ABAS.USUARIOS);
    const user = rows.find(r => 
      String(r.USUARIO).trim().toLowerCase() === String(usuario).trim().toLowerCase() &&
      String(r.SENHA).trim() === String(senha).trim()
    );
    if (!user) return null;
    return { nome: user.NOME, usuario: user.USUARIO, email: user.EMAIL, perfil: user.PERFIL };
  } catch(e) {
    logErro('validarLogin: ' + e.message);
    return null;
  }
}

// ============================================================
// LOOKUP POR CÓDIGO
// ============================================================
function buscarCodigo(tipo, codigo) {
  try {
    const mapa = {
      fornecedor:     { aba: ABAS.FORNECEDORES,   col: 'COD' },
      materia:        { aba: ABAS.MATERIAS,        col: 'COD' },
      transportadora: { aba: ABAS.TRANSPORTADORAS, col: 'COD' },
      filial:         { aba: ABAS.FILIAIS,         col: 'COD' }
    };
    if (!mapa[tipo]) return null;
    const rows = sheetToArray(mapa[tipo].aba);
    return rows.find(r => String(r[mapa[tipo].col]).trim() === String(codigo).trim()) || null;
  } catch(e) {
    logErro('buscarCodigo: ' + e.message);
    return null;
  }
}

// ============================================================
// LISTAR TODOS (para dropdowns)
// ============================================================
function listarTodos(tipo) {
  try {
    const mapa = {
      fornecedor:     ABAS.FORNECEDORES,
      materia:        ABAS.MATERIAS,
      transportadora: ABAS.TRANSPORTADORAS,
      filial:         ABAS.FILIAIS,
      usuario:        ABAS.USUARIOS
    };
    if (!mapa[tipo]) return [];
    return sheetToArray(mapa[tipo]);
  } catch(e) {
    logErro('listarTodos: ' + e.message);
    return [];
  }
}

// ============================================================
// CADASTROS
// ============================================================
function salvarCadastro(tipo, dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const mapa = {
      fornecedor:     { aba: ABAS.FORNECEDORES,    cols: ['COD','NOME','EMAIL','CONTATO','ENDERECO','CIDADE','ESTADO'] },
      materia:        { aba: ABAS.MATERIAS,         cols: ['COD','DESCRICAO','UNIDADE','CATEGORIA'] },
      transportadora: { aba: ABAS.TRANSPORTADORAS,  cols: ['COD','NOME','CONTATO','PRAZO','OBSERVACAO'] },
      filial:         { aba: ABAS.FILIAIS,           cols: ['COD','NOME','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL'] },
      usuario:        { aba: ABAS.USUARIOS,          cols: ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL'] }
    };
    if (!mapa[tipo]) return { ok: false, msg: 'Tipo inválido' };

    const cfg = mapa[tipo];
    const sh = getSheet(cfg.aba);

    // Verifica duplicidade de COD
    const existentes = sheetToArray(cfg.aba);
    if (existentes.find(r => String(r.COD).trim() === String(dados.COD).trim())) {
      // Atualiza linha existente
      const allData = sh.getDataRange().getValues();
      const headers = allData[0];
      const codIdx = headers.indexOf('COD');
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][codIdx]).trim() === String(dados.COD).trim()) {
          const row = cfg.cols.map(c => dados[c] !== undefined ? dados[c] : '');
          sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
          return { ok: true, msg: 'Atualizado com sucesso' };
        }
      }
    }

    // Novo registro
    const row = cfg.cols.map(c => dados[c] !== undefined ? dados[c] : '');
    sh.appendRow(row);
    return { ok: true, msg: 'Cadastrado com sucesso' };
  } catch(e) {
    logErro('salvarCadastro: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function excluirCadastro(tipo, cod) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const mapa = {
      fornecedor:     ABAS.FORNECEDORES,
      materia:        ABAS.MATERIAS,
      transportadora: ABAS.TRANSPORTADORAS,
      filial:         ABAS.FILIAIS,
      usuario:        ABAS.USUARIOS
    };
    if (!mapa[tipo]) return { ok: false, msg: 'Tipo inválido' };
    const sh = getSheet(mapa[tipo]);
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const codIdx = headers.indexOf('COD');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codIdx]).trim() === String(cod).trim()) {
        sh.deleteRow(i + 1);
        return { ok: true, msg: 'Excluído com sucesso' };
      }
    }
    return { ok: false, msg: 'Registro não encontrado' };
  } catch(e) {
    logErro('excluirCadastro: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// PEDIDOS
// ============================================================
function salvarPedido(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const shPedidos = getSheet(ABAS.PEDIDOS);
    const shItens   = getSheet(ABAS.ITENS_PEDIDO);

    // Gera ID do pedido
    const totalPedidos = shPedidos.getLastRow();
    const idPedido = 'PED-' + String(totalPedidos).padStart(5, '0');

    const dataHoje = new Date();

    // Grava cabeçalho do pedido
    shPedidos.appendRow([
      idPedido,
      dataHoje,
      dados.filialCod,
      dados.filialNome,
      dados.fornecedorCod,
      dados.fornecedorNome,
      dados.transportadoraCod,
      dados.transportadoraNome,
      dados.prazoEntrega,
      dados.observacao,
      dados.usuarioLogado,
      dados.valorTotal,
      'ENVIADO'
    ]);

    // Grava itens
    dados.itens.forEach(item => {
      shItens.appendRow([
        idPedido,
        item.cod,
        item.descricao,
        item.quantidade,
        item.unidade,
        item.preco,
        item.subtotal
      ]);
    });

    // Dispara email
    const fornecedor = buscarCodigo('fornecedor', dados.fornecedorCod);
    if (!fornecedor || !fornecedor.EMAIL) {
      logErro('salvarPedido: email do fornecedor ausente para ' + dados.fornecedorCod);
      return { ok: true, msg: 'Pedido salvo, mas email não enviado (fornecedor sem email)', id: idPedido };
    }

    const emailValido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(fornecedor.EMAIL);
    if (!emailValido) {
      logErro('salvarPedido: email inválido: ' + fornecedor.EMAIL);
      return { ok: true, msg: 'Pedido salvo, mas email não enviado (email inválido)', id: idPedido };
    }

    const htmlEmail = montarEmailHTML(idPedido, dataHoje, dados);
    
    MailApp.sendEmail({
  to: fornecedor.EMAIL,
  cc: dados.emailUsuario || '',
  replyTo: 'marco@marfim.ind.br',
  subject: `Pedido de Compra ${idPedido} — ${dados.filialNome}`,
  htmlBody: htmlEmail
});

    return { ok: true, msg: 'Pedido salvo e email enviado com sucesso', id: idPedido };

  } catch(e) {
    logErro('salvarPedido: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function montarEmailHTML(idPedido, data, dados) {
  const dataFmt = Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let linhasItens = dados.itens.map(item => `
    <tr>
      <td style="padding:8px;border:1px solid #ddd;">${item.cod}</td>
      <td style="padding:8px;border:1px solid #ddd;">${item.descricao}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.quantidade} ${item.unidade}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:right;">R$ ${parseFloat(item.preco).toFixed(2)}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:right;">R$ ${parseFloat(item.subtotal).toFixed(2)}</td>
    </tr>
  `).join('');

  return `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">
    <div style="background:#1a3c5e;color:white;padding:20px 30px;">
      <h2 style="margin:0;">Pedido de Compra</h2>
      <p style="margin:4px 0 0;">${idPedido} — ${dataFmt}</p>
    </div>
    <div style="padding:20px 30px;background:#f9f9f9;">
      <table style="width:100%;border-collapse:collapse;margin-bottom:16px;">
        <tr>
          <td style="padding:6px 0;"><strong>Filial:</strong> ${dados.filialNome} (${dados.filialCod})</td>
          <td style="padding:6px 0;"><strong>Fornecedor:</strong> ${dados.fornecedorNome}</td>
        </tr>
        <tr>
          <td style="padding:6px 0;"><strong>Transportadora:</strong> ${dados.transportadoraNome || '—'}</td>
          <td style="padding:6px 0;"><strong>Prazo de Entrega:</strong> ${dados.prazoEntrega || '—'}</td>
        </tr>
        <tr>
          <td colspan="2" style="padding:6px 0;"><strong>Solicitante:</strong> ${dados.usuarioLogado}</td>
        </tr>
      </table>
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#1a3c5e;color:white;">
            <th style="padding:10px;text-align:left;">Código</th>
            <th style="padding:10px;text-align:left;">Descrição</th>
            <th style="padding:10px;text-align:center;">Qtd/Un</th>
            <th style="padding:10px;text-align:right;">Preço Unit.</th>
            <th style="padding:10px;text-align:right;">Subtotal</th>
          </tr>
        </thead>
        <tbody>${linhasItens}</tbody>
        <tfoot>
          <tr style="background:#eef2f7;">
            <td colspan="4" style="padding:10px;text-align:right;font-weight:bold;">TOTAL</td>
            <td style="padding:10px;text-align:right;font-weight:bold;">R$ ${parseFloat(dados.valorTotal).toFixed(2)}</td>
          </tr>
        </tfoot>
      </table>
      ${dados.observacao ? `<p style="margin-top:16px;"><strong>Observações:</strong> ${dados.observacao}</p>` : ''}
    </div>
    <div style="background:#eee;padding:12px 30px;font-size:12px;color:#666;">
      Email gerado automaticamente pelo Sistema de Compras.
    </div>
  </div>`;
}

// ============================================================
// HISTÓRICO
// ============================================================
function getHistorico(tipo, cod) {
  try {
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    const colMapa = { filial: 'COD_FILIAL', fornecedor: 'COD_FORNECEDOR' };
    if (!colMapa[tipo]) return [];
    const filtrados = pedidos.filter(p => String(p[colMapa[tipo]]).trim() === String(cod).trim());
    return filtrados.map(p => {
      const itens = sheetToArray(ABAS.ITENS_PEDIDO).filter(i => i.ID_PEDIDO === p.ID_PEDIDO);
      return { ...p, itens };
    });
  } catch(e) {
    logErro('getHistorico: ' + e.message);
    return [];
  }
}

function getTodosPedidos() {
  try {
    return sheetToArray(ABAS.PEDIDOS);
  } catch(e) {
    logErro('getTodosPedidos: ' + e.message);
    return [];
  }
}

// ============================================================
// SETUP — rode UMA VEZ para criar as abas
// ============================================================
function setupPlanilha() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const estrutura = {
    USUARIOS:       ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL'],
    FORNECEDORES:   ['COD','NOME','EMAIL','CONTATO','ENDERECO','CIDADE','ESTADO'],
    MATERIAS_PRIMAS:['COD','DESCRICAO','UNIDADE','CATEGORIA'],
    TRANSPORTADORAS:['COD','NOME','CONTATO','PRAZO','OBSERVACAO'],
    FILIAIS:        ['COD','NOME','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL'],
    PEDIDOS:        ['ID_PEDIDO','DATA','COD_FILIAL','NOME_FILIAL','COD_FORNECEDOR','NOME_FORNECEDOR','COD_TRANSPORTADORA','NOME_TRANSPORTADORA','PRAZO_ENTREGA','OBSERVACAO','USUARIO','VALOR_TOTAL','STATUS'],
    ITENS_PEDIDO:   ['ID_PEDIDO','COD_MP','DESCRICAO','QUANTIDADE','UNIDADE','PRECO_UNIT','SUBTOTAL'],
    LOG_ERROS:      ['DATA','MENSAGEM']
  };

  Object.entries(estrutura).forEach(([nome, headers]) => {
    let sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length)
        .setBackground('#1a3c5e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });

  // Cria usuário admin padrão se USUARIOS estiver vazia
  const shUser = ss.getSheetByName('USUARIOS');
  if (shUser.getLastRow() <= 1) {
    shUser.appendRow(['USR001','Administrador','admin','admin123','admin@empresa.com','ADMIN']);
  }

  Logger.log('Setup concluído.');
}
