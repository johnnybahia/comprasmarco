// ============================================================
// CONFIGURAÇÃO — altere apenas esta seção
// ============================================================
const SHEET_ID = '1jJkc5LWZnTK4am6oAg_gPHdDJ-NqwT4fYSbY3-r2S_E';

const ABAS = {
  USUARIOS:         'USUARIOS',
  FORNECEDORES:     'FORNECEDORES',
  MATERIAS:         'MATERIAS_PRIMAS',
  TRANSPORTADORAS:  'TRANSPORTADORAS',
  FILIAIS:          'FILIAIS',
  PEDIDOS:          'PEDIDOS',
  ITENS_PEDIDO:     'ITENS_PEDIDO',
  PRECO_FORNECEDOR:    'PRECO_FORNECEDOR',
  TRANSP_FORN_FILIAL:  'TRANSP_FORN_FILIAL',
  RECEBIMENTOS:        'RECEBIMENTOS',
  ITENS_RECEBIMENTO:'ITENS_RECEBIMENTO',
  LOG:              'LOG_ERROS'
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
  if (!sh) return [];
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
// LISTAR TODOS (para dropdowns e autocomplete)
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
// CADASTROS — coluna-aware (adiciona colunas faltantes automaticamente)
// ============================================================
function salvarCadastro(tipo, dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const mapa = {
      fornecedor:     { aba: ABAS.FORNECEDORES,   cols: ['COD','NOME','CNPJ','EMAIL','CONTATO','COND_PAGAMENTO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'] },
      materia:        { aba: ABAS.MATERIAS,        cols: ['COD','DESCRICAO','UNIDADE','CATEGORIA'] },
      transportadora: { aba: ABAS.TRANSPORTADORAS, cols: ['COD','NOME','CNPJ','CONTATO','PRAZO','OBSERVACAO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'] },
      filial:         { aba: ABAS.FILIAIS,         cols: ['COD','NOME','CNPJ','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL','COD_TRANSPORTADORA'] },
      usuario:        { aba: ABAS.USUARIOS,        cols: ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL'] }
    };
    if (!mapa[tipo]) return { ok: false, msg: 'Tipo inválido' };

    const cfg = mapa[tipo];
    const sh = getSheet(cfg.aba);
    const allData = sh.getDataRange().getValues();
    let headers = allData[0].map(String);

    // Adiciona colunas faltantes na planilha
    const missingCols = cfg.cols.filter(c => !headers.includes(c));
    if (missingCols.length > 0) {
      missingCols.forEach((col, i) => {
        const colIdx = headers.length + i + 1;
        const cell = sh.getRange(1, colIdx);
        cell.setValue(col);
        cell.setBackground('#1a3c5e').setFontColor('#ffffff').setFontWeight('bold');
      });
      headers = headers.concat(missingCols);
    }

    const codIdx = headers.indexOf('COD');

    // Verifica se já existe
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][codIdx]).trim() === String(dados.COD).trim()) {
        // Atualiza preservando colunas não mapeadas
        const newRow = headers.map((h, idx) => {
          if (dados[h] !== undefined) return dados[h];
          return allData[i][idx] !== undefined ? allData[i][idx] : '';
        });
        sh.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        return { ok: true, msg: 'Atualizado com sucesso' };
      }
    }

    // Novo registro
    const newRow = headers.map(h => dados[h] !== undefined ? dados[h] : '');
    sh.appendRow(newRow);
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
// PREÇOS POR FORNECEDOR
// ============================================================
function salvarPrecoFornecedor(codForn, codMP, preco) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet(ABAS.PRECO_FORNECEDOR);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === String(codForn).trim() &&
          String(allData[i][1]).trim() === String(codMP).trim()) {
        sh.getRange(i + 1, 3).setValue(preco);
        return { ok: true, msg: 'Preço atualizado' };
      }
    }
    sh.appendRow([codForn, codMP, preco]);
    return { ok: true, msg: 'Preço cadastrado' };
  } catch(e) {
    logErro('salvarPrecoFornecedor: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function excluirPrecoFornecedor(codForn, codMP) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet(ABAS.PRECO_FORNECEDOR);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === String(codForn).trim() &&
          String(allData[i][1]).trim() === String(codMP).trim()) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, msg: 'Não encontrado' };
  } catch(e) {
    logErro('excluirPrecoFornecedor: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function buscarPrecoFornecedor(codForn, codMP) {
  try {
    const rows = sheetToArray(ABAS.PRECO_FORNECEDOR);
    const found = rows.find(r =>
      String(r.COD_FORNECEDOR).trim() === String(codForn).trim() &&
      String(r.COD_MP).trim() === String(codMP).trim()
    );
    return found ? found.PRECO : null;
  } catch(e) {
    logErro('buscarPrecoFornecedor: ' + e.message);
    return null;
  }
}

function listarPrecosPorMateria(codMP) {
  try {
    const precos = sheetToArray(ABAS.PRECO_FORNECEDOR)
      .filter(r => String(r.COD_MP).trim() === String(codMP).trim());
    const fornecedores = sheetToArray(ABAS.FORNECEDORES);
    return precos.map(p => {
      const forn = fornecedores.find(f => String(f.COD).trim() === String(p.COD_FORNECEDOR).trim());
      return {
        COD_FORNECEDOR: p.COD_FORNECEDOR,
        COD_MP: p.COD_MP,
        PRECO: p.PRECO,
        NOME_FORNECEDOR: forn ? forn.NOME : p.COD_FORNECEDOR
      };
    });
  } catch(e) {
    logErro('listarPrecosPorMateria: ' + e.message);
    return [];
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

    const totalPedidos = shPedidos.getLastRow();
    const idPedido = 'PED-' + String(totalPedidos).padStart(5, '0');
    const dataHoje = new Date();

    shPedidos.appendRow([
      idPedido, dataHoje,
      dados.filialCod, dados.filialNome,
      dados.fornecedorCod, dados.fornecedorNome,
      dados.frete || 'CIF',
      dados.transportadoraCod, dados.transportadoraNome,
      dados.prazoEntrega, dados.condPagamento || '',
      dados.observacao, dados.usuarioLogado, dados.valorTotal, 'ENVIADO'
    ]);

    dados.itens.forEach(item => {
      shItens.appendRow([
        idPedido, item.cod, item.descricao,
        item.quantidade, item.unidade, item.preco, item.subtotal
      ]);
    });

    // Dispara email para todos os endereços cadastrados
    const fornecedor = buscarCodigo('fornecedor', dados.fornecedorCod);
    if (!fornecedor || !fornecedor.EMAIL) {
      logErro('salvarPedido: email do fornecedor ausente para ' + dados.fornecedorCod);
      return { ok: true, msg: 'Pedido salvo, mas email não enviado (fornecedor sem email)', id: idPedido };
    }

    const emailsList = String(fornecedor.EMAIL)
      .split(';')
      .map(e => e.trim())
      .filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));

    if (emailsList.length === 0) {
      logErro('salvarPedido: nenhum email válido para ' + dados.fornecedorCod);
      return { ok: true, msg: 'Pedido salvo, mas email não enviado (email inválido)', id: idPedido };
    }

    const filial = buscarCodigo('filial', dados.filialCod) || {};
    const emailsFilial = String(filial.EMAIL_RESPONSAVEL || '')
      .split(';').map(e => e.trim()).filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
    const ccList = [...new Set([
      ...emailsFilial,
      'marco@marfim.ind.br',
      ...(dados.emailUsuario ? [dados.emailUsuario] : [])
    ])].join(',');

    const dadosEmail = Object.assign({}, dados, {
      filialCNPJ:      filial.CNPJ      || '',
      filialEndereco:  [filial.ENDERECO, filial.BAIRRO, filial.CIDADE, filial.ESTADO].filter(Boolean).join(', '),
      fornecedorCNPJ:  fornecedor.CNPJ  || '',
      fornecedorEndereco: [fornecedor.ENDERECO, fornecedor.BAIRRO, fornecedor.CIDADE, fornecedor.ESTADO].filter(Boolean).join(', ')
    });
    const htmlEmail  = montarEmailHTML(idPedido, dataHoje, dadosEmail);
    const textoEmail = montarEmailTexto(idPedido, dataHoje, dadosEmail);
    MailApp.sendEmail({
      to: emailsList.join(','),
      cc: ccList,
      replyTo: 'marco@marfim.ind.br',
      subject: `Pedido de Compra ${idPedido} — ${dados.filialNome}`,
      body: textoEmail,
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
  const linhasItens = dados.itens.map(item => `
    <tr>
      <td style="padding:8px;border:1px solid #ddd;">${item.cod}</td>
      <td style="padding:8px;border:1px solid #ddd;">${item.descricao}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.quantidade} ${item.unidade}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:right;">R$ ${parseFloat(item.preco).toFixed(2)}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:right;">R$ ${parseFloat(item.subtotal).toFixed(2)}</td>
    </tr>
  `).join('');

  return `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;border:1px solid #dde3ea;border-radius:6px;overflow:hidden;">

    <!-- Cabeçalho -->
    <div style="background:#1a3c5e;padding:20px 28px;display:flex;align-items:center;gap:18px;">
      <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Marfim" style="height:52px;width:auto;border-radius:4px;flex-shrink:0;">
      <div>
        <div style="font-size:11px;font-weight:600;letter-spacing:3px;color:#e8a020;text-transform:uppercase;margin-bottom:4px;">Pedido de Compra</div>
        <div style="font-size:20px;font-weight:700;color:white;">${idPedido}</div>
        <div style="font-size:12px;color:rgba(255,255,255,0.6);margin-top:2px;">Emitido em ${dataFmt} · Solicitante: ${dados.usuarioLogado}</div>
      </div>
    </div>

    <!-- Aviso NF -->
    <div style="background:#fff8e1;border-left:4px solid #e8a020;padding:10px 28px;font-size:13px;color:#5a4000;">
      <strong>Para a filial:</strong> ao receber esta entrega, informe o número <strong style="font-family:monospace;">${idPedido}</strong> no sistema para registrar o recebimento da NF.
    </div>

    <!-- Partes: Comprador × Fornecedor -->
    <table style="width:100%;border-collapse:collapse;background:#f4f7fb;">
      <tr>
        <td style="width:50%;padding:16px 28px;vertical-align:top;border-right:1px solid #dde3ea;">
          <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1a3c5e;margin-bottom:8px;">Comprador</div>
          <div style="font-size:14px;font-weight:700;color:#1a1a1a;">${dados.filialNome}</div>
          ${dados.filialCNPJ ? `<div style="font-size:12px;color:#555;margin-top:3px;">CNPJ: ${dados.filialCNPJ}</div>` : ''}
          ${dados.filialEndereco ? `<div style="font-size:12px;color:#777;margin-top:3px;">${dados.filialEndereco}</div>` : ''}
        </td>
        <td style="width:50%;padding:16px 28px;vertical-align:top;">
          <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1a3c5e;margin-bottom:8px;">Fornecedor</div>
          <div style="font-size:14px;font-weight:700;color:#1a1a1a;">${dados.fornecedorNome}</div>
          ${dados.fornecedorCNPJ ? `<div style="font-size:12px;color:#555;margin-top:3px;">CNPJ: ${dados.fornecedorCNPJ}</div>` : ''}
          ${dados.fornecedorEndereco ? `<div style="font-size:12px;color:#777;margin-top:3px;">${dados.fornecedorEndereco}</div>` : ''}
        </td>
      </tr>
    </table>

    <!-- Entrega -->
    <div style="background:#fff;padding:12px 28px;border-top:1px solid #dde3ea;border-bottom:1px solid #dde3ea;display:flex;gap:32px;flex-wrap:wrap;">
      <div><span style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;">Frete</span><br>
        <span style="display:inline-block;background:${dados.frete==='CIF'?'#e8f5e9':'#fff3e0'};border:1px solid ${dados.frete==='CIF'?'#a5d6a7':'#ffcc80'};border-radius:4px;padding:1px 10px;font-size:13px;font-weight:700;color:${dados.frete==='CIF'?'#2e7d32':'#e65100'};">${dados.frete || 'CIF'}</span>
      </div>
      ${dados.frete !== 'CIF' ? `<div><span style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;">Transportadora</span><br><span style="font-size:13px;color:#1a1a1a;">${dados.transportadoraNome || '—'}</span></div>` : ''}
      <div><span style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;">Prazo de Entrega</span><br><span style="font-size:13px;color:#1a1a1a;">${dados.prazoEntrega || '—'}</span></div>
      ${dados.condPagamento ? `<div><span style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;">Condição de Pagamento</span><br><span style="font-size:13px;color:#1a1a1a;">${dados.condPagamento}</span></div>` : ''}
    </div>

    <!-- Itens -->
    <div style="padding:20px 28px;background:#fff;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#1a3c5e;color:white;">
            <th style="padding:10px 12px;text-align:left;font-size:12px;font-weight:600;">Código</th>
            <th style="padding:10px 12px;text-align:left;font-size:12px;font-weight:600;">Descrição</th>
            <th style="padding:10px 12px;text-align:center;font-size:12px;font-weight:600;">Qtd / Un</th>
            <th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:600;">Preço Unit.</th>
            <th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:600;">Subtotal</th>
          </tr>
        </thead>
        <tbody>${linhasItens}</tbody>
        <tfoot>
          <tr style="background:#eef2f7;">
            <td colspan="4" style="padding:12px;text-align:right;font-weight:700;font-size:13px;color:#1a3c5e;">TOTAL DO PEDIDO</td>
            <td style="padding:12px;text-align:right;font-weight:700;font-size:15px;color:#1a3c5e;">R$ ${parseFloat(dados.valorTotal).toFixed(2)}</td>
          </tr>
        </tfoot>
      </table>
      ${dados.observacao ? `<div style="margin-top:16px;padding:12px 14px;background:#f9f9f9;border-left:3px solid #1a3c5e;border-radius:3px;font-size:13px;color:#333;"><strong>Observações:</strong> ${dados.observacao}</div>` : ''}
    </div>

    <!-- Responder -->
    <div style="background:#f4f7fb;padding:14px 28px;border-top:1px solid #dde3ea;text-align:center;">
      <span style="font-size:13px;color:#555;">Dúvidas ou confirmações? Responda diretamente para </span>
      <a href="mailto:marco@marfim.ind.br?subject=Re: Pedido ${idPedido}" style="font-size:13px;font-weight:700;color:#1a3c5e;text-decoration:none;">marco@marfim.ind.br</a>
    </div>

    <!-- Assinatura -->
    <div style="background:#1a3c5e;padding:18px 28px;">
      <table style="width:100%;border-collapse:collapse;">
        <tr>
          <td style="vertical-align:middle;width:90px;">
            <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Marfim" style="height:52px;width:auto;display:block;border-radius:4px;">
          </td>
          <td style="vertical-align:middle;padding-left:16px;border-left:1px solid rgba(255,255,255,0.2);">
            <div style="font-size:14px;font-weight:700;color:white;">Marco Aurélio Bonalume</div>
          </td>
        </tr>
      </table>
    </div>

  </div>`;
}

function montarEmailTexto(idPedido, data, dados) {
  const dataFmt = Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const sep  = '─'.repeat(60);
  const sep2 = '═'.repeat(60);

  const linhasItens = dados.itens.map((item, i) =>
    `  ${String(i+1).padStart(2,'0')}. ${item.descricao}\n` +
    `      Qtd: ${item.quantidade} ${item.unidade}  |  Unit: R$ ${parseFloat(item.preco).toFixed(2)}  |  Subtotal: R$ ${parseFloat(item.subtotal).toFixed(2)}`
  ).join('\n');

  const frete = dados.frete || 'CIF';
  const transpLinha = frete !== 'CIF' && dados.transportadoraNome
    ? `Transportadora : ${dados.transportadoraNome}\n` : '';

  return [
    sep2,
    `PEDIDO DE COMPRA  ${idPedido}`,
    `Emitido em ${dataFmt}  |  Solicitante: ${dados.usuarioLogado}`,
    sep2,
    '',
    'COMPRADOR (FILIAL)',
    `  ${dados.filialNome}${dados.filialCNPJ ? '  |  CNPJ: ' + dados.filialCNPJ : ''}`,
    dados.filialEndereco ? `  ${dados.filialEndereco}` : '',
    '',
    'FORNECEDOR',
    `  ${dados.fornecedorNome}${dados.fornecedorCNPJ ? '  |  CNPJ: ' + dados.fornecedorCNPJ : ''}`,
    dados.fornecedorEndereco ? `  ${dados.fornecedorEndereco}` : '',
    '',
    sep,
    `Modalidade de Frete : ${frete}`,
    transpLinha + `Prazo de Entrega    : ${dados.prazoEntrega || '—'}`,
    dados.condPagamento ? `Condição Pagamento  : ${dados.condPagamento}` : '',
    sep,
    '',
    'ITENS DO PEDIDO',
    '',
    linhasItens,
    '',
    sep,
    `TOTAL DO PEDIDO : R$ ${parseFloat(dados.valorTotal).toFixed(2)}`,
    sep,
    dados.observacao ? `\nObservações: ${dados.observacao}\n` : '',
    '',
    'Dúvidas ou confirmações? Responda para: marco@marfim.ind.br',
    sep2,
    'Marco Aurélio Bonalume — Marfim',
    sep2,
  ].filter(l => l !== null && l !== undefined).join('\n');
}

// ============================================================
// TRANSPORTADORA POR FORNECEDOR × FILIAL
// ============================================================
function salvarTranspFornFilial(codForn, codFilial, codTransp) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet(ABAS.TRANSP_FORN_FILIAL);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === String(codForn).trim() &&
          String(allData[i][1]).trim() === String(codFilial).trim()) {
        sh.getRange(i + 1, 3).setValue(codTransp);
        return { ok: true, msg: 'Transportadora atualizada' };
      }
    }
    sh.appendRow([codForn, codFilial, codTransp]);
    return { ok: true, msg: 'Transportadora vinculada' };
  } catch(e) {
    logErro('salvarTranspFornFilial: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function excluirTranspFornFilial(codForn, codFilial) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet(ABAS.TRANSP_FORN_FILIAL);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === String(codForn).trim() &&
          String(allData[i][1]).trim() === String(codFilial).trim()) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, msg: 'Não encontrado' };
  } catch(e) {
    logErro('excluirTranspFornFilial: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function buscarTranspFornFilial(codForn, codFilial) {
  try {
    const rows = sheetToArray(ABAS.TRANSP_FORN_FILIAL);
    const found = rows.find(r =>
      String(r.COD_FORNECEDOR).trim() === String(codForn).trim() &&
      String(r.COD_FILIAL).trim()     === String(codFilial).trim()
    );
    return found ? found.COD_TRANSPORTADORA : null;
  } catch(e) {
    logErro('buscarTranspFornFilial: ' + e.message);
    return null;
  }
}

function listarTranspPorFornecedor(codForn) {
  try {
    const rows = sheetToArray(ABAS.TRANSP_FORN_FILIAL)
      .filter(r => String(r.COD_FORNECEDOR).trim() === String(codForn).trim());
    const filiais       = sheetToArray(ABAS.FILIAIS);
    const transportadoras = sheetToArray(ABAS.TRANSPORTADORAS);
    return rows.map(r => {
      const fil  = filiais.find(f => String(f.COD).trim() === String(r.COD_FILIAL).trim());
      const trp  = transportadoras.find(t => String(t.COD).trim() === String(r.COD_TRANSPORTADORA).trim());
      return {
        COD_FORNECEDOR:    r.COD_FORNECEDOR,
        COD_FILIAL:        r.COD_FILIAL,
        NOME_FILIAL:       fil  ? fil.NOME  : r.COD_FILIAL,
        COD_TRANSPORTADORA:r.COD_TRANSPORTADORA,
        NOME_TRANSPORTADORA: trp ? trp.NOME : r.COD_TRANSPORTADORA
      };
    });
  } catch(e) {
    logErro('listarTranspPorFornecedor: ' + e.message);
    return [];
  }
}

// ============================================================
// RECEBIMENTO DE NF
// ============================================================
function buscarPedidoParaRecebimento(idPedido) {
  try {
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    const ped = pedidos.find(p => String(p.ID_PEDIDO).trim().toUpperCase() === String(idPedido).trim().toUpperCase());
    if (!ped) return null;
    const itens = sheetToArray(ABAS.ITENS_PEDIDO)
      .filter(i => String(i.ID_PEDIDO).trim() === String(ped.ID_PEDIDO).trim());
    const recebimentos = sheetToArray(ABAS.RECEBIMENTOS)
      .filter(r => String(r.ID_PEDIDO).trim() === String(ped.ID_PEDIDO).trim());
    return { pedido: ped, itens, jaRecebido: recebimentos.length > 0, recebimentos };
  } catch(e) {
    logErro('buscarPedidoParaRecebimento: ' + e.message);
    return null;
  }
}

function salvarRecebimento(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sh      = getSheet(ABAS.RECEBIMENTOS);
    const shItens = getSheet(ABAS.ITENS_RECEBIMENTO);

    const total   = sh.getLastRow();
    const idRec   = 'REC-' + String(total).padStart(5, '0');
    const dataHoje = new Date();

    let temDivQtd   = false;
    let temDivPreco = false;
    let valorTotalRec = 0;

    dados.itens.forEach(item => {
      const qtdPed  = parseFloat(item.qtdPedida)    || 0;
      const qtdRec  = parseFloat(item.qtdRecebida)  || 0;
      const precPed = parseFloat(item.precoPedido)  || 0;
      const precRec = parseFloat(item.precoRecebido)|| 0;
      if (Math.abs(qtdRec  - qtdPed)  > 0.001) temDivQtd   = true;
      if (Math.abs(precRec - precPed) > 0.001) temDivPreco = true;
      valorTotalRec += qtdRec * precRec;
    });

    sh.appendRow([
      idRec, dados.idPedido, dados.nfNumero, dataHoje,
      dados.codFilial, dados.nomeFilial,
      dados.codFornecedor, dados.nomeFornecedor,
      dados.usuarioLogado,
      parseFloat(dados.valorTotalPedido) || 0,
      valorTotalRec,
      temDivQtd   ? 'SIM' : 'NÃO',
      temDivPreco ? 'SIM' : 'NÃO',
      dados.observacao || ''
    ]);

    dados.itens.forEach(item => {
      const qtdPed  = parseFloat(item.qtdPedida)    || 0;
      const qtdRec  = parseFloat(item.qtdRecebida)  || 0;
      const precPed = parseFloat(item.precoPedido)  || 0;
      const precRec = parseFloat(item.precoRecebido)|| 0;
      const divQtd   = Math.abs(qtdRec  - qtdPed)  > 0.001;
      const divPreco = Math.abs(precRec - precPed) > 0.001;
      shItens.appendRow([
        idRec, dados.idPedido, item.codMP, item.descricao,
        qtdPed, qtdRec,
        precPed, precRec,
        qtdPed  * precPed,
        qtdRec  * precRec,
        divQtd   ? 'SIM' : 'NÃO',
        divPreco ? 'SIM' : 'NÃO'
      ]);
    });

    return {
      ok: true,
      msg: 'NF lançada com sucesso',
      id: idRec,
      divQtd: temDivQtd,
      divPreco: temDivPreco
    };
  } catch(e) {
    logErro('salvarRecebimento: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getRecebimentosDoPedido(idPedido) {
  try {
    const recs = sheetToArray(ABAS.RECEBIMENTOS)
      .filter(r => String(r.ID_PEDIDO).trim() === String(idPedido).trim());
    return recs.map(r => {
      const itens = sheetToArray(ABAS.ITENS_RECEBIMENTO)
        .filter(i => String(i.ID_RECEBIMENTO).trim() === String(r.ID_RECEBIMENTO).trim());
      return { ...r, itens };
    });
  } catch(e) {
    logErro('getRecebimentosDoPedido: ' + e.message);
    return [];
  }
}

// ============================================================
// HISTÓRICO
// ============================================================
function _enriquecerPedidosComNF(pedidos) {
  try {
    const recs = sheetToArray(ABAS.RECEBIMENTOS);
    return pedidos.map(p => {
      const recPed = recs.filter(r => String(r.ID_PEDIDO).trim() === String(p.ID_PEDIDO).trim());
      const temNF        = recPed.length > 0;
      const temDivQtd    = recPed.some(r => String(r.DIVERGENCIA_QTD).trim()   === 'SIM');
      const temDivPreco  = recPed.some(r => String(r.DIVERGENCIA_PRECO).trim() === 'SIM');
      return { ...p, temNF, temDivQtd, temDivPreco };
    });
  } catch(e) {
    return pedidos;
  }
}

function getHistorico(tipo, cod) {
  try {
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    const colMapa = { filial: 'COD_FILIAL', fornecedor: 'COD_FORNECEDOR' };
    if (!colMapa[tipo]) return [];
    const filtrados = pedidos.filter(p => String(p[colMapa[tipo]]).trim() === String(cod).trim());
    const enriquecidos = _enriquecerPedidosComNF(filtrados);
    return enriquecidos.map(p => {
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
    return _enriquecerPedidosComNF(sheetToArray(ABAS.PEDIDOS));
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
    USUARIOS:         ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL'],
    FORNECEDORES:     ['COD','NOME','CNPJ','EMAIL','CONTATO','COND_PAGAMENTO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'],
    MATERIAS_PRIMAS:  ['COD','DESCRICAO','UNIDADE','CATEGORIA'],
    TRANSPORTADORAS:  ['COD','NOME','CNPJ','CONTATO','PRAZO','OBSERVACAO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'],
    FILIAIS:          ['COD','NOME','CNPJ','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL','COD_TRANSPORTADORA'],
    PEDIDOS:          ['ID_PEDIDO','DATA','COD_FILIAL','NOME_FILIAL','COD_FORNECEDOR','NOME_FORNECEDOR','FRETE','COD_TRANSPORTADORA','NOME_TRANSPORTADORA','PRAZO_ENTREGA','COND_PAGAMENTO','OBSERVACAO','USUARIO','VALOR_TOTAL','STATUS'],
    ITENS_PEDIDO:     ['ID_PEDIDO','COD_MP','DESCRICAO','QUANTIDADE','UNIDADE','PRECO_UNIT','SUBTOTAL'],
    PRECO_FORNECEDOR:   ['COD_FORNECEDOR','COD_MP','PRECO'],
    TRANSP_FORN_FILIAL: ['COD_FORNECEDOR','COD_FILIAL','COD_TRANSPORTADORA'],
    RECEBIMENTOS:      ['ID_RECEBIMENTO','ID_PEDIDO','NF_NUMERO','DATA_RECEBIMENTO','COD_FILIAL','NOME_FILIAL','COD_FORNECEDOR','NOME_FORNECEDOR','USUARIO','VALOR_TOTAL_PEDIDO','VALOR_TOTAL_RECEBIDO','DIVERGENCIA_QTD','DIVERGENCIA_PRECO','OBSERVACAO'],
    ITENS_RECEBIMENTO: ['ID_RECEBIMENTO','ID_PEDIDO','COD_MP','DESCRICAO','QTD_PEDIDA','QTD_RECEBIDA','PRECO_PEDIDO','PRECO_RECEBIDO','SUBTOTAL_PEDIDO','SUBTOTAL_RECEBIDO','DIV_QTD','DIV_PRECO'],
    LOG_ERROS:         ['DATA','MENSAGEM']
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
    } else {
      // Migração: adiciona colunas faltantes
      const existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const missing = headers.filter(h => !existingHeaders.includes(h));
      missing.forEach((col, i) => {
        const colIdx = existingHeaders.length + i + 1;
        const cell = sh.getRange(1, colIdx);
        cell.setValue(col);
        cell.setBackground('#1a3c5e').setFontColor('#ffffff').setFontWeight('bold');
      });
    }
  });

  const shUser = ss.getSheetByName('USUARIOS');
  if (shUser.getLastRow() <= 1) {
    shUser.appendRow(['USR001','Administrador','admin','admin123','admin@empresa.com','ADMIN']);
  }

  Logger.log('Setup concluído.');
}
