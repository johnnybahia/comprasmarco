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
  LOG:              'LOG_ERROS',
  LOG_NF:           'LOG_NF'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Compras Marfim')
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
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'") : v;
    });
    return obj;
  });
}

function logErro(msg) {
  try {
    getSheet(ABAS.LOG).appendRow([new Date(), msg]);
  } catch(e) {}
}

function _escHTML(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// Retorna o perfil real do usuário consultando a planilha; null se não encontrado.
// Espera o LOGIN (coluna USUARIO) — nunca o nome de exibição.
function _getPerfilReal(usuarioLogin) {
  if (!usuarioLogin) return null;
  const users = sheetToArray(ABAS.USUARIOS);
  const user = users.find(u =>
    String(u.USUARIO).trim().toLowerCase() === String(usuarioLogin).trim().toLowerCase()
  );
  return user ? String(user.PERFIL || '').toUpperCase().trim() : null;
}

// Gera próximo ID sequencial baseado no máximo existente (imune a deleções de linha).
function _proximoId(sh, colNome, prefixo) {
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return prefixo + '00001';
  const headers = data[0].map(h => String(h).trim());
  const col = headers.indexOf(colNome);
  if (col < 0) return prefixo + '00001';
  const re = new RegExp('^' + prefixo.replace('-', '\\-') + '(\\d+)$');
  const nums = data.slice(1)
    .map(r => { const m = re.exec(String(r[col] || '')); return m ? parseInt(m[1], 10) : 0; })
    .filter(n => n > 0);
  const max = nums.length > 0 ? Math.max(...nums) : 0;
  return prefixo + String(max + 1).padStart(5, '0');
}

// Previne injeção de fórmula no Sheets: prefixa com ' valores que começam com = + - @.
function _sanVal(v) {
  if (typeof v !== 'string') return v;
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

// Gera salt aleatório de 16 bytes em hex
function _gerarSalt() {
  const bytes = [];
  for (let i = 0; i < 16; i++) bytes.push(Math.floor(Math.random() * 256));
  return bytes.map(b => b.toString(16).padStart(2, '0')).join('');
}

// SHA-256(senha + salt) → hex string
function _hashSenha(senha, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(senha) + String(salt),
    Utilities.Charset.UTF_8
  );
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// Retorna "hash|salt" pronto para gravar na planilha
function _encriptarSenha(senha) {
  const salt = _gerarSalt();
  return _hashSenha(senha, salt) + '|' + salt;
}

// ============================================================
// AUTH
// ============================================================
function validarLogin(usuario, senha) {
  try {
    const rows = sheetToArray(ABAS.USUARIOS);
    const user = rows.find(r =>
      String(r.USUARIO).trim().toLowerCase() === String(usuario).trim().toLowerCase()
    );
    if (!user) return null;

    const senhaArmazenada = String(user.SENHA).trim();
    let senhaOk = false;
    if (senhaArmazenada.includes('|')) {
      // Formato seguro: "hash|salt"
      const partes = senhaArmazenada.split('|');
      const hashArmazenado = partes[0];
      const salt = partes[1];
      senhaOk = _hashSenha(senha, salt) === hashArmazenado;
    } else {
      // Legado: texto puro — ainda aceita para não travar usuários existentes
      senhaOk = senhaArmazenada === String(senha).trim();
    }

    if (!senhaOk) return null;
    const filiaisRaw = String(user.FILIAIS_LIBERADAS || '');
    const filiaisLiberadas = filiaisRaw ? filiaisRaw.split(',').map(f => f.trim()).filter(Boolean) : [];
    return { nome: user.NOME, usuario: user.USUARIO, email: user.EMAIL, perfil: user.PERFIL, filiaisLiberadas };
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
      filial:         { aba: ABAS.FILIAIS,         col: 'COD' },
      usuario:        { aba: ABAS.USUARIOS,        col: 'COD' }
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
function salvarCadastro(tipo, dados, usuarioLogado) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Impede criação/edição de usuário ADMIN por quem não é ADMIN
    if (tipo === 'usuario' && String(dados.PERFIL || '').toUpperCase() === 'ADMIN') {
      const perfilReal = _getPerfilReal(usuarioLogado);
      if (perfilReal !== 'ADMIN') {
        return { ok: false, msg: 'Apenas administradores podem criar ou editar usuários ADMIN' };
      }
    }

    const mapa = {
      fornecedor:     { aba: ABAS.FORNECEDORES,   cols: ['COD','NOME','CNPJ','EMAIL','CONTATO','COND_PAGAMENTO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'] },
      materia:        { aba: ABAS.MATERIAS,        cols: ['COD','DESCRICAO','UNIDADE','CATEGORIA'] },
      transportadora: { aba: ABAS.TRANSPORTADORAS, cols: ['COD','NOME','CNPJ','CONTATO','PRAZO','OBSERVACAO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'] },
      filial:         { aba: ABAS.FILIAIS,         cols: ['COD','NOME','CNPJ','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL','COD_TRANSPORTADORA'] },
      usuario:        { aba: ABAS.USUARIOS,        cols: ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL','FILIAIS_LIBERADAS'] }
    };
    if (!mapa[tipo]) return { ok: false, msg: 'Tipo inválido' };

    const cfg = mapa[tipo];
    const sh = getSheet(cfg.aba);
    const allData = sh.getDataRange().getValues();
    let headers = allData[0].map(h => String(h).trim());

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

    // Para usuários: hasheia a senha se uma nova for fornecida
    if (tipo === 'usuario') {
      const novaSenha = String(dados.SENHA || '').trim();
      if (novaSenha) {
        dados = Object.assign({}, dados, { SENHA: _encriptarSenha(novaSenha) });
      } else {
        // Sem nova senha → não sobrescreve o campo
        dados = Object.assign({}, dados);
        delete dados.SENHA;
      }
    }

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

    // Novo registro — senha obrigatória
    if (tipo === 'usuario' && !dados.SENHA) {
      return { ok: false, msg: 'Informe uma senha para o novo usuário' };
    }
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

function excluirCadastro(tipo, cod, usuarioLogado) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Apenas ADMIN pode excluir usuários
    if (tipo === 'usuario') {
      const perfilReal = _getPerfilReal(usuarioLogado);
      if (perfilReal !== 'ADMIN') {
        return { ok: false, msg: 'Apenas administradores podem excluir usuários' };
      }
    }

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
    const headers = data[0].map(h => String(h).trim());
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
    const hdrs = allData[0].map(h => String(h).trim());
    const iCF = hdrs.indexOf('COD_FORNECEDOR'), iMP = hdrs.indexOf('COD_MP'), iPR = hdrs.indexOf('PRECO');
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][iCF]).trim() === String(codForn).trim() &&
          String(allData[i][iMP]).trim() === String(codMP).trim()) {
        sh.getRange(i + 1, iPR + 1).setValue(preco);
        return { ok: true, msg: 'Preço atualizado' };
      }
    }
    _appendRowMapeado(sh, { COD_FORNECEDOR: codForn, COD_MP: codMP, PRECO: preco });
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

// Insere linha no sheet mapeando por nome de coluna (ignora colunas extras).
// Sanitiza strings para evitar injeção de fórmula.
function _appendRowMapeado(sh, dadosMap) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const row = headers.map(h => {
    if (!dadosMap.hasOwnProperty(h)) return '';
    const v = dadosMap[h];
    return _sanVal(v);
  });
  sh.appendRow(row);
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
    if (!shPedidos) throw new Error('Aba PEDIDOS não encontrada — execute setupPlanilha().');
    if (!shItens)   throw new Error('Aba ITENS_PEDIDO não encontrada — execute setupPlanilha().');

    // Valida itens ANTES de gravar o cabeçalho do pedido (evita pedido órfão sem itens)
    if (!Array.isArray(dados.itens) || dados.itens.length === 0) {
      return { ok: false, msg: 'O pedido deve ter pelo menos um item' };
    }

    const idPedido = _proximoId(shPedidos, 'ID_PEDIDO', 'PED-');
    const dataHoje = new Date();

    _appendRowMapeado(shPedidos, {
      ID_PEDIDO:          idPedido,
      DATA:               dataHoje,
      COD_FILIAL:         dados.filialCod,
      NOME_FILIAL:        dados.filialNome,
      COD_FORNECEDOR:     dados.fornecedorCod,
      NOME_FORNECEDOR:    dados.fornecedorNome,
      FRETE:              dados.frete || 'CIF',
      COD_TRANSPORTADORA: dados.transportadoraCod,
      NOME_TRANSPORTADORA:dados.transportadoraNome,
      PRAZO_ENTREGA:      dados.prazoEntrega,
      COND_PAGAMENTO:     dados.condPagamento || '',
      OBSERVACAO:         dados.observacao,
      USUARIO:            dados.usuarioLogado,
      VALOR_TOTAL:        dados.valorTotal,
      STATUS:             'ENVIADO'
    });

    dados.itens.forEach(item => {
      _appendRowMapeado(shItens, {
        ID_PEDIDO:   idPedido,
        COD_MP:      item.cod,
        DESCRICAO:   item.descricao,
        QUANTIDADE:  item.quantidade,
        UNIDADE:     item.unidade,
        PRECO_UNIT:  item.preco,
        SUBTOTAL:    item.subtotal
      });
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

    const nomeRemetente = dados.usuarioLogado || '';

    const dadosEmail = Object.assign({}, dados, {
      filialCNPJ:      filial.CNPJ      || '',
      filialEndereco:  [filial.ENDERECO, filial.BAIRRO, filial.CIDADE, filial.ESTADO].filter(Boolean).join(', '),
      fornecedorCNPJ:  fornecedor.CNPJ  || '',
      fornecedorEndereco: [fornecedor.ENDERECO, fornecedor.BAIRRO, fornecedor.CIDADE, fornecedor.ESTADO].filter(Boolean).join(', '),
      nomeRemetente
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

function montarEmailHTML(idPedido, data, dados, retif) {
  const dataFmt = Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const linhasItens = dados.itens.map(item => {
    const descricaoLimpa = _escHTML(String(item.descricao || '').replace(/[\n\r]+/g, ' ').trim());
    return `
    <tr>
      <td width="110" style="padding:9px 10px;border-bottom:1px solid #eee;font-family:monospace;font-size:11px;word-break:break-all;color:#1a3c5e;">${_escHTML(item.cod)}</td>
      <td style="padding:9px 10px;border-bottom:1px solid #eee;font-size:12px;word-break:break-word;">${descricaoLimpa}</td>
      <td width="50" style="padding:9px 10px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap;font-size:12px;">${_escHTML(item.quantidade)}</td>
      <td width="35" style="padding:9px 10px;border-bottom:1px solid #eee;text-align:left;white-space:nowrap;font-size:11px;color:#666;">${_escHTML(item.unidade)}</td>
      <td width="90" style="padding:9px 10px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap;font-size:12px;">R$&nbsp;${parseFloat(item.preco).toFixed(2)}</td>
      <td width="90" style="padding:9px 10px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap;font-size:12px;font-weight:600;">R$&nbsp;${parseFloat(item.subtotal).toFixed(2)}</td>
    </tr>
  `;
  }).join('');

  return `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;border:1px solid #dde3ea;border-radius:6px;overflow:hidden;">

    <!-- Cabeçalho -->
    <div style="background:#1a3c5e;padding:20px 28px;display:flex;align-items:center;gap:18px;">
      <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Marfim" style="height:52px;width:auto;border-radius:4px;flex-shrink:0;">
      <div>
        <div style="font-size:11px;font-weight:600;letter-spacing:3px;color:#e8a020;text-transform:uppercase;margin-bottom:4px;">${retif ? 'Pedido Retificado' : 'Pedido de Compra'}</div>
        <div style="font-size:20px;font-weight:700;color:white;">${_escHTML(idPedido)}</div>
        <div style="font-size:12px;color:rgba(255,255,255,0.6);margin-top:2px;">Emitido em ${dataFmt} · Solicitante: ${_escHTML(dados.usuarioLogado)}</div>
      </div>
    </div>

    ${retif ? `
    <!-- Aviso de retificação -->
    <div style="background:#fff3e0;border-left:4px solid #e65100;padding:10px 28px;font-size:13px;color:#5a3000;">
      <strong>Este pedido foi corrigido e substitui a versão anterior enviada para ${_escHTML(idPedido)}.</strong>${retif.motivo ? ' Motivo: ' + _escHTML(retif.motivo) : ''}
    </div>` : ''}

    <!-- Aviso NF -->
    <div style="background:#fff8e1;border-left:4px solid #e8a020;padding:10px 28px;font-size:13px;color:#5a4000;">
      <strong>Para a filial:</strong> ao receber esta entrega, informe o número <strong style="font-family:monospace;">${_escHTML(idPedido)}</strong> no sistema para registrar o recebimento da NF.
    </div>

    <!-- Partes: Comprador × Fornecedor -->
    <table style="width:100%;border-collapse:collapse;background:#f4f7fb;">
      <tr>
        <td style="width:50%;padding:16px 28px;vertical-align:top;border-right:1px solid #dde3ea;">
          <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1a3c5e;margin-bottom:8px;">Comprador</div>
          <div style="font-size:14px;font-weight:700;color:#1a1a1a;">${_escHTML(dados.filialNome)}</div>
          ${dados.filialCNPJ ? `<div style="font-size:12px;color:#555;margin-top:3px;">CNPJ: ${_escHTML(dados.filialCNPJ)}</div>` : ''}
          ${dados.filialEndereco ? `<div style="font-size:12px;color:#777;margin-top:3px;">${_escHTML(dados.filialEndereco)}</div>` : ''}
        </td>
        <td style="width:50%;padding:16px 28px;vertical-align:top;">
          <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1a3c5e;margin-bottom:8px;">Fornecedor</div>
          <div style="font-size:14px;font-weight:700;color:#1a1a1a;">${_escHTML(dados.fornecedorNome)}</div>
          ${dados.fornecedorCNPJ ? `<div style="font-size:12px;color:#555;margin-top:3px;">CNPJ: ${_escHTML(dados.fornecedorCNPJ)}</div>` : ''}
          ${dados.fornecedorEndereco ? `<div style="font-size:12px;color:#777;margin-top:3px;">${_escHTML(dados.fornecedorEndereco)}</div>` : ''}
        </td>
      </tr>
    </table>

    <!-- Entrega -->
    <table width="100%" style="border-collapse:collapse;background:#fff;border-top:1px solid #dde3ea;border-bottom:1px solid #dde3ea;" cellpadding="0" cellspacing="0">
      <tr>
        <td width="80" style="padding:14px 20px;vertical-align:top;border-right:1px solid #dde3ea;white-space:nowrap;">
          <div style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Frete</div>
          <span style="display:inline-block;background:${dados.frete==='CIF'?'#e8f5e9':'#fff3e0'};border:1px solid ${dados.frete==='CIF'?'#a5d6a7':'#ffcc80'};border-radius:4px;padding:2px 12px;font-size:13px;font-weight:700;color:${dados.frete==='CIF'?'#2e7d32':'#e65100'};">${dados.frete || 'CIF'}</span>
        </td>
        ${dados.frete !== 'CIF' ? `
        <td style="padding:14px 20px;vertical-align:top;border-right:1px solid #dde3ea;">
          <div style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Transportadora</div>
          <div style="font-size:13px;color:#1a1a1a;">${_escHTML(dados.transportadoraNome || '—')}</div>
        </td>` : ''}
        <td width="140" style="padding:14px 20px;vertical-align:top;white-space:nowrap;${dados.condPagamento ? 'border-right:1px solid #dde3ea;' : ''}">
          <div style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Prazo de Entrega</div>
          <div style="font-size:13px;color:#1a1a1a;">${_escHTML(dados.prazoEntrega || '—')}</div>
        </td>
        ${dados.condPagamento ? `
        <td width="140" style="padding:14px 20px;vertical-align:top;white-space:nowrap;">
          <div style="font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">Condição de Pagamento</div>
          <div style="font-size:13px;color:#1a1a1a;">${_escHTML(dados.condPagamento)}</div>
        </td>` : ''}
      </tr>
    </table>

    <!-- Itens -->
    <div style="padding:20px 20px;background:#fff;">
      <table width="100%" style="border-collapse:collapse;" cellpadding="0" cellspacing="0">
        <thead>
          <tr style="background:#1a3c5e;color:white;">
            <th width="110" style="padding:10px 10px;text-align:left;font-size:11px;font-weight:600;">Código</th>
            <th style="padding:10px 10px;text-align:left;font-size:11px;font-weight:600;">Descrição</th>
            <th width="50" style="padding:10px 10px;text-align:right;font-size:11px;font-weight:600;white-space:nowrap;">Qtd</th>
            <th width="35" style="padding:10px 10px;text-align:left;font-size:11px;font-weight:600;white-space:nowrap;">Un</th>
            <th width="90" style="padding:10px 10px;text-align:right;font-size:11px;font-weight:600;white-space:nowrap;">Preço Unit.</th>
            <th width="90" style="padding:10px 10px;text-align:right;font-size:11px;font-weight:600;white-space:nowrap;">Subtotal</th>
          </tr>
        </thead>
        <tbody>${linhasItens}</tbody>
        <tfoot>
          <tr style="background:#1a3c5e;">
            <td colspan="5" style="padding:14px 12px;text-align:right;font-weight:700;font-size:13px;color:rgba(255,255,255,0.85);letter-spacing:1px;text-transform:uppercase;">TOTAL DO PEDIDO</td>
            <td style="padding:14px 12px;text-align:right;font-weight:800;font-size:18px;color:#e8a020;white-space:nowrap;">R$&nbsp;${parseFloat(dados.valorTotal).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
          </tr>
        </tfoot>
      </table>
      ${dados.observacao ? `<div style="margin-top:16px;padding:12px 14px;background:#f9f9f9;border-left:3px solid #1a3c5e;border-radius:3px;font-size:13px;color:#333;"><strong>Observações:</strong> ${_escHTML(dados.observacao)}</div>` : ''}
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
            <div style="font-size:14px;font-weight:700;color:white;">${dados.nomeRemetente || dados.usuarioLogado}</div>
          </td>
        </tr>
      </table>
    </div>

  </div>`;
}

function montarEmailTexto(idPedido, data, dados, retif) {
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
    retif ? `PEDIDO RETIFICADO  ${idPedido}` : `PEDIDO DE COMPRA  ${idPedido}`,
    `Emitido em ${dataFmt}  |  Solicitante: ${dados.usuarioLogado}`,
    retif ? `Este pedido substitui a versão anterior enviada.${retif.motivo ? ' Motivo: ' + retif.motivo : ''}` : null,
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
    `${dados.nomeRemetente || dados.usuarioLogado} — Marfim`,
    sep2,
  ].filter(l => l !== null && l !== undefined).join('\n');
}

// Converte string ISO (vinda de sheetToArray) para dd/MM/yyyy.
// Se já for texto comum (ex: "150", "29/05/2026"), devolve como está.
function _formatarDataISO(val) {
  if (!val) return '';
  const s = String(val);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    try {
      return Utilities.formatDate(new Date(s), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } catch(e) { return s; }
  }
  return s;
}

// ============================================================
// TESTE DE LAYOUT DE EMAIL
// Execute esta função diretamente no Apps Script Editor para
// receber os pedidos 1–4 em johnny@marfim-bahia.ind.br
// ============================================================
function enviarEmailTestePedidos() {
  const EMAIL_TESTE  = 'johnny@marfim-bahia.ind.br';
  const NUMEROS_ALVO = ['1', '2', '3', '4'];

  const todosPedidos = sheetToArray(ABAS.PEDIDOS);
  const todosItens   = sheetToArray(ABAS.ITENS_PEDIDO);

  if (todosPedidos.length === 0) {
    Logger.log('Nenhum pedido encontrado na planilha.');
    return;
  }

  const erros    = [];
  let   enviados = 0;

  NUMEROS_ALVO.forEach(function(num) {
    const ped = todosPedidos.find(function(p) {
      const id = String(p.ID_PEDIDO || '').trim().toUpperCase();
      const n  = id.replace(/^PED-?0*/i, '') || '0';
      return n === num;
    });

    if (!ped) {
      erros.push('Pedido #' + num + ' não encontrado');
      return;
    }

    const idPedido   = String(ped.ID_PEDIDO).trim();
    const itensPed   = todosItens.filter(function(i) {
      return String(i.ID_PEDIDO || '').trim() === idPedido;
    });
    const fornecedor = buscarCodigo('fornecedor', ped.COD_FORNECEDOR) || {};
    const filial     = buscarCodigo('filial',     ped.COD_FILIAL)     || {};

    const dados = {
      filialNome:         String(ped.NOME_FILIAL      || ''),
      filialCNPJ:         String(filial.CNPJ          || ''),
      filialEndereco:     [filial.ENDERECO, filial.BAIRRO, filial.CIDADE, filial.ESTADO].filter(Boolean).join(', '),
      fornecedorNome:     String(ped.NOME_FORNECEDOR  || ''),
      fornecedorCNPJ:     String(fornecedor.CNPJ      || ''),
      fornecedorEndereco: [fornecedor.ENDERECO, fornecedor.BAIRRO, fornecedor.CIDADE, fornecedor.ESTADO].filter(Boolean).join(', '),
      frete:              String(ped.FRETE             || 'CIF'),
      transportadoraNome: String(ped.NOME_TRANSPORTADORA || ''),
      prazoEntrega:       _formatarDataISO(ped.PRAZO_ENTREGA),
      condPagamento:      String(ped.COND_PAGAMENTO    || ''),
      observacao:         String(ped.OBSERVACAO        || ''),
      usuarioLogado:      String(ped.USUARIO           || ''),
      nomeRemetente:      String(ped.USUARIO           || ''),
      valorTotal:         parseFloat(ped.VALOR_TOTAL)  || 0,
      itens: itensPed.map(function(i) {
        return {
          cod:        String(i.COD_MP    || ''),
          descricao:  String(i.DESCRICAO || ''),
          quantidade: i.QUANTIDADE,
          unidade:    String(i.UNIDADE   || ''),
          preco:      parseFloat(i.PRECO_UNIT) || 0,
          subtotal:   parseFloat(i.SUBTOTAL)   || 0
        };
      })
    };

    // DATA vem como string ISO do sheetToArray; converte de volta para Date
    const dataPedido = ped.DATA ? new Date(ped.DATA) : new Date();

    try {
      const htmlEmail  = montarEmailHTML(idPedido, dataPedido, dados);
      const textoEmail = montarEmailTexto(idPedido, dataPedido, dados);
      MailApp.sendEmail({
        to:       EMAIL_TESTE,
        subject:  '[TESTE LAYOUT] Pedido ' + idPedido + ' — ' + (dados.filialNome || 'Filial'),
        body:     textoEmail,
        htmlBody: htmlEmail
      });
      enviados++;
      Logger.log('OK — ' + idPedido + ' enviado para ' + EMAIL_TESTE);
    } catch(e) {
      erros.push(idPedido + ': ' + e.message);
      Logger.log('ERRO ao enviar ' + idPedido + ': ' + e.message);
    }
  });

  const resumo = 'Enviados: ' + enviados + '/' + NUMEROS_ALVO.length +
    (erros.length ? ' | Erros: ' + erros.join('; ') : ' | Tudo OK');
  Logger.log(resumo);
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
    const hdrs = allData[0].map(h => String(h).trim());
    const iCF = hdrs.indexOf('COD_FORNECEDOR'), iFI = hdrs.indexOf('COD_FILIAL'), iCT = hdrs.indexOf('COD_TRANSPORTADORA');
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][iCF]).trim() === String(codForn).trim() &&
          String(allData[i][iFI]).trim() === String(codFilial).trim()) {
        sh.getRange(i + 1, iCT + 1).setValue(codTransp);
        return { ok: true, msg: 'Transportadora atualizada' };
      }
    }
    _appendRowMapeado(sh, { COD_FORNECEDOR: codForn, COD_FILIAL: codFilial, COD_TRANSPORTADORA: codTransp });
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
function listarIdsPedidos() {
  try {
    return sheetToArray(ABAS.PEDIDOS)
      .map(p => ({ id: String(p.ID_PEDIDO||'').trim(), status: String(p.STATUS||''), fornecedor: String(p.NOME_FORNECEDOR||'') }))
      .filter(p => p.id);
  } catch(e) {
    return [];
  }
}

function buscarPedidoParaRecebimento(idPedido) {
  try {
    const sh = getSheet(ABAS.PEDIDOS);
    if (!sh) return { erro: 'Aba PEDIDOS não encontrada — execute setupPlanilha() no Apps Script.' };
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    if (pedidos.length === 0) return { erro: 'Nenhum pedido cadastrado ainda.' };
    const query = String(idPedido).trim().toUpperCase();
    const numQuery = query.replace(/^PED-?/, '').replace(/^0+/, '') || '0';
    const ped = pedidos.find(p => {
      const id = String(p.ID_PEDIDO).trim().toUpperCase();
      if (id === query) return true;
      const num = id.replace(/^PED-?/, '').replace(/^0+/, '') || '0';
      return num === numQuery;
    });
    if (!ped) return { erro: 'Pedido "' + idPedido + '" não encontrado. IDs disponíveis: ' + pedidos.slice(0,5).map(p=>p.ID_PEDIDO).join(', ') + (pedidos.length > 5 ? '...' : '') };
    const itens = sheetToArray(ABAS.ITENS_PEDIDO)
      .filter(i => String(i.ID_PEDIDO).trim() === String(ped.ID_PEDIDO).trim());
    const cancelado = String(ped.STATUS || '').trim().toUpperCase() === 'CANCELADO';
    if (cancelado) {
      return { cancelado: true, pedido: ped, itens, jaRecebido: false, recebimentos: [] };
    }
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
    if (!sh)      throw new Error('Aba RECEBIMENTOS não encontrada — execute setupPlanilha().');
    if (!shItens) throw new Error('Aba ITENS_RECEBIMENTO não encontrada — execute setupPlanilha().');

    // Bloqueia NF duplicada para o mesmo pedido
    const recsExist = sheetToArray(ABAS.RECEBIMENTOS);
    const nfDup = recsExist.find(r =>
      String(r.ID_PEDIDO).trim()  === String(dados.idPedido).trim() &&
      String(r.NF_NUMERO).trim()  === String(dados.nfNumero).trim()
    );
    if (nfDup) {
      return { ok: false, msg: `NF ${dados.nfNumero} já lançada para este pedido (${nfDup.ID_RECEBIMENTO})` };
    }

    const idRec   = _proximoId(sh, 'ID_RECEBIMENTO', 'REC-');
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

    const nfDataObj = dados.nfData
      ? new Date(dados.nfData.replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$3-$2-$1'))
      : '';
    _appendRowMapeado(sh, {
      ID_RECEBIMENTO:    idRec,
      ID_PEDIDO:         dados.idPedido,
      NF_NUMERO:         dados.nfNumero,
      DATA_NF:           nfDataObj,
      DATA_RECEBIMENTO:  dataHoje,
      COD_FILIAL:        dados.codFilial,
      NOME_FILIAL:       dados.nomeFilial,
      COD_FORNECEDOR:    dados.codFornecedor,
      NOME_FORNECEDOR:   dados.nomeFornecedor,
      USUARIO:           dados.usuarioLogado,
      VALOR_TOTAL_PEDIDO:  parseFloat(dados.valorTotalPedido) || 0,
      VALOR_TOTAL_RECEBIDO: valorTotalRec,
      DIVERGENCIA_QTD:   temDivQtd   ? 'SIM' : 'NÃO',
      DIVERGENCIA_PRECO: temDivPreco ? 'SIM' : 'NÃO',
      DIVERGENCIA_PAGTO: dados.divPagto ? 'SIM' : 'NÃO',
      PAGTO_PRAZO:       dados.condPagamento  || '',
      PAGTO_ESPERADO:    dados.pagtoEsperado  || '',
      PAGTO_NF:          dados.pagtoNF        || '',
      OBSERVACAO:        dados.observacao     || ''
    });

    dados.itens.forEach(item => {
      const qtdPed  = parseFloat(item.qtdPedida)    || 0;
      const qtdRec  = parseFloat(item.qtdRecebida)  || 0;
      const precPed = parseFloat(item.precoPedido)  || 0;
      const precRec = parseFloat(item.precoRecebido)|| 0;
      const divQtd   = Math.abs(qtdRec  - qtdPed)  > 0.001;
      const divPreco = Math.abs(precRec - precPed) > 0.001;
      _appendRowMapeado(shItens, {
        ID_RECEBIMENTO:   idRec,
        ID_PEDIDO:        dados.idPedido,
        COD_MP:           item.codMP,
        DESCRICAO:        item.descricao,
        QTD_PEDIDA:       qtdPed,
        QTD_RECEBIDA:     qtdRec,
        PRECO_PEDIDO:     precPed,
        PRECO_RECEBIDO:   precRec,
        SUBTOTAL_PEDIDO:  qtdPed * precPed,
        SUBTOTAL_RECEBIDO: qtdRec * precRec,
        DIV_QTD:          divQtd   ? 'SIM' : 'NÃO',
        DIV_PRECO:        divPreco ? 'SIM' : 'NÃO'
      });
    });

    _logNF('LANÇAMENTO', idRec, dados.idPedido, dados.nfNumero, dados.codFilial, dados.nomeFilial, dados.usuarioLogado || '');

    const temDivPagto = !!dados.divPagto;
    if (temDivQtd || temDivPreco || temDivPagto) {
      try {
        _enviarEmailDivergencia(idRec, dados, temDivQtd, temDivPreco, temDivPagto, valorTotalRec);
      } catch(eEmail) {
        logErro('Email divergência: ' + eEmail.message);
      }
    }

    return {
      ok: true,
      msg: 'NF lançada com sucesso',
      id: idRec,
      divQtd: temDivQtd,
      divPreco: temDivPreco,
      divPagto: temDivPagto
    };
  } catch(e) {
    logErro('salvarRecebimento: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _enviarEmailDivergencia(idRec, dados, temDivQtd, temDivPreco, temDivPagto, valorTotalRec) {
  const EMAIL_FIXO = 'marco@marfim.ind.br';
  const emailUsuario = String(dados.emailUsuario || '').trim().toLowerCase();
  const destinatarios = [EMAIL_FIXO];
  if (emailUsuario && emailUsuario !== EMAIL_FIXO.toLowerCase()) {
    destinatarios.push(dados.emailUsuario.trim());
  }

  const dataFmt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const divLabels = [temDivQtd && 'Quantidade', temDivPreco && 'Preço', temDivPagto && 'Prazo de Pagamento']
    .filter(Boolean).join(', ');

  const linhasItens = dados.itens.map(item => {
    const qtdPed  = parseFloat(item.qtdPedida)    || 0;
    const qtdRec  = parseFloat(item.qtdRecebida)  || 0;
    const precPed = parseFloat(item.precoPedido)  || 0;
    const precRec = parseFloat(item.precoRecebido)|| 0;
    const divQ = Math.abs(qtdRec  - qtdPed)  > 0.001;
    const divP = Math.abs(precRec - precPed) > 0.001;
    const bg    = (divQ || divP) ? 'background:#fff3cd;' : '';
    return `<tr style="${bg}">
      <td style="padding:8px 10px;border-bottom:1px solid #eee;font-family:monospace;font-size:12px;">${_escHTML(item.codMP)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px;">${_escHTML(item.descricao)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;">${qtdPed.toFixed(3)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;${divQ?'color:#c0392b;font-weight:bold;':''}">${qtdRec.toFixed(3)}${divQ?' ⚠':''}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;">R$ ${precPed.toFixed(2)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;${divP?'color:#c0392b;font-weight:bold;':''}">${'R$ '+precRec.toFixed(2)}${divP?' ⚠':''}</td>
    </tr>`;
  }).join('');

  const secaoPagto = temDivPagto ? (() => {
    const esperadas        = String(dados.pagtoEsperado        || '').split(',').filter(Boolean);
    const esperadasEntrega = String(dados.pagtoEsperadoEntrega || '').split(',').filter(Boolean);
    const informadas       = String(dados.pagtoNF              || '').split(',').filter(Boolean);

    function _isoToDisplay(isoStr) {
      if (!isoStr) return null;
      const [y, m, d] = isoStr.split('-').map(Number);
      if (!y) return null;
      return { str: String(d).padStart(2,'0') + '/' + String(m).padStart(2,'0') + '/' + y,
               ms: Date.UTC(y, m - 1, d) };
    }

    const rowsParcelas = esperadas.map((espISO, i) => {
      const infStr  = informadas[i] || '—';
      const esp     = _isoToDisplay(espISO);
      const espEnt  = _isoToDisplay(esperadasEntrega[i] || '');
      const espStr  = esp ? esp.str : '—';

      // Linha de referência alternativa (entrega) quando diferente da emissão
      const altRefStr = (espEnt && espEnt.str !== espStr)
        ? ` <span style="color:#666;font-size:11px">ou <strong>${espEnt.str}</strong> (pela entrega)</span>`
        : '';

      let diffStr = '';
      if (esp && infStr !== '—' && /^\d{2}\/\d{2}\/\d{4}$/.test(infStr)) {
        const [dd, mm, aaaa] = infStr.split('/').map(Number);
        const infMs = Date.UTC(aaaa, mm - 1, dd);
        const diff  = Math.round((infMs - esp.ms) / 86400000);
        if (diff > 0)      diffStr = ` <span style="color:#c0392b;font-size:11px">(+${diff} dias — prazo maior)</span>`;
        else if (diff < 0) diffStr = ` <span style="color:#e8a020;font-size:11px">(${Math.abs(diff)} dias — prazo menor)</span>`;
      }
      return `<tr>
        <td style="padding:4px 0;color:#666;width:90px;vertical-align:top">Parcela ${i+1}:</td>
        <td style="padding:4px 0">Esperado <strong>${espStr}</strong>${altRefStr} &nbsp;→&nbsp; NF <strong style="color:#c0392b">${infStr}</strong>${diffStr}</td>
      </tr>`;
    }).join('');
    return `
    <div style="margin:20px 28px;padding:14px 16px;background:#fff3cd;border-left:4px solid #e8a020;border-radius:4px;">
      <div style="font-weight:700;color:#5a4000;margin-bottom:8px;">⚠ Divergência de Prazo de Pagamento</div>
      <table style="font-size:13px;width:100%;border-collapse:collapse;">
        <tr><td style="padding:4px 0;color:#666;width:90px">Condição:</td><td style="font-weight:600;">${_escHTML(dados.condPagamento||'—')}</td></tr>
        ${rowsParcelas}
      </table>
    </div>`;
  })() : '';

  const htmlBody = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;border:1px solid #dde3ea;border-radius:6px;overflow:hidden;">
    <div style="background:#1a3c5e;padding:20px 28px;">
      <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Marfim" style="height:48px;width:auto;border-radius:4px;margin-bottom:10px;display:block;">
      <div style="font-size:11px;font-weight:600;letter-spacing:3px;color:#e8a020;text-transform:uppercase;margin-bottom:4px;">Alerta de Divergência — Recebimento NF</div>
      <div style="font-size:18px;font-weight:700;color:white;">${_escHTML(dados.idPedido)} · NF ${_escHTML(dados.nfNumero)}</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.65);margin-top:3px;">Registrado em ${dataFmt} por ${_escHTML(dados.usuarioLogado)}</div>
    </div>
    <div style="background:#fdecea;border-left:4px solid #c0392b;padding:12px 28px;font-size:13px;color:#7b1a1a;">
      <strong>Divergência detectada em:</strong> ${divLabels}
    </div>
    <div style="padding:16px 28px 0;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr style="background:#1a3c5e;color:white;">
          <th style="padding:8px 10px;text-align:left;">Código</th>
          <th style="padding:8px 10px;text-align:left;">Descrição</th>
          <th style="padding:8px 10px;text-align:right;">Qtd Ped.</th>
          <th style="padding:8px 10px;text-align:right;">Qtd Rec.</th>
          <th style="padding:8px 10px;text-align:right;">Preço Ped.</th>
          <th style="padding:8px 10px;text-align:right;">Preço Rec.</th>
        </tr>
        ${linhasItens}
      </table>
    </div>
    ${secaoPagto}
    <div style="padding:16px 28px;">
      <table style="font-size:13px;width:100%;">
        <tr><td style="color:#666;width:180px;">Pedido:</td><td><strong>${_escHTML(dados.idPedido)}</strong></td></tr>
        <tr><td style="color:#666;">Recebimento:</td><td><strong>${_escHTML(idRec)}</strong></td></tr>
        <tr><td style="color:#666;">NF Nº:</td><td>${_escHTML(dados.nfNumero)}</td></tr>
        <tr><td style="color:#666;">Data NF:</td><td>${_escHTML(dados.nfData||'—')}</td></tr>
        <tr><td style="color:#666;">Filial:</td><td>${_escHTML(dados.nomeFilial)} (${_escHTML(dados.codFilial)})</td></tr>
        <tr><td style="color:#666;">Fornecedor:</td><td>${_escHTML(dados.nomeFornecedor)}</td></tr>
        <tr><td style="color:#666;">Total Recebido:</td><td><strong>R$ ${valorTotalRec.toFixed(2)}</strong></td></tr>
        ${dados.observacao ? `<tr><td style="color:#666;">Observação:</td><td>${_escHTML(dados.observacao)}</td></tr>` : ''}
      </table>
    </div>
    <div style="background:#f4f7fb;padding:12px 28px;font-size:11px;color:#888;text-align:center;">
      Sistema de Compras Marfim · Este é um email automático.
    </div>
  </div>`;

  const textBody =
    `ALERTA DE DIVERGÊNCIA — ${dados.idPedido} · NF ${dados.nfNumero}\n` +
    `Registrado em ${dataFmt} por ${dados.usuarioLogado}\n` +
    `Divergência em: ${divLabels}\n\n` +
    dados.itens.map(item => {
      const qtdPed = parseFloat(item.qtdPedida)||0, qtdRec = parseFloat(item.qtdRecebida)||0;
      const pPed   = parseFloat(item.precoPedido)||0, pRec = parseFloat(item.precoRecebido)||0;
      return `${item.codMP} ${item.descricao} | Qtd: ${qtdPed}→${qtdRec} | Preço: ${pPed.toFixed(2)}→${pRec.toFixed(2)}`;
    }).join('\n') +
    (temDivPagto ? (() => {
      const esperadas  = String(dados.pagtoEsperado || '').split(',').filter(Boolean);
      const informadas = String(dados.pagtoNF       || '').split(',').filter(Boolean);
      const linhas = esperadas.map((espISO, i) => {
        const infStr = informadas[i] || '—';
        let espStr = espISO;
        if (espISO) {
          const [y, m, d] = espISO.split('-').map(Number);
          espStr = String(d).padStart(2,'0') + '/' + String(m).padStart(2,'0') + '/' + y;
        }
        let diffStr = '';
        if (espISO && infStr !== '—' && /^\d{2}\/\d{2}\/\d{4}$/.test(infStr)) {
          const [y, m, d] = espISO.split('-').map(Number);
          const [dd, mm, aaaa] = infStr.split('/').map(Number);
          const diff = Math.round((Date.UTC(aaaa, mm-1, dd) - Date.UTC(y, m-1, d)) / 86400000);
          if (diff > 0)      diffStr = ` (+${diff}d — prazo maior)`;
          else if (diff < 0) diffStr = ` (${Math.abs(diff)}d — prazo menor)`;
        }
        return `  Parcela ${i+1}: Esperado ${espStr} → NF ${infStr}${diffStr}`;
      }).join('\n');
      return `\n\nDivergência Pagto:\nCondição: ${dados.condPagamento}\n${linhas}`;
    })() : '');

  MailApp.sendEmail({
    to:      destinatarios.join(','),
    replyTo: EMAIL_FIXO,
    subject: `⚠ Divergência NF — ${dados.idPedido} · ${dados.nomeFornecedor}`,
    body:    textBody,
    htmlBody: htmlBody
  });
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

function _logNF(acao, idRec, idPedido, nfNumero, codFilial, nomeFilial, usuario, detalhe) {
  try {
    _appendRowMapeado(getSheet(ABAS.LOG_NF), {
      DATA_HORA:   new Date(),
      ACAO:        acao,
      ID_RECEBIMENTO: idRec   || '',
      ID_PEDIDO:   idPedido   || '',
      NF_NUMERO:   nfNumero   || '',
      COD_FILIAL:  codFilial  || '',
      NOME_FILIAL: nomeFilial || '',
      USUARIO:     usuario    || '',
      DETALHE:     detalhe    || ''
    });
  } catch(e) {
    logErro('_logNF: ' + e.message);
  }
}

function getLogNF(idPedido) {
  try {
    return sheetToArray(ABAS.LOG_NF)
      .filter(r => String(r.ID_PEDIDO).trim() === String(idPedido).trim())
      .map(r => ({ ...r, DATA_HORA: r.DATA_HORA ? new Date(r.DATA_HORA).toISOString() : '' }))
      .reverse(); // mais recente primeiro
  } catch(e) {
    logErro('getLogNF: ' + e.message);
    return [];
  }
}

function excluirRecebimento(idRec, usuarioLogado, filiaisLiberadas, perfil, usuarioLogin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // Busca dados do recebimento antes de apagar (para log e verificação de permissão)
    const recs = sheetToArray(ABAS.RECEBIMENTOS);
    const rec  = recs.find(r => String(r.ID_RECEBIMENTO).trim() === String(idRec).trim());
    if (!rec) return { ok: false, msg: 'Recebimento não encontrado' };

    // Busca perfil real do usuário na planilha (não confia no param enviado pelo frontend)
    const loginReal = usuarioLogin || usuarioLogado;
    const perfilReal = _getPerfilReal(loginReal);
    if (!perfilReal) return { ok: false, msg: 'Usuário não encontrado' };

    if (perfilReal !== 'ADMIN') {
      // Usuário não-ADMIN só pode excluir NFs das suas filiais liberadas
      const users = sheetToArray(ABAS.USUARIOS);
      const user  = users.find(u => String(u.USUARIO).trim().toLowerCase() === String(loginReal).trim().toLowerCase());
      const liberadas = String(user?.FILIAIS_LIBERADAS || '').split(',').map(f => f.trim()).filter(Boolean);
      if (!liberadas.includes(String(rec.COD_FILIAL).trim())) {
        return { ok: false, msg: 'Sem permissão para excluir NF desta filial' };
      }
    }

    const _deleteRows = (sh, colName) => {
      const data = sh.getDataRange().getValues();
      const hdrs = data[0].map(h => String(h).trim());
      const col  = hdrs.indexOf(colName);
      if (col < 0) return;
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][col]).trim() === String(idRec).trim()) {
          sh.deleteRow(i + 1);
        }
      }
    };

    _deleteRows(getSheet(ABAS.ITENS_RECEBIMENTO), 'ID_RECEBIMENTO');
    _deleteRows(getSheet(ABAS.RECEBIMENTOS),       'ID_RECEBIMENTO');

    _logNF('EXCLUSÃO', idRec, rec.ID_PEDIDO, rec.NF_NUMERO, rec.COD_FILIAL, rec.NOME_FILIAL, usuarioLogado || '');

    return { ok: true, msg: 'NF excluída — pedido liberado para novo lançamento' };
  } catch(e) {
    logErro('excluirRecebimento: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getHistoricoMultiFiliais(codFiliais, usuarioLogin) {
  try {
    let filiais = (codFiliais || []).map(f => String(f).trim()).filter(Boolean);
    // Valida filiais no servidor contra as permissões reais do usuário
    if (usuarioLogin) {
      const users = sheetToArray(ABAS.USUARIOS);
      const user = users.find(u => String(u.USUARIO).trim().toLowerCase() === String(usuarioLogin).trim().toLowerCase());
      if (user && String(user.PERFIL).toUpperCase() === 'FILIAL') {
        const liberadas = String(user.FILIAIS_LIBERADAS || '').split(',').map(f => f.trim()).filter(Boolean);
        filiais = filiais.filter(f => liberadas.includes(f));
      }
    }
    if (filiais.length === 0) return [];
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    const filtrados = pedidos.filter(p => filiais.includes(String(p.COD_FILIAL).trim()));
    const enriquecidos = _enriquecerPedidosComNF(filtrados);
    const todosItens = sheetToArray(ABAS.ITENS_PEDIDO);
    return enriquecidos.map(p => {
      const itens = todosItens.filter(i => String(i.ID_PEDIDO).trim() === String(p.ID_PEDIDO).trim());
      return { ...p, itens };
    });
  } catch(e) {
    logErro('getHistoricoMultiFiliais: ' + e.message);
    throw e;
  }
}

function cancelarPedido(idPedido, observacao, usuarioNome, emailUsuario) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const sh = getSheet(ABAS.PEDIDOS);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const iStatus = headers.indexOf('STATUS');
    const iId     = headers.indexOf('ID_PEDIDO');
    const iObs    = headers.indexOf('OBSERVACAO');
    const iCodFilial  = headers.indexOf('COD_FILIAL');
    const iNomeFilial = headers.indexOf('NOME_FILIAL');

    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iId]).trim() === String(idPedido).trim()) {
        rowIdx = i;
        break;
      }
    }
    if (rowIdx < 0) return { ok: false, msg: 'Pedido não encontrado' };

    const statusAtual = String(data[rowIdx][iStatus] || '').trim().toUpperCase();
    if (statusAtual === 'CANCELADO') return { ok: false, msg: 'Pedido já está cancelado' };

    // Bloqueia cancelamento se há NFs lançadas
    const nfsLancadas = sheetToArray(ABAS.RECEBIMENTOS)
      .filter(r => String(r.ID_PEDIDO).trim() === String(idPedido).trim());
    if (nfsLancadas.length > 0) {
      const ids = nfsLancadas.map(r => r.ID_RECEBIMENTO).join(', ');
      return { ok: false, msg: `Pedido possui ${nfsLancadas.length} NF(s) lançada(s) (${ids}). Exclua as NFs antes de cancelar.` };
    }

    // Atualiza STATUS
    sh.getRange(rowIdx + 1, iStatus + 1).setValue('CANCELADO');

    // Atualiza OBSERVACAO
    if (iObs >= 0 && observacao) {
      const obsAtual = String(data[rowIdx][iObs] || '').trim();
      const novaObs = obsAtual ? obsAtual + ' | CANCELADO: ' + observacao : 'CANCELADO: ' + observacao;
      sh.getRange(rowIdx + 1, iObs + 1).setValue(novaObs);
    }

    const codFilial  = String(data[rowIdx][iCodFilial]  || '');
    const nomeFilial = String(data[rowIdx][iNomeFilial] || '');
    const rowData    = data[rowIdx];

    _logNF('CANCELAMENTO', '', idPedido, '', codFilial, nomeFilial, usuarioNome, observacao || '');

    _enviarEmailCancelamento(rowData, headers, usuarioNome, emailUsuario, observacao || '');

    return { ok: true, msg: 'Pedido cancelado com sucesso' };
  } catch(e) {
    logErro('cancelarPedido: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _enviarEmailCancelamento(rowData, headers, usuarioNome, emailUsuario, observacao) {
  try {
    const get = col => {
      const idx = headers.indexOf(col);
      return idx >= 0 ? String(rowData[idx] || '') : '';
    };

    const codFornecedor  = get('COD_FORNECEDOR');
    const codFilial      = get('COD_FILIAL');
    const nomeFilial     = get('NOME_FILIAL');
    const nomeFornecedor = get('NOME_FORNECEDOR');
    const idPedido       = get('ID_PEDIDO');
    const valorTotal     = get('VALOR_TOTAL');

    const dataRaw = get('DATA');
    let dataFmt = '—';
    try {
      if (dataRaw) {
        const d = new Date(dataRaw);
        if (!isNaN(d.getTime())) dataFmt = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    } catch(e) {}

    // Emails do fornecedor
    const fornecedor = buscarCodigo('fornecedor', codFornecedor) || {};
    const emailsForn = String(fornecedor.EMAIL || '')
      .split(';').map(e => e.trim()).filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));

    // Emails da filial
    const filial = buscarCodigo('filial', codFilial) || {};
    const emailsFilial = String(filial.EMAIL_RESPONSAVEL || '')
      .split(';').map(e => e.trim()).filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));

    if (emailsForn.length === 0) {
      logErro('_enviarEmailCancelamento: nenhum email válido para fornecedor ' + codFornecedor);
      return;
    }

    const ccSet = new Set([...emailsFilial, 'marco@marfim.ind.br']);
    if (emailUsuario && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailUsuario.trim())) {
      ccSet.add(emailUsuario.trim());
    }
    const ccList = [...ccSet].filter(Boolean).join(',');

    const valorFmt = (() => {
      const n = parseFloat(valorTotal);
      return isNaN(n) ? '—' : 'R$ ' + n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    })();

    const htmlBody = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;border:1px solid #dde3ea;border-radius:6px;overflow:hidden;">
    <div style="background:#c0392b;padding:20px 28px;">
      <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Marfim" style="height:48px;width:auto;border-radius:4px;margin-bottom:10px;display:block;">
      <div style="font-size:11px;font-weight:600;letter-spacing:3px;color:#fdecea;text-transform:uppercase;margin-bottom:4px;">Pedido Cancelado</div>
      <div style="font-size:20px;font-weight:700;color:white;">${_escHTML(idPedido)} — ${_escHTML(nomeFornecedor)}</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.7);margin-top:2px;">Cancelado por: ${_escHTML(usuarioNome)}</div>
    </div>
    <div style="background:#fdecea;border-left:4px solid #c0392b;padding:16px 28px;text-align:center;">
      <div style="font-size:28px;margin-bottom:8px;">🚫</div>
      <div style="font-size:18px;font-weight:700;color:#c0392b;margin-bottom:4px;">PEDIDO CANCELADO</div>
      <div style="font-size:13px;color:#7b1a1a;">Este pedido foi cancelado e não terá entrega.</div>
    </div>
    <div style="padding:20px 28px;background:#fff;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr><td style="padding:6px 0;color:#666;width:180px;">ID do Pedido:</td><td><strong style="font-family:monospace">${_escHTML(idPedido)}</strong></td></tr>
        <tr><td style="padding:6px 0;color:#666;">Fornecedor:</td><td>${_escHTML(nomeFornecedor)}</td></tr>
        <tr><td style="padding:6px 0;color:#666;">Filial:</td><td>${_escHTML(nomeFilial)} (${_escHTML(codFilial)})</td></tr>
        <tr><td style="padding:6px 0;color:#666;">Valor Total:</td><td><strong>${_escHTML(valorFmt)}</strong></td></tr>
        <tr><td style="padding:6px 0;color:#666;">Data do Pedido:</td><td>${_escHTML(dataFmt)}</td></tr>
        <tr><td style="padding:6px 0;color:#666;">Cancelado por:</td><td>${_escHTML(usuarioNome)}</td></tr>
        ${observacao ? `<tr><td style="padding:6px 0;color:#666;vertical-align:top">Motivo:</td><td style="color:#c0392b;font-weight:600">${_escHTML(observacao)}</td></tr>` : ''}
      </table>
    </div>
    <div style="background:#f4f7fb;padding:12px 28px;font-size:11px;color:#888;text-align:center;">
      Sistema de Compras Marfim · Este é um email automático.
    </div>
    <div style="background:#1a3c5e;padding:18px 28px;">
      <div style="font-size:14px;font-weight:700;color:white;">${_escHTML(usuarioNome)} — Marfim</div>
    </div>
  </div>`;

    const textBody =
      `PEDIDO CANCELADO — ${idPedido}\n` +
      `Fornecedor: ${nomeFornecedor}\n` +
      `Filial: ${nomeFilial} (${codFilial})\n` +
      `Valor: ${valorFmt}\n` +
      `Data do Pedido: ${dataFmt}\n` +
      `Cancelado por: ${usuarioNome}\n` +
      (observacao ? `Motivo: ${observacao}\n` : '');

    MailApp.sendEmail({
      to:       emailsForn.join(','),
      cc:       ccList,
      replyTo:  'marco@marfim.ind.br',
      subject:  `Pedido Cancelado: ${idPedido} — ${nomeFornecedor}`,
      body:     textBody,
      htmlBody: htmlBody
    });
  } catch(e) {
    logErro('_enviarEmailCancelamento: ' + e.message);
  }
}

// ============================================================
// EDIÇÃO E RETIFICAÇÃO DE PEDIDOS
// ============================================================

// Correção interna: atualiza dados/itens do pedido. NÃO notifica o fornecedor.
function editarPedido(idPedido, dados, usuarioLogado, motivo, usuarioLogin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const perfil = _getPerfilReal(usuarioLogin || usuarioLogado);
    if (perfil !== 'ADMIN' && perfil !== 'COMPRAS') {
      return { ok: false, msg: 'Sem permissão para editar pedidos' };
    }

    if (!Array.isArray(dados.itens) || dados.itens.length === 0) {
      return { ok: false, msg: 'O pedido deve ter pelo menos um item' };
    }

    const sh = getSheet(ABAS.PEDIDOS);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const iId     = headers.indexOf('ID_PEDIDO');
    const iStatus = headers.indexOf('STATUS');

    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iId]).trim() === String(idPedido).trim()) { rowIdx = i; break; }
    }
    if (rowIdx < 0) return { ok: false, msg: 'Pedido não encontrado' };

    const statusAtual = String(data[rowIdx][iStatus] || '').trim().toUpperCase();
    if (statusAtual === 'CANCELADO') return { ok: false, msg: 'Pedido cancelado não pode ser editado' };

    // Bloqueia edição se já há NF lançada, para não divergir do que já foi recebido
    const nfsLancadas = sheetToArray(ABAS.RECEBIMENTOS)
      .filter(r => String(r.ID_PEDIDO).trim() === String(idPedido).trim());
    if (nfsLancadas.length > 0) {
      const ids = nfsLancadas.map(r => r.ID_RECEBIMENTO).join(', ');
      return { ok: false, msg: `Pedido possui ${nfsLancadas.length} NF(s) lançada(s) (${ids}). Exclua as NFs antes de editar.` };
    }

    const setCol = (colName, val) => {
      const idx = headers.indexOf(colName);
      if (idx >= 0) sh.getRange(rowIdx + 1, idx + 1).setValue(val);
    };
    setCol('COD_FILIAL',          dados.filialCod);
    setCol('NOME_FILIAL',         dados.filialNome);
    setCol('COD_FORNECEDOR',      dados.fornecedorCod);
    setCol('NOME_FORNECEDOR',     dados.fornecedorNome);
    setCol('FRETE',               dados.frete || 'CIF');
    setCol('COD_TRANSPORTADORA',  dados.transportadoraCod);
    setCol('NOME_TRANSPORTADORA', dados.transportadoraNome);
    setCol('PRAZO_ENTREGA',       dados.prazoEntrega);
    setCol('COND_PAGAMENTO',      dados.condPagamento || '');
    setCol('OBSERVACAO',          dados.observacao);
    setCol('VALOR_TOTAL',         dados.valorTotal);

    // Substitui os itens do pedido pelos itens corrigidos
    const shItens   = getSheet(ABAS.ITENS_PEDIDO);
    const dataItens = shItens.getDataRange().getValues();
    const hdrItens  = dataItens[0].map(h => String(h).trim());
    const colId     = hdrItens.indexOf('ID_PEDIDO');
    for (let i = dataItens.length - 1; i >= 1; i--) {
      if (String(dataItens[i][colId]).trim() === String(idPedido).trim()) shItens.deleteRow(i + 1);
    }
    dados.itens.forEach(item => {
      _appendRowMapeado(shItens, {
        ID_PEDIDO:  idPedido,
        COD_MP:     item.cod,
        DESCRICAO:  item.descricao,
        QUANTIDADE: item.quantidade,
        UNIDADE:    item.unidade,
        PRECO_UNIT: item.preco,
        SUBTOTAL:   item.subtotal
      });
    });

    _logNF('EDIÇÃO', '', idPedido, '', dados.filialCod, dados.filialNome, usuarioLogado, motivo || 'Correção interna do pedido');

    return { ok: true, msg: 'Pedido corrigido internamente. O fornecedor ainda não foi notificado.' };
  } catch(e) {
    logErro('editarPedido: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Reenvia ao fornecedor a versão atual (já corrigida) do pedido, marcada como retificação.
// Ação separada e explícita — só dispara email se o usuário chamar esta função.
function reenviarPedidoRetificado(idPedido, usuarioLogado, emailUsuario, motivo, usuarioLogin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const perfil = _getPerfilReal(usuarioLogin || usuarioLogado);
    if (perfil !== 'ADMIN' && perfil !== 'COMPRAS') {
      return { ok: false, msg: 'Sem permissão para reenviar pedidos' };
    }

    const sh = getSheet(ABAS.PEDIDOS);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const iId = headers.indexOf('ID_PEDIDO');

    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iId]).trim() === String(idPedido).trim()) { rowIdx = i; break; }
    }
    if (rowIdx < 0) return { ok: false, msg: 'Pedido não encontrado' };

    const get = col => { const idx = headers.indexOf(col); return idx >= 0 ? data[rowIdx][idx] : ''; };
    const statusAtual = String(get('STATUS') || '').trim().toUpperCase();
    if (statusAtual === 'CANCELADO') return { ok: false, msg: 'Pedido cancelado não pode ser reenviado' };

    const logPedido = sheetToArray(ABAS.LOG_NF)
      .filter(r => String(r.ID_PEDIDO).trim() === String(idPedido).trim());
    const ultimaEdicao = logPedido
      .filter(r => String(r.ACAO).trim() === 'EDIÇÃO')
      .map(r => new Date(r.DATA_HORA).getTime())
      .reduce((max, t) => Math.max(max, t), 0);
    const ultimaRetificacao = logPedido
      .filter(r => String(r.ACAO).trim() === 'RETIFICAÇÃO')
      .map(r => new Date(r.DATA_HORA).getTime())
      .reduce((max, t) => Math.max(max, t), 0);
    if (ultimaEdicao === 0 || ultimaEdicao <= ultimaRetificacao) {
      return { ok: false, msg: 'Não há correção pendente para reenviar — edite o pedido antes de retificar.' };
    }

    const fornecedorCod = String(get('COD_FORNECEDOR') || '');
    const filialCod     = String(get('COD_FILIAL') || '');
    const nomeFilial    = String(get('NOME_FILIAL') || '');

    const fornecedor = buscarCodigo('fornecedor', fornecedorCod);
    if (!fornecedor || !fornecedor.EMAIL) {
      return { ok: false, msg: 'Fornecedor sem email cadastrado — não é possível reenviar' };
    }
    const emailsList = String(fornecedor.EMAIL).split(';').map(e => e.trim())
      .filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
    if (emailsList.length === 0) {
      return { ok: false, msg: 'Email do fornecedor inválido — não é possível reenviar' };
    }

    const filial = buscarCodigo('filial', filialCod) || {};
    const emailsFilial = String(filial.EMAIL_RESPONSAVEL || '')
      .split(';').map(e => e.trim()).filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
    const ccList = [...new Set([
      ...emailsFilial,
      'marco@marfim.ind.br',
      ...(emailUsuario ? [emailUsuario] : [])
    ])].join(',');

    const itens = sheetToArray(ABAS.ITENS_PEDIDO)
      .filter(i => String(i.ID_PEDIDO).trim() === String(idPedido).trim())
      .map(i => ({
        cod: i.COD_MP, descricao: i.DESCRICAO, quantidade: i.QUANTIDADE,
        unidade: i.UNIDADE, preco: i.PRECO_UNIT, subtotal: i.SUBTOTAL
      }));
    if (itens.length === 0) return { ok: false, msg: 'Pedido sem itens — não é possível reenviar' };

    const dataPedidoRaw = get('DATA');
    const dataPedido = dataPedidoRaw ? new Date(dataPedidoRaw) : new Date();

    const dadosEmail = {
      filialNome:         nomeFilial,
      filialCNPJ:         filial.CNPJ || '',
      filialEndereco:     [filial.ENDERECO, filial.BAIRRO, filial.CIDADE, filial.ESTADO].filter(Boolean).join(', '),
      fornecedorNome:     get('NOME_FORNECEDOR'),
      fornecedorCNPJ:     fornecedor.CNPJ || '',
      fornecedorEndereco: [fornecedor.ENDERECO, fornecedor.BAIRRO, fornecedor.CIDADE, fornecedor.ESTADO].filter(Boolean).join(', '),
      frete:              get('FRETE'),
      transportadoraNome: get('NOME_TRANSPORTADORA'),
      prazoEntrega:       get('PRAZO_ENTREGA'),
      condPagamento:      get('COND_PAGAMENTO'),
      observacao:         get('OBSERVACAO'),
      usuarioLogado:      usuarioLogado,
      nomeRemetente:      usuarioLogado,
      valorTotal:         get('VALOR_TOTAL'),
      itens
    };

    const retif = { motivo: motivo || '' };
    const htmlEmail  = montarEmailHTML(idPedido, dataPedido, dadosEmail, retif);
    const textoEmail = montarEmailTexto(idPedido, dataPedido, dadosEmail, retif);

    MailApp.sendEmail({
      to:       emailsList.join(','),
      cc:       ccList,
      replyTo:  'marco@marfim.ind.br',
      subject:  `Pedido de Compra RETIFICADO ${idPedido} — ${nomeFilial}`,
      body:     textoEmail,
      htmlBody: htmlEmail
    });

    _logNF('RETIFICAÇÃO', '', idPedido, '', filialCod, nomeFilial, usuarioLogado, motivo || '');

    return { ok: true, msg: 'Pedido retificado reenviado ao fornecedor com sucesso' };
  } catch(e) {
    logErro('reenviarPedidoRetificado: ' + e.message);
    return { ok: false, msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// HISTÓRICO
// ============================================================
function _enriquecerPedidosComNF(pedidos) {
  try {
    const recs  = sheetToArray(ABAS.RECEBIMENTOS);
    const itRec = sheetToArray(ABAS.ITENS_RECEBIMENTO);
    return pedidos.map(p => {
      const idP   = String(p.ID_PEDIDO).trim();
      const recPed = recs.filter(r => String(r.ID_PEDIDO).trim() === idP);
      const temNF  = recPed.length > 0;
      // Check divergence from header flags OR from item-level comparison (handles corrupted columns)
      let temDivQtd   = recPed.some(r => String(r.DIVERGENCIA_QTD).trim()   === 'SIM');
      let temDivPreco = recPed.some(r => String(r.DIVERGENCIA_PRECO).trim() === 'SIM');
      const temDivPagto = recPed.some(r => String(r.DIVERGENCIA_PAGTO).trim() === 'SIM');
      if (!temDivQtd || !temDivPreco) {
        const recIds = new Set(recPed.map(r => String(r.ID_RECEBIMENTO).trim()));
        const itensRec = itRec.filter(i => recIds.has(String(i.ID_RECEBIMENTO).trim()));
        if (!temDivQtd)   temDivQtd   = itensRec.some(i => Math.abs(parseFloat(i.QTD_RECEBIDA||0) - parseFloat(i.QTD_PEDIDA||0)) > 0.001);
        if (!temDivPreco) temDivPreco = itensRec.some(i => Math.abs(parseFloat(i.PRECO_RECEBIDO||0) - parseFloat(i.PRECO_PEDIDO||0)) > 0.001);
      }
      return { ...p, temNF, temDivQtd, temDivPreco, temDivPagto };
    });
  } catch(e) {
    logErro('_enriquecerPedidosComNF: ' + e.message);
    return pedidos;
  }
}

function getHistorico(tipo, cod, usuarioLogin) {
  try {
    // Valida que usuário existe antes de retornar dados
    if (!_getPerfilReal(usuarioLogin)) throw new Error('Autenticação inválida');
    const pedidos = sheetToArray(ABAS.PEDIDOS);
    const colMapa = { filial: 'COD_FILIAL', fornecedor: 'COD_FORNECEDOR' };
    if (!colMapa[tipo]) return [];
    const filtrados = pedidos.filter(p => String(p[colMapa[tipo]]).trim() === String(cod).trim());
    const enriquecidos = _enriquecerPedidosComNF(filtrados);
    const todosItens = sheetToArray(ABAS.ITENS_PEDIDO);
    const resultado = enriquecidos.map(p => {
      const itens = todosItens.filter(i => String(i.ID_PEDIDO).trim() === String(p.ID_PEDIDO).trim());
      return { ...p, itens };
    });
    return resultado.sort((a, b) => new Date(b.DATA || 0) - new Date(a.DATA || 0));
  } catch(e) {
    logErro('getHistorico: ' + e.message);
    throw e;
  }
}

function getTodosPedidos(usuarioLogin) {
  try {
    const sh = getSheet(ABAS.PEDIDOS);
    if (!sh) throw new Error('Aba PEDIDOS não encontrada — execute setupPlanilha() no Apps Script Editor.');
    if (!_getPerfilReal(usuarioLogin)) throw new Error('Autenticação inválida');
    const pedidos = _enriquecerPedidosComNF(sheetToArray(ABAS.PEDIDOS));
    const todosItens = sheetToArray(ABAS.ITENS_PEDIDO);
    const resultado = pedidos.map(p => {
      const itens = todosItens.filter(i => String(i.ID_PEDIDO).trim() === String(p.ID_PEDIDO).trim());
      return { ...p, itens };
    });
    return resultado.sort((a, b) => new Date(b.DATA || 0) - new Date(a.DATA || 0));
  } catch(e) {
    logErro('getTodosPedidos: ' + e.message);
    throw e;
  }
}

// ============================================================
// SETUP — rode UMA VEZ para criar as abas
// ============================================================
function setupPlanilha() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const estrutura = {
    USUARIOS:         ['COD','NOME','USUARIO','SENHA','EMAIL','PERFIL','FILIAIS_LIBERADAS'],
    FORNECEDORES:     ['COD','NOME','CNPJ','EMAIL','CONTATO','COND_PAGAMENTO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'],
    MATERIAS_PRIMAS:  ['COD','DESCRICAO','UNIDADE','CATEGORIA'],
    TRANSPORTADORAS:  ['COD','NOME','CNPJ','CONTATO','PRAZO','OBSERVACAO','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO'],
    FILIAIS:          ['COD','NOME','CNPJ','CEP','BAIRRO','ENDERECO','CIDADE','ESTADO','EMAIL_RESPONSAVEL','COD_TRANSPORTADORA'],
    PEDIDOS:          ['ID_PEDIDO','DATA','COD_FILIAL','NOME_FILIAL','COD_FORNECEDOR','NOME_FORNECEDOR','FRETE','COD_TRANSPORTADORA','NOME_TRANSPORTADORA','PRAZO_ENTREGA','COND_PAGAMENTO','OBSERVACAO','USUARIO','VALOR_TOTAL','STATUS'],
    ITENS_PEDIDO:     ['ID_PEDIDO','COD_MP','DESCRICAO','QUANTIDADE','UNIDADE','PRECO_UNIT','SUBTOTAL'],
    PRECO_FORNECEDOR:   ['COD_FORNECEDOR','COD_MP','PRECO'],
    TRANSP_FORN_FILIAL: ['COD_FORNECEDOR','COD_FILIAL','COD_TRANSPORTADORA'],
    RECEBIMENTOS:      ['ID_RECEBIMENTO','ID_PEDIDO','NF_NUMERO','DATA_NF','DATA_RECEBIMENTO','COD_FILIAL','NOME_FILIAL','COD_FORNECEDOR','NOME_FORNECEDOR','USUARIO','VALOR_TOTAL_PEDIDO','VALOR_TOTAL_RECEBIDO','DIVERGENCIA_QTD','DIVERGENCIA_PRECO','DIVERGENCIA_PAGTO','PAGTO_PRAZO','PAGTO_ESPERADO','PAGTO_NF','OBSERVACAO'],
    ITENS_RECEBIMENTO: ['ID_RECEBIMENTO','ID_PEDIDO','COD_MP','DESCRICAO','QTD_PEDIDA','QTD_RECEBIDA','PRECO_PEDIDO','PRECO_RECEBIDO','SUBTOTAL_PEDIDO','SUBTOTAL_RECEBIDO','DIV_QTD','DIV_PRECO'],
    LOG_ERROS:         ['DATA','MENSAGEM'],
    LOG_NF:            ['DATA_HORA','ACAO','ID_RECEBIMENTO','ID_PEDIDO','NF_NUMERO','COD_FILIAL','NOME_FILIAL','USUARIO','DETALHE']
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
      const existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
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
    shUser.appendRow(['USR001','Administrador','admin',_encriptarSenha('admin123'),'admin@empresa.com','ADMIN']);
  }

  Logger.log('Setup concluído.');
}

// ============================================================
// DIAGNÓSTICO — rode no editor do Apps Script e veja o log
// Mostra a ordem real das colunas de cada aba chave.
// ============================================================
function diagnosticarPlanilha() {
  const abas = [ABAS.PEDIDOS, ABAS.RECEBIMENTOS, ABAS.ITENS_RECEBIMENTO, ABAS.ITENS_PEDIDO];
  abas.forEach(nomeAba => {
    const sh = getSheet(nomeAba);
    if (!sh) { Logger.log(nomeAba + ': ABA NÃO ENCONTRADA'); return; }
    const data = sh.getDataRange().getValues();
    const headers = data[0].map((h, i) => i + ':' + String(h).trim());
    Logger.log('=== ' + nomeAba + ' ===');
    Logger.log('Colunas: ' + headers.join(' | '));
    data.slice(1, 4).forEach((row, i) => {
      const linha = row.map((v, j) => data[0][j] + '=' + (v instanceof Date ? v.toISOString() : v)).join(' | ');
      Logger.log('Linha ' + (i + 2) + ': ' + linha);
    });
  });
}

// ============================================================
// MIGRAÇÃO — rode uma única vez no editor do Apps Script
// Corrige linhas que foram gravadas com colunas trocadas.
// Rode primeiro diagnosticarPlanilha() e confirme o resultado.
// ============================================================
function migrarPedidosCorretos() {
  const sh = getSheet(ABAS.PEDIDOS);
  if (!sh) { Logger.log('Aba PEDIDOS não encontrada.'); return; }
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const idx = name => headers.indexOf(name);
  const iStatus = idx('STATUS'), iValor = idx('VALOR_TOTAL');
  const iUsuario = idx('USUARIO'), iNomeTrans = idx('NOME_TRANSPORTADORA');
  const iPrazo = idx('PRAZO_ENTREGA'), iObs = idx('OBSERVACAO');
  const iCond = idx('COND_PAGAMENTO');

  if (iStatus < 0 || iValor < 0) { Logger.log('Colunas essenciais não encontradas.'); return; }

  const statusValidos = ['ENVIADO','RECEBIDO','PENDENTE','CANCELADO','EM ANÁLISE'];

  let corrigidas = 0;
  for (let r = 1; r < data.length; r++) {
    const row    = data[r];
    const status = String(row[iStatus]).trim();
    const valor  = String(row[iValor]).trim();
    const statusErrado = !statusValidos.includes(status.toUpperCase());
    const valorErrado  = isNaN(parseFloat(valor));

    if (!statusErrado && !valorErrado) continue; // linha OK

    Logger.log('Linha ' + (r+1) + ' (' + row[0] + ') parece corrompida. Valores atuais:');
    headers.forEach((h, j) => Logger.log('  ' + h + ' [col ' + (j+1) + ']: ' + row[j]));

    // Encontra onde "ENVIADO"/"RECEBIDO" foi parar (esse é o STATUS real)
    const posStatus = row.findIndex((v, j) => statusValidos.includes(String(v).trim().toUpperCase()));
    // Encontra onde um número positivo > 0 foi parar que possa ser o total
    const posValor  = row.findIndex((v, j) => !isNaN(parseFloat(v)) && parseFloat(v) > 0
                                               && j !== headers.indexOf('DATA') && j > 5);
    if (posStatus >= 0 && posValor >= 0 && posStatus !== iStatus) {
      sh.getRange(r+1, iStatus+1).setValue(row[posStatus]);
      sh.getRange(r+1, posStatus+1).setValue(row[iStatus]);
      Logger.log('  → STATUS movido da col ' + (posStatus+1) + ' para col ' + (iStatus+1));
    }
    if (posValor >= 0 && posValor !== iValor) {
      const valorReal = parseFloat(row[posValor]);
      sh.getRange(r+1, iValor+1).setValue(valorReal);
      sh.getRange(r+1, posValor+1).setValue(row[iValor]);
      Logger.log('  → VALOR_TOTAL movido da col ' + (posValor+1) + ' para col ' + (iValor+1));
    }
    corrigidas++;
  }
  Logger.log('Migração concluída. ' + corrigidas + ' linhas processadas.');
}

// ============================================================
// MIGRAÇÃO DE SENHAS — rode UMA VEZ no editor do Apps Script
// Converte todas as senhas em texto puro para hash+salt.
// Após rodar, qualquer senha que não seja "hash|salt" vira hash.
// ============================================================
function migrarSenhas() {
  const sh = getSheet(ABAS.USUARIOS);
  if (!sh) { Logger.log('Aba USUARIOS não encontrada.'); return; }
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const colSenha = headers.indexOf('SENHA');
  if (colSenha < 0) { Logger.log('Coluna SENHA não encontrada.'); return; }

  let migradas = 0;
  for (let i = 1; i < data.length; i++) {
    const senhaAtual = String(data[i][colSenha]).trim();
    if (!senhaAtual || senhaAtual.includes('|')) continue; // vazia ou já hasheada
    sh.getRange(i + 1, colSenha + 1).setValue(_encriptarSenha(senhaAtual));
    Logger.log('Usuário linha ' + (i + 1) + ': senha migrada.');
    migradas++;
  }
  Logger.log('migrarSenhas concluída. ' + migradas + ' senha(s) convertida(s).');
}
