const PLANILHA_ID = '1dK85kPoRzeoWtCQh0dGwTWWDALN67_uvhRDg7ofwmSQ';

const SHEETS = Object.freeze({
  LOGIN: 'LOGIN',
  SETORES: 'SETORES',
  PACIENTES: 'PACIENTES'
});

const STATUS = Object.freeze({
  ATIVO: 'ATIVO',
  INATIVO: 'INATIVO'
});

const NIVEL_ACESSO = Object.freeze({
  ADMIN: 0,
  PROFISSIONAL: 1,
  CONSULTA: 2
});

const HEADERS = Object.freeze({
  LOGIN: ['Nome', 'Matricula', 'Setor', 'NivelAcesso', 'SenhaHash', 'Status', 'CriadoEm', 'AtualizadoEm'],
  SETORES: ['ID', 'Descricao', 'NivelAcesso', 'Ativo', 'CriadoEm', 'AtualizadoEm'],
  PACIENTES: [
    'DataRegistro',
    'HoraRegistro',
    'Profissional',
    'Setor',
    'NomePaciente',
    'Prontuario',
    'DataNascimento',
    'PressaoArterial',
    'FrequenciaCardiaca',
    'Temperatura',
    'Saturacao',
    'Glicemia',
    'QueixaPrincipal',
    'Observacoes'
  ]
});

let cachedSpreadsheet;

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Central de Triagem ISGH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  if (!cachedSpreadsheet) {
    cachedSpreadsheet = SpreadsheetApp.openById(PLANILHA_ID);
  }
  return cachedSpreadsheet;
}

function ensureSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn() || headers.length).getValues()[0];
  const needsHeader = !currentHeaders || currentHeaders.join('') === '' || currentHeaders.length !== headers.length;
  if (needsHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1d4ed8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
  return sheet;
}

function hashPassword(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return raw.map(byte => ('0' + (byte & 0xff).toString(16)).slice(-2)).join('');
}

function configurarSistema() {
  const ss = getSpreadsheet();
  const agora = new Date();

  const loginSheet = ensureSheet(SHEETS.LOGIN, HEADERS.LOGIN);
  const setoresSheet = ensureSheet(SHEETS.SETORES, HEADERS.SETORES);
  ensureSheet(SHEETS.PACIENTES, HEADERS.PACIENTES);

  if (loginSheet.getLastRow() < 2) {
    const senha = hashPassword('admin123');
    loginSheet.appendRow([
      'Administrador',
      'admin',
      'Diretoria',
      NIVEL_ACESSO.ADMIN,
      senha,
      STATUS.ATIVO,
      agora,
      agora
    ]);
  }

  if (setoresSheet.getLastRow() < 2) {
    setoresSheet.appendRow([1, 'Diretoria', NIVEL_ACESSO.ADMIN, true, agora, agora]);
    setoresSheet.appendRow([2, 'Triagem', NIVEL_ACESSO.PROFISSIONAL, true, agora, agora]);
  }

  return { success: true, message: 'Sistema configurado.' };
}

function getSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }
  return values.slice(1).filter(row => row.join('').trim() !== '');
}

function localizarUsuario(matricula) {
  if (!matricula) {
    return null;
  }
  const sheet = ensureSheet(SHEETS.LOGIN, HEADERS.LOGIN);
  const rows = getSheetData(sheet);
  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    if (String(row[1]).toLowerCase() === String(matricula).toLowerCase()) {
      return {
        rowIndex: index + 2,
        dados: {
          nome: row[0],
          matricula: row[1],
          setor: row[2],
          nivelAcesso: Number(row[3]),
          senhaHash: row[4],
          status: row[5],
          criadoEm: row[6],
          atualizadoEm: row[7]
        }
      };
    }
  }
  return null;
}

function exigirAdmin(matricula) {
  const usuario = localizarUsuario(matricula);
  if (!usuario || usuario.dados.nivelAcesso !== NIVEL_ACESSO.ADMIN || usuario.dados.status !== STATUS.ATIVO) {
    throw new Error('Acesso restrito aos administradores.');
  }
  return usuario;
}

function fazerLogin(credenciais) {
  try {
    if (!credenciais || !credenciais.matricula || !credenciais.senha) {
      return { success: false, message: 'Informe matrícula e senha.' };
    }

    configurarSistema();

    const usuario = localizarUsuario(credenciais.matricula);
    if (!usuario || usuario.dados.status !== STATUS.ATIVO) {
      return { success: false, message: 'Usuário não encontrado ou inativo.' };
    }

    const hash = hashPassword(credenciais.senha);
    if (hash !== usuario.dados.senhaHash) {
      return { success: false, message: 'Senha inválida.' };
    }

    return {
      success: true,
      user: {
        nome: usuario.dados.nome,
        matricula: usuario.dados.matricula,
        setor: usuario.dados.setor,
        nivelAcesso: usuario.dados.nivelAcesso
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function listarPacientes() {
  try {
    const sheet = ensureSheet(SHEETS.PACIENTES, HEADERS.PACIENTES);
    const rows = getSheetData(sheet);
    const pacientes = rows.map(row => ({
      dataRegistro: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy') : row[0],
      horaRegistro: row[1],
      profissional: row[2],
      setor: row[3],
      nomePaciente: row[4],
      prontuario: row[5],
      dataNascimento: row[6],
      pressaoArterial: row[7],
      frequenciaCardiaca: row[8],
      temperatura: row[9],
      saturacao: row[10],
      glicemia: row[11],
      queixaPrincipal: row[12],
      observacoes: row[13]
    }));
    return { success: true, pacientes };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function registrarPaciente(payload) {
  try {
    if (!payload || !payload.nomePaciente || !payload.queixaPrincipal) {
      return { success: false, message: 'Preencha os campos obrigatórios do paciente.' };
    }

    const sheet = ensureSheet(SHEETS.PACIENTES, HEADERS.PACIENTES);
    const agora = new Date();
    const dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const horaFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm');

    sheet.appendRow([
      dataFormatada,
      horaFormatada,
      payload.profissional || '',
      payload.setor || '',
      payload.nomePaciente,
      payload.prontuario || '',
      payload.dataNascimento || '',
      payload.pressaoArterial || '',
      payload.frequenciaCardiaca || '',
      payload.temperatura || '',
      payload.saturacao || '',
      payload.glicemia || '',
      payload.queixaPrincipal || '',
      payload.observacoes || ''
    ]);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function obterEstatisticas() {
  try {
    const resultado = listarPacientes();
    if (!resultado.success) {
      return resultado;
    }

    const pacientes = resultado.pacientes || [];
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const estatisticas = {
      success: true,
      total: pacientes.length,
      atendimentosHoje: 0,
      porSetor: {}
    };

    pacientes.forEach(paciente => {
      if (paciente.dataRegistro === hoje) {
        estatisticas.atendimentosHoje += 1;
      }
      const setor = paciente.setor || 'Não informado';
      estatisticas.porSetor[setor] = (estatisticas.porSetor[setor] || 0) + 1;
    });

    return estatisticas;
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function listarSetores() {
  try {
    const sheet = ensureSheet(SHEETS.SETORES, HEADERS.SETORES);
    const rows = getSheetData(sheet);
    const setores = rows.map(row => ({
      id: Number(row[0]),
      descricao: row[1],
      nivelAcesso: Number(row[2]),
      ativo: row[3] === true || String(row[3]).toLowerCase() === 'true',
      criadoEm: row[4],
      atualizadoEm: row[5]
    }));

    return { success: true, setores };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function salvarSetor(payload) {
  try {
    if (!payload || !payload.solicitante || !payload.setor) {
      return { success: false, message: 'Solicitação inválida.' };
    }

    exigirAdmin(payload.solicitante);

    const setor = payload.setor;
    if (!setor.descricao) {
      return { success: false, message: 'Informe a descrição do setor.' };
    }

    const sheet = ensureSheet(SHEETS.SETORES, HEADERS.SETORES);
    const rows = getSheetData(sheet);
    const agora = new Date();
    let proximoId = 1;

    rows.forEach(row => {
      if (Number(row[0]) >= proximoId) {
        proximoId = Number(row[0]) + 1;
      }
    });

    if (setor.id) {
      const index = rows.findIndex(row => Number(row[0]) === Number(setor.id));
      if (index === -1) {
        return { success: false, message: 'Setor não encontrado.' };
      }
      const rowNumber = index + 2;
      sheet.getRange(rowNumber, 2, 1, 5).setValues([[
        setor.descricao,
        Number(setor.nivelAcesso),
        Boolean(setor.ativo),
        rows[index][4] || agora,
        agora
      ]]);
    } else {
      sheet.appendRow([
        proximoId,
        setor.descricao,
        Number(setor.nivelAcesso),
        Boolean(setor.ativo),
        agora,
        agora
      ]);
    }

    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function listarUsuarios(matriculaSolicitante) {
  try {
    exigirAdmin(matriculaSolicitante);
    const sheet = ensureSheet(SHEETS.LOGIN, HEADERS.LOGIN);
    const rows = getSheetData(sheet);
    const usuarios = rows.map(row => ({
      nome: row[0],
      matricula: row[1],
      setor: row[2],
      nivelAcesso: Number(row[3]),
      status: row[5],
      criadoEm: row[6],
      atualizadoEm: row[7]
    }));

    return { success: true, usuarios };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function salvarUsuario(payload) {
  try {
    if (!payload || !payload.solicitante || !payload.usuario) {
      return { success: false, message: 'Solicitação inválida.' };
    }

    exigirAdmin(payload.solicitante);

    const usuario = payload.usuario;
    if (!usuario.nome || !usuario.matricula || !usuario.setor) {
      return { success: false, message: 'Preencha os dados obrigatórios do usuário.' };
    }

    const sheet = ensureSheet(SHEETS.LOGIN, HEADERS.LOGIN);
    const rows = getSheetData(sheet);
    const agora = new Date();
    const matricula = String(usuario.matricula).trim();
    const existente = rows.findIndex(row => String(row[1]).toLowerCase() === matricula.toLowerCase());

    if (existente >= 0) {
      const rowNumber = existente + 2;
      const values = sheet.getRange(rowNumber, 1, 1, HEADERS.LOGIN.length).getValues()[0];
      values[0] = usuario.nome;
      values[2] = usuario.setor;
      values[3] = Number(usuario.nivelAcesso);
      values[5] = usuario.status || STATUS.ATIVO;
      values[7] = agora;

      if (usuario.senha) {
        values[4] = hashPassword(usuario.senha);
      }

      sheet.getRange(rowNumber, 1, 1, HEADERS.LOGIN.length).setValues([values]);
    } else {
      if (!usuario.senha) {
        return { success: false, message: 'Defina uma senha provisória.' };
      }
      sheet.appendRow([
        usuario.nome,
        matricula,
        usuario.setor,
        Number(usuario.nivelAcesso),
        hashPassword(usuario.senha),
        usuario.status || STATUS.ATIVO,
        agora,
        agora
      ]);
    }

    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
