// Código.gs - Versão 3.0: Reformulação Completa com Listagem de Todos os Pacientes + Tema Dark e Profissional

// CONFIGURAÇÕES
const PLANILHA_ID = '1dK85kPoRzeoWtCQh0dGwTWWDALN67_uvhRDg7ofwmSQ';
const ADMIN_EMAIL = 'lukyam.lmm@isgh.org.br';

const SHEET_NAMES = Object.freeze({
  LOGIN: 'LOGIN',
  CADASTRO: 'CADASTRO',
  BASE: 'BASE'
});

const STATUS = Object.freeze({
  ATIVO: 'Ativo',
  INATIVO: 'Inativo'
});

const DEFAULT_SETORES = Object.freeze([
  { id: 1, descricao: 'Acesso Total', setor: 'Administração', nivelAcesso: 0 },
  { id: 2, descricao: 'Acesso Técnico', setor: 'Enfermagem', nivelAcesso: 1 },
  { id: 3, descricao: 'Acesso Médico', setor: 'Médico', nivelAcesso: 2 }
]);

let cachedSpreadsheet;

// Função principal
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Sistema de Saúde ISGH - Profissional');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPlanilha() {
  try {
    if (!cachedSpreadsheet) {
      cachedSpreadsheet = SpreadsheetApp.openById(PLANILHA_ID);
    }
    return cachedSpreadsheet;
  } catch (error) {
    throw new Error('Erro ao acessar planilha: ' + error.message);
  }
}

function getSheet(sheetName) {
  const ss = getPlanilha();
  return ss.getSheetByName(sheetName);
}

function withSheet(sheetName, onSuccess, onMissing) {
  const sheet = getSheet(sheetName);
  if (!sheet) {
    return typeof onMissing === 'function' ? onMissing() : onMissing;
  }
  return onSuccess(sheet);
}

function setHeaders(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatarCabecalho(sheet);
}

function safeNumber(value) {
  const parsed = parseFloat(value);
  return Number.isFinite(parsed) ? parsed : '';
}

function asDate(value) {
  return value instanceof Date ? value : new Date(value);
}

function gerarSenhaTemporaria() {
  return Math.random().toString(36).slice(2, 10).toUpperCase();
}

// Função para calcular hash MD5
function calcularHashMD5(senha) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, senha)
    .map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
    .join('');
}

// Criar estrutura inicial
function criarEstruturaInicial() {
  const ss = getPlanilha();
  const abas = [SHEET_NAMES.LOGIN, SHEET_NAMES.CADASTRO, SHEET_NAMES.BASE];
  const resultado = [];

  abas.forEach(aba => {
    let sheet = ss.getSheetByName(aba);
    if (!sheet) {
      sheet = ss.insertSheet(aba);
      resultado.push(`✓ Aba ${aba} criada`);
    }

    if (aba === SHEET_NAMES.LOGIN) {
      const headers = ['Nome', 'Matricula', 'Setor', 'SenhaHash', 'DataCriacao', 'Status', 'UltimaAlteracao'];
      setHeaders(sheet, headers);
      if (sheet.getLastRow() === 1) {
        const hashAdmin = calcularHashMD5('admin');
        sheet.getRange(2, 1, 1, headers.length)
          .setValues([[
            'Administrador',
            'admin',
            'Administração',
            hashAdmin,
            new Date(),
            STATUS.ATIVO,
            new Date()
          ]]);
        resultado.push('✓ Admin criado (admin/admin)');
      }
    } else if (aba === SHEET_NAMES.CADASTRO) {
      const headers = ['ID', 'Descricao', 'Setor', 'NivelAcesso', 'DataCriacao'];
      setHeaders(sheet, headers);
      if (sheet.getLastRow() === 1) {
        const agora = new Date();
        const linhas = DEFAULT_SETORES.map(setor => [
          setor.id,
          setor.descricao,
          setor.setor,
          setor.nivelAcesso,
          agora
        ]);
        sheet.getRange(2, 1, linhas.length, linhas[0].length).setValues(linhas);
        resultado.push('✓ Setores padrão criados');
      }
    } else if (aba === SHEET_NAMES.BASE) {
      const headers = [
        'Nome',
        'Prontuario',
        'DataNascimento',
        'Peso',
        'Altura',
        'PressaoArterial',
        'Temperatura',
        'Saturacao',
        'Glicemia',
        'DataRegistro',
        'UsuarioRegistro',
        'Observacoes'
      ];
      setHeaders(sheet, headers);
      resultado.push('✓ Aba BASE configurada');
    }
  });

  return resultado.join('\n');
}

function formatarCabecalho(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range
    .setBackground('#2563eb')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

// LOGIN
function fazerLogin(matricula, senha) {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        const hash = calcularHashMD5(senha);
        const usuario = rows.find(row => row[1] === matricula && row[3] === hash && row[5] === STATUS.ATIVO);

        if (!usuario) {
          return { success: false, message: 'Credenciais inválidas ou usuário inativo' };
        }

        const tipo = determinarTipoUsuario(usuario[2]);
        return {
          success: true,
          user: {
            nome: usuario[0],
            matricula: usuario[1],
            setor: usuario[2],
            tipo
          }
        };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro interno: ' + error.message };
  }
}

function determinarTipoUsuario(setor) {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        const encontrado = rows.find(row => row[2] === setor);
        if (!encontrado) return 'TECNICO';

        const nivel = encontrado[3];
        if (nivel === 0) return 'ADM';
        if (nivel === 1) return 'TECNICO';
        if (nivel === 2) return 'MEDICO';
        return 'TECNICO';
      },
      'TECNICO'
    );
  } catch (error) {
    return 'TECNICO';
  }
}

// Recuperação de senha
function recuperarSenha(matricula) {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => {
        const data = sheet.getDataRange().getValues();
        const index = data.findIndex((row, i) => i > 0 && row[1] === matricula);
        if (index === -1) {
          return { success: false, message: 'Matrícula não encontrada' };
        }

        const novaSenha = gerarSenhaTemporaria();
        const hash = calcularHashMD5(novaSenha);
        const rowIndex = index + 1;

        sheet.getRange(rowIndex, 4).setValue(hash);
        sheet.getRange(rowIndex, 7).setValue(new Date());

        const usuario = data[index][0];
        const emailUsuario = usuario.toLowerCase().replace(/\s+/g, '.') + '@isgh.org.br';
        const corpo = `Recuperação de Senha - ISGH\n\nUsuário: ${usuario}\nNova Senha: ${novaSenha}\nData: ${new Date().toLocaleString('pt-BR')}`;

        MailApp.sendEmail(emailUsuario, 'Nova Senha - Sistema ISGH', corpo);
        MailApp.sendEmail(ADMIN_EMAIL, 'Recuperação Solicitada', `Usuário ${matricula} solicitou recuperação.`);

        return { success: true, message: `Senha resetada para ${usuario}. Nova senha enviada por email.` };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Cadastrar usuário
function cadastrarUsuario(nome, matricula, setor, senha) {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        const existe = rows.some(row => row[1] === matricula);
        if (existe) {
          return { success: false, message: 'Matrícula já existe' };
        }

        const hash = calcularHashMD5(senha);
        sheet.appendRow([nome, matricula, setor, hash, new Date(), STATUS.ATIVO, new Date()]);
        return { success: true, message: 'Usuário cadastrado com sucesso' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Obter setores
function obterSetores() {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => sheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map(row => row[2])
        .filter(Boolean),
      DEFAULT_SETORES.map(setor => setor.setor)
    );
  } catch (error) {
    return DEFAULT_SETORES.map(setor => setor.setor);
  }
}

// Obter setores com detalhes
function obterSetoresComDetalhes() {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => sheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map(row => ({
          id: row[0],
          descricao: row[1],
          setor: row[2],
          nivelAcesso: row[3]
        })),
      []
    );
  } catch (error) {
    return [];
  }
}

// Adicionar setor
function adicionarSetor(novoSetor, nivelAcesso) {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        const existe = rows.some(row => row[2] === novoSetor);
        if (existe) {
          return { success: false, message: 'Setor já cadastrado' };
        }

        const descricao = nivelAcesso === 0
          ? 'Acesso Total'
          : nivelAcesso === 1
            ? 'Acesso Técnico'
            : 'Acesso Médico';

        const novoId = rows.reduce((max, row) => {
          const id = parseInt(row[0], 10);
          return Number.isNaN(id) ? max : Math.max(max, id);
        }, 0) + 1;
        sheet.appendRow([novoId, descricao, novoSetor, parseInt(nivelAcesso, 10), new Date()]);
        return { success: true, message: 'Setor adicionado com sucesso' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Deletar setor
function deletarSetor(id) {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => {
        const rowIndex = parseInt(id, 10) + 1;
        if (Number.isNaN(rowIndex) || rowIndex <= 1 || rowIndex > sheet.getLastRow()) {
          return { success: false, message: 'Setor inválido' };
        }
        sheet.deleteRow(rowIndex);
        return { success: true, message: 'Setor deletado' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Editar setor
function editarSetor(id, novosDados) {
  try {
    return withSheet(
      SHEET_NAMES.CADASTRO,
      sheet => {
        const row = parseInt(id, 10) + 1;
        if (Number.isNaN(row) || row <= 1 || row > sheet.getLastRow()) {
          return { success: false, message: 'Setor inválido' };
        }

        const { descricao, setor, nivelAcesso } = novosDados;
        if (descricao) sheet.getRange(row, 2).setValue(descricao);
        if (setor) sheet.getRange(row, 3).setValue(setor);
        if (nivelAcesso !== undefined) sheet.getRange(row, 4).setValue(nivelAcesso);

        return { success: true, message: 'Setor atualizado' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Registrar paciente
function registrarDadosPaciente(dados) {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const usuario = Session.getActiveUser().getEmail().split('@')[0] || 'Sistema';
        sheet.appendRow([
          dados.nomeCompleto,
          dados.prontuario,
          dados.dataNascimento,
          safeNumber(dados.peso),
          safeNumber(dados.altura),
          dados.pressaoArterial,
          safeNumber(dados.temperatura),
          safeNumber(dados.saturacao),
          safeNumber(dados.glicemia),
          new Date(),
          usuario,
          dados.observacoes || ''
        ]);

        return { success: true, message: 'Registro salvo com sucesso' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

function mapPaciente(row) {
  return {
    nome: row[0] || '',
    prontuario: row[1] || '',
    dataNascimento: row[2] || '',
    peso: row[3] || '',
    altura: row[4] || '',
    pressaoArterial: row[5] || '',
    temperatura: row[6] || '',
    saturacao: row[7] || '',
    glicemia: row[8] || '',
    dataRegistro: row[9] || '',
    usuarioRegistro: row[10] || '',
    observacoes: row[11] || ''
  };
}

// NOVA FUNÇÃO: Obter todos os pacientes (todos os registros)
function getAllPacientes() {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        return rows
          .map(mapPaciente)
          .sort((a, b) => {
            const dataA = a.dataRegistro ? asDate(a.dataRegistro) : null;
            const dataB = b.dataRegistro ? asDate(b.dataRegistro) : null;
            if (!dataA && !dataB) return 0;
            if (!dataA) return 1;
            if (!dataB) return -1;
            return dataB - dataA;
          });
      },
      []
    );
  } catch (error) {
    console.error('Erro ao obter todos os pacientes:', error);
    return [];
  }
}

function buscarPaciente(termo) {
  if (!termo) return [];
  const termoNormalizado = termo.toString().trim().toLowerCase();
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        return rows
          .filter(row => {
            const nome = (row[0] || '').toString().toLowerCase();
            const prontuario = (row[1] || '').toString().toLowerCase();
            return nome.includes(termoNormalizado) || prontuario.includes(termoNormalizado);
          })
          .map(mapPaciente);
      },
      []
    );
  } catch (error) {
    return [];
  }
}

// Obter usuários
function obterUsuarios() {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => sheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map(row => ({
          nome: row[0],
          matricula: row[1],
          setor: row[2],
          ultimaAlteracao: row[6] ? new Date(row[6]) : 'N/A',
          status: row[5]
        })),
      []
    );
  } catch (error) {
    return [];
  }
}

// Deletar usuário
function deletarUsuario(matricula) {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => {
        const data = sheet.getDataRange().getValues();
        const index = data.findIndex((row, i) => i > 0 && row[1] === matricula);
        if (index === -1) {
          return { success: false, message: 'Usuário não encontrado' };
        }

        const rowIndex = index + 1;
        sheet.getRange(rowIndex, 6).setValue(STATUS.INATIVO);
        sheet.getRange(rowIndex, 7).setValue(new Date());
        return { success: true, message: 'Usuário inativado com sucesso' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Editar usuário
function editarUsuario(matricula, novosDados) {
  try {
    return withSheet(
      SHEET_NAMES.LOGIN,
      sheet => {
        const data = sheet.getDataRange().getValues();
        const index = data.findIndex((row, i) => i > 0 && row[1] === matricula);
        if (index === -1) {
          return { success: false, message: 'Usuário não encontrado' };
        }

        const rowIndex = index + 1;
        const registroAtual = data[index];
        sheet.getRange(rowIndex, 1).setValue(novosDados.nome || registroAtual[0]);
        sheet.getRange(rowIndex, 2).setValue(novosDados.matricula || registroAtual[1]);
        sheet.getRange(rowIndex, 3).setValue(novosDados.setor || registroAtual[2]);
        if (novosDados.senha) {
          sheet.getRange(rowIndex, 4).setValue(calcularHashMD5(novosDados.senha));
        }
        sheet.getRange(rowIndex, 7).setValue(new Date());
        return { success: true, message: 'Usuário atualizado' };
      },
      { success: false, message: 'Sistema em configuração' }
    );
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Estatísticas
function obterEstatisticas() {
  try {
    const ss = getPlanilha();
    const loginSheet = getSheet(SHEET_NAMES.LOGIN);
    const baseSheet = getSheet(SHEET_NAMES.BASE);

    const totalUsuarios = loginSheet ? Math.max(loginSheet.getLastRow() - 1, 0) : 0;
    const totalRegistros = baseSheet ? Math.max(baseSheet.getLastRow() - 1, 0) : 0;

    if (!baseSheet) {
      return { totalUsuarios, totalRegistros, registrosHoje: 0, mediaTemperatura: 0 };
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);

    const [, ...rows] = baseSheet.getDataRange().getValues();
    const resumoHoje = rows.reduce(
      (acc, row) => {
        const dataReg = row[9] ? asDate(row[9]) : null;
        if (dataReg && dataReg >= hoje && dataReg < amanha) {
          acc.registrosHoje += 1;
          if (row[6]) acc.somaTemperaturas += parseFloat(row[6]);
        }
        return acc;
      },
      { registrosHoje: 0, somaTemperaturas: 0 }
    );

    const mediaTemperatura = resumoHoje.registrosHoje > 0
      ? resumoHoje.somaTemperaturas / resumoHoje.registrosHoje
      : 0;

    return {
      totalUsuarios,
      totalRegistros,
      registrosHoje: resumoHoje.registrosHoje,
      mediaTemperatura
    };
  } catch (error) {
    return { totalUsuarios: 0, totalRegistros: 0, registrosHoje: 0, mediaTemperatura: 0 };
  }
}

// Registros de hoje
function obterRegistrosHoje() {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);

        const [, ...rows] = sheet.getDataRange().getValues();
        return rows
          .filter(row => {
            const dataReg = row[9] ? asDate(row[9]) : null;
            return dataReg && dataReg >= hoje && dataReg < amanha;
          })
          .slice(0, 5)
          .map(row => ({
            nome: row[0],
            prontuario: row[1],
            dataRegistro: row[9],
            temperatura: row[6]
          }));
      },
      []
    );
  } catch (error) {
    return [];
  }
}

// Gerar relatório
function gerarRelatorio(dataInicio, dataFim) {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const inicio = new Date(`${dataInicio}T00:00:00`);
        const fim = new Date(`${dataFim}T23:59:59`);

        const [, ...rows] = sheet.getDataRange().getValues();
        const resumo = rows.reduce(
          (acc, row) => {
            const dataReg = row[9] ? asDate(row[9]) : null;
            if (dataReg && dataReg >= inicio && dataReg <= fim) {
              acc.total += 1;
              acc.usuarios.add(row[10]);
              if (row[6]) acc.somaTemperaturas += parseFloat(row[6]);
            }
            return acc;
          },
          { total: 0, usuarios: new Set(), somaTemperaturas: 0 }
        );

        return {
          total: resumo.total,
          usuariosUnicos: resumo.usuarios.size,
          mediaTemp: resumo.total > 0 ? resumo.somaTemperaturas / resumo.total : 0,
          data: `${dataInicio} a ${dataFim}`
        };
      },
      { total: 0, usuariosUnicos: 0, mediaTemp: 0, data: '' }
    );
  } catch (error) {
    return { total: 0, usuariosUnicos: 0, mediaTemp: 0, data: '' };
  }
}

// FUNÇÕES DE DEBUG
function debugProntuarios() {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        const [, ...rows] = sheet.getDataRange().getValues();
        const detalhes = rows
          .map((row, index) => {
            const linha = index + 2;
            return `Linha ${linha}: "${row[0]}" → Prontuário: "${row[1]}"`;
          })
          .join('\n');

        return [
          '=== DEBUG PRONTUÁRIOS ===',
          '',
          `Total de registros: ${rows.length}`,
          '',
          detalhes
        ].join('\n');
      },
      '❌ Aba BASE não encontrada'
    );
  } catch (error) {
    return '❌ Erro no debug: ' + error.message;
  }
}

function criarDadosTeste() {
  try {
    return withSheet(
      SHEET_NAMES.BASE,
      sheet => {
        if (sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
        }

        const agora = new Date();
        const dadosTeste = [
          ['João Silva', '12345', '1990-05-15', 70.5, 175, '120/80', 36.5, 98, 95, agora, 'Admin', 'Paciente teste 1'],
          ['Maria Santos', '67890', '1985-08-20', 65.2, 165, '110/70', 36.8, 99, 100, agora, 'Admin', 'Paciente teste 2'],
          ['Pedro Oliveira', '11111', '1978-12-10', 80.0, 180, '130/85', 37.1, 97, 105, agora, 'Admin', 'Paciente teste 3']
        ];

        sheet.getRange(2, 1, dadosTeste.length, dadosTeste[0].length).setValues(dadosTeste);
        return '✅ Dados de teste criados! Prontuários: 12345, 67890, 11111';
      },
      '❌ Aba BASE não encontrada'
    );
  } catch (error) {
    return '❌ Erro ao criar dados teste: ' + error.message;
  }
}

// TESTE DIRETO - Busca específica
function testeBuscaDireta() {
  console.log('=== TESTE DIRETO DA BUSCA ===');

  const resultado = buscarPaciente('1');

  console.log('Resultado do teste direto:', resultado);
  console.log('Tipo:', typeof resultado);
  console.log('É array?', Array.isArray(resultado));
  console.log('É null?', resultado === null);
  console.log('Quantidade:', resultado ? resultado.length : 'null');

  if (resultado && resultado.length > 0) {
    resultado.forEach((reg, idx) => {
      console.log(`Registro ${idx + 1}:`, reg.nome, '-', reg.prontuario);
    });
  }

  return `Teste concluído. Resultados: ${resultado ? resultado.length : 'null'}`;
}
