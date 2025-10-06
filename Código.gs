// Código.gs - Versão 3.0: Reformulação Completa com Listagem de Todos os Pacientes + Tema Dark e Profissional

// CONFIGURAÇÕES
const PLANILHA_ID = '1dK85kPoRzeoWtCQh0dGwTWWDALN67_uvhRDg7ofwmSQ';
const ADMIN_EMAIL = 'lukyam.lmm@isgh.org.br';

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
    return SpreadsheetApp.openById(PLANILHA_ID);
  } catch (error) {
    throw new Error('Erro ao acessar planilha: ' + error.message);
  }
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
  const abas = ['LOGIN', 'CADASTRO', 'BASE'];
  let resultado = [];

  abas.forEach(aba => {
    let sheet = ss.getSheetByName(aba);
    if (!sheet) {
      sheet = ss.insertSheet(aba);
      resultado.push(`✓ Aba ${aba} criada`);
    }

    if (aba === 'LOGIN') {
      const headers = ['Nome', 'Matricula', 'Setor', 'SenhaHash', 'DataCriacao', 'Status', 'UltimaAlteracao'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (sheet.getLastRow() === 1) {
        const hashAdmin = calcularHashMD5('admin');
        sheet.getRange(2, 1, 1, 7).setValues([['Administrador', 'admin', 'Administração', hashAdmin, new Date(), 'Ativo', new Date()]]);
        resultado.push('✓ Admin criado (admin/admin)');
      }
      formatarCabecalho(sheet);
    } else if (aba === 'CADASTRO') {
      const headers = ['ID', 'Descricao', 'Setor', 'NivelAcesso', 'DataCriacao'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (sheet.getLastRow() === 1) {
        const agora = new Date();
        sheet.getRange(2, 1, 3, 5).setValues([
          [1, 'Acesso Total', 'Administração', 0, agora],
          [2, 'Acesso Técnico', 'Enfermagem', 1, agora],
          [3, 'Acesso Médico', 'Médico', 2, agora]
        ]);
        resultado.push('✓ Setores padrão criados');
      }
      formatarCabecalho(sheet);
    } else if (aba === 'BASE') {
      const headers = ['Nome', 'Prontuario', 'DataNascimento', 'Peso', 'Altura', 'PressaoArterial', 'Temperatura', 'Saturacao', 'Glicemia', 'DataRegistro', 'UsuarioRegistro', 'Observacoes'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatarCabecalho(sheet);
      resultado.push('✓ Aba BASE configurada');
    }
  });

  return resultado.join('\n');
}

function formatarCabecalho(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground('#2563eb').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
}

// LOGIN
function fazerLogin(matricula, senha) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    if (!sheet) return { success: false, message: 'Sistema em configuração' };

    const hash = calcularHashMD5(senha);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === matricula && data[i][3] === hash && data[i][5] === 'Ativo') {
        const tipo = determinarTipoUsuario(data[i][2]);
        return {
          success: true,
          user: {
            nome: data[i][0],
            matricula: data[i][1],
            setor: data[i][2],
            tipo: tipo
          }
        };
      }
    }
    return { success: false, message: 'Credenciais inválidas ou usuário inativo' };
  } catch (error) {
    return { success: false, message: 'Erro interno: ' + error.message };
  }
}

function determinarTipoUsuario(setor) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    if (!sheet) return 'TECNICO';

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === setor) {
        const nivel = data[i][3];
        if (nivel === 0) return 'ADM';
        if (nivel === 1) return 'TECNICO';
        if (nivel === 2) return 'MEDICO';
      }
    }
    return 'TECNICO';
  } catch (error) {
    return 'TECNICO';
  }
}

// Recuperação de senha
function recuperarSenha(matricula) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === matricula) {
        const novaSenha = Math.random().toString(36).substring(2, 10).toUpperCase();
        const hash = calcularHashMD5(novaSenha);

        sheet.getRange(i + 1, 4).setValue(hash);
        sheet.getRange(i + 1, 7).setValue(new Date());

        const emailUsuario = data[i][0].toLowerCase().replace(/\s+/g, '.') + '@isgh.org.br';
        const corpo = `Recuperação de Senha - ISGH\n\nUsuário: ${data[i][0]}\nNova Senha: ${novaSenha}\nData: ${new Date().toLocaleString('pt-BR')}`;
        
        MailApp.sendEmail(emailUsuario, 'Nova Senha - Sistema ISGH', corpo);
        MailApp.sendEmail(ADMIN_EMAIL, 'Recuperação Solicitada', `Usuário ${matricula} solicitou recuperação.`);

        return { success: true, message: `Senha resetada para ${data[i][0]}. Nova senha enviada por email.` };
      }
    }
    return { success: false, message: 'Matrícula não encontrada' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Cadastrar usuário
function cadastrarUsuario(nome, matricula, setor, senha) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === matricula) return { success: false, message: 'Matrícula já existe' };
    }

    const hash = calcularHashMD5(senha);
    sheet.appendRow([nome, matricula, setor, hash, new Date(), 'Ativo', new Date()]);

    return { success: true, message: 'Usuário cadastrado com sucesso' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Obter setores
function obterSetores() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    if (!sheet) return ['Administração', 'Enfermagem', 'Médico'];

    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(row => row[2]).filter(Boolean);
  } catch (error) {
    return ['Administração', 'Enfermagem', 'Médico'];
  }
}

// Obter setores com detalhes
function obterSetoresComDetalhes() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(row => ({
      id: row[0],
      descricao: row[1],
      setor: row[2],
      nivelAcesso: row[3]
    }));
  } catch (error) {
    return [];
  }
}

// Adicionar setor
function adicionarSetor(novoSetor, nivelAcesso) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    const id = sheet.getLastRow();
    const descricao = nivelAcesso === 0 ? 'Acesso Total' : nivelAcesso === 1 ? 'Acesso Técnico' : 'Acesso Médico';
    
    sheet.appendRow([id, descricao, novoSetor, parseInt(nivelAcesso), new Date()]);
    return { success: true, message: 'Setor adicionado com sucesso' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Deletar setor
function deletarSetor(id) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    sheet.deleteRow(parseInt(id) + 1);
    return { success: true, message: 'Setor deletado' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Editar setor
function editarSetor(id, novosDados) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('CADASTRO');
    const row = parseInt(id) + 1;
    sheet.getRange(row, 2).setValue(novosDados.descricao);
    sheet.getRange(row, 3).setValue(novosDados.setor);
    sheet.getRange(row, 4).setValue(novosDados.nivelAcesso);
    return { success: true, message: 'Setor atualizado' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Registrar paciente
function registrarDadosPaciente(dados) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    const usuario = Session.getActiveUser().getEmail().split('@')[0] || 'Sistema';

    sheet.appendRow([
      dados.nomeCompleto,
      dados.prontuario,
      dados.dataNascimento,
      parseFloat(dados.peso) || '',
      parseFloat(dados.altura) || '',
      dados.pressaoArterial,
      parseFloat(dados.temperatura) || '',
      parseFloat(dados.saturacao) || '',
      parseFloat(dados.glicemia) || '',
      new Date(),
      usuario,
      dados.observacoes || ''
    ]);

    return { success: true, message: 'Registro salvo com sucesso' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// NOVA FUNÇÃO: Obter todos os pacientes (todos os registros)
function getAllPacientes() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    
    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const registros = [];
    
    for (let i = 1; i < data.length; i++) {
      registros.push({
        nome: data[i][0] || '',
        prontuario: data[i][1] || '',
        dataNascimento: data[i][2] || '',
        peso: data[i][3] || '',
        altura: data[i][4] || '',
        pressaoArterial: data[i][5] || '',
        temperatura: data[i][6] || '',
        saturacao: data[i][7] || '',
        glicemia: data[i][8] || '',
        dataRegistro: data[i][9] || '',
        usuarioRegistro: data[i][10] || '',
        observacoes: data[i][11] || ''
      });
    }
    
    // Ordenar por data mais recente
    registros.sort((a, b) => {
      try {
        const dataA = new Date(a.dataRegistro);
        const dataB = new Date(b.dataRegistro);
        return dataB - dataA;
      } catch (e) {
        return 0;
      }
    });
    
    return registros;
    
  } catch (error) {
    console.error('Erro ao obter todos os pacientes:', error);
    return [];
  }
}

// Obter usuários
function obterUsuarios() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(row => ({
      nome: row[0],
      matricula: row[1],
      setor: row[2],
      ultimaAlteracao: row[6] ? new Date(row[6]) : 'N/A',
      status: row[5]
    }));
  } catch (error) {
    return [];
  }
}

// Deletar usuário
function deletarUsuario(matricula) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === matricula) {
        sheet.getRange(i + 1, 6).setValue('Inativo');
        sheet.getRange(i + 1, 7).setValue(new Date());
        return { success: true, message: 'Usuário inativado com sucesso' };
      }
    }
    return { success: false, message: 'Usuário não encontrado' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Editar usuário
function editarUsuario(matricula, novosDados) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('LOGIN');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === matricula) {
        sheet.getRange(i + 1, 1).setValue(novosDados.nome || data[i][0]);
        sheet.getRange(i + 1, 2).setValue(novosDados.matricula || data[i][1]);
        sheet.getRange(i + 1, 3).setValue(novosDados.setor || data[i][2]);
        if (novosDados.senha) {
          const hash = calcularHashMD5(novosDados.senha);
          sheet.getRange(i + 1, 4).setValue(hash);
        }
        sheet.getRange(i + 1, 7).setValue(new Date());
        return { success: true, message: 'Usuário atualizado' };
      }
    }
    return { success: false, message: 'Usuário não encontrado' };
  } catch (error) {
    return { success: false, message: 'Erro: ' + error.message };
  }
}

// Estatísticas
function obterEstatisticas() {
  try {
    const ss = getPlanilha();
    const loginSheet = ss.getSheetByName('LOGIN');
    const baseSheet = ss.getSheetByName('BASE');

    const totalUsuarios = (loginSheet ? loginSheet.getLastRow() - 1 : 0);
    const totalRegistros = (baseSheet ? baseSheet.getLastRow() - 1 : 0);

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);
    let registrosHoje = 0;
    let somaTemp = 0;

    if (baseSheet) {
      const data = baseSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const dataReg = new Date(data[i][9]);
        if (dataReg >= hoje && dataReg < amanha) {
          registrosHoje++;
          if (data[i][6]) somaTemp += parseFloat(data[i][6]);
        }
      }
    }

    const mediaTemperatura = registrosHoje > 0 ? somaTemp / registrosHoje : 0;

    return {
      totalUsuarios,
      totalRegistros,
      registrosHoje,
      mediaTemperatura
    };
  } catch (error) {
    return { totalUsuarios: 0, totalRegistros: 0, registrosHoje: 0, mediaTemperatura: 0 };
  }
}

// Registros de hoje
function obterRegistrosHoje() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    if (!sheet) return [];

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);

    const data = sheet.getDataRange().getValues();
    const hojeRegs = [];

    for (let i = 1; i < data.length; i++) {
      const dataReg = new Date(data[i][9]);
      if (dataReg >= hoje && dataReg < amanha) {
        hojeRegs.push({
          nome: data[i][0],
          prontuario: data[i][1],
          dataRegistro: data[i][9],
          temperatura: data[i][6]
        });
      }
    }

    return hojeRegs.slice(0, 5);
  } catch (error) {
    return [];
  }
}

// Gerar relatório
function gerarRelatorio(dataInicio, dataFim) {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    const inicio = new Date(dataInicio + 'T00:00:00');
    const fim = new Date(dataFim + 'T23:59:59');

    const data = sheet.getDataRange().getValues();
    let total = 0;
    const usuarios = new Set();
    let somaTemp = 0;

    for (let i = 1; i < data.length; i++) {
      const dataReg = new Date(data[i][9]);
      if (dataReg >= inicio && dataReg <= fim) {
        total++;
        usuarios.add(data[i][10]);
        if (data[i][6]) somaTemp += parseFloat(data[i][6]);
      }
    }

    return {
      total,
      usuariosUnicos: usuarios.size,
      mediaTemp: total > 0 ? somaTemp / total : 0,
      data: `${dataInicio} a ${dataFim}`
    };
  } catch (error) {
    return { total: 0, usuariosUnicos: 0, mediaTemp: 0, data: '' };
  }
}

// FUNÇÕES DE DEBUG
function debugProntuarios() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    if (!sheet) return '❌ Aba BASE não encontrada';
    
    const data = sheet.getDataRange().getValues();
    let resultado = '=== DEBUG PRONTUÁRIOS ===\n\n';
    resultado += `Total de registros: ${data.length - 1}\n\n`;
    
    for (let i = 1; i < data.length; i++) {
      const prontuario = data[i][1];
      const nome = data[i][0];
      resultado += `Linha ${i+1}: "${nome}" → Prontuário: "${prontuario}"\n`;
    }
    
    return resultado;
  } catch (error) {
    return '❌ Erro no debug: ' + error.message;
  }
}

function criarDadosTeste() {
  try {
    const ss = getPlanilha();
    const sheet = ss.getSheetByName('BASE');
    
    // Limpar dados existentes (exceto cabeçalho)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clear();
    }
    
    // Adicionar dados de teste
    const dadosTeste = [
      ['João Silva', '12345', '1990-05-15', 70.5, 175, '120/80', 36.5, 98, 95, new Date(), 'Admin', 'Paciente teste 1'],
      ['Maria Santos', '67890', '1985-08-20', 65.2, 165, '110/70', 36.8, 99, 100, new Date(), 'Admin', 'Paciente teste 2'],
      ['Pedro Oliveira', '11111', '1978-12-10', 80.0, 180, '130/85', 37.1, 97, 105, new Date(), 'Admin', 'Paciente teste 3']
    ];
    
    sheet.getRange(2, 1, dadosTeste.length, dadosTeste[0].length).setValues(dadosTeste);
    
    return '✅ Dados de teste criados! Prontuários: 12345, 67890, 11111';
  } catch (error) {
    return '❌ Erro ao criar dados teste: ' + error.message;
  }
}

// TESTE DIRETO - Busca específica
function testeBuscaDireta() {
  console.log('=== TESTE DIRETO DA BUSCA ===');
  
  // Testar com prontuário "1" que sabemos que existe
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
