// ID da planilha do Google Sheets criada manualmente
const SPREADSHEET_ID = "1L9Cvup216xjr4AQqXIUGMxrIophmyJut3YK6K2M4N0U";

/**
 * Função principal que serve a interface web do Apps Script
 */
function doGet() {
  try {
    initRootUser();
  } catch (e) {
    Logger.log("Aviso na inicialização do Usuário Root: " + e.message);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Sistema de Biblioteca Escolar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função helper para incluir arquivos no HTML (ex: Style.html/Style.css)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Algoritmo de Criptografia SHA-256 Nativo do Google Apps Script
 */
function hashSHA256(text) {
  if (!text) return "";
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  let txtHash = "";
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    let byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = "0" + byteHex;
    txtHash += byteHex;
  }
  return txtHash;
}

/**
 * Retorna a aba da planilha com base no nome da tabela
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`A aba "${sheetName}" não foi encontrada. Crie a aba na planilha com esse exato nome.`);
  }
  return sheet;
}

/**
 * Inicializa a verificação e criação do usuário Root (Admin@email.com / 123)
 */
function initRootUser() {
  const sheet = getSheet("Usuarios");
  const data = sheet.getDataRange().getValues();

  let rootExists = false;

  for (let i = 1; i < data.length; i++) {
    const userEmail = data[i][2];
    if (userEmail && userEmail.toString().toLowerCase() === "admin@email.com") {
      rootExists = true;
      break;
    }
  }

  if (!rootExists) {
    const rootPassHash = hashSHA256("123");
    sheet.appendRow([
      "USR-ROOT",
      "Administrador Root",
      "Admin@email.com",
      rootPassHash,
      "(00) 00000-0000",
      "Root",
      "",
      "Ativo"
    ]);
  }
}

/**
 * Processa a autenticação de login do usuário
 */
function loginUser(email, password) {
  try {
    const sheet = getSheet("Usuarios");
    const data = sheet.getDataRange().getValues();
    const passHash = hashSHA256(password);

    for (let i = 1; i < data.length; i++) {
      const uEmail = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
      const uPass = data[i][3] ? data[i][3].toString() : "";
      const uStatus = data[i][7] ? data[i][7].toString() : "";

      if (uEmail === email.trim().toLowerCase()) {
        if (uPass === passHash) {
          if (uStatus === "Suspenso") {
            return { success: false, message: "Conta suspensa. Entre em contato com a administração." };
          }
          return {
            success: true,
            user: {
              id: data[i][0],
              nome: data[i][1],
              email: data[i][2],
              perfil: data[i][5],
              turma: data[i][6]
            }
          };
        } else {
          return { success: false, message: "Senha incorreta." };
        }
      }
    }
    return { success: false, message: "Usuário não cadastrado." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Cria novos usuários no sistema
 */
function createUser(userData) {
  try {
    const sheet = getSheet("Usuarios");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === userData.email.trim().toLowerCase()) {
        return { success: false, message: "Este e-mail já está cadastrado no sistema." };
      }
    }

    const passHash = hashSHA256(userData.senha);
    const newId = userData.id_usuario || "USR-" + new Date().getTime();

    sheet.appendRow([
      newId,
      userData.nome_completo,
      userData.email.trim(),
      passHash,
      userData.telefone || "",
      userData.perfil || "Aluno",
      userData.id_turma || "",
      userData.status_conta || "Ativo"
    ]);

    return { success: true, message: `Usuário ${userData.nome_completo} cadastrado com sucesso!` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Função de Recuperação de Senha ("Esqueci minha senha")
 */
function forgotPassword(email) {
  try {
    const sheet = getSheet("Usuarios");
    const data = sheet.getDataRange().getValues();
    let userRow = -1;
    let userName = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === email.trim().toLowerCase()) {
        userRow = i + 1;
        userName = data[i][1];
        break;
      }
    }

    if (userRow === -1) {
      return { success: false, message: "E-mail não encontrado no cadastro." };
    }

    const tempPass = "bib" + Math.floor(100000 + Math.random() * 900000);
    const tempHash = hashSHA256(tempPass);

    sheet.getRange(userRow, 4).setValue(tempHash);

    try {
      MailApp.sendEmail({
        to: email,
        subject: "Recuperação de Senha - Biblioteca Escolar",
        htmlBody: `<h3>Recuperação de Senha</h3> <p>Olá <b>${userName}</b>,</p> <p>Sua nova senha temporária de acesso é: <strong style="font-size:16px; color:#2c3e50;">${tempPass}</strong></p> <p>Recomendamos realizar o login e alterar sua senha posteriormente.</p>`
      });
      return { success: true, message: "Uma nova senha temporária foi enviada para o seu e-mail." };
    } catch (mailErr) {
      return {
        success: true,
        message: `Senha redefinida com sucesso! Sua senha temporária é: ${tempPass}`
      };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Consulta dados das 8 tabelas aplicando regras estritas de privilégio para a tabela "Usuarios"
 */
function getTableData(tableName, userEmail, userPerfil) {
  try {
    const sheet = getSheet(tableName);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    let rows = data.slice(1);

    // Regras de segurança para a tabela "Usuarios"
    if (tableName === "Usuarios") {
      
      // 1. Perfil ALUNO: visualiza APENAS o seu próprio registro
      if (userPerfil === "Aluno") {
        const emailColIndex = headers.indexOf("email");
        rows = rows.filter(row => {
          return row[emailColIndex] && row[emailColIndex].toString().trim().toLowerCase() === (userEmail || "").trim().toLowerCase();
        });

        return rows.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            if (header !== "senha_hash") { // Oculta a senha hash
              obj[header] = row[index];
            }
          });
          return obj;
        });
      }

      // 2. Perfil ADMIN: visualiza todos os registros, mas apenas colunas permitidas
      if (userPerfil === "Admin") {
        const allowedFields = ["email", "telefone", "nome_completo", "status_conta", "id_turma"];
        return rows.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            if (allowedFields.indexOf(header) !== -1) {
              obj[header] = row[index];
            }
          });
          return obj;
        });
      }

      // 3. Perfil ROOT: visualiza todas as colunas de todos os usuários
      if (userPerfil === "Root") {
        return rows.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
      }

      // Demais perfis padrão (ex: Professor/Funcionário): veem apenas seus próprios dados
      const emailColIndex = headers.indexOf("email");
      rows = rows.filter(row => {
        return row[emailColIndex] && row[emailColIndex].toString().trim().toLowerCase() === (userEmail || "").trim().toLowerCase();
      });

      return rows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
          if (header !== "senha_hash") {
            obj[header] = row[index];
          }
        });
        return obj;
      });
    }

    // Consulta normal sem restrições de linha/coluna para as demais tabelas
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

  } catch (e) {
    return [];
  }
}