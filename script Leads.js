function importarLeadsDoGmail() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads");
  const threads = GmailApp.search('subject:"Lead" newer_than:1d in:sent OR in:inbox');
  const emails = GmailApp.getMessagesForThreads(threads);

  emails.forEach(thread => {
    thread.forEach(email => {
      const corpo = email.getPlainBody();
      const dados = extrairDados(corpo);
      if (dados) {
        const ultimaLinha = planilha.getLastRow() + 1; 
        planilha.getRange(ultimaLinha, 2, 1, dados.length).setValues([dados]); 
      }
    });
  });
}

function extrairDados(texto) {
  const regexEmpresa = /Empresa:\s*(.*)/;
  const regexResponsavel = /Responsável:\s*(.*)/;
  const regexCargo = /Cargo:\s*(.*)/;
  const regexTelefone = /Telefone:\s*(.*)/;
  const regexEmail = /Email:\s*(.*)/;
  const regexProduto = /Produto:\s*(.*)/;
  const regexData = /Data:\s*(.*)/;
  const regexSite = /Site:\s*(.*)/;
  const regexObs = /Observações:\s*(.*)/;

  const empresa = texto.match(regexEmpresa)?.[1]?.trim() || "Não informado";
  const contato = texto.match(regexResponsavel)?.[1]?.trim() || "Não informado";
  const cargo = texto.match(regexCargo)?.[1]?.trim() || "Não informado";
  const telefone = texto.match(regexTelefone)?.[1]?.trim() || "Não informado";
  const email = texto.match(regexEmail)?.[1]?.trim() || "Não informado";
  const produto = texto.match(regexProduto)?.[1]?.trim() || "Não informado";
  const data = texto.match(regexData)?.[1]?.trim() || "Não informado";
  const site = texto.match(regexSite)?.[1]?.trim() || "Não informado";
  const observacoes = texto.match(regexObs)?.[1]?.trim() || "Não informado";

  return [empresa, contato, cargo, telefone, email, produto, data, site, observacoes];
}
