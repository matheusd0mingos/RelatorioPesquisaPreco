function createNewGoogleDocs() {
  // Get the current sheet and its data
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("Pesquisa");
  var start_sheet = spreadsheet.getSheetByName('Start')
  var row_number = sheet.getLastRow().toString();
  var row_number = row_number.replace(".0","");
  var file_name = start_sheet.getRange('B5').getValue().toString();
  var obra = start_sheet.getRange('B6').getValue().toString();
  var agente = start_sheet.getRange('B7').getValue().toString();
  var funcao = start_sheet.getRange('B8').getValue().toString();
  var posto = start_sheet.getRange('B9').getValue().toString();
  var endereco = start_sheet.getRange('B10').getValue().toString();
  var hoje = new Date();
  hoje = hoje.getDate() + '/' + ( +hoje.getMonth() + +1).toLocaleString('en-US', {minimumIntegerDigits: 2, useGrouping:false}) + '/' +hoje.getFullYear();
  Logger.log(hoje);


  //Logger.log(file_name);
// Get your template and destination folder
  var template = DriveApp.getFileById("PREENCHA O SEU");
  var destinationFolder = DriveApp.getFolderById('PREENCHA O SEU');
  const copy = template.makeCopy(file_name, destinationFolder);
  var doc = DocumentApp.openById(copy.getId());
  
  var body = doc.getBody(); 
  body.replaceText("{{OBRA}}", obra);
  body.replaceText("{{ENDERECO}}", endereco);
  body.replaceText("{{AGENTE}}", agente);
  body.replaceText("{{POSTO}}", posto);
  body.replaceText("{{FUNCAO}}", funcao);
  body.replaceText("{{HOJE}}", hoje)
  var tables = body.getTables();
  var table = tables[1];
  //Logger.log(table);

  // Loop through each row in the sheet
  for(i=3; i<=row_number;i++){
    var j = +i-2;
    j = j.toString().replace(".0","");
    //Logger.log(j);
    var equipamento = sheet.getRange('A'+i).getValue().toString();
    var nome1 = sheet.getRange('B'+i).getValue().toString();
    var cnpj1 = sheet.getRange('C'+i).getValue().toString();
    var preco1 = sheet.getRange('D'+i).getValue().toString();
    var nome2 = sheet.getRange('E'+i).getValue().toString();
    var cnpj2 = sheet.getRange('F'+i).getValue().toString();
    var preco2 = sheet.getRange('G'+i).getValue().toString();
    var nome3 = sheet.getRange('H'+i).getValue();
    var cnpj3 = sheet.getRange('I'+i).getValue().toString();
    var preco3 = sheet.getRange('J'+i).getValue().toString();

    var arrayEmpresas = [];
    if (preco1!=0){
      const emp1 = new Empresa(nome1, cnpj1, preco1);
      arrayEmpresas.push(emp1);
    }
    if (preco2!=0){
      const emp2 = new Empresa(nome2, cnpj2, preco2);
      arrayEmpresas.push(emp2)
    }
    if (preco3!=0){
      const emp3 = new Empresa(nome3, cnpj3, preco3);
      arrayEmpresas.push(emp3);
    }
    const pesquisa = new Pesquisa(equipamento, arrayEmpresas);
    pesquisa.addPesquisa(j,table);      
  }
  var style = {};
  style[DocumentApp.Attribute.BORDER_COLOR] = '#000000';
  style[DocumentApp.Attribute.BORDER_WIDTH] = 0.5;
  table.setAttributes(style);
  //table = table.merge();
  doc.saveAndClose();

  var url = doc.getUrl();
  start_sheet.getRange('E4').setValue(url);
}

class Empresa{
  constructor(nomeEmpresa, cnpjEmpresa, precoEmpresa){
    this.nomeEmpresa = nomeEmpresa;
    this.cnpjEmpresa = cnpjEmpresa;
    this.precoEmpresa = precoEmpresa;

  }
  getPrice(){
    return this.precoEmpresa;
  }

}
class Pesquisa{
  constructor(equipamento, empresas = []){
    this.equipamento = equipamento;
    this.empresas = empresas;
  }
  getMedian(){
    var arrayPrices = [];
    for(const element of this.empresas){
      if(element.getPrice()!=0){
        arrayPrices.push(element.getPrice());
      }
    }
    var median = findMedian(arrayPrices);
    return median;
  }
  addPesquisa(count, table){

    var median = this.getMedian();
    Logger.log(median);
    median = Math.round(median * 100) / 100;
    for(const element of this.empresas){
      const row = [count, this.equipamento, element.nomeEmpresa, element.cnpjEmpresa, 'R$ '+ element.precoEmpresa.toString().replace('.',','), 'R$ ' + median.toString().replace('.',',')];

      var tr = table.appendTableRow();
      for (const amigo of row){
        tr.appendTableCell(amigo);
      }
    }    

  }
}
function findMedian(arr) {
    arr.sort((a, b) => a - b);
    const middleIndex = Math.floor(arr.length / 2);

    if (arr.length % 2 === 0) {
        return (+arr[middleIndex - 1] + +arr[middleIndex]) / 2;
    } else {
        return arr[middleIndex];
    }
}
