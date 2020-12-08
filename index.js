const axios = require('axios');
const qs = require('qs');
const xl = require('excel4node');
const codigos = require('./constants');

const wb = new xl.Workbook();
const ws = wb.addWorksheet('Zonas Extendidas');
const headingColumnNames = [
  "Estado",
  "Municipio",
  "Código Postal",
  "DHL",
  "Fedex",
  "Estafeta",
  "RedPack",
];

let finalArray = [];

async function getFile() {
  for(const codigo of codigos.codigospruebas) {
    await primero(codigo)
  }
  //console.log(finalArray)
  createExcel(finalArray)
}

function primero(codigo) {
  return axios({
    method: 'post',
    url: 'http://zonaextendida.com/consultarGuia.php',
    data: qs.stringify({
      numero: codigo
    }),
    headers: { 'content-type': 'application/x-www-form-urlencoded;charset=utf-8' }
  })
    .then(async (res) => {
      //console.log(res.data);
      await segundo(res.data);
    })
    .catch(error => {
      console.error(error);
    });
}

const segundo = (data) => axios({
  method: 'get',
  url: `http://zonaextendida.com/consultarGuia.php?${data}`,
  headers: {'content-type': 'application/x-www-form-urlencoded;charset=utf-8'}
})
.then(res => {
  let registro = {}
  //console.log('Segundapeticion: ', res.data)
  registro.estado = res.data.informacion.estado
  registro.municipio = res.data.informacion.municipio
  registro.cp = res.data.informacion.cp
  registro.DHL = res.data.informacion.DHL.zonaExtendida = 'N' ? 'No' : 'Si'
  registro.Fedex = res.data.informacion.Fedex.zonaExtendida = '0' ? 'No' : 'Si'
  registro.Estafeta = res.data.informacion.Estafeta.zonaExtendida = '0' ? 'No' : 'Si'
  registro.RedPack = res.data.informacion.RedPack.zonaExtendida = '0' ? 'No' : 'Si'

  finalArray.push(registro);

})
.catch(error => {
  console.error(error)
})

const createExcel = (data) => {
  let headingColumnIndex = 1;
  headingColumnNames.forEach(heading => {
    ws.cell(1, headingColumnIndex++).string(heading)
  });

  let rowIndex = 2;
  data.forEach( record => {
    let columnIndex = 1;
    Object.keys(record).forEach(columnName => {
      ws.cell(rowIndex, columnIndex++).string(record [columnName])
    });
    rowIndex++;
  });
  wb.write('ZonasExtendidas.xlsx');
}

getFile();
