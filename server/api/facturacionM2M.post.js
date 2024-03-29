import ExcelJS from "exceljs" // Esto importa la libreria para manipular excel
import { readFiles } from "h3-formidable" // Esto importa la librearia para leer archivos

// Esta constante tiene el path de el archivo template de facturacion
const templateFile = {
  filepath: "server/template/plantilla",
}

let tasaSol = 0 // Declaracion de variable que obtiene mas adelante la tasa del sol con respecto al CLP
let tasaDolar = 0
let tasaMXN = 0
let tasaCOP_CLP = 0
let tasaCOP_Sol = 0
let tasaCOP_Dolar = 0
let tasaCOP_MXN = 0
let tasaARS_Dolar = 0
let tasaARS_CLP = 0
let tasaMXN_Dolar = 0
let facturationType = ''

export default defineEventHandler(async (event) => { 
  // with fields
  const { fields, files } = await readFiles(event, { // Obteniendo los campos y los archivos que me llegan de la peticion
    includeFields: true,
    // other formidable options here
  })

  const S = fields.S[0]              //Declaracion de constantes a traves de arreglos que definen los campos del archivo de excel
  const USD = fields.USD[0]
  const MXN = fields.MXN[0]
  const COP_CLP = fields.COP_CLP[0]
  const COP_SOL = fields.COP_SOL[0]
  const COP_DOLAR = fields.COP_DOLAR[0]
  const COP_MXN = fields.COP_MXN[0]
  const ARS_DOLAR = fields.ARS_DOLAR[0]
  const ARS_CLP = fields.ARS_CLP[0] 
  const MXN_DOLAR = fields.MXN_DOLAR[0] 
  const fechaFactura = fields.fechaFactura[0]
  const fechaVencimiento = fields.fechaVencimiento[0]
  const consumoMes = fields.consumoMes[0]
  const pais = fields.pais[0]
  const fType = fields.facturationType[0]

  try { // Inicia un bloque de manejo de errores

    //Validacion de campos
    if (!S || !USD || !MXN || !COP_CLP || !COP_SOL || !COP_DOLAR || !COP_MXN || !ARS_DOLAR || !ARS_CLP || !MXN_DOLAR ||!fechaFactura || !fechaVencimiento || !consumoMes || !pais) {
      throw new Error(
        "No se encontraron los campos S (tasa soles) o USD (tasa dolár)"
      )
    }

    if(!["Facturable Manual", "Facturable Masivo"].includes(fType)) {
      throw new Error(
        "facturationType es invalido"
      )
    }
    //Parseo de tasas
    tasaSol = parseFloat(S)
    tasaDolar = parseFloat(USD)
    tasaMXN = parseFloat(MXN)
    tasaCOP_CLP = parseFloat(COP_CLP)
    tasaCOP_Sol = parseFloat(COP_SOL)
    tasaCOP_Dolar = parseFloat(COP_DOLAR)
    tasaCOP_MXN = parseFloat(COP_MXN)
    tasaARS_Dolar = parseFloat(ARS_DOLAR)
    tasaARS_CLP = parseFloat(ARS_CLP)
    tasaMXN_Dolar = parseFloat(MXN_DOLAR)
    facturationType = fType

    const companies = {}
    const worksheets = await getWorksheetsFromFiles(files) // Obteniendo las hojas de excel
    const data2 = buildData2(worksheets.data2)
    const data3 = buildData3(worksheets.data3)
    const dataExtraCharges = buildDataCobrosExtras(worksheets.data5)
    const dataDiscounts = buildDataDescuentos(worksheets.data4)
    const paisRow = getNamedPaisRow(pais)
    worksheets.data1.eachRow(function (row, rowNumber) { //Ejecucion para encontrar las columnas correspondientes
      if ([1].includes(rowNumber)) {                     //y cruzar los valores de las mismas a traves de arreglos
        return
      }
      const values = row.values
      if (values[3] === 0) {
        return
      }
      let objToPush = { // se aplica un arreglo de objetos para indicar la columna que queremos cruzar
        country: values[2],
        companyId: values[3],
        companyName: values[4],
        totalChargePlan: values[13],
        totalChargeOverconsumption: values[14],
        totalChargeSMS: values[15],
        totalChargeVoice: values[16],
        currency: values[18],
      }
      objToPush["totalCharge"] =        //se obtiene la totalidad de cada columna realizando un arreglo de objetos
        objToPush.totalChargePlan +
        objToPush.totalChargeOverconsumption +
        objToPush.totalChargeSMS +
        objToPush.totalChargeVoice

      if (!paisRow.includes(objToPush.country)) {  //se aplica condiciones para obtener en este caso
        return
      }

      if (pais === 'chile' && objToPush.currency !== "CLP") {
        objToPush = convertCurrency(objToPush, objToPush.currency, "CLP")
      }

      else if (pais === 'global' && objToPush.currency !== "USD") {
        objToPush = convertCurrency(objToPush, objToPush.currency, "USD")
      }
      else if (pais === 'argentina' && objToPush.currency !== "AR $") {
        objToPush = convertCurrency(objToPush, objToPush.currency, "AR $")
      }

      else if((pais === 'peruSd' || pais === 'peruCd') && objToPush.currency !== "S") {
        objToPush = convertCurrency(objToPush, objToPush.currency, "S")
      }

      else if(pais === 'colombia' && objToPush.currency !== "COP") {
        objToPush = convertCurrency(objToPush, objToPush.currency, "COP")
      }

      const aditionalData2 = data2[objToPush.companyId]
      const aditionalData3 = data3[objToPush.companyId]
      const extraCharges = dataExtraCharges[objToPush.companyId]
      const discount = dataDiscounts[objToPush.companyId]
      if (aditionalData2) {
        objToPush = { ...objToPush, ...aditionalData2 }
      }

      if (objToPush.facturableType !== facturationType) {
        return
      }
      if (aditionalData3) {
        objToPush = { ...objToPush, ...aditionalData3 }
      }
      if (extraCharges) {
        if (objToPush.currency !== extraCharges.currency) {
          const currencyRate = getCurrencyRate(extraCharges.currency, objToPush.currency)
          extraCharges.value = customRound(extraCharges.value * currencyRate)
          extraCharges.currency = objToPush.currency
          if(objToPush.companyId == 4504) {
            console.log({extraCharges, objToPush})
          }
        }
        objToPush["extraCharge"] = extraCharges.value
      }
      if (discount) {
        objToPush["discount"] = discount.value
      }
      pushKeyOrAddToExistingKey(companies, objToPush)
    })

    await writeRowsInTemplateFile(
      companies,
      fechaFactura,
      fechaVencimiento,
      consumoMes,
      pais
    )
    return companies
  } catch (error) {
    return {
      error: error.message,
    }
  }
})

const getWorksheetsFromFiles = async (files) => {
  const obj = {}
  let index = 1
  for (const key of Object.keys(files)) {
    const file = files[key][0]
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(file.filepath)
    obj[`data${index}`] = workbook.getWorksheet(getWorksheetNameByData(index))
    index++
  }
  return obj
}

const getWorksheetNameByData = (index) => {
  if ([1, 3, 6].includes(index)) return "Hoja1"
  if (index === 4) return "Descuentos"
  if (index === 5) return "Cobros Extras"
  return "Sheet1"
}

const pushKeyOrAddToExistingKey = (companies, objToPush) => {
  const objInCompanies = companies[objToPush.companyId]
  if (objInCompanies) {
    /*  if (objToPush.currency !== objInCompanies.currency) {
      objToPush = convertCurrency(objToPush, objInCompanies.currency)
    } */
    objInCompanies.totalCharge += customRound(objToPush.totalCharge)
    objInCompanies.totalChargePlan += customRound(objToPush.totalChargePlan)
    objInCompanies.totalChargeOverconsumption += customRound(
      objToPush.totalChargeOverconsumption
    )
    objInCompanies.totalChargeSMS += customRound(objToPush.totalChargeSMS)
    objInCompanies.totalChargeVoice += customRound(objToPush.totalChargeVoice)

    if (!objInCompanies.extraCharge && objToPush.extraCharge) {
      objInCompanies.totalCharge += customRound(objToPush.extraCharges)
    }

    companies[objToPush.companyId] = objInCompanies
  } else {
    if (objToPush.extraCharge) {
      objToPush.totalCharge += customRound(objToPush.extraCharge)
    }
    companies[objToPush.companyId] = objToPush
  }
}

const convertCurrency = (objToPush, currency, baseCurrency) => {
  const currencyRate = getCurrencyRate(currency, baseCurrency)
  objToPush.totalCharge = customRound(objToPush.totalCharge * currencyRate)
  objToPush.totalChargePlan = customRound(
    objToPush.totalChargePlan * currencyRate
  )

  objToPush.totalChargeOverconsumption = customRound(
    objToPush.totalChargeOverconsumption * currencyRate
  )
  objToPush.totalChargeSMS = customRound(
    objToPush.totalChargeSMS * currencyRate
  )
  objToPush.totalChargeVoice = customRound(
    objToPush.totalChargeVoice * currencyRate
  )
  objToPush.currency = baseCurrency
  return objToPush
}

const getCurrencyRate = (currency, baseCurrency = 'CLP') => {
  const currencies = {
    'CLP': {
      CLP: 1,
      S: tasaSol,
      USD: tasaDolar,
      'MXN $': tasaMXN
    },
    'USD': {
      USD: 1,
      CLP: 1 / tasaDolar,
      'MXN $': 1 / tasaMXN_Dolar
    },
    'S': {
      S: 1,
      CLP: 1 / tasaSol
    },
    'COP': {
      COP: 1,
      CLP: tasaCOP_CLP,
      S: tasaCOP_Sol,
      USD: tasaCOP_Dolar,
      'MXN $': tasaCOP_MXN
    }, 
    'AR $': {
      ARS: 1,
      USD: tasaARS_Dolar,
      CLP: tasaARS_CLP
    },
  }
  return currencies[baseCurrency][currency]
}

const buildData2 = (data2) => {
  const data = {}
  data2.eachRow(function (row, rowNumber) {
    if ([1].includes(rowNumber)) {
      return
    }
    const values = row.values
    if (!values[1]) {
      return
    }
    data[values[1]] = {
      odooId: values[2],
      facturableType: values[4],
      fantasyName: values[5],
      data2Name: values[6],
      consumer: values[16] 
    }
  })
  return data
}

const buildData3 = (data3) => {
  const data = {}
  data3.eachRow(function (row, rowNumber) {
    if ([1].includes(rowNumber)) {
      return
    }
    const values = row.values
    if (!values[1]) {
      return
    }
    data[values[1]] = {
      showName: values[2],
    }
  })
  return data
}

const buildDataCobrosExtras = (data5) => {
  const data = {}
  data5.eachRow(function (row, rowNumber) {
    if ([1].includes(rowNumber)) {
      return
    }
    const values = row.values
    if (!values[1]) {
      return
    }
    let objToPush = {
      companyId: values[1],
      value: values[4],
      currency: values[5],
    }

    const objInData = data[objToPush.companyId]
    if (objInData) {
      if (objToPush.currency !== objInData.currency) {
        const currencyRate = getCurrencyRate(objInData.currency)
        objToPush.value = customRound(objToPush.value * currencyRate)
      }
      objInData.value += objToPush.value

      data[objToPush.companyId] = objInData
    } else {
      data[objToPush.companyId] = objToPush
    }
  })
  return data
}

const buildDataDescuentos = (data4) => {
  const data = {}
  data4.eachRow(function (row, rowNumber) {
    if ([1].includes(rowNumber)) {
      return
    }
    const values = row.values
    if (!values[1]) {
      return
    }
    data[values[1]] = {
      companyId: values[1],
      companyName: values[2],
      value: values[6],
    }
  })
  return data
}

const customRound = (value, decimals = 2) => {
  return parseFloat(value.toFixed(decimals))
}

const getNamedPaisRow = (pais) => {
  const namedPais = {
    'chile': ['Chile (2)'],
    'ecuador': ['Ecuador (5)'],
    'mexico': ['MEXICO (8)'],
    'global': ['Bolivia (44)', 'Brasil (353)', 'Canadá (3740)', 'Costa Rica (160)', 'Estados Unidos (9)', 'Guatemala (860)', 'Honduras (4141)', 'Nicaragua (4140)', 'Panama (648)', 'Paraguay (1229)', 'Uruguay (667)', 'Venezuela (3224)', 'Colombia (4)'],
    'peruSd' : ['Peru (11)'],
    'peruCd' : ['Peru (11)'],
    'colombia' : ['Colombia (4)'],
    'argentina' : ['Argentina M2M (14)']
  }
  return namedPais[pais]
}

const writeRowsInTemplateFile = async (
  companies,
  fechaFactura,
  fechaVencimiento,
  consumoMes, 
  pais
) => {
  const templateWorkbook = new ExcelJS.Workbook()
  await templateWorkbook.xlsx.readFile(`${templateFile.filepath}_${pais}.xlsx`)
  const template = templateWorkbook.getWorksheet("Hoja1")

  const keys = Object.keys(companies)
  for await (const key of keys) {
    const company = companies[key]
    let writeCompany = true
    if(pais === 'peruSd' || pais === 'peruCd') {
      writeCompany = verifyDetractionPeru(pais, company)
    }
    if(writeCompany) {
      const rowToWrite = getRowToWriteByCountry(pais, company, fechaFactura, fechaVencimiento, consumoMes)
      template.addRow(rowToWrite)
    }
  }
  await templateWorkbook.xlsx.writeFile("server/template/facturacion.xlsx")
}

const getRowToWriteByCountry = (pais, company, fechaFactura, fechaVencimiento, consumoMes) => {
  if (pais === 'chile') {
    return [
      company.fantasyName || company.showName || company.companyName,
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "Facturas de cliente",
      "CLP",
      "__export__.product_product_725_194cd9a3",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "IVA 19% Vta",
      0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
    ]
  }
  if (pais === 'ecuador') {
    return [
      company.fantasyName || company.showName || company.companyName,
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "Facturas de cliente",
      "USD",
      "Servicios de Conectividad",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_ec.6_tax_vat_411",
      0,
      "Conectividad Gestionada",
      "Debo y pagaré incondicionalmente y sin protesto esta factura a M2M DATAGOBAL LATAM ECUADOR CIA LTDA. El valor de la factura (s) pendiente (s) EN EFECTIVO/DEPOSITO/TRANSFERENCIA A LA CUENTA CORRIENTE BANCO PICHINCHA N° 2100202222 DE M2M DATAGLOBAL. El recibo y retención lo enviare al correo electrónico alejandra.valencia@m2mdataglobal.com. Declaro, que los datos para la generación de la presente factura son verídicos y autorizo en forma expresa a; M2M DATAGLOBAL LATAM ECUADOR CIA LTDA. A Solicitar o publicar toda la información crediticia o de mora en cualquier fuente de información o publicaciones, incluidos los Burós de Crédito legalmente autorizados por la Superintendencia de Compañías y que para gestionar cobros y demás no será requisito que las facturas tengan firma alguna"
    ]
  }
  if (pais === 'mexico') {
    return [
      company.odooId,
      fechaFactura,
      "10 Días",
      "Facturas de cliente",
      "MXN",
      "Servicios Conectividad Gestionada",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_mx.5_tax12",
      0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
      "Unidades",
      "Gastos en General"
    ]
  }
  if (pais === 'peruSd') {
    return [
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "PEN",
      "__export__.product_product_879_a63f9941",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_pe.3_sale_tax_igv_18",
      0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
      "l10n_pe.document_type01",
      "Unidades",
    ]
  }
  if (pais === 'peruCd') {
    return [
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "PEN",
      "_export_.product_product_879_a63f9941",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_pe.3_sale_tax_igv_18",
      0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
      "l10n_pe.document_type01",
      "Unidades",
      "[1001] Operación Sujeta a Detracción"
    ]
  }
  if (pais === 'global') {
    return [
      company.fantasyName || company.showName || company.companyName,
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "Facturas de cliente",
      "USD",
      "__export__.product_product_725_194cd9a3",
      "Servicios de Conectividad Gestionada consumos de Febrero 2024",
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
      "Factura de Exportación Electrónica",
    ]
  }
  if (pais === 'colombia') {
    return [
      company.fantasyName || company.showName || company.companyName,
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "Facturas de cliente",
      "COP",
      "Servicios Conectividad Gestionada",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_co.7_l10n_co_tax_8",
      0,
      "Conectividad Gestionada",
      "",
      "[0101] Internal sale",
      "Instrumento no definido",
    ]
  }
  if (pais === 'argentina') {
    return [
      company.data2Name,
      fechaFactura,
      fechaVencimiento,
      "Factura Electrónica",
      company.consumer === 'Consumidor Final' ? "FACTURAS B" : "FACTURAS A",
      "ARS",
      "Servicios M2M",
      consumoMes,
      1,
      company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge,
      "l10n_ar.4_ri_tax_vat_21_ventas",
      "",
      "Conectividad Gestionada",
      "",
    ]
  }
}

const verifyDetractionPeru = (pais, company) => {
  const totalCharge = company?.discount ? customRound(company.totalCharge - company.discount) : company.totalCharge
  let writeCompany = false
  if(pais === 'peruSd' && totalCharge <= 582) {
    writeCompany = true
  }
  if(pais === 'peruCd' && totalCharge >= 583) 
  {
    writeCompany = true
  }
  return writeCompany
}
