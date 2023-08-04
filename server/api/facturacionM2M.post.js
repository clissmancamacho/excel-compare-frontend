import ExcelJS from "exceljs"
import { readFiles } from "h3-formidable"

const templateFile = {
  filepath: "server/template/plantilla.xlsx",
}

let tasaSol = 0
let tasaDolar = 0
export default defineEventHandler(async (event) => {
  // with fields
  const { fields, files } = await readFiles(event, {
    includeFields: true,
    // other formidable options here
  })

  const S = fields.S[0]
  const USD = fields.USD[0]
  const fechaFactura = fields.fechaFactura[0]
  const fechaVencimiento = fields.fechaVencimiento[0]
  const consumoMes = fields.consumoMes[0]

  try {
    if (!S || !USD || !fechaFactura || !fechaVencimiento || !consumoMes) {
      throw new Error(
        "No se encontraron los campos S (tasa soles) o USD (tasa dolár)"
      )
    }
    tasaSol = parseFloat(S)
    tasaDolar = parseFloat(USD)

    const companies = {}
    const worksheets = await getWorksheetsFromFiles(files)
    const data2 = buildData2(worksheets.data2)
    const data3 = buildData3(worksheets.data3)
    const dataExtraCharges = buildDataCobrosExtras(worksheets.data5)
    const dataDiscounts = buildDataDescuentos(worksheets.data4)

    worksheets.data1.eachRow(function (row, rowNumber) {
      if ([1].includes(rowNumber)) {
        return
      }
      const values = row.values
      if (values[3] === 0) {
        return
      }
      let objToPush = {
        country: values[2],
        companyId: values[3],
        companyName: values[4],
        totalChargePlan: values[13],
        totalChargeOverconsumption: values[14],
        totalChargeSMS: values[15],
        totalChargeVoice: values[16],
        currency: values[18],
      }
      objToPush["totalCharge"] =
        objToPush.totalChargePlan +
        objToPush.totalChargeOverconsumption +
        objToPush.totalChargeSMS +
        objToPush.totalChargeVoice

      if (objToPush.country !== "Chile (2)") {
        return
      }

      if (objToPush.currency !== "CLP") {
        objToPush = convertCurrency(objToPush, objToPush.currency)
      }

      const aditionalData2 = data2[objToPush.companyId]
      const aditionalData3 = data3[objToPush.companyId]
      const extraCharges = dataExtraCharges[objToPush.companyId]
      const discount = dataDiscounts[objToPush.companyId]
      if (aditionalData2) {
        objToPush = { ...objToPush, ...aditionalData2 }
      }

      if (objToPush.facturableType !== "Facturable Masivo") {
        return
      }
      if (aditionalData3) {
        objToPush = { ...objToPush, ...aditionalData3 }
      }
      if (extraCharges) {
        if (objToPush.currency !== extraCharges.currency) {
          const currencyRate = getCurrencyRate(extraCharges.currency)
          extraCharges.value = customRound(extraCharges.value * currencyRate)
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
      consumoMes
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
  if ([1, 6].includes(index)) return "Hoja1"
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

const convertCurrency = (objToPush, currency) => {
  const currencyRate = getCurrencyRate(currency)
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
  objToPush.currency = "CLP"
  return objToPush
}

const getCurrencyRate = (currency) => {
  const currencies = {
    CLP: 1,
    S: tasaSol,
    USD: tasaDolar,
  }
  return currencies[currency]
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

const writeRowsInTemplateFile = async (
  companies,
  fechaFactura,
  fechaVencimiento,
  consumoMes
) => {
  const templateWorkbook = new ExcelJS.Workbook()
  await templateWorkbook.xlsx.readFile(templateFile.filepath)
  const template = templateWorkbook.getWorksheet("Hoja1")

  const keys = Object.keys(companies)
  for await (const key of keys) {
    const company = companies[key]
    const rowToWrite = [
      company.showName || company.fantasyName || company.companyName,
      company.odooId,
      fechaFactura,
      fechaVencimiento,
      "Facturas de cliente",
      "CLP",
      "__export__.product_product_725_194cd9a3",
      consumoMes,
      1,
      company.totalCharge,
      "IVA 19% Vta",
      company?.discount ? company.discount : 0,
      "Conectividad Gestionada",
      "Servicios de Conectividad Gestionada",
    ]
    template.addRow(rowToWrite)
  }
  await templateWorkbook.xlsx.writeFile("server/template/facturacion.xlsx")
}
