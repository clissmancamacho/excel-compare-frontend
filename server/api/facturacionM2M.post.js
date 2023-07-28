import ExcelJS from "exceljs"
import { readFiles } from "h3-formidable"

export default defineEventHandler(async (event) => {
  // with fields
  const { fields, files } = await readFiles(event, {
    includeFields: true,
    // other formidable options here
  })

  try {
    const companies = {}
    const worksheets = await getWorksheetsFromFiles(files)
    const data2 = buildData2(worksheets.data2)
    const data3 = buildData3(worksheets.data3)
    const dataExtraCharges = buildDataCobrosExtras(worksheets.data5)
    const dataDiscounts = buildDataDescuentos(worksheets.data4)

    worksheets.data1.eachRow(function (row, rowNumber) {
      if ([1, 2, 3].includes(rowNumber)) {
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

      if (objToPush.country === "Chile (2)" && objToPush.currency !== "CLP") {
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
          const currencyRate = getCurrencyRate(objToPush.currency)
          extraCharges.value = customRound(extraCharges.value * currencyRate)
        }
        objToPush["totalCharge"] = objToPush.totalCharge + extraCharges.value
      }
      if (discount) {
        objToPush["discount"] = discount.value
      }
      pushKeyOrAddToExistingKey(companies, objToPush)
    })
    // debugger

    return response
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
    if (objToPush.currency !== objInCompanies.currency) {
      objToPush = convertCurrency(objToPush, objInCompanies.currency)
    }
    objInCompanies.totalCharge += objToPush.totalCharge
    objInCompanies.totalChargePlan += objToPush.totalChargePlan
    objInCompanies.totalChargeOverconsumption +=
      objToPush.totalChargeOverconsumption
    objInCompanies.totalChargeSMS += objToPush.totalChargeSMS
    objInCompanies.totalChargeVoice += objToPush.totalChargeVoice

    companies[objToPush.companyId] = objInCompanies
  } else {
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
  objToPush.currency = currency
  return objToPush
}

const getCurrencyRate = (currency) => {
  const currencies = {
    CLP: 1,
    S: 220.68,
    USD: 801.66,
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
