// Importing 'xlsx' to read xls/xlsx files:
const xlsx = require('xlsx')

// Getting current month:
const months = [
    "JANEIRO",
    "FEVEREIRO",
    "MARÇO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEBRO",
    "DEZEMBRO"
]

let currentMonth = months[new Date().getMonth()]
let currentDay = new Date().getDate()

// Main method to read our xls/xlsx file:
const parseExcel = (fileName) => {
    const excelData = xlsx.readFile(fileName)

    return Object.keys(excelData.Sheets).map(name => ({
        name,
        data: xlsx.utils.sheet_to_json(excelData.Sheets[name]),
    }))
}

// This is the offset to get our current day in the .xls file:
let currentXlsDay = currentDay - 3

parseExcel("./files/Cardápio-AGOSTO.xlsx").forEach(element => {
    
    console.log(
        "Data: " +
        element.data[currentXlsDay]["__EMPTY"] + // day name
        ", " +
        element.data[currentXlsDay][`${currentMonth}`] + // day
        " de " +
        `${currentMonth}`[0] + `${currentMonth}`.toLowerCase().slice(1) // month
    )

    // Remove acentuation:
    // exampleString.normalize('NFD').replace(/[\u0300-\u036f]/g, "")

    console.log(
        "- CARDÁPIO - " +
        "\n" +
        "Principal: " + element.data[currentXlsDay]["__EMPTY_1"] + " e " + element.data[currentXlsDay]["__EMPTY_2"] +
        "\n" +
        "Prato Protéico: " + element.data[currentXlsDay]["__EMPTY_3"] +
        "\n" +
        "Vegetariana: " + element.data[currentXlsDay]["__EMPTY_4"] +
        "\n" +
        "Vegana: " + element.data[currentXlsDay]["__EMPTY_5"] +
        "\n" +
        "Guarnição: " + element.data[currentXlsDay]["__EMPTY_6"] + 
        "\n" +
        "Salada - Folhas: " + element.data[currentXlsDay]["__EMPTY_7"] +
        "\n" +
        "Salada - Legumes/Acompanhamentos: " + element.data[currentXlsDay]["__EMPTY_8"] +
        "\n" +
        "Salada - Cozidos: " + element.data[currentXlsDay]["__EMPTY_9"] +
        "\n" +
        "Salada - Composta: " + element.data[currentXlsDay]["__EMPTY_10"] +
        "\n" +
        "Sobremesa: " + element.data[currentXlsDay]["__EMPTY_11"]
    )
})
