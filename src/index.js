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

// Getting day name:
const weekDays = [
    "Domingo",
    "Segunda-feira",
    "Terçã-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado"
]

// Setting a fixed date to test:
const date = "2022-08-08" // the menu starts in the 8th day of the month

let weekDay = new Date(date).getDay() + 1
let monthDay = new Date(date).getUTCDate()
let currentMonth = months[new Date(date).getUTCMonth()]

console.log(monthDay)

// Treating month writing pattern (Ex.: "AGOSTO" to "Agosto"):
currentMonth = currentMonth[0] + currentMonth.toLowerCase().slice(1)

// Main method to read our xls/xlsx file:
const parseExcel = (fileName) => {
    const excelData = xlsx.readFile(fileName)

    return Object.keys(excelData.Sheets).map(name => ({
        name,
        data: xlsx.utils.sheet_to_json(excelData.Sheets[name])
    }))
}

// This is the offset to get our current day in the .xls file:
let currentXlsDay = monthDay - 3

parseExcel(`./files/Cardápio-${currentMonth}.xlsx`).forEach(element => {
    /*
    // Uncomment this to see the whole data from the .xls file:
    console.log(element.data)
    */

    console.log(
        "Data: " + weekDays[weekDay] + ", " + monthDay + " de " + currentMonth
    )
 
    console.log(
        "- Cardápio -" +
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
