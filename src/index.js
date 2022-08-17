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

parseExcel("./files/Cardápio-AGOSTO.xlsx").forEach(element => {
    
    console.log(
        "Data: " +
        //element.data[5]["__EMPTY"], // day name
        //element.data[14][`${currentMonth}`][0] + // day
        element.data[currentDay - 3][`${currentMonth}`] + // day
        " de " +
        `${currentMonth}`[0] + `${currentMonth}`.toLowerCase().slice(1) // month
    )

    // Remove acentuation:
    // exampleString.normalize('NFD').replace(/[\u0300-\u036f]/g, "")

    console.log(
        "Cardápio: " +
        "\n" +
        "Principal: " + element.data[currentDay - 3]["__EMPTY_1"] + " e " + element.data[currentDay - 3]["__EMPTY_2"] +
        "\n" +
        "Prato Protéico: " + element.data[currentDay - 3]["__EMPTY_3"] +
        "\n" +
        "Vegetariana: " + element.data[currentDay - 3]["__EMPTY_4"] +
        "\n" +
        "Vegana: " + element.data[currentDay - 3]["__EMPTY_currentDay - 3"] +
        "\n" +
        "Guarnição: " + element.data[currentDay - 3]["__EMPTY_6"] + 
        "\n" +
        "Salada - Folhas: " + element.data[currentDay - 3]["__EMPTY_7"] +
        "\n" +
        "Salada - Legumes/Acompanhamentos: " + element.data[currentDay - 3]["__EMPTY_8"] +
        "\n" +
        "Salada - Cozidos: " + element.data[currentDay - 3]["__EMPTY_9"] +
        "\n" +
        "Salada - Composta: " + element.data[currentDay - 3]["__EMPTY_10"] +
        "\n" +
        "Sobremesa: " + element.data[currentDay - 3]["__EMPTY_11"]
    )

    //console.log(element.data[14])
})
