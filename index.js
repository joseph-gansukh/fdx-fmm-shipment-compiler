const xlsx = require('xlsx');
const fs = require('fs');

const files = fs.readdirSync('./files')
console.log('files: ', files);

// Iterate through each files

const run = async _ => {

    for (let i = 0; i < files.length; ++i) {
        console.log(files[i])
    }
    const fileName = './files/CMA CGM LITANI 4525 - AVAILABLE.xlsm'
    const wb = xlsx.readFile(fileName, { cellDates: true })
    const wsName = wb.SheetNames[0]
    const ws = wb.Sheets[wsName]

    // ENCODING
    const ec = (r, c) => {
        return xlsx.utils.encode_cell({ r: r, c: c })
    }

    const delete_row = (ws, row_index) => {
        let range = xlsx.utils.decode_range(ws["!ref"])
        for (var R = row_index; R <= range.e.r; ++R) {
            for (var C = range.s.c; C <= range.e.c; ++C) {
                ws[ec(R, C)] = ws[ec(R + 1, C)]
            }
        }
        range.e.r--
            ws['!ref'] = xlsx.utils.encode_range(range.s, range.e)
    }

    const devanDate = 'H4'
    const devanDateValue = ws[devanDate].v.toLocaleDateString()
    console.log('devanDateValue: ', devanDateValue);

    for (let i = 0; i < 11; i++) {
        delete_row(ws, 0)
    }
    xlsx.writeFile(wb, fileName)


    const json = xlsx.utils.sheet_to_json(ws)

    // Remove empty rows
    const cleanData = []

    json.map(el => {
        if (el["HBL#"]) {
            delete el['Days past 7']
            delete el['Initial Storage']
            delete el['Days past 12']
            delete el['Additional Storage']
            delete el['Total Storage']
            delete el[' TOTAL: ']
            delete el[' STORAGE ']
            delete el[' WHSE FEE ']
            delete el['Shipper Pallets']
            delete el.Rate
            delete el.MIN
            delete el.MAX
            delete el['FTN Pallets']
            delete el.__EMPTY
            delete el.LOCATION
            cleanData.push(el)
        }
    })
    console.log(cleanData)

    // Write JSON data to compiled spreadsheet
    const newWb = xlsx.utils.book_new();
    const newWs = xlsx.utils.json_to_sheet(cleanData)

    xlsx.utils.book_append_sheet(newWb, newWs, "Compiled Data")
    xlsx.writeFile(newWb, "./files/Compilation.xlsx")

}

// run()