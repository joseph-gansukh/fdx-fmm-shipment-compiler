const xlsx = require('xlsx');
const fs = require('fs');

const filePath = './files/Compilation.xlsx';
fs.unlinkSync(filePath);

const filePath2 = './files/CMA CGM LITANI 4525 - AVAILABLE.xlsm';
fs.unlinkSync(filePath2)

fs.copyFileSync('./files/CMA CGM LITANI 4525 - AVAILABLE - Copy.xlsm', './files/CMA CGM LITANI 4525 - AVAILABLE.xlsm')

const run = async _ => {
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
    console.log(json)

    const newWb = xlsx.utils.book_new();
    const newWs = xlsx.utils.json_to_sheet(json)

    xlsx.utils.book_append_sheet(newWb, newWs, "Compiled Data")
    xlsx.writeFile(newWb, "./files/Compilation.xlsx")

}

// run()