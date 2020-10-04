const fs = require('fs');

const reset = () => {
    const filePath = './files/Compilation.xlsx';
    fs.unlinkSync(filePath);

    const filePath2 = './files/CMA CGM LITANI 4525 - AVAILABLE.xlsm';
    fs.unlinkSync(filePath2)

    fs.copyFileSync('./files/CMA CGM LITANI 4525 - AVAILABLE - Copy.xlsm', './files/CMA CGM LITANI 4525 - AVAILABLE.xlsm')
}

reset()