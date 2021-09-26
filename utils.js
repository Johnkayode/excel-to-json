const readexcel = require('xlsx')

const excelFilter = (req, file, cb) => {
    if (
      file.mimetype.includes("excel") ||
      file.mimetype.includes("spreadsheetml")
    ) {
      cb(null, true);
    } else {
      cb(new Error('Not an Excel File'), false);
    }
};

const excelTojson = (file) => {
    var workbook = readexcel.read(file.buffer, {type:"buffer"});
    return readexcel.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]])
    
    
}


module.exports = {
    excelFilter: excelFilter,
    excelTojson: excelTojson
}