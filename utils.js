const { json } = require('express');
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
    var json_response = []
    for (sheet of workbook.SheetNames){
      json_response = json_response.concat(readexcel.utils.sheet_to_json(workbook.Sheets[sheet]))
    }
    return json_response
    
    
}


module.exports = {
    excelFilter: excelFilter,
    excelTojson: excelTojson
}