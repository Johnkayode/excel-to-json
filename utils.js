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

const sheet_to_json = (sheet) => {

  const outputArray = []

 
  // Remove unused properties
  delete sheet['!margins']
  delete sheet['!ref']
  delete sheet['!rows']

  
  
  Object.keys(sheet).forEach(key=>{
    const [_, column, row] = key.match(/([a-zA-Z]+)(\d+)/)

    if (row !== '1') {
      const parent = column + '1';

      //We can get the array index we'll put it at by subtracting 2 from the row
      //Eg A2 => 0, A4 => 2, BB2 => 0
      const targetRow = parseInt(row) - 2;
      //Put an empty object in that row if there isn't one already
      if(!outputArray[targetRow]){
        outputArray[targetRow] = {}
      }

      //Then add our item to that row
      outputArray[targetRow][sheet[parent]['v']] = sheet[key]['v']

    }
  

    
  })
  
  
   
  return outputArray


}

const excelTojson = (file) => {
    var workbook = readexcel.read(file.buffer, {type:"buffer"});
    var json_response = []
    for (sheet of workbook.SheetNames){
      json_response = json_response.concat(sheet_to_json(workbook.Sheets[sheet]))
    }
    return json_response
    
    
}


module.exports = {
    excelFilter: excelFilter,
    excelTojson: excelTojson
}