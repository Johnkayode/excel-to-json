const express = require('express')
const multer = require('multer')

const utils = require('./utils')




const app = express()
const upload = multer({
    storage: multer.memoryStorage(),
    fileFilter: utils.excelFilter 
})



var excel_upload = upload.single("xlsx")

app.post('/', (req, res) => {
    excel_upload(req, res, (err)=>{
        if (err) {
            res.send(err)
        }
        else {
            try{
                json_response = utils.excelTojson(req.file)
            }
            catch(err){
                json_response = {
                    'error': 'There was an error processing this file'
                }
            }
            res.send(json_response)
        }
    })
})



app.listen(3000)