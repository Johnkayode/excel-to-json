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
            json_response = utils.excelTojson(req.file)
            res.send(json_response)
        }
    })
})



app.listen(3000)