const express = require('express')
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const port = 3000;


app.use(cors());

// Configuring body parser middleware
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

app.post('/getjson', (req, res) => {
        const xlsx = require('xlsx')
        const workbook = xlsx.readFile(req.body.excellocation)
        let worksheet = workbook.Sheets[req.body.sheetname]
        let json_value = xlsx.utils.sheet_to_json(worksheet)
        res.json(json_value)
    
});

app.listen(port, () => console.log(`Excel to JSON converter app listening on port ${port}!`));