const express = require("express");
const app = express();
const bodyParser = require("body-parser");
var xlsx = require("xlsx");
const upload = require('express-fileupload');
const Excel = require('exceljs');


app.use(upload());

//Kay3awni njbed les éléments diol body b7al chkel diyal Jquery selector
app.use(bodyParser.urlencoded({ extended: true }));

// Bach manb9ach nkteb dima xxx.ejs
app.set("view engine", "ejs");

// add CSS
app.use(express.static(__dirname + "/public"));

app.get("/", (req,res) =>{
    res.render("index.ejs");
})

function getref(){
    var wb = xlsx.readFile('./uploads/WI PC AVD - steps.xlsx');
    var ws = wb.Sheets["ref"];
    var reference = [];
    for (x in ws) {
        if(x=='!ref' || x=='!margins'){
            x++;
        } else {
            reference.push((ws[x]['v']));
        }
    }
    // console.log(reference);
    return reference;
}

function postref(reference){
    console.log(reference.length);
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile('./uploads/WI PC AVD - steps.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet("Static Work Instruction (WI)");
        var row = worksheet.getRow(10); // les lignes
        for(var i=0; i<reference.length;i++){
            row.getCell(i+24).value = reference[i]; // A5's value set to 5 columnes
            row.commit();
        }
        
        return workbook.xlsx.writeFile('new.xlsx');
    })
}

function getfil(){
    var wb = xlsx.readFile('./uploads/WI PC AVD - steps.xlsx');
    var ws = wb.Sheets['WI Step 0 To fullfil'];
    var fils = [];
    for (x in ws) {
        for(var i=13; i<86;i++){
            if(x=='!ref' || x=='!margins' || x=='POSTE DE TRAVAIL:' || x=='Fil'){
                x++;
            } 
            else if (x==`K${i}`){
                fils.push((ws[x]['v']));
            }
        }
    }
    // console.log(fils);
    return fils;
}

async function getOptionsFromDataFils(i){
    var workbook = new Excel.Workbook();
    var option = [];
    await workbook.xlsx.readFile('./uploads/WI PC AVD - steps.xlsx')
        var worksheet = workbook.getWorksheet('Data Fils ');
        const row = worksheet.getRow(i);
        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
            if(colNumber >= 23){
                var value = cell.value;
                option.push(value);
            }
        });
    return option;
}

function RechercheFils(Fils, reference){
    var workbook = new Excel.Workbook();
    var option = [];
    workbook.xlsx.readFile('./uploads/WI PC AVD - steps.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet('Data Fils ');
    })    

}

app.post('/upload', function(req, res) {
    let sampleFile;
    let uploadPath;
  
    if (!req.files || Object.keys(req.files).length === 0) {
      return res.status(400).send('No files were uploaded.');
    }
  
    // The name of the input field (i.e. "sampleFile") is used to retrieve the uploaded file
    sampleFile = req.files.sampleFile;
    uploadPath = __dirname + '/uploads/' + sampleFile.name;
  
    // Use the mv() method to place the file somewhere on your server
    sampleFile.mv(uploadPath, function(err) {
      if (err)
        return res.status(500).send(err);
  
      res.send('File uploaded!');
    });

    // var reference = getref();
    // postref(reference);
    // var Fils = getfil();
    var h;
    const result = getOptionsFromDataFils(7);
    result.then(res => {
    h = res;
    })
    console.log(h);
    // RechercheFils(Fils, reference)
  });





const PORT = process.env.PORT || 3000;

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
