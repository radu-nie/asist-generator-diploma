const express = require('express');
const { json } = require('express/lib/response');
const excelToJson = require('convert-excel-to-json');
const JSZip = require("jszip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const app = express();
const port = 3000;

app.get('/', (req, res) => {
  res.send('Hello World!')
});

app.get('/generate', (req, res) => {
   // Load the docx file as binary content
  const content = fs.readFileSync(
      path.resolve(__dirname, "src/template.docx"),
      "binary"
  );
  
  //Convert xls to json
  const result = excelToJson({
    sourceFile: 'uploads/a.xlsx',
    header:{
      rows: 1
    }, 
    columnToKey: {
      A: 'numar_diploma',
      B: 'nume_cursant'
    }
  });
  
  //Save result in array of objects from the FIRST SHEET
  const resultFormatted = result[Object.keys(result)[0]];

  //Iterate for everty record in xls
  resultFormatted.forEach(individual => {
    // setup zip file
    var zip = new JSZip(content);

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
    doc.render({
        full_name: individual.nume_cursant,
        nr: individual.numar_diploma,
    });

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });

    // buf is a nodejs Buffer, you can either write it to a
    // file or res.send it with express for example.
    fs.writeFileSync(path.resolve(__dirname, `output/diploma_nr_${individual.numar_diploma}_${individual.nume_cursant}.docx`), buf);
  });

  //Respond with success
  res.send(resultFormatted);
});


app.listen(port, () => {
  console.log(`Example app listening on port ${port}!`)
});


