const express = require('express');
const { json } = require('express/lib/response');
const excelToJson = require('convert-excel-to-json');
const JSZip = require("jszip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path")
const multer = require("multer")
const download = require('download');
const FileSaver = require('file-saver');
const app = express();
const port = 3000;

// View Engine Setup
app.set("views",path.join(__dirname,"views"))
app.set("view engine","ejs")

    
// var upload = multer({ dest: "Upload_folder_name" })
// If you do not want to use diskStorage then uncomment it
    
var storage = multer.diskStorage({
  destination: function (req, file, cb) {

      // Uploads is the Upload_folder_name
      cb(null, "uploads")
  },
  filename: function (req, file, cb) {
    //Renaming original file name
    //cb(null, Date.now() + "-" + file.originalname)
    cb(null,"a.xlsx")
  }
})
  
// Define the maximum size for uploading
// picture i.e. 1 MB. it is optional
const maxSize = 10 * 1000 * 1000;
    
var upload = multer({ 
    storage: storage,
    limits: { fileSize: maxSize },
    fileFilter: function (req, file, cb){
    
        // Set the filetypes, it is optional
        var filetypes = /xls|xlsx|application\/vnd.openxmlformats-officedocument.spreadsheetml.sheet/;
        var mimetype = filetypes.test(file.mimetype);
  
        var extname = filetypes.test(path.extname(
                    file.originalname).toLowerCase());
        
        if (mimetype && extname) {
            return cb(null, true);
        }
      
        cb("Error: File upload only supports the "
                + "following filetypes - " + filetypes);
      } 
  
// mypic is the name of file attribute
}).single("fileObj");   

app.get('/', (req, res) => {
  res.render("uploadPage");
});

app.post("/uploadXlsFile", function (req, res, next) {
  upload(req,res,function(err){
    if(err) {
      res.send(err)
    } 
    else {
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


      var generatedList = [];
      var zipFolder = new JSZip();

      // Run generation
      fs.readdir(
        path.resolve(__dirname, 'output'),
        (err, files) => {
          if (err) throw err;
          for (let file of files) {
            generatedList.push(file);

            var fileContent = fs.readFileSync(
                path.resolve(__dirname, "output/" + file),
                "binary"
            );
            zipFolder.file(file, fileContent);
          }
          //zipFolder.generate
          var buff = zipFolder.generate({type:"nodebuffer"});
          fs.writeFileSync(path.resolve(__dirname, `output/DiplomeGenerate.zip`), buff)
          //saveAs(t, "Diplome Generate.zip");
          
          //console.log(generatedList);
          res.render('finishPage', {data: generatedList});
        }
      );
    }
  })
})

app.get("/download", (req, res) => {
  res.download(path.resolve(__dirname, `output/DiplomeGenerate.zip`));
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}!`)
});


