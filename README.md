# pdf-parse-and-execl-convert

```

const fs = require("fs");
const pdf = require("pdf-parse");
const excel = require("exceljs");

let dataBuffer = fs.readFileSync("a1.pdf");

exports.pdfExtratct = (req, res) => {
  pdf(dataBuffer).then(function (data) {
    // number of pages

       let extracted_text = data.text;

       // Step 3: Write the extracted text to an Excel file
       let workbook = new excel.Workbook();
       let worksheet = workbook.addWorksheet("MCQ");
       worksheet.columns = [
         { header: "Question", key: "question", width: 30 },
         { header: "Option A", key: "optionA", width: 20 },
         { header: "Option B", key: "optionB", width: 20 },
         { header: "Option C", key: "optionC", width: 20 },
         { header: "Option D", key: "optionD", width: 20 },
         { header: "Answer", key: "answer", width: 10 },
       ];

       // Step 4: Format the text into a table with columns for the question, options, and answer in the Excel file
       let lines = extracted_text.split("\n");
       for (let i = 0; i < lines.length; i++) {
         let line = lines[i];
         if (line.startsWith("Q.")) {
           let question = line.substring(2);
           let optionA = lines[i + 1];
           let optionB = lines[i + 2];
           let optionC = lines[i + 3];
           let optionD = lines[i + 4];
           let answer = lines[i + 5];
           worksheet.addRow({
             question,
             optionA,
             optionB,
             optionC,
             optionD,
             answer,
           });
         }
       }

       // Save the workbook to disk
       workbook.xlsx.writeFile("./MCQ.xlsx").then(function () {
         console.log("Excel file created!");
       });
    
    

    res.status(200).json({ editor: data.text });
  });
};

```
