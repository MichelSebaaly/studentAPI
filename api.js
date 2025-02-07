const express = require("express");
const app = express();
const cors = require("cors");
const XLSX = require("xlsx");
const fs = require("fs");

const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(cors());

app.get("/", function (req, res) {
  res.send("Hello World");
});

app.get("/students", (req, res) => {
  const fileBuffer = fs.readFileSync("data.xlsx");
  const workbook = XLSX.read(fileBuffer, { type: "buffer" });

  const sheetName = workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(sheet);

  console.log(data);
  res.send(data);
});

app.post("/createStudent", function (req, res) {
  // const filePath = "data.xlsx";
  // const workbook = XLSX.readFile(filePath);

  // const sheetName = workbook.SheetNames[0];
  // const sheet = workbook.Sheets[sheetName];

  // let data = XLSX.utils.sheet_to_json(sheet);
  // // console.log(req.body);
  // data.push(req.body);

  // const updatedSheet = XLSX.utils.json_to_sheet(data);

  // workbook.Sheets[sheetName] = updatedSheet;

  // XLSX.writeFile(workbook, filePath);

  // console.log("Record added successfully!");

  // res.send("OK");
  const filePath = "data.xlsx";
  const workbook = XLSX.readFile(filePath);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  let data = XLSX.utils.sheet_to_json(sheet);
  console.log("Before Update:", data);

  data.push(req.body);

  console.log("After Update:", data);

  const updatedSheet = XLSX.utils.json_to_sheet(data);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, sheetName);

  // Write file using fs to ensure proper saving
  const buffer = XLSX.write(newWorkbook, { bookType: "xlsx", type: "buffer" });
  fs.writeFileSync(filePath, buffer);

  console.log("Record added successfully!");
  res.send("OK");
});

app.listen(3000);
