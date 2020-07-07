const express = require("express"); //step 01 : import express module
const router = express.Router(); //step 02 : create router object
const app = express(); //step 03 : create express app
const bodyParser = require("body-parser");
const xl = require("excel4node");

const wb = new xl.Workbook();
const ws = wb.addWorksheet("Sheet 1");
const ws2 = wb.addWorksheet("Sheet 2");

app.use("/", router);

router.use(
  bodyParser.urlencoded({
    extended: true,
  })
);

router.use(
  bodyParser.json({
    extended: true,
  })
);

router.get("/", function (req, res) {
  console.log(res);
  res.sendFile("submit.html", { root: __dirname });
});

router.post("/submit", function (req, res) {
  console.log(req.body.serialNumber);

  res.sendFile("submit.html", { root: __dirname });

  var style = wb.createStyle({
    font: {
      color: "#FF0800",
      size: 12,
    },
    numberFormat: "$#,##0.00; ($#,##0.00); -",
  });

  ws.cell(1, 1).number(Number(req.body.serialNumber)).style(style);

  wb.write("Excel.xlsx");
});

app.listen(8080, function () {
  //step 04 : make app listen via an specific port.
  console.log("express server is up! : http://localhost:8080");
});
