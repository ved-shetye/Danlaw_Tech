import path from "path";
import fs from "fs-extra";
import fss from "fs";
import express from "express";
import bodyParser from "body-parser";
import multer from "multer";
("use strict");
import excelToJson from "convert-excel-to-json";
import mongoXlsx from "mongo-xlsx";
import mongoose from "mongoose";

const app = express();
// require('dotenv').config;
import "dotenv/config";
import connectDB from "./connectMongo.js";
connectDB();

const excelSchema = new mongoose.Schema({
  DATE: String,
  SHIFT: String,
  PRODUCT: String,
  TYPE: String,
  ACTIVITY: String,
  OPERATOR: String,
  QTY: Number,
});
const Excel = mongoose.model("Excel", excelSchema);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    return cb(null, "./uploads");
  },
  filename: function (req, file, cb) {
    return cb(null, "excelsheet.xlsx");
  },
});
const upload = multer({ storage: storage });

app.use(express.static("public"));

app.set("view engine", "ejs");
app.set("views", path.resolve("./views"));

app.use(express.urlencoded({ extended: false }));
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.render("home");
});

app.post("/", upload.single("excelsheet"), (req, res) => {
  // let doc = parser.parseXls2Json(req.file.path);
  Excel.deleteMany()
    .then((items) => {
      console.log("Deleted", items);
    })
    .catch((error) => {
      console.log(error);
    });

  const doc = excelToJson({
    sourceFile: "./uploads/excelsheet.xlsx",
    header: {
      rows: 1,
    },
    sheets: ["Assembly_Main"],
  });
  // console.log(Object.keys(doc["Assembly_Main"]).length);
  // res.send(doc["Assembly_Main"]);
  // console.log(doc["Assembly_Main"][0].G);
  Excel.find()
    .then(function (foundItems) {
      for (let i = 0; i < Object.keys(doc["Assembly_Main"]).length; i++) {
        let excel = new Excel({
          DATE: doc["Assembly_Main"][i].A,
          SHIFT: doc["Assembly_Main"][i].B,
          PRODUCT: doc["Assembly_Main"][i].C,
          TYPE: doc["Assembly_Main"][i].D,
          ACTIVITY: doc["Assembly_Main"][i].E,
          OPERATOR: doc["Assembly_Main"][i].F,
          QTY: doc["Assembly_Main"][i].G,
        });
        excel.save();
        //  console.log(doc[0][i].name);
      }
      console.log("Success");
    })
    .catch(function (error) {
      console.log(error);
    });
  // const pathToFile = "./uploads/excelsheet.xlsx";
  // fss.unlink(pathToFile, function (err) {
  //   if (err) {
  //     console.log(err);
  //   } else {
  //     // console.log("Successfully deleted excelsheet file.");
  //   }
  // });
  res.redirect("/");
});

app.post("/filterby", (req, res) => {
  const dd = req.body.DATE;
  const dv  = req.body.SHIFT;
  const dp = req.body.PRODUCT;
  const dt = req.body.TYPE;
  const da = req.body.ACTIVITY;
  const dop = req.body.OPERATOR;
  const pathToFile2 = "./downloads/file.xlsx";
  if (fss.existsSync(pathToFile2)) {
    fs.unlink("./downloads/file.xlsx", function (err) {
      if (err) console.log(err);
      // if no error, file has been deleted successfully
      console.log("File deleted!");
    });
  }
  const obj = {
    DATE:dd,
    SHIFT:dv,
    PRODUCT:dp,
    TYPE:dt,
    ACTIVITY:da,
    OPERATOR:dop
  }
  const parameters ={};
  for (let i in obj) {
if (obj[i]!=undefined) {
  const key = i;
  const value =obj[i];
  parameters[key]=value;
}
  }
  // console.log(obj);
  async function datef() {
    try {
      const totems = await Excel.find(parameters).select(
        "-_id -__v"
      );
      const number = await Excel.find(parameters).exec();
      const model = mongoXlsx.buildDynamicModel(totems);
      mongoXlsx.mongoData2Xlsx(totems, model, function (err, totems) {
        console.log("File saved at:", totems.fullPath);
        fs.move(totems.fullPath, "./downloads/file.xlsx")
          .then(() => {
            console.log("successfully shifted to downloads!");
          })
          .catch((err) => {
            console.error(err);
          });
      });
      res.render("download",{
         number:number
      });
    } catch (error) {
      console.log(error);
    }
  }
  const userd = datef();
});

app.get("/downloadFile", function (req, res) {
  console.log("Downloaded");
  res.download("./downloads/file.xlsx", function (err) {
    if (err) {
      console.log(err);
    }
  });
});

const PORT = process.env.PORT;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
