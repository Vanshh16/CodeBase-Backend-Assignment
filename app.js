var express = require("express");
var bodyParser = require("body-parser");
var ejs = require("ejs");
var mongoose = require("mongoose");
var multer = require("multer");
var excelToJson = require("convert-excel-to-json");
const upload = multer({ dest:"./public/uploads/" });
var async = require("async");
var eachSeries = require("async/eachSeries");

var app = express();

mongoose.connect("mongodb://127.0.0.1/excelDB",{useNewUrlParser:true});

app.set("view engine","ejs");

app.use(bodyParser.urlencoded({extended:true}));

app.use(express.static("public"));

var excelSchema = new mongoose.Schema({
    name: String,
    email: String,
    mobile: Number,
    dob: String,
    workExp: String,
    resumeTitle: String,
    curr_location: String,
    postal_add: String,
    curr_designation: String,
    curr_employer: String
});

const User = mongoose.model("User",excelSchema);

app.get("/", function(req, res){
  res.sendFile(__dirname + "/index.html");
});

app.post("/uploadfile", upload.single("uploadfile"), function(req, res){
    var flag = excelToMongo(__dirname + "/public/uploads/" + req.file.filename);
    if(flag===true){
      res.render("message", {message:"Success!"});
    }
    else{
      res.render("message", {message:"Error!"});
    }
});

const insertFun = async function(sheetData){
    await User.findOne({email: sheetData.Email }).then(async function(docs){
     if(docs===null)
     {
       console.log(docs);
       const user = new User({
           name: sheetData["Name of the Candidate"],
           email: sheetData.Email,
           mobile: sheetData['Mobile No.'],
           dob: sheetData['Date of Birth'],
           workExp: sheetData['Work Experience'],
           resumeTitle: sheetData['Resume Title'],
           curr_location: sheetData['Current Location'],
           postal_add: sheetData['Postal Address'],
           curr_designation: sheetData['Current Employer'],
           curr_employer: sheetData['Current Designation']
         });
        await user.save();
     }
 }).catch((err)=>{
     console.log(err);
 });
}

function excelToMongo(filePath){
    var flag = false;
    // Converting excel to json
    const excelData = excelToJson({
        sourceFile: filePath,
        sheets:[{
            name: "Sheet1",
            header:{
               rows: 1
            },
            // Mapping columns to keys
            columnToKey: {
                A: "Name of the Candidate",
                B: "Email",
                C: "Mobile No.",
                D: "Date of Birth",
                E: "Work Experience",
                F: "Resume Title",
                G: "Current Location",
                H: "Postal Address",
                I: "Current Employer",
                J: "Current Designation"
            }
        }]
    });
    var sheet = excelData.Sheet1;
    console.log(sheet);
    async.eachSeries(sheet, insertFun, function(err,results) {
    if( err ) {
        console.log(err);
    } else {
        console.log('All files have been added successfully');
        flag = true;
    }
 });
 return true;
}

app.listen(3000,function(){
  console.log("Server started on port 3000");
});
