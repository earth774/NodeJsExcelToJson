var express = require('express')
var app = express()
var xlsxtojson = require("xlsx-to-json");
var xlstojson = require("xls-to-json");

var mysql = require('mysql');
const util = require('util');

var con_commerce = mysql.createPool({
  host: "localhost",
  socketPath: '/Applications/MAMP/tmp/mysql/mysql.sock',
  user: "root",
  password: "root",
  database: "sampran_ecommerce_cloud",
  charset: 'utf8'
});

const query_commerce = util.promisify(con_commerce.query).bind(con_commerce);

app.use(function (req, res, next) { //allow cross origin requests
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "POST, PUT, OPTIONS, DELETE, GET");
  res.header("Access-Control-Max-Age", "3600");
  res.header("Access-Control-Allow-Headers", "Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");
  next();
});

// configuration
app.use(express.static(__dirname + '/public'));
app.use('/public/uploads', express.static(__dirname + '/public/uploads'));

app.get('/', function (req, res) {
  res.send('Hello World')
})

app.get('/api/xlstojson', function (req, res) {
  xlsxtojson({
    input: "./event.xlsx", // input xls 
    output: "output.json", // output json 
    lowerCaseHeaders: true
  }, (err, excel) => {
    if (err) {
      res.json(err);
    } else {
      excel.map(async (result, index) => {
        console.log(index)
        if (index >= 2 && result.name_th != "") {
          let image_path = (result.image_path != "") ? result.image_path.split(";") : "-";
          let info = [
            /*name_th:*/
            result.name_th,
            /*name_en:*/
            (result.name_en == "") ? "-" : result.name_en,
            /*image_path:*/
            (image_path[0] != undefined) ? image_path[0] : "-",
            /*detail_th:*/
            result.detail_th,
            /*detail_en:*/
            result.detail_en,
            /*activity_category_id:*/
            (result.course_type == "ประชุมกลุ่ม") ? 6 : 7,
            /*remark_th:*/
            ((result.speaker != "") ? ("ผู้บรรยาย " + result.speaker) + " ; " : "") + result.remark_th,
            /*location_th:*/
            result.location_th,
            /*location_en:*/
            "-",
            /*contact:*/
            result.contact,
            /*latitude:*/
            "-",
            /*longitude:*/
            "-",
            /*website:*/
            (result.website != "") ? result.website : ((result.facebook != "") ? result.facebook : "-"),
            /*open_time_th:*/
            result.start_date + ((result.start_date != "" && result.end_date != "") ? " - " : "") + result.end_date + "  " + result.start_time + ((result.start_time != "" && result.end_time != "") ? " - " : "") + result.end_time,
            /*open_time_en:*/
            result.start_date + ((result.start_date != "" && result.end_date != "") ? " - " : "") + result.end_date + "  " + result.start_time + ((result.start_time != "" && result.end_time != "") ? " - " : "") + result.end_time,
          ]

          let sql1 = "INSERT INTO activity (name_th,name_en,image_path,detail_th,detail_en,activity_category_id,remark_th,location_th,location_en,contact,latitude,longitude,website,open_time_th,open_time_en) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
          let activity = await query_commerce(sql1, info);
          console.log(activity.insertId)
          if (image_path.length > 1) {
            for (let data of image_path) {
              let sql2 = "INSERT INTO activity_image (image_path, activity_id) VALUES (?,?)";
              await query_commerce(sql2, [data, activity.insertId]);

            }
          }
        }
      })
      res.json(excel);

    }
  });
});

app.get('/api/xlstojson1', function (req, res) {
  xlsxtojson({
    input: "./event2.xlsx", // input xls 
    output: "output.json", // output json 
    lowerCaseHeaders: true
  }, (err, excel) => {
    if (err) {
      res.json(err);
    } else {
      excel.map(async (result, index) => {
        console.log(index)
        if (index >= 2 && result.name_th != "") {
          let image_path = (result.image_path != "") ? result.image_path.split(";") : "-";
          var search_index = result.google_map_location.indexOf("@");
          var res = result.google_map_location.substring(search_index + 1, 148);
          let split_data = res.split(",");
          let info = [
            /*name_th:*/
            result.name_th,
            /*name_en:*/
            (result.name_en == "") ? "-" : result.name_en,
            /*image_path:*/
            (image_path[0] != undefined) ? image_path[0] : "-",
            /*detail_th:*/
            (result.detail_th == "") ? "-" : result.detail_th,
            /*detail_en:*/
            (result.detail_en == "") ? "-" : result.detail_en,
            /*activity_category_id:*/
            8,
            /*remark_th:*/
            result.remark_th,
            /*location_th:*/
            result.location_th,
            /*location_en:*/
            "-",
            /*contact:*/
            result.contact,
            /*latitude:*/
            split_data[0],
            /*longitude:*/
            split_data[1],
            /*website:*/
            (result.website != "") ? result.website : ((result.facebook != "") ? result.facebook : "-"),
            /*open_time_th:*/
            "-",
            /*open_time_en:*/
            "-",
          ]

          let sql1 = "INSERT INTO activity (name_th,name_en,image_path,detail_th,detail_en,activity_category_id,remark_th,location_th,location_en,contact,latitude,longitude,website,open_time_th,open_time_en) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
          let activity = await query_commerce(sql1, info);
          console.log(activity.insertId)
          if (image_path.length > 1) {
            for (let data of image_path) {
              let sql2 = "INSERT INTO activity_image (image_path, activity_id) VALUES (?,?)";
              await query_commerce(sql2, [data, activity.insertId]);

            }
          }
        }
      })
      res.json(excel);

    }
  });
});
app.listen(3000)