var express = require('express');
var router = express.Router();
var mammoth = require("mammoth");
var nodeXlsx = require('node-xlsx');
var xlsx = require('xlsx');

/* GET home page. */
// router.get('/', function(req, res, next) {
// 	mammoth.convertToHtml({path: "W版集中交易场内业务外围接口文档(V5.5.3已启用)_6063.docx"})
// 		.then(function(result){
// 			var html = result.value; // The generated HTML
// 			var messages = result.messages; // Any messages, such as warnings during conversion
// 			res.render('index',{title:html});
// 		})
// 		.done();
// });

// router.get('/', function(req, res, next) {
// 	var obj = nodeXlsx.parse('渠道研发团队个人工作周报-陈彦霖.xlsx');
// 	console.log(obj[0]);
// 	res.render('index',{title:JSON.stringify(obj[0])});
// });
router.get('/', function(req, res, next) {
	var workbook = xlsx.readFile('test.xlsx');
	console.log(workbook);
	res.render('index', {title: JSON.stringify(workbook)});
})
module.exports = router;
