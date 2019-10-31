const ExcelJS = require('exceljs');
const Excel = require('exceljs/modern.nodejs');
const Underscore = require('underscore')
var workbook = new Excel.Workbook();
var elasticsearch = require('elasticsearch/src/elasticsearch');
var elasticClient = new elasticsearch.Client({
   host: 'localhost:9200',
   log:"trace"
});
var start1 = new Date();
var simulateTime = 1000
console.log('start time----->', start1);

elasticClient.search({
   index: ".http*",
   track_total_hits:true,
   size: 50
});
