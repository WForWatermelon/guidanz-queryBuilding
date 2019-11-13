const ExcelJS = require('exceljs');
const esUrl = "localhost:9200";
var moment = require('moment');
const Excel = require('exceljs/lib/exceljs.nodejs');
var get_ES_without_aggs = function (index, size, download_type) {
   console.log('hihihihi', index);

   return new Promise((resolve, reject) => {
      var workbook = new Excel.Workbook();
      var Underscore = require('underscore')
      var elasticsearch = require('elasticsearch/src/elasticsearch');
      var bodybuilder = require('bodybuilder');
      var converter = require('number-to-words');
      var elasticClient = new elasticsearch.Client({
         host: esUrl
      });
      var start1 = new Date();
      var simulateTime = 1000
      console.log('start time----->', start1);

      elasticClient.search(
         {
            index: index,
            size: size
         }
      ).then(function (resp) {
         console.log('Please wait while computing the execution time');
         workbook.creator = 'Me',
            workbook.lastModifiedBy = 'Him',
            workbook.created = new Date(2019, 10, 3),
            workbook.modified = new Date(),
            workbook.lastPrinted = new Date(2019, 10, 1),
            workbook.views = [
               {
                  x: 0, y: 0, width: 10000, height: 20000,
                  firstSheet: 0, activeTab: 1, visibility: 'visible'
               }
            ]
         var sheet = workbook.addWorksheet('My Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });
         sheet.pageSetup.margins = {
            left: 0.7, right: 0.7,
            top: 0.75, bottom: 0.75,
            header: 0.3, footer: 0.3
         };
         var j = 0;
         let ID = [];
         var lookup = {};
         var startResponse = new Date();
         //console.log('1111111111', (startResponse - start1) / 1000)
         for (let j = 0; j < resp.hits.hits.length; j++) {
            var ID1 = Underscore.map(
               Underscore.uniq(
                  Underscore.map(Object.keys(resp.hits.hits[j]._source), function (obj) {
                     if (ID.indexOf(obj) == -1)
                        ID.push(obj);
                     return JSON.stringify(obj);
                  })
               ), function (obj) {
                  return JSON.parse(obj);
               }
            );


         }
         //console.log('iddddddddddddd', ID, ID.length)
         var start3 = new Date();
         // console.log('22222222222', (start3 - start1) / 1000)
         let colArr = [];
         for (let i in ID) {
            colArr.push({ header: ID[i], key: ID[i], width: 30 });
         }
         sheet.columns = colArr;
         var start4 = new Date();
         //console.log('333333333', (start4 - start3) / 1000)
         for (let j = 0; j < resp.hits.hits.length; j++) {
            let element = resp.hits.hits[j];
            sheet.addRow(element._source);
         }
         //console.log(ID, ID.length);
         // var end1 = new Date() - start1;
         // console.log("Execution time from start till before adding to excel/csv sheet:", end1 / 1000);
         // var start2 = new Date();

         // const cell = sheet.row(5).cell(3);
         // console.log('1111111111111111111', cell)
         // sheet.getColumn('B').font = {
         //    name: 'Times new roman',
         //    color: { argb: 'black' },
         //    family: 2,
         //    size: 14,
         //    italic: true
         // };
         // sheet.getColumn('Z').fill = {
         //    type: 'pattern',
         //    pattern: 'darkVertical',
         //    fgColor: { argb: 'Red' }
         // };




         if (download_type == "excel") {

            var filename = 'sampleExcel' + '.xlsx';
            var fullpath = __dirname + '/Reports/' + filename;
            workbook.xlsx.writeFile('./Reports/' + filename).then(function () {
               var end2 = new Date();
               console.log("Total time taken:" + (end2 - start1) / 1000);
               resolve({ status: 'success', path: fullpath, fileName: filename });
            });
         }
         else {
            var filename = 'sampleCsv' + '.csv';
            var fullpath = __dirname + "/Reports/" + filename;
            workbook.csv.writeFile('./Reports/' + filename).then(function () {
               var end2 = new Date();
               console.log("Total time taken:" + (end2 - start1 / 1000));
               resolve({ status: 'success', path: fullpath, fileName: filename });
            });
         }
      });
   });
}

exports.get_ES_without_aggs = get_ES_without_aggs;