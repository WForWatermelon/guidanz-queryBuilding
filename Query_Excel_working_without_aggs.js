const ExcelJS = require('exceljs');
const Excel = require('exceljs/modern.nodejs');
const Underscore = require('underscore')
var workbook = new Excel.Workbook();
var elasticsearch = require('elasticsearch/src/elasticsearch');
const    build  = require("json-schema-to-es-mapping");
var elasticClient = new elasticsearch.Client({
   host: 'localhost:9200'
   //log:"trace"
});
var start1 = new Date();
var simulateTime = 1000
console.log('start time----->', start1);

elasticClient.search({
   index: ".data*",
   size:5000
}).then(function (resp) {
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
//    elasticClient.indices.getMapping({  
//       index: '.data*'
      
//     },
//   function (error,response) {  
//       if (error){
//         console.log(error.message);
//       }
//       else {
//          console.log(response)
//         // console.log("Mappings:\n",response.gov.mappings.constituencies.properties);
//         //console.log("Mappings:\n",JSON.stringify(response));
//         for(res in response){
//             var b=response[res];
//             var c=Object.entries(b.mappings.properties);
//             //const mappings = buildMappingsFor("people", schema);
//             console.log({ mappings });

         
//             if(c.type=="long" ||c.type== "integer" || c.type== "short"|| c.type== "byte"|| c.type== " double"|| c.type== " float"|| c.type== " half_float"|| c.type== " scaled_float"){
//                console.log(c)
//          }
         
         

//       }
//    }
//   });
   //console.log('iddddddddddddd', ID, ID.length)
   var start3 = new Date();
   // console.log('22222222222', (start3 - startResponse) / 1000)
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
   var end1 = new Date() - start1;
   console.log("Execution time from start till adding:", end1 / 1000);
   var start2 = new Date();

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

   workbook.xlsx.writeFile('./sample_bool_query1.xlsx').then(function () {
      var end2 = new Date() - start2;
      console.log("Execution time:", end2 / 1000);
      var endFinal = new Date() - start1;
      console.log("Execution time:", endFinal / 1000);
      // workbook = null, sheet = null;
   })
});
