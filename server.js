const ExcelJS = require('exceljs');
const Excel = require('exceljs/modern.nodejs');
const express = require('express');
var bodyParser = require('body-parser');
const swaggerJsDoc = require('swagger-jsdoc');
const swaggerui = require('swagger-ui-express');
const esUrl = "localhost:9200";
const swaggerOptions = {
   swaggerDefinition: {
      info: {
         title: 'QueryBuilding',
         description: 'Query building index:".data*"from kibana'
      },
      servers: ["http://localhost:3035"]
   },
   apis: ["server.js"]
};
const swaggerDocs = swaggerJsDoc(swaggerOptions);
var path = require('path');
const app = express();
app.use(express.static(path.join(__dirname + './Reports')));
app.use('/', express.static('app', { redirect: false }));
app.use('/api-docs', swaggerui.serve, swaggerui.setup(swaggerDocs));


app.use(bodyParser.json({
   limit: '10mb',
   parameterLimit: 10000
}));
app.use(bodyParser.urlencoded({
   limit: '10mb',
   parameterLimit: 10000,
   extended: true
}));



/**
 * @swagger
 * /api/v1/agg/excel:
*   post:
 *     description: export as excel from elasticsearch with aggregation
 *     produces:
 *       - application/xlsx
 *     parameters:
 *       - name: id
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           type: object
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */

app.post('/api/v1/agg/excel', function (req, res) {
   get_Esquery(req.body, 'excel').then(result => {
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./Reports/' + result.fileName);

      }

   })

});

//Routes
/**
/**
 * @swagger
 * /api/v1/agg/csv:
 *   post:
 *     description: export as CSV from elasticsearch with aggregation
 *     produces:
 *       - application/csv
 *     parameters:
 *       - name: id
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           type: object
 *     responses:
 *       201:
 *         description: Created successfully
 */
app.post('/api/v1/agg/csv', function (req, res) {
   get_Esquery(req.body, 'csv').then(result => {
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./' + result.fileName);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
      }
   })
});



var get_Esquery = function (metaData, download_type) {
   return new Promise((resolve, reject) => {
      var workbook = new Excel.Workbook();
      var elasticsearch = require('elasticsearch/src/elasticsearch');
      var bodybuilder = require('bodybuilder');
      var converter = require('number-to-words');
      var elasticClient = new elasticsearch.Client({
         host: esUrl
      });

      //var metaData = { "title": "data-table3", "type": "table", "params": { "dimensions": { "metrics": [{ "accessor": 0, "format": { "id": "number" }, "params": {}, "aggType": "avg" }], "buckets": [] }, "perPage": 10, "showMetricsAtAllLevels": false, "showPartialRows": false, "showTotal": false, "sort": { "columnIndex": null, "direction": null }, "totalFunc": "sum" }, "aggs": [{ "id": "3", "enabled": true, "type": "count", "schema": "metric", "params": {} }, { "id": "4", "enabled": true, "type": "terms", "schema": "bucket", "params": { "field": "age", "orderBy": "_key", "order": "desc", "size": 1000, "otherBucket": false, "otherBucketLabel": "Other", "missingBucket": false, "missingBucketLabel": "Missing" } }] }

      if (metaData.aggs[0].type == "count") {
         if (metaData.aggs.length > 1) {

            var aggs = {};
            aggs.order = {
            }
            var orderBy = metaData.aggs[1].params.orderBy;
            if (orderBy == metaData.aggs[0].id) {
               aggs.order['_' + metaData.aggs[0].type] = metaData.aggs[1].params.order;

            }
            else {
               aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
            }
            var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
               order: aggs.order,
               size: metaData.aggs[1].params.size
            }).build();
         }
         else {
            var body = {
               aggs: {
               }
            }
         }
      }
      else if (metaData.aggs[0].type == "median") {
         var body = bodybuilder()
            .aggregation('percentiles', metaData.aggs[0].params.field, {
               percents: [50], keyed: false
            }, metaData.aggs[0].id)
            .build();
      }
      else if (metaData.aggs[0].type == "std_dev") {
         var body = bodybuilder()
            .aggregation('extended_stats', metaData.aggs[0].params.field, metaData.aggs[0].id)
            .build();
      }
      else if (metaData.aggs[0].type == "percentiles") {
         var body = bodybuilder()
            .aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, {
               percents: metaData.aggs[0].params.percents, keyed: false
            }, metaData.aggs[0].id)
            .build();
      }
      else if (metaData.aggs[0].type == "percentile_ranks") {
         var body = bodybuilder()
            .aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, {
               values: metaData.aggs[0].params.values, keyed: false
            }, metaData.aggs[0].id)
            .build();
      }
      else if (metaData.aggs[0].type == 'top_hits') {
         var aggs = {};
         var sortField = metaData.aggs[0].params.sortField
         aggs[sortField] = { order: metaData.aggs[0].params.sortOrder }
         var body = bodybuilder()
            .aggregation('top_hits', {
               docvalue_fields: [{
                  field: metaData.aggs[0].params.field,
                  format: 'use_field_mapping'
               }],
               _source: metaData.aggs[0].params.field,
               size: metaData.aggs[0].params.size,
               sort: [
                  aggs
               ]
            }, metaData.aggs[0].id)
            .build();
      }
      else {
         var body = bodybuilder().aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)
            .build();
      }
      var query2 = {
         body: {
            size: 100,
            track_total_hits: true,
            query: {
               bool: {
                  must: [
                     {
                        range: {
                           Timestamp: {
                              format: "strict_date_optional_time",
                              gte: "2018-10-23T06:14:21.272Z",
                              lte: "2019-10-23T06:14:21.272Z"
                           }
                        }
                     }
                  ],
                  filter: [
                     {
                        match_all: {}
                     }
                  ],
                  should: [],
                  must_not: []
               }
            },
            aggs: body.aggs
         }
      }
      console.log(JSON.stringify(query2));

      elasticClient.search(query2).then(function (resp) {
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
         var ID1 = [];
         if (Object.keys(query2.body.aggs).length == 0) {
            sheet.columns = [{ header: "Count", key: "Count", width: 20 }];
            sheet.addRow([resp.hits.total.value]);
         }
         else {
            for (const type in query2.body.aggs) {
               var a = query2.body.aggs[type];
               for (const type1 in a) {
                  if (type1 == 'avg') {
                     var ID = 'Average ' + metaData.aggs[0].params.field;
                  }
                  else if (metaData.aggs[0].type == 'median') {
                     var ID = converter.toOrdinal(metaData.aggs[0].params.percents) + ' percentile of ' + metaData.aggs[0].params.field
                  }
                  else if (metaData.aggs[0].type == 'percentiles') {
                     var ID = metaData.aggs[0].params.percents;
                  }
                  else if (metaData.aggs[0].type == 'percentile_ranks') {
                     var ID = metaData.aggs[0].params.values;
                  }
                  else if (metaData.aggs[0].type == 'top_hits') {
                     var ID = 'Last ' + metaData.aggs[0].params.size + ' ' + metaData.aggs[0].params.field;
                  }
                  else if (type1 == "sum") {
                     var ID = "Sum of " + metaData.aggs[0].params.field
                  }
                  else if (type1 == "cardinality") {
                     var ID = "Unique count of " + metaData.aggs[0].params.field
                  }
                  else if (type1 == 'extended_stats') {
                     ID1.push('Lower standard Deviation of ' + metaData.aggs[0].params.field);
                     ID1.push('Upper standard Deviation of ' + metaData.aggs[0].params.field);
                  }
                  else if (type1 == 'terms') {
                     if (metaData.aggs[1].params.order == "desc") {
                        ID1.push(metaData.aggs[1].params.field + ':Descending');
                     }
                     else {
                        ID1.push(metaData.aggs[1].params.field + ":Ascending");
                     }
                     ID1.push('Count');
                  }
                  else {
                     var ID = type1 + ' ' + metaData.aggs[0].params.field;
                  }

               }
            }
            if (ID1.length != 0) {
               let colArr = [];
               for (let i in ID1) {
                  colArr.push({ header: ID1[i], key: ID1[i], width: 30 });
               }
               sheet.columns = colArr;
               for (const key in resp.aggregations) {
                  var b = resp.aggregations[key];
               }
               if (metaData.aggs[0].type == 'std_dev') {
                  sheet.addRow([b.std_deviation_bounds.lower, b.std_deviation_bounds.upper])
               }
               else {
                  for (var i = 0; i < b.buckets.length; i++) {
                     if (b.buckets[i].key == 1) {
                        sheet.addRow(['true', b.buckets[i].doc_count]);
                     }
                     else if (b.buckets[i].key == 0) {
                        sheet.addRow(['false', b.buckets[i].doc_count])
                     }
                     else {
                        if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'historicTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                           var _date = JSON.stringify(new Date(b.buckets[i].key));
                           sheet.addRow([_date, b.buckets[i].doc_count]);
                        }
                        else {
                           sheet.addRow([b.buckets[i].key, b.buckets[i].doc_count])
                        }


                     }

                  }
               }
            }
            else {
               if (typeof ID == 'object' && ID.length) {
                  let colArr = [];
                  if (metaData.aggs[0].type == 'percentile_ranks') {
                     for (let i in ID) {
                        colArr.push({ header: 'Percentile rank ' + ID[i] + ' of ' + metaData.aggs[0].params.field, key: ID[i], width: 30 });
                     }
                  }
                  else {
                     for (let i in ID) {
                        colArr.push({ header: converter.toOrdinal(ID[i]) + ' percentile of ' + metaData.aggs[0].params.field, key: converter.toOrdinal(ID[i]), width: 30 });
                     }
                  }
                  sheet.columns = colArr;
               }
               else {
                  sheet.columns = [{ header: ID, key: ID, width: 30 }];
               }

               for (const key in resp.aggregations) {
                  var b = resp.aggregations[key];
               }

               if (metaData.aggs[0].params.field == 'endedTime' || metaData.aggs[0].params.field == 'historicTimestamp' || metaData.aggs[0].params.field == 'startedTime' || metaData.aggs[0].params.field == 'Timestamp') {
                  if (metaData.aggs[0].type == 'cardinality') {
                     sheet.addRow([b.value]);
                  }
                  else {
                     var _date = JSON.stringify(new Date(b.value));
                     sheet.addRow([_date]);
                  }
               }
               else if (metaData.aggs[0].type == 'percentiles') {
                  var ID2 = [];
                  for (let i = 0; i < b.values.length; i++) {
                     ID2.push(b.values[i].value);
                  }
                  sheet.addRow(ID2);
               }
               else if (metaData.aggs[0].type == 'percentile_ranks') {
                  var ID2 = [];
                  for (let i = 0; i < b.values.length; i++) {
                     ID2.push(b.values[i].value);
                  }
                  sheet.addRow(ID2);
               }
               else if (metaData.aggs[0].type == 'top_hits') {
                  var c = metaData.aggs[0].params.field;
                  let ID3 = [];
                  var element = b.hits;
                  for (let i = 0; i < element.hits.length; i++) {
                     ID3.push(Object.values(element.hits[i]._source));
                  }
                  sheet.addRow([ID3]);
               }

               else if (metaData.aggs[0].type == 'median') {
                  sheet.addRow([b.values[0].value])
               }
               else {
                  sheet.addRow([b.value])
               }
            }
         }
         var date_time = new Date();
         if (download_type == "excel") {
            var filename = metaData.title + '.xlsx';
            var fullpath = __dirname + "/Reports/" + filename;
            workbook.xlsx.writeFile('./Reports/' + filename).then(function () {
               resolve({ status: 'success', Timestamp: date_time, path: fullpath, fileName: filename });
            });
         }
         else {
            var filename = metaData.title + '.csv';
            var fullpath = __dirname + "/Reports/" + filename;
            workbook.csv.writeFile('./Reports/' + filename).then(function () {
               resolve({ status: 'success', Timestamp: date_time, path: fullpath, fileName: filename });
            });
         }
      }), function (err) {
         console.log("error")
         reject(err.message);
      };

   })
}


/**
 * @swagger
 * definitions:
 *    index:
 *       type: object
 *       properties:
 *          index_name:
 *             type: string
 */

/**
 * @swagger
 * /api/v1/basic/excel:
*   post:
 *     description: export as excel from elasticsearch without aggregation
 *     produces:
 *       - application/xlsx
 *     parameters:
 *       - name: id
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           $ref: '#/definitions/index'
 *           type: object
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */

app.post('/api/v1/basic/excel', function (req, res) {
   get_ES_without_aggs(req.body.index_name, 'excel').then(result => {
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./Reports/' + result.fileName);

      }
   })

});

/**
 * @swagger
 * /api/v1/basic/csv:
*   post:
 *     description: export as excel from elasticsearch without aggregation
 *     produces:
 *       - application/csv
 *     parameters:
 *       - name: id
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           $ref: '#/definitions/index'
 *           type: object
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */

app.post('/api/v1/basic/csv', function (req, res) {
   get_ES_without_aggs(req.body.index_name, 'csv').then(result => {
      console.log(req.body.index)
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./Reports/' + result.fileName);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
      }
   })
});



var get_ES_without_aggs = function (index, download_type) {
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
            index: "." + index + "*",
            size: 5000
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




         if (download_type == "excel") {

            var filename = 'sampleExcel' + '.xlsx';
            var fullpath = __dirname + '/Reports/' + filename;
            workbook.xlsx.writeFile('./Reports/' + filename).then(function () {
               resolve({ status: 'success', path: fullpath, fileName: filename });
            });
         }
         else {
            var filename = 'sampleCsv' + '.csv';
            var fullpath = __dirname + "/Reports/" + filename;
            workbook.csv.writeFile('./Reports/' + filename).then(function () {
               resolve({ status: 'success', path: fullpath, fileName: filename });
            });
         }
      });
   });
}
app.listen(3035);
