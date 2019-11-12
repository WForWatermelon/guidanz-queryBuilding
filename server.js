
const ExcelJS = require('exceljs');
//test git
const Excel = require('exceljs/lib/exceljs.nodejs');
const express = require('express');
var bodyParser = require('body-parser');
const swaggerJsDoc = require('swagger-jsdoc');
const swaggerui = require('swagger-ui-express');
var moment = require('moment');
var nodemon = require('nodemon');
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
      var ID2 = [];
      //var metaData = { "title": "data-table3", "type": "table", "params": { "dimensions": { "metrics": [{ "accessor": 0, "format": { "id": "number" }, "params": {}, "aggType": "avg" }], "buckets": [] }, "perPage": 10, "showMetricsAtAllLevels": false, "showPartialRows": false, "showTotal": false, "sort": { "columnIndex": null, "direction": null }, "totalFunc": "sum" }, "aggs": [{ "id": "3", "enabled": true, "type": "count", "schema": "metric", "params": {} }, { "id": "4", "enabled": true, "type": "terms", "schema": "bucket", "params": { "field": "age", "orderBy": "_key", "order": "desc", "size": 1000, "otherBucket": false, "otherBucketLabel": "Other", "missingBucket": false, "missingBucketLabel": "Missing" } }] }
      if (metaData.aggs.length == 1) {
         switch (metaData.aggs[0].type) {
            case "count":
               var body = {
                  aggs: {
                  }
               }
               break;
            case "median":
               var body = bodybuilder()
                  .aggregation('percentiles', metaData.aggs[0].params.field, {
                     percents: [50], keyed: false
                  }, metaData.aggs[0].id)
                  .build();
               break;

            case "std_dev":
               var body = bodybuilder()
                  .aggregation('extended_stats', metaData.aggs[0].params.field, metaData.aggs[0].id)
                  .build();
               break;

            case "percentiles":
               var body = bodybuilder()
                  .aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, {
                     percents: metaData.aggs[0].params.percents, keyed: false
                  }, metaData.aggs[0].id)
                  .build();
               break;

            case "percentile_ranks":
               var body = bodybuilder()
                  .aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, {
                     values: metaData.aggs[0].params.values, keyed: false
                  }, metaData.aggs[0].id)
                  .build();
               break;

            case 'top_hits':
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
               break;

            default://avg,max,min, sum,unique count
               var body = bodybuilder().aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)
                  .build();

         }
      }
      else {//aggs.length>1

         var aggs = {};
         aggs.order = {
         }
         switch (metaData.aggs[0].type) {
            case 'count':
               if (metaData.aggs[1].type == 'terms') {
                  var orderBy = metaData.aggs[1].params.orderBy;
                  if (orderBy == 'custom') {
                     if (metaData.aggs[1].params.orderAgg.type == 'count') {
                        aggs.order['_' + metaData.aggs[0].type] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }).build();
                     }
                     else {
                        aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }, agg => agg.aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();

                     }
                  }
                  else if (orderBy == metaData.aggs[0].id) {
                     aggs.order['_' + metaData.aggs[0].type] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }).build();

                  }
                  else if (orderBy == "_key") {
                     aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }).build();
                  }

               }
               else {//if buckets!=terms
                  switch (metaData.aggs[1].type) {
                     case 'date_histogram':
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           calendar_interval: "1w",
                           time_zone: "Asia/Calcutta",
                           min_doc_count: metaData.aggs[1].params.min_doc_count
                        }).build();
                        break;

                     case 'date_range':
                        var ranges = metaData.aggs[1].params.ranges;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           ranges: ranges,
                           time_zone: "Asia/Calcutta"
                        }).build();
                        break;

                     case 'range':
                        var ranges = metaData.aggs[1].params.ranges;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           ranges: ranges,
                           keyed: 'true'
                        }).build();
                        break;

                     case 'geohash_grid':
                     case "geotile_grid":
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           precision: metaData.aggs[1].params.precision
                        }, agg => agg.aggregation('geo_centroid', metaData.aggs[1].params.field, metaData.aggs[1].id + '1')).build();
                        break;

                     case "histogram":
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           interval: metaData.aggs[1].params.interval,
                           min_doc_count: "1"
                        }).build();
                        break;
                     //two more cases to be added here namely significant termns and filters
                     default:
                        if (metaData.aggs[1].params.ipRangeType == 'fromTo') {
                           var ranges = metaData.aggs[1].params.ranges.fromTo;
                           var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                              ranges: ranges
                           }).build();
                        }
                        else {
                           var ranges = metaData.aggs[1].params.ranges.mask;
                           var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                              ranges: ranges
                           }).build();
                        }
                        break;
                  }
               }
               break;

            case 'avg':
            case 'max':
            case 'min':
            case 'sum':
            case 'cardinality':
               if (metaData.aggs[1].type == 'terms') {
                  var orderBy = metaData.aggs[1].params.orderBy;
                  if (orderBy == 'custom') {
                     if (metaData.aggs[1].params.orderAgg.type == 'count') {
                        aggs.order['_' + metaData.aggs[1].params.orderAgg.type] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }, agg => agg.aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)).build();
                     }
                     else {
                        aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }, agg => agg.aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)
                           .aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();
                     }


                  }
                  else if (orderBy == metaData.aggs[0].id) {
                     aggs.order[metaData.aggs[0].id] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)).build();

                  }
                  else {
                     aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation(metaData.aggs[0].type, metaData.aggs[0].params.field, metaData.aggs[0].id)).build();
                  }

               }
               break;

            case 'std_dev':
               if (metaData.aggs[1].type == 'terms') {
                  var orderBy = metaData.aggs[1].params.orderBy;
                  if (orderBy == 'custom') {
                     if (metaData.aggs[1].params.orderAgg.type == 'count') {
                        aggs.order['_' + metaData.aggs[1].params.orderAgg.type] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }, agg => agg.aggregation('extended_stats', metaData.aggs[0].params.field, metaData.aggs[0].id)).build();
                     }
                     else {
                        aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                        var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                           order: aggs.order,
                           size: metaData.aggs[1].params.size
                        }, agg => agg.aggregation('extended_stats', metaData.aggs[0].params.field, metaData.aggs[0].id)
                           .aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();
                     }
                  }
                  else {
                     aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('extended_stats', metaData.aggs[0].params.field, metaData.aggs[0].id))
                        .build();
                  }
               }
               break;

            case 'median':
            case 'percentiles':
               var orderBy = metaData.aggs[1].params.orderBy;
               if (orderBy == 'custom') {
                  if (metaData.aggs[1].params.orderAgg.type == 'count') {
                     aggs.order['_' + metaData.aggs[1].params.orderAgg.type] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('percentiles', metaData.aggs[0].params.field, {
                        percents: metaData.aggs[0].params.percents, keyed: false
                     }, metaData.aggs[0].id)).build();
                  }
                  else {
                     aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('percentiles', metaData.aggs[0].params.field, {
                        percents: metaData.aggs[0].params.percents, keyed: false
                     }, metaData.aggs[0].id)
                        .aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();
                  }
               }
               else {
                  aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
                  var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                     order: aggs.order,
                     size: metaData.aggs[1].params.size
                  }, agg => agg.aggregation('percentiles', metaData.aggs[0].params.field, {
                     percents: metaData.aggs[0].params.percents, keyed: false
                  }, metaData.aggs[0].id)).build();
               }
               break;

            case 'percentile_ranks':
               var orderBy = metaData.aggs[1].params.orderBy;
               if (orderBy == 'custom') {
                  if (metaData.aggs[1].params.orderAgg.type == 'count') {
                     aggs.order['_' + metaData.aggs[1].params.orderAgg.type] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('percentile_ranks', metaData.aggs[0].params.field, {
                        values: metaData.aggs[0].params.values, keyed: false
                     }, metaData.aggs[0].id)).build();
                  }
                  else {
                     aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('percentile_ranks', metaData.aggs[0].params.field, {
                        values: metaData.aggs[0].params.values, keyed: false
                     }, metaData.aggs[0].id)
                        .aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();
                  }
               }
               else {
                  // if (metaData.aggs[1].params.orderBy == '_key') {
                  aggs.order[metaData.aggs[1].params.orderBy] = metaData.aggs[1].params.order;
                  var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                     order: aggs.order,
                     size: metaData.aggs[1].params.size
                  }, agg => agg.aggregation('percentile_ranks', metaData.aggs[0].params.field, {
                     values: metaData.aggs[0].params.values, keyed: false
                  }, metaData.aggs[0].id)).build();
               }
               break;

            case 'top_hits':
               var aggs1 = {};
               var sortField = metaData.aggs[0].params.sortField
               aggs1[sortField] = { order: metaData.aggs[0].params.sortOrder }
               var orderBy = metaData.aggs[1].params.orderBy;
               if (orderBy == 'custom') {
                  if (metaData.aggs[1].params.orderAgg.type == 'count') {
                     aggs.order['_' + metaData.aggs[1].params.orderAgg.type] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('top_hits', {
                        docvalue_fields: [{
                           field: metaData.aggs[0].params.field,
                           format: 'use_field_mapping'
                        }],
                        _source: metaData.aggs[0].params.field,
                        size: metaData.aggs[0].params.size,
                        sort: [
                           aggs1
                        ]
                     }, metaData.aggs[0].id))
                        .build();
                  }
                  else {
                     aggs.order[metaData.aggs[1].params.orderAgg.id] = metaData.aggs[1].params.order;
                     var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                        order: aggs.order,
                        size: metaData.aggs[1].params.size
                     }, agg => agg.aggregation('top_hits', {
                        docvalue_fields: [{
                           field: metaData.aggs[0].params.field,
                           format: 'use_field_mapping'
                        }],
                        _source: metaData.aggs[0].params.field,
                        size: metaData.aggs[0].params.size,
                        sort: [
                           aggs1
                        ]
                     }, metaData.aggs[0].id).aggregation(metaData.aggs[1].params.orderAgg.type, metaData.aggs[1].params.orderAgg.params.field, metaData.aggs[1].params.orderAgg.id)).build();
                  }
               }
               else {
                  // if (metaData.aggs[1].params.orderBy == '_key') {
                  aggs.order[metaData.aggs[1].params.field] = metaData.aggs[1].params.order;
                  var body = bodybuilder().aggregation(metaData.aggs[1].type, metaData.aggs[1].params.field, metaData.aggs[1].id, {
                     order: aggs.order,
                     size: metaData.aggs[1].params.size
                  }, agg => agg.aggregation('top_hits', {
                     docvalue_fields: [{
                        field: metaData.aggs[0].params.field,
                        format: 'use_field_mapping'
                     }],
                     _source: metaData.aggs[0].params.field,
                     size: metaData.aggs[0].params.size,
                     sort: [
                        aggs1
                     ]
                  }, metaData.aggs[0].id))
                     .build();
               }
               break;
            default:
               console.log('iieieieieieieie');
         }
      }


      var query2 = {
         body: {
            size: 0,
            track_total_hits: true,
            query: {
               bool: {
                  must: [
                     {
                        range: {
                           Timestamp: {
                              format: "strict_date_optional_time",
                              gte: "2018-10-23T06:14:21.272Z",
                              lte: "2019-11-04T06:14:21.272Z"
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
                  else if (type1 == 'percentiles') {
                     if (metaData.aggs[0].type == 'median') {
                        var ID = converter.toOrdinal(metaData.aggs[0].params.percents) + ' percentile of ' + metaData.aggs[0].params.field
                     }
                     else {
                        var ID = metaData.aggs[0].params.percents;
                     }
                  }
                  else if (type1 == 'percentile_ranks') {
                     var ID = metaData.aggs[0].params.values;
                  }
                  else if (type1 == 'top_hits') {
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
                     switch (metaData.aggs[0].type) {
                        case 'count':
                           ID1.push('Count');
                           break;

                        case 'avg':
                           ID1.push('Average ' + metaData.aggs[0].params.field);
                           break;

                        case 'sum':
                           ID1.push('Sum of ' + metaData.aggs[0].params.field);
                           break;

                        case 'max':
                           ID1.push('Max ' + metaData.aggs[0].params.field);
                           break;

                        case 'min':
                           ID1.push('Min ' + metaData.aggs[0].params.field);
                           break;

                        case 'cardinality':
                           ID1.push('Unique count of ' + metaData.aggs[0].params.field);
                           break;
                        case 'std_dev':

                           ID1.push('Lower standard Deviation of ' + metaData.aggs[0].params.field);
                           ID1.push('Upper standard Deviation of ' + metaData.aggs[0].params.field);
                           break;
                        case 'median':
                           ID2 = converter.toOrdinal(metaData.aggs[0].params.percents) + ' percentile of ' + metaData.aggs[0].params.field
                           ID1.push(ID2);
                           break;
                        case 'percentiles':
                           ID2 = metaData.aggs[0].params.percents;
                           break;
                        case 'percentile_ranks':
                           ID1.push('Percentile rank ' + metaData.aggs[0].params.values + ' of ' + metaData.aggs[0].params.field);
                           break;
                        case 'top_hits':
                           ID1.push('Last ' + metaData.aggs[0].params.field);
                        default:
                           console.log('hahahah');
                     }
                  }
                  else if (type1 == "date_histogram") {
                     ID1.push(metaData.aggs[1].params.field + " per week");
                     ID1.push('Count');
                  }
                  else if (type1 == 'date_range') {
                     ID1.push(metaData.aggs[1].params.field + " per ranges");
                     ID1.push('Count');
                  }
                  else if (type1 == 'range') {
                     ID1.push(metaData.aggs[1].params.field + " range");
                     ID1.push('Count');
                  }
                  else if (type1 == 'geohash_grid' || type1 == 'geotile_grid') {
                     ID1.push(metaData.aggs[1].type);
                     ID1.push('Count');
                  }
                  else if (type1 == 'histogram') {
                     ID1.push(metaData.aggs[1].params.field);
                     ID1.push('Count');
                  }
                  else if (type1 == 'ip_range') {
                     ID1.push(metaData.aggs[1].params.field + ' IP ranges');
                     ID1.push('Count');
                  }

                  else {
                     var ID = type1 + ' ' + metaData.aggs[0].params.field;
                  }

               }
            }

            if (ID1.length == 0) {//if only metrics present
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
                     var _date = moment(b.value).format('YYYY MM DD, h:mm:ss a');
                     sheet.addRow([_date]);
                  }
               }
               else if (metaData.aggs[0].type == 'percentiles') {

                  for (let i = 0; i < b.values.length; i++) {
                     ID2.push(b.values[i].value);
                  }
                  sheet.addRow(ID2);
               }
               else if (metaData.aggs[0].type == 'percentile_ranks') {
                  for (let i = 0; i < b.values.length; i++) {
                     ID2.push(b.values[i].value + "%");
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
            else {//if buckets also present
               let colArr = [];
               for (let i in ID1) {
                  colArr.push({ header: ID1[i], key: ID1[i], width: 30 });
               }
               if (ID2.length) {
                  for (let i in ID2) {
                     colArr.push({ header: converter.toOrdinal(ID2[i]) + ' percentile of ' + metaData.aggs[0].params.field, key: converter.toOrdinal(ID2[i]), width: 30 });
                  }
               }
               sheet.columns = colArr;
               for (const key in resp.aggregations) {
                  var b = resp.aggregations[key];
               }
               if (metaData.aggs[0].type == 'std_dev') {
                  if (metaData.aggs.length == 1) {
                     sheet.addRow([b.std_deviation_bounds.lower, b.std_deviation_bounds.upper]);
                  }
                  else {
                     var bucket_arr = {};
                     var lower_values = [];
                     var upper_values = [];
                     var keys = [];
                     var c = function (bucket_arr) {
                        if (bucket_arr.length == 0) {
                           return 'stop';
                        }
                        else {
                           var id = metaData.aggs[0].id;
                           if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                              var _date = moment(b.buckets[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                              upper_values.push(bucket_arr[0][id].std_deviation_bounds.upper);
                              lower_values.push(bucket_arr[0][id].std_deviation_bounds.lower);
                              keys.push(_date);

                           }
                           else {
                              var id = metaData.aggs[0].id;
                              upper_values.push(bucket_arr[0][id].std_deviation_bounds.upper);
                              lower_values.push(bucket_arr[0][id].std_deviation_bounds.lower);
                              keys.push(bucket_arr[0].key);
                           }
                           bucket_arr.shift();
                           c(bucket_arr);

                        }

                     }
                     c(b.buckets);
                     for (var i = 0; i < lower_values.length; i++) {
                        sheet.addRow([keys[i], lower_values[i], upper_values[i]]);
                     }

                  }

               }
               else {
                  switch (metaData.aggs[0].type) {
                     case 'count':
                        if (metaData.aggs[1].type == 'geohash_grid' || metaData.aggs[1].type == 'geotile_grid') {
                           for (key in b) { var c = b[key]; }
                           for (var i = 0; i < c.length; i++) {
                              sheet.addRow([c[i].key, c[i].doc_count]);
                           }
                        }
                        else if (metaData.aggs[1].type == "range") {
                           console.log(metaData.aggs[1].type);
                           for (key in b) {
                              var c = b[key];
                              for (key in c) {
                                 sheet.addRow([c[key].from + ' to ' + c[key].to, c[key].doc_count]);
                              }
                           }
                        }
                        else {//everything other than range, geohash_grid and geotilegrid

                           for (var i = 0; i < b.buckets.length; i++) {
                              if (b.buckets[i].key == 1) {
                                 sheet.addRow(['true', b.buckets[i].doc_count]);
                              }//terms=flag
                              else if (b.buckets[i].key == 0) {
                                 sheet.addRow(['false', b.buckets[i].doc_count])
                              }//terms=flag
                              else {
                                 //terms=date based
                                 if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                    if (metaData.aggs[1].type == "date_range") {
                                       var from = moment(b.buckets[i].key.from).format('MMM D, YYYY @ H:mm:ss.SSS')
                                       var to = moment(b.buckets[i].key.to).format('MMM D, YYYY @ H:mm:ss.SSS')
                                       var _date = from + ' to ' + to;
                                    }
                                    else if (metaData.aggs[1].type == 'date_histogram') {
                                       var _date = moment(b.buckets[i].key).format('YYYY-MM-DD')
                                    }//type=date_histogram
                                    else {
                                       var _date = moment(b.buckets[i].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                    }//terms

                                    sheet.addRow([_date, b.buckets[i].doc_count]);
                                 }
                                 else {
                                    sheet.addRow([b.buckets[i].key, b.buckets[i].doc_count])
                                 }//field =string,IPV4 range


                              }

                           }
                        }
                        break;

                     case 'avg':
                     case 'max':
                     case 'sum':
                     case 'min':
                     case 'cardinality':
                        var bucket_arr = {};
                        var values = []
                        var keys = [];
                        var c = function (bucket_arr) {
                           if (bucket_arr.length == 0) {
                              return 'stop';
                           }
                           else {
                              var id = metaData.aggs[0].id;
                              if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                 var _date = moment(b.buckets[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                 values.push(bucket_arr[0][id].value);
                                 keys.push(_date);

                              }
                              else {
                                 var id = metaData.aggs[0].id;
                                 values.push(bucket_arr[0][id].value);
                                 keys.push(bucket_arr[0].key);
                              }
                              bucket_arr.shift();
                              c(bucket_arr);

                           }
                        }

                        for (var key in resp.aggregations) {
                           var b = resp.aggregations[key]
                        }
                        c(b.buckets);
                        for (var i = 0; i < values.length; i++) {
                           sheet.addRow([keys[i], values[i]]);
                        }
                        break;
                     case 'median':
                        var bucket_arr = {};
                        var values = []
                        var keys = [];
                        var c = function (bucket_arr) {
                           if (bucket_arr.length == 0) {
                              return 'stop';
                           }
                           else {
                              var id = metaData.aggs[0].id
                              if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                 var _date = moment(bucket_arr[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                 values.push(bucket_arr[0][id].values[0].value);
                                 keys.push(_date);

                              }
                              else {
                                 values.push(bucket_arr[0][id].values[0].value);
                                 keys.push(bucket_arr[0].key);
                              }
                              bucket_arr.shift();
                              c(bucket_arr);
                           }

                        }

                        for (var key in resp.aggregations) {
                           var b = resp.aggregations[key]
                        }
                        c(b.buckets);
                        for (var i = 0; i < values.length; i++) {
                           sheet.addRow([keys[i], values[i]]);
                        }
                        break;
                     case 'percentiles':
                        var bucket_arr = {};
                        var values = []
                        var keys = [];
                        var c = function (bucket_arr) {
                           if (bucket_arr.length == 0) {
                              return 'stop';
                           }
                           else {
                              var id = metaData.aggs[0].id;
                              if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                 var _date = moment(bucket_arr[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                 values.push(_date);
                                 for (var i = 0; i < bucket_arr[0][id].values.length; i++) {
                                    values.push(bucket_arr[0][id].values[i].value);
                                 }
                                 sheet.addRow(values);
                                 values.length = 0;
                                 keys.length = 0;
                                 bucket_arr.shift();
                                 c(bucket_arr);
                              }
                              else {
                                 values.push(bucket_arr[0].key);
                                 for (var i = 0; i < bucket_arr[0][id].values.length; i++) {
                                    values.push(bucket_arr[0][id].values[i].value);
                                 }
                                 sheet.addRow(values);
                                 values.length = 0;
                                 keys.length = 0;
                                 bucket_arr.shift();
                                 c(bucket_arr);
                              }

                           }

                        }

                        for (var key in resp.aggregations) {
                           var b = resp.aggregations[key]
                        }
                        c(b.buckets);
                        break;

                     case 'percentile_ranks':
                        var bucket_arr = {};
                        var values = []
                        var keys = [];
                        var c = function (bucket_arr) {
                           if (bucket_arr.length == 0) {
                              return 'stop';
                           }
                           else {
                              var id = metaData.aggs[0].id;
                              if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                 var _date = moment(b.buckets[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                 values.push(bucket_arr[0][id].values[0].value + '%');
                                 keys.push(_date);

                              }
                              else {
                                 values.push(bucket_arr[0][id].values[0].value + '%');
                                 keys.push(bucket_arr[0].key);
                              }
                              bucket_arr.shift();
                              c(bucket_arr);

                           }
                        }
                        for (var key in resp.aggregations) {
                           var b = resp.aggregations[key]
                        }
                        c(b.buckets);
                        for (var i = 0; i < values.length; i++) {
                           sheet.addRow([keys[i], values[i]]);
                        }
                        break;
                     case 'top_hits':
                        var bucket_arr = {};
                        var values = [];
                        var keys = [];
                        var c = function (bucket_arr) {
                           if (bucket_arr.length == 0) {
                              return 'stop';
                           }
                           else {
                              var id = metaData.aggs[0].id
                              var field = metaData.aggs[0].params.field;
                              if (metaData.aggs[1].params.field == 'endedTime' || metaData.aggs[1].params.field == 'HistoricTimestamp' || metaData.aggs[1].params.field == 'startedTime' || metaData.aggs[1].params.field == 'Timestamp') {
                                 var _date = moment(bucket_arr[0].key).format('MMM D, YYYY @ H:mm:ss.SSS');
                                 values.push(bucket_arr[0][id].hits.hits[0]._source[field]);
                                 keys.push(_date);

                              }
                              else {
                                 values.push(bucket_arr[0][id].hits.hits[0]._source[field]);
                                 keys.push(bucket_arr[0].key);
                              }
                              bucket_arr.shift();
                              c(bucket_arr);
                           }

                        }

                        for (var key in resp.aggregations) {
                           var b = resp.aggregations[key]
                        }
                        c(b.buckets);
                        for (var i = 0; i < values.length; i++) {
                           sheet.addRow([keys[i], values[i]]);
                        }
                        break;

                  }

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
 *          size:
 *             type: number
 */

/**
 * @swagger
 * /api/v1/basic/excel:
*   post:
 *     description: export as excel from elasticsearch without aggregation
 *     produces:
 *       - application/xlsx
 *     parameters:
 *       - name: body
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           $ref: '#/definitions/index'
 *           type: object
 *         properties:
 *           index_name:
 *             type: string
 *           size:
 *             type: number
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */

app.post('/api/v1/basic/excel', function (req, res) {
   get_ES_without_aggs(req.body.index_name, req.body.size, 'excel').then(result => {
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
   get_ES_without_aggs(req.body.index_name, req.body.size, 'csv').then(result => {
      console.log(req.body.index)
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./Reports/' + result.fileName);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
      }
   })
});



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
