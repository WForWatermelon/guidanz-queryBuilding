const ExcelJS = require('exceljs');
const Excel = require('exceljs/lib/exceljs.nodejs');
const esUrl = "localhost:9200";
var moment = require('moment');
var tuple = require('tuple');
var elasticsearch = require('elasticsearch/src/elasticsearch');
var bodybuilder = require('bodybuilder');
var converter = require('number-to-words');
var serializer = require('./response.serializer');
var get_Esquery = function (dashboard_name, metaData, workbook, timeField, download_type) {
   return new Promise((resolve, reject) => {

      var elasticClient = new elasticsearch.Client({
         host: esUrl
      });
      var status = false;
      var metrics = [];
      var buckets = {};
      var bucket_list = [];
      var body1 = [];
      var curr_bucket = {};
      for (let i in metaData.aggs) {
         if (metaData.aggs[i].schema == 'metric') {
            metrics.push(metaData.aggs[i]);
         }
         else {
            curr_bucket['aggs'] = metaData.aggs[i];
            if (Object.keys(buckets).length === 0) {
               buckets = curr_bucket;
            }
            curr_bucket = curr_bucket['aggs'];
            bucket_list.push(metaData.aggs[i]);
         }
      }
      var metrics_data = {
         aggs: {}
      };
      for (let i = 0; i < metrics.length; i++) {
         switch (metrics[i].type) {
            case 'count':
               break;
            case "median":
            case 'percentiles':
               var body = bodybuilder()
                  .aggregation('percentiles', metrics[i].params.field, {
                     percents: metrics[i].params.percents, keyed: false
                  }, metrics[i].id)
                  .build();
               break;

            case "std_dev":
               var body = bodybuilder()
                  .aggregation('extended_stats', metrics[i].params.field, metrics[i].id)
                  .build();
               break;


            case "percentile_ranks":
               var body = bodybuilder()
                  .aggregation(metrics[i].type, metrics[i].params.field, {
                     values: metrics[i].params.values, keyed: false
                  }, metrics[i].id)
                  .build();
               break;

            case 'top_hits':
               var aggs1 = {};
               var sortField = metrics[i].params.sortField
               aggs1[sortField] = { order: metrics[i].params.sortOrder }
               var body = bodybuilder()
                  .aggregation('top_hits', {
                     docvalue_fields: [{
                        field: metrics[i].params.field,
                        format: 'use_field_mapping'
                     }],
                     _source: metrics[i].params.field,
                     size: metrics[i].params.size,
                     sort: [
                        aggs1
                     ]
                  }, metrics[i].id)
                  .build();
               break;

            default://avg,max,min, sum,unique count
               var body = bodybuilder().aggregation(metrics[i].type, metrics[i].params.field, metrics[i].id)
                  .build();

         }
         metrics_data['aggs'] = Object.assign(metrics_data.aggs, body.aggs);
      }
      var last_id = (parseInt(bucket_list[bucket_list.length - 1].id) + 1).toString();

      var func = function (main_body, curr_body, buckets) {
         var curr_type = '';
         var type_body = {};
         var curr_id = ''
         var curr_body_of_curr_type = {};
         if (Object.keys(buckets).length == 0) {
            return main_body;
         }
         for (var key_iter in buckets) {
            var c = buckets[key_iter];
            var key = key_iter;
            switch (key) {
               case 'aggs':
                  curr_body['aggs'] = {};
                  if (Object.keys(main_body).length == 0) {
                     main_body = curr_body;
                  }
                  curr_body = curr_body['aggs'];
                  func(main_body, curr_body, buckets[key]);
                  break;

               case 'id':
                  curr_body[c] = {};
                  if (Object.keys(main_body).length == 0) {
                     main_body = curr_body;
                  }
                  curr_body = curr_body[c];
                  curr_id = c;
                  break;
               case 'type':
                  curr_type = c;
                  break;

               case 'params':
                  curr_type = buckets.type
                  switch (curr_type) {
                     case 'terms':
                        var aggs = {};
                        aggs.order = {
                        }
                        var orderBy = c.orderBy;
                        if (orderBy == 'custom') {
                           if (c.orderAgg.type == 'count') {
                              aggs.order['_' + c.orderAgg.type] = c.order;
                              curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                                 order: aggs.order,
                                 size: c.size
                              }).build().aggs[curr_id][curr_type];
                           }
                           else {
                              aggs.order[c.orderAgg.id] = c.order;
                              curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                                 order: aggs.order,
                                 size: c.size
                              }).build().aggs[curr_id][curr_type];
                              body1.push(bodybuilder().aggregation(c.orderAgg.type, c.orderAgg.params.field, c.orderAgg.id).build());
                           }

                        }
                        else if (orderBy == "_key") {
                           aggs.order[c.orderBy] = c.order;
                           curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                              order: aggs.order,
                              size: c.size
                           }).build().aggs[curr_id][curr_type];
                        }
                        else {
                           aggs.order[c.orderBy] = c.order;
                           curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                              order: aggs.order,
                              size: c.size
                           }).build().aggs[curr_id][curr_type];

                        }
                        break;
                     case 'date_histogram':
                        curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                           calendar_interval: "1w",
                           time_zone: "Asia/Calcutta",
                           min_doc_count: c.min_doc_count
                        }).build().aggs[curr_id][curr_type];

                        break;
                     case 'aggs':
                        curr_body['aggs'] = {};
                        if (Object.keys(main_body).length == 0) {
                           main_body = curr_body;
                        }
                        curr_body = curr_body['aggs'];
                        func(main_body, curr_body, buckets[key]);
                        break;

                     case 'date_range':
                        var ranges = c.ranges;
                        curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                           ranges: ranges,
                           time_zone: "Asia/Calcutta"
                        }).build().aggs[curr_id][curr_type];
                        break;

                     case 'range':
                        var ranges = c.ranges;
                        curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                           ranges: ranges,
                           keyed: 'true'
                        }).build().aggs[curr_id][curr_type];
                        break;
                     case 'geohash_grid':
                     case "geotile_grid":
                        status = true;
                        curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                           precision: c.precision
                        }).build().aggs[curr_id][curr_type];
                        body = bodybuilder().aggregation('geo_centroid', c.field, last_id).build();
                        break;

                     case "histogram":
                        curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                           interval: c.interval,
                           min_doc_count: "1"
                        }).build().aggs[curr_id][curr_type];
                        break;
                     //two more cases to be added here namely significant termns and filters
                     case 'ip_range':
                        if (c.ipRangeType == 'fromTo') {
                           var ranges = c.ranges.fromTo;
                           curr_body[curr_type] = bodybuilder().aggregation(curr_type, c.field, curr_id, {
                              ranges: ranges
                           }).build().aggs[curr_id][curr_type];
                        }
                        else {
                           var ranges = c.ranges.mask;
                           curr_body[curr_type] = bodybuilder().aggregation(curr_types, c.field, curr_id, {
                              ranges: ranges
                           }).build().aggs[curr_id][curr_type];
                        }
                        break;
                     default:
                        break;
                  }
                  break;

            }
         }
         return main_body;
      }
      var combine_metrics_buckets_helper = function construct_query_helper(main_body) {
         var final_body = {};
         for (var i in main_body) {
            if (i == 'aggs' || main_body[i].hasOwnProperty('aggs')) {
               return combine_metrics_buckets_helper(main_body[i]);
            }
            final_body = main_body[i];
         }
         return final_body;
      }

      var combine_custom_metrics = function (main_body) {
         var curr_body1 = main_body;
         var temp_body = curr_body1;
         for (var i = body1.length - 1; i >= 0; i--) {
            if (i == body1.length - 1) {
               console.log('body123', body1[i]);
               Object.assign(metrics_data.aggs, body1[i].aggs);
            }
            else {
               console.log('hoooooo');
               temp_body = curr_body1;
               for (var j in temp_body) {
                  if (j == 'aggs' || temp_body[j].hasOwnProperty('aggs')) {
                     Object.assign(temp_body[j], body1[i].aggs);
                     curr_body1 = temp_body
                     break;
                  }
               }
            }

         }
      }
      var combine_metrics_buckets = function construct_query(main_body) {
         if (status == true) {
            metrics_data['aggs'] = Object.assign(metrics_data.aggs, body.aggs);
         }
         if (body1.length != 0) {
            combine_custom_metrics(main_body);

         }
         var final_body = combine_metrics_buckets_helper(main_body);
         Object.assign(final_body, metrics_data);
      }
      var main_body = {};
      var curr_body = {};
      main_body = func(main_body, curr_body, buckets);

      if (Object.keys(main_body).length == 0) {
         main_body = metrics_data;
      }
      else {
         combine_metrics_buckets(main_body);
         for (var i = 0; i < body1.length; i = i + 2) {
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
                           [timeField]: {
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
            aggs: main_body.aggs
         }
      }
      console.log("\n\n\nQuery=======", JSON.stringify(query2));
      elasticClient.search(query2).then(function (resp) {
         var sheet = workbook.addWorksheet(metaData.title, { properties: { tabColor: { argb: 'FFC0000' } } });
         sheet.pageSetup.margins = {
            left: 0.7, right: 0.7,
            top: 0.75, bottom: 0.75,
            header: 0.3, footer: 0.3
         };
         var col_list = [];
         for (let i = 0; i < bucket_list.length; i++) {
            switch (bucket_list[i].type) {
               case 'terms':
                  if (bucket_list[i].params.order == "desc") {
                     col_list.push({ key: bucket_list[i].params.field + ':Descending', id: bucket_list[i].id, type: bucket_list[i].type });
                  }
                  else {
                     col_list.push({ key: bucket_list[i].params.field + ":Ascending", id: bucket_list[i].id, type: bucket_list[i].type });
                  }
                  break;

               case 'date_histogram':
                  col_list.push({ key: bucket_list[i].params.field + " per week", id: bucket_list[i].id, type: bucket_list[i].type });
                  break;

               case 'date_range':
                  col_list.push({ key: bucket_list[i].params.field + " per ranges", id: bucket_list[i].id, type: bucket_list[i].type });
                  break;

               case 'range':
                  col_list.push({ key: bucket_list[i].params.field + " range", id: bucket_list[i].id, type: bucket_list[i].type });
                  break;

               case 'geohash_grid':
               case 'geotile_grid':
                  col_list.push({ key: bucket_list[i].type, id: bucket_list[i].id, type: bucket_list[i].type });
                  break;

               case 'histogram':
                  col_list.push({ key: bucket_list[i].params.field, id: bucket_list[i].id, type: bucket_list[i].type });
                  break;

               case 'ip_range':
                  col_list.push({ key: bucket_list[i].params.field + ' IP ranges', id: bucket_list[i].id, type: bucket_list[i].type });
                  break;
               default:
                  console.log('hi look here');
                  break;
            }
         }
         for (let i = 0; i < metrics.length; i++) {
            switch (metrics[i].type) {
               case 'count':
                  col_list.push({ key: 'Count', id: metrics[i].id, type: bucket_list[i].type });
                  break;

               case 'avg':
                  col_list.push({ key: 'Average ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'sum':
                  col_list.push({ key: 'Sum of ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'max':
                  col_list.push({ key: 'Max ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'min':
                  col_list.push({ key: 'Min ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'cardinality':
                  col_list.push({ key: 'Unique count of ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'std_dev':
                  col_list.push({ key: 'Lower Standard Deviation of ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  col_list.push({ key: 'Upper Standard Deviation of ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'top_hits':
                  col_list.push({ key: 'Last ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'percentile_ranks':
                  col_list.push({ key: 'Percentile rank ' + metrics[i].params.values + ' of ' + metrics[i].params.field, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'median':
                  ID2 = converter.toOrdinal(metrics[i].params.percents) + ' percentile of ' + metrics[i].params.field;
                  col_list.push({ key: ID2, id: metrics[i].id, type: metrics[i].type });
                  break;

               case 'percentiles':
                  for (let j = 0; j < metrics[i].params.percents.length; j++) {
                     ID2 = converter.toOrdinal(metrics[i].params.percents[j]) + ' percentile of ' + metrics[i].params.field
                     col_list.push({ key: ID2, id: metrics[i].id, type: metrics[i].type });
                  }
                  break;

               default:
                  console.log('Look here');
            }
         }
         var colArr = [];
         var rowArr = [];
         for (let i in col_list) {
            colArr.push({ header: col_list[i].key, key: col_list[i].key, width: 30 });
         }
         sheet.columns = colArr;
         var result = serializer(resp);
         // console.log('collist:', col_list);
         for (let i = 0; i < result.length; i++) {
            for (let j = 0; j < col_list.length; j++) {
               if (col_list[j].key.match('Timestamp') || col_list[j].key.match('endedTime') || col_list[j].key.match('HistoricTimestamp') || col_list[j].key.match('startedTime') || col_list[j].key.match('@timestamp') || col_list[j].key.match('timestamp') || col_list[j].key.match('relatedContent.article: modified_time')) {
                  if (col_list[j].type == 'date_range') {
                     var _str = result[i][col_list[j].id].split('Z-');
                     _str[0] = _str[0] + 'Z';
                     var from = moment(_str[0]).format('MMM D, YYYY @ H:mm:ss.SSS');
                     var to = moment(_str[1]).format('MMM D, YYYY @ H:mm:ss.SSS');
                     var _date = from + 'to' + to;

                  }
                  else if (col_list[j].type == 'date_histogram') {
                     var _date = moment(result[i][col_list[j].id]).format('YYYY-MM-DD');
                  }
                  else {
                     var _date = moment(result[i][col_list[j].id]).format('MMM D, YYYY @ H:mm:ss.SSS');
                  }//terms
                  rowArr.push(_date);

               }
               else {
                  rowArr.push(result[i][col_list[j].id]);
               }

            }
            sheet.addRow(rowArr);
            rowArr = [];
         }
         var date_time = new Date();
         if (download_type == "excel") {
            var filename = dashboard_name + '.xlsx';
            // console.log('dashboard name=', dashboard_name)
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

   });
}
exports.get_Esquery = get_Esquery;