//test git
const get_Esquery = require('./utilFunctions/queryBuilding_kibana');
const get_ES_without_aggs = require('./utilFunctions/queryBuilding_basic');
const getTimestamp = require('./utilFunctions/getTimefield');
const getVizList = require('./utilFunctions/getVizList');
const getMeta = require('./utilFunctions/getMetaData');
const getDash=require('./utilFunctions/getDashboardList');
const express = require('express');
var bodyParser = require('body-parser');
const swaggerJsDoc = require('swagger-jsdoc');
const swaggerui = require('swagger-ui-express');
var nodemon = require('nodemon');
var elasticsearch = require('elasticsearch');
var Excel = require('exceljs');
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
 * /api/v1/getDashboardList:
*   get:
 *     description: get all dashboards
 *     produces:
 *       - application/json
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */
var dash_list=[];
 app.get('/api/v1/getDashboardList',async function(req,res){
   console.log('hooo');
    dash_list=await getDash.getDashboardList();
    res.send(dash_list);
 })

/**
 * @swagger
 * definitions:
 *    visualization:
 *       type: object
 *       properties:
 *          dashboard_name:
 *             type: string
 */

/**
 * @swagger
 * /api/v1/getVisualizationList:
*   post:
 *     description: Give dashboard name as input and get list of visualisations
 *     produces:
 *       - application/xlsx
 *     parameters:
 *       - name: id
 *         description: data to post
 *         in: body
 *         required: true
 *         schema:
 *           $ref: '#/definitions/visualization'
 *           type: object
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */
let vis_list = [];
let dashboard_name;
app.post('/api/v1/getVisualizationList', async function (req, res) {
   dashboard_name = req.body.dashboard_name;
   vis_list = await getVizList.getVisualizationList(dashboard_name);
   res.send(vis_list);
})

/**
 * @swagger
 * /api/v1/getSelectedVisualizations/excel:
*   post:
 *     description: Set dashboard UID
 *     produces:
 *       - application/json
 *     parameters:
 *       - name: UID
 *         description: Unique ID of the dashboard
 *         in: body
 *         required: true
 *         schema:
 *           type: object
 *           properties:
 *              value:
 *                  type: array
 *                  items:
 *                      type: string
 *     responses:
 *       200:
 *         description: success
 *
 *
 */

app.post('/api/v1/getSelectedVisualizations/excel', async function (req, res) {
   var workbook = new Excel.Workbook();
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
   var metaData = [];
   for (let i = 0; i < req.body.value.length; i++) {
      metaData.push(await getMeta.getMetaData(req.body.value[i]));
   }
   Promise.all(metaData).then(val => {
      var result = val.map(async (element, i) => {
         let time = await getTimestamp.getTimestamp(req.body.value[i]);
         return get_Esquery.get_Esquery(dashboard_name, element, workbook, time, 'excel');
      });
      Promise.all(result).then(val => {
         let flag = true;
         for (let i = 0; i < val.length; i++) {
            if (val[i].status != 'success') {
               flag = false;
            }
         }
         if (flag == true) {
            console.log("File successfully Downloaded in ", val[0].path);
            res.download('./Reports/' + val[0].fileName);
         }
         else {
            console.log('File cannot be downloaded');
         }
      })
   })
})
/**
 * @swagger
 * definitions:
 *    visualization:
 *       type: object
 *       properties:
 *          dashboard_name:
 *             type: string
 */

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
 *           $ref: '#/definitions/visualization'
 *           type: object
 *     responses:
 *       200:
 *         description: success
 *         schema:
 *           type: file
 *
 */

app.post('/api/v1/agg/excel', async function (req, res) {
   var workbook = new Excel.Workbook();
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
   var id = [];
   var metaData = [];
   var elasticClient = new elasticsearch.Client({
      host: esUrl
   });
   var query = {
      body: {
         query: {
            match_phrase: {
               "dashboard.title": req.body.dashboard_name
            }
         }
      }
   }

   elasticClient.search(query).then(async function (resp) {
      var references = resp.hits.hits[0]._source.references;
      for (var i = 0; i < references.length; i++) {
         id.push('visualization:' + references[i].id);
      }
      console.log('iddddd', id.length);
      for (let i = 0; i < id.length; i++) {
         metaData.push(await getMeta.getMetaData(id[i]));
      }
      Promise.all(metaData).then(val => {
         var result = val.map(async (element, i) => {
            let time = await getTimestamp.getTimestamp(id[i]);
            return get_Esquery.get_Esquery(req.body.dashboard_name, element, workbook, time, 'excel');
         });
         Promise.all(result).then(val => {
            let flag = true;
            for (let i = 0; i < val.length; i++) {
               if (val[i].status != 'success') {
                  flag = false;
               }
            }
            if (flag == true) {
               console.log("File successfully Downloaded in ", val[0].path);
               res.download('./Reports/' + val[0].fileName);
            }
            else {
               console.log('File cannot be downloaded');
            }
         })
      })
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
 *           $ref: '#/definitions/visualization'
 *           type: object
 *     responses:
 *       201:
 *         description: Created successfully
 */
app.post('/api/v1/agg/csv', function (req, res) {
   var workbook = new Excel.Workbook();
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
   var id = [];
   var metaData = [];
   var elasticClient = new elasticsearch.Client({
      host: esUrl
   });
   var query = {
      body: {
         query: {
            match_phrase: {
               "dashboard.title": req.body.dashboard_name
            }
         }
      }
   }

   elasticClient.search(query).then(async function (resp) {
      var references = resp.hits.hits[0]._source.references;
      for (var i = 0; i < references.length; i++) {
         id.push('visualization:' + references[i].id);
      }
      for (let i = 0; i < id.length; i++) {
         var query = {
            body: {
               query: {
                  match_phrase: {
                     "_id": id[i]
                  }
               }
            }
         }
         await elasticClient.search(query).then(function (resp) {
            metaData.push(JSON.parse(resp.hits.hits[0]._source.visualization.visState));
         });
      }
      var result = metaData.map(async (element, i) => {
         let time = await getTimestamp.getTimestamp(id[i]);
         return get_Esquery.get_Esquery(req.body.dashboard_name, element, workbook, time, 'csv');
      });
      Promise.all(result).then(val => {
         let flag = true;
         for (let i = 0; i < val.length; i++) {
            if (val[i].status != 'success') {
               flag = false;
            }
         }
         if (flag == true) {
            console.log("File successfully Downloaded in ", val[0].path);
            res.download('./Reports/' + val[0].fileName);
         }
         else {
            console.log('File cannot be downloaded');
         }
      })
      // console.log(result);


   })
});
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
   get_ES_without_aggs.get_ES_without_aggs(req.body.index_name, req.body.size, 'excel').then(result => {
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
   get_ES_without_aggs.get_ES_without_aggs(req.body.index_name, req.body.size, 'csv').then(result => {
      console.log(req.body.index)
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./Reports/' + result.fileName);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
      }
   })
});
app.listen(3035);
