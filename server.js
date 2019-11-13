//test git
const get_Esquery = require('./utilFunctions/queryBuilding_kibana');
const get_ES_without_aggs = require('./utilFunctions/queryBuilding_basic');
const express = require('express');
var bodyParser = require('body-parser');
const swaggerJsDoc = require('swagger-jsdoc');
const swaggerui = require('swagger-ui-express');
var nodemon = require('nodemon');

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
   get_Esquery.get_Esquery(req.body, 'excel').then(result => {
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
   get_Esquery.get_Esquery(req.body, 'csv').then(result => {
      if (result.status == "success") {
         console.log('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
         res.download('./' + result.fileName);
         // res.send('File created successfully\nFilename:' + result.fileName + '\nPath:' + result.path);
      }
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
