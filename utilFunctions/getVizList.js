var elasticsearch = require('elasticsearch/src/elasticsearch');
const esUrl = require('../server');
var getVisualizationList = function (dashboard_name) {
   let viz_list = [];
   let elasticClient = new elasticsearch.Client({
      host: esUrl.esUrl
   });
   let query = {
      body: {
         query: {
            match_phrase: {
               "dashboard.title": dashboard_name
            }
         }
      }
   }
   return elasticClient.search(query).then(async function (resp) {
      let references = resp.hits.hits[0]._source.references;
      for (let i = 0; i < references.length; i++) {
         let query1 = {
            body: {
               query: {
                  match_phrase: {
                     "_id": 'visualization:' + references[i].id
                  }
               }
            }
         }
         await elasticClient.search(query1).then(function (resp1) {
            viz_list.push({ [resp1.hits.hits[0]._source.visualization.title]: 'visualization:' + references[i].id });
            // console.log(viz_list);
         })
      }
      return viz_list;
   })
}

module.exports.getVisualizationList = getVisualizationList;
