var elasticsearch = require('elasticsearch/src/elasticsearch');
const esUrl = require('../server');

var getMetaData = function (vis_id) {
   let elasticClient = new elasticsearch.Client({
      host: esUrl.esUrl
   });
   var query = {
      body: {
         query: {
            match_phrase: {
               "_id": vis_id
            }
         }
      }
   }
   return elasticClient.search(query).then(function (resp) {
      let metaData = JSON.parse(resp.hits.hits[0]._source.visualization.visState);
      return metaData;
   });
}

module.exports.getMetaData = getMetaData;