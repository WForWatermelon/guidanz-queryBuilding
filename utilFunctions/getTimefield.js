var elasticsearch = require('elasticsearch/src/elasticsearch');
const esUrl = require('../server');
var index_pattern = [];
let time;
var getTimestamp = function (vis_id) {
   var elasticClient = new elasticsearch.Client({
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
   return elasticClient.search(query).then(async function (resp) {
      var query2 = {
         body: {
            query: {
               match_phrase: {
                  "_id": 'index-pattern:' + resp.hits.hits[0]._source.references[0].id
               }
            }
         }
      }
      return elasticClient.search(query2)
      // return time;
   }).then(response => {
      time = response.hits.hits[0]._source['index-pattern'].timeFieldName;
      return time;
   })

   // console.log('time', time);

   // "Timestamp"
   // console.log(resp);
}
exports.getTimestamp = getTimestamp;