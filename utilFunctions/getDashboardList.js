var elasticsearch = require('elasticsearch/src/elasticsearch');
const esUrl = require('../server');
var getDashboardList = async function () {
  console.log('hihihi');
  let dash_list = [];
  let elasticClient = new elasticsearch.Client({
     host: esUrl.esUrl
  });
  let query = {
    body: {
       query: {
          match_phrase: {
             type: 'dashboard'
          }
       }
    }
 }

 await elasticClient.search(query).then(function(resp){
   var response=resp.hits.hits;
   for(let i=0;i<response.length;i++){
     dash_list.push({[response[i]._source.dashboard.title]:response[i]._id});
   }
 })
 return dash_list;
}

module.exports.getDashboardList = getDashboardList;