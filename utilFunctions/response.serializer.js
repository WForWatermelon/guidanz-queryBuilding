/* eslint-disable prefer-rest-params */
/* eslint-disable func-names */
var merge = require('xtend');
const reduceObj = function reduceObj(predicate, initial, obj, pastAgg) {
  return Object.keys(obj).reduce((acc, key) => predicate(acc, obj[key], key, pastAgg), initial);
};

function createKeyValueObject(key, value) {
  const obj = {};
  obj[key] = value;
  return obj;
}

function ensureDimensionIsComplete(agg) {
  if ((!agg.dimensionName || !agg.dimensionValue) && !(agg.dimensionValue === 0)) return agg;
  const newAgg = merge(createKeyValueObject(agg.dimensionName, agg.dimensionValue), agg);
  delete newAgg.dimensionName;
  delete newAgg.dimensionValue;
  return newAgg;
}

function Aggregation(state) {
  this.state = state || [{}];
}

Aggregation.prototype.map = function map(iterator) {
  return new Aggregation(this.state.map(iterator));
};

Aggregation.prototype.setDimensionName = function setDimensionName(dimensionName) {
  return this.map(agg => ensureDimensionIsComplete(merge(createKeyValueObject('dimensionName', dimensionName), agg)));
};

Aggregation.prototype.setDimensionValue = function setDimensionValue(dimensionValue) {
  return this.map(agg => ensureDimensionIsComplete(merge(createKeyValueObject('dimensionValue', dimensionValue), agg)));
};

Aggregation.prototype.addMetric = function addMetric(name, val) {
  return this.map(agg => merge(agg, createKeyValueObject(name, val)));
};

Aggregation.prototype.concat = function concat(agg) {
  return new Aggregation(this.state.concat(agg.state));
};

Aggregation.prototype.addSubAggregation = function addSubAggregation(agg) {
  return new Aggregation(this.state.reduce((acc, left) => acc.concat((agg.state
    || []).map(right => merge(left, right))), []));
};

Aggregation.prototype.get = function get() {
  return this.state;
};

function EmptyAggregation() {
}
EmptyAggregation.prototype.map = function map() {
  return this;
};
EmptyAggregation.prototype.setDimensionName = function setDimensionName(dimensionName) {
  return new Aggregation().setDimensionName(dimensionName);
};
EmptyAggregation.prototype.setDimensionValue = function setDimensionValue(dimensionValue) {
  return new Aggregation().setDimensionValue(dimensionValue);
};
EmptyAggregation.prototype.addMetric = function addMetric(name, val) {
  return new Aggregation().addMetric(name, val);
};
EmptyAggregation.prototype.concat = function concat(agg) {
  return agg;
};
EmptyAggregation.prototype.addSubAggregation = function addSubAggregation(agg) {
  return agg;
};
EmptyAggregation.prototype.get = function get() {
  return [];
};

function isBucket(bucket) {
  return typeof bucket === 'object' && Object.prototype.hasOwnProperty.call(bucket, 'key');
}

function isSubAgg(subAgg) {
  return typeof subAgg === 'object' && Object.prototype.hasOwnProperty.call(subAgg, 'buckets');
}

function handleMetrics(next, first, agg, metric, metricName, pastAggregation) {
  if (!Object.prototype.hasOwnProperty.call(metric, 'value')) return next(agg, metric, metricName, pastAggregation);
  return agg.addMetric(metricName, metric.value);
}

function handleOneBucket(next, first, agg, bucket, key, pastAggregation) {
  if (!isBucket(bucket)) return next(agg, bucket, key, pastAggregation);
  return reduceObj(first, agg.setDimensionValue(bucket.key),
    bucket, agg.setDimensionValue(bucket.key));
}

function handleBuckets(next, first, agg, buckets, key, pastAggregation) {
  if (!Array.isArray(buckets) || key !== 'buckets') return next(agg, buckets, key, pastAggregation);
  return buckets.map((bucket, idx) => first(agg, bucket, idx, pastAggregation)).reduce(
    (acc, aggregations) => acc.concat(aggregations), new EmptyAggregation()
  );
}

function addSubOrMissingAgg(first, agg, subAgg, subAggName,
  pastAggregation = new EmptyAggregation()) {
  if (agg.state && agg.state.length && Object.keys(agg.state[0]).length > 1) {
    return agg.concat(pastAggregation.addSubAggregation(reduceObj(first,
      new EmptyAggregation().setDimensionName(subAggName), subAgg, agg)));
  }
  return (agg.state && agg.state.length ? agg : pastAggregation).addSubAggregation(reduceObj(first,
    new EmptyAggregation().setDimensionName(subAggName), subAgg, agg));
}

function handleSubAggregation(next, first, agg, subAgg, subAggName, pastAggregation) {
  if (!isSubAgg(subAgg)) return next(agg, subAgg, subAggName, pastAggregation);
  return addSubOrMissingAgg(first, agg, subAgg, subAggName, pastAggregation);
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function isMissingAggregation(missingAggName, missingAggs) {
  return (missingAggs.doc_count && (missingAggName && !isNumeric(missingAggName))
    && typeof missingAggs === 'object' && missingAggName.indexOf('missing_') >= 0);
}

function createSubAggWithMissingAgg(missingAggs) {
  return createKeyValueObject('buckets', [merge(createKeyValueObject('key', 'null'), missingAggs)]);
}


function handleMissingAggregations(next, first, agg, missingAggs, missingAggName, pastAggregation) {
  if (!isMissingAggregation(missingAggName, missingAggs)) {
    return next(agg, missingAggs, missingAggName, pastAggregation);
  }
  return addSubOrMissingAgg(first, agg, createSubAggWithMissingAgg(missingAggs),
    missingAggName.slice(8), pastAggregation);
}

const defaultHandlers = [
  handleMetrics,
  handleOneBucket,
  handleBuckets,
  handleSubAggregation,
  handleMissingAggregations
];

function createCor(handlers, fallBack) {
  // eslint-disable-next-line prefer-arrow-callback
  return handlers.reduce(function (nextHandler, handler) {
    return handler.bind(null, nextHandler, function first() {
      // eslint-disable-next-line prefer-spread
      return createCor(handlers, fallBack).apply(null, arguments);
    });
  }, fallBack);
}

function serializer(response, options = {}) {
  if (!response.aggregations) return [];
  return reduceObj(createCor(options.handlers || defaultHandlers, handler => handler),
    new EmptyAggregation(), response.aggregations).get();
}
module.exports = serializer
