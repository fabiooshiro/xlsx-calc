// webpack.config.js

var webpack = require('webpack');
var path = require('path');
var libraryName = 'xlsx-calc';
var outputFile = libraryName + '.js';

var config = {
  entry: __dirname + '/src/index.js',
  devtool: 'source-map',
  output: {
    path: __dirname + '/lib',
    filename: outputFile,
    libraryTarget: "umd",
    globalObject: 'this'
  }
};

module.exports = config;
