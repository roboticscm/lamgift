
const nodeExternals = require('webpack-node-externals');
const path = require('path');
const alias = {

};

module.exports = {
  target: 'node', // ignore built-in modules like path, fs, etc.
  externals: [nodeExternals()],
  // entry: {
  //   bundle: ['./src/index.js'],
  // },
  // output: {
  //   path: path.resolve(__dirname, 'dist'),
  //   filename: '[name][hash].js',
  //   chunkFilename: '[name].[id][hash].js'
  // },
};



