
const nodeExternals = require('webpack-node-externals');

const alias = {

};

module.exports = {
  target: 'node', // ignore built-in modules like path, fs, etc.
  externals: [nodeExternals()]
};



