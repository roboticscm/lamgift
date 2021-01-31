
const nodeExternals = require('webpack-node-externals');
const path = require('path');
const CopyPlugin = require('copy-webpack-plugin');
const alias = {
  '$/src': path.resolve(__dirname, '/src'),
};

module.exports = {
  // resolve: {
  // alias
  // },
  target: 'node', // ignore built-in modules like path, fs, etc.
  // externals: [nodeExternals()],
  // entry: {
  //   bundle: ['./src/index.js'],
  // },
  // output: {
  //   path: path.resolve(__dirname, 'dist'),
  //   filename: '[name][hash].js',
  //   chunkFilename: '[name].[id][hash].js'
  // },
  module: {
    rules: [
      {
        test: /\.(html|cs|md|svg|ico|jpg|jpeg|png|gif|eot|otf|webp|ttf|woff|woff2|cur|ani|pdf|ogg)(\?.*)?$/,
        loader: 'file-loader'
      }
    ]
  },
  plugins: [
    new CopyPlugin({
      patterns: [
        { from: 'public', to: 'public' },
        { from: 'views', to: 'views' },
      ],
    }),
  ]
};



