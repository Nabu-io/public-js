const path = require('path');

module.exports = {
  entry: './index.ts',
  output: {
    filename: 'parseImportSeed.js',
    path: path.resolve(__dirname, 'dist'),
    library: 'parseImportSeed',
    libraryTarget: 'umd',
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js'],
    modules: [path.resolve(__dirname), 'node_modules']
  },
  mode: 'production',
  target: 'web',
};
