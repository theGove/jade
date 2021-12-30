module.exports = {
    mode: 'development',
    entry: ['whatwg-fetch', './src/index.js'],
    output: {
      filename: 'bundle.js'
    },
    module: {
      rules: [
        {
          test: /\.m?js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
          }
        }
      ]
    }
  };
  