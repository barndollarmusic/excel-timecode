const devCerts = require('office-addin-dev-certs');
const CleanWebpackPlugin = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const CustomFunctionsMetadataPlugin = require('custom-functions-metadata-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const config = {
    devtool: 'source-map',
    entry: {
      functions: './src/functions/functions.js',
      polyfill: '@babel/polyfill',
      taskpane: './src/taskpane/taskpane.js',
      commands: './src/commands/commands.js'
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js']
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader', 
            options: {
              presets: ['@babel/preset-env']
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: 'html-loader'
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: 'file-loader'
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin({
        cleanOnceBeforeBuildPatterns: dev ? [] : ['**/*']
      }),
      new CustomFunctionsMetadataPlugin({
        output: 'functions.json',
        input: './src/functions/functions.js'
      }),
      new HtmlWebpackPlugin({
        filename: 'functions.html',
        template: './src/functions/functions.html',
        chunks: ['polyfill', 'functions']
      }),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
        {
          from: './src/taskpane/taskpane.css',
          to: 'taskpane.css'
        },
        {
          from: './assets/**/*',
          to: './'
        }
      ]),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['polyfill', 'commands']
      })
    ],
    devServer: {
      headers: {
        'Access-Control-Allow-Origin': '*'
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
