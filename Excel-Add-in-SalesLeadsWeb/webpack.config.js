"use strict";

module.exports = {
  entry: "./home.js",
  output: {
    filename: "bundle.js"
  },
  module: {
    loaders: [
        {
          test: /\.js$/,
          loader: "babel-loader",
          exclude: /node_modules/,
          query: {
            presets: ["es2015", "es2016"]
          }
        }
    ]
  }
};