const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const webpack = require("webpack");
const fs = require('fs');

module.exports = {
  entry: "./src/index.tsx",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "bundle.js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
    fallback: {
      "process": require.resolve("process/browser"),
    },
  },
  module: {
    rules: [
      {
        test: /\.(ts|tsx)$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, "public", "index.html"),
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: path.resolve(__dirname, "public", "assets"),
          to: path.resolve(__dirname, "dist", "assets"),
        },
        {
          from: path.resolve(__dirname, "environments.json"),
          to: path.resolve(__dirname, "dist", "environments.json"),
        },
      ],
    }),
    new webpack.ProvidePlugin({
      process: "process/browser",
    }),
    new webpack.DefinePlugin({
      'process.env': JSON.stringify({
        // Only include NODE_ENV - all other env vars come from environments.json at runtime
        NODE_ENV: process.env.NODE_ENV || 'development',
      }),
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "public"),
    },
    port: 3001,
    server: {
      type: "https",
      options: {
        key: fs.readFileSync(
          path.resolve(__dirname, "certs", "localhost-key.pem")
        ),
        cert: fs.readFileSync(
          path.resolve(__dirname, "certs", "localhost-cert.pem")
        ),
      },
    },
    open: true,
    hot: true,
  },
  mode: "development",
  devtool: 'source-map'
};
