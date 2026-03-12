const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const os = require("os");

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");

module.exports = {
  entry: { taskpane: "./src/taskpane.js" },
  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "dist"),
    clean: true
  },
  resolve: { extensions: [".js"] },
  module: {
    rules: [
      { test: /\.css$/, use: ["style-loader", "css-loader"] }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"]
    }),
    new CopyPlugin({
      patterns: [
        { from: "manifest.xml", to: "manifest.xml" },
        { from: "src/commands.html", to: "commands.html" },
        { from: "src/esg-calculator.js", to: "esg-calculator.js" },
        { from: "src/taskpane.css", to: "taskpane.css" }
      ]
    })
  ],
  devServer: {
    port: 3000,
    hot: true,
    server: {
      type: "https",
      options: {
        cert: path.join(certDir, "localhost.crt"),
        key:  path.join(certDir, "localhost.key"),
        ca:   path.join(certDir, "ca.crt")
      }
    },
    headers: { "Access-Control-Allow-Origin": "*" }
  }
};
