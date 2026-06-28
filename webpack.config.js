const path = require("path");
const os = require("os");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane.ts",
    commands: "./src/commands.ts",
  },
  output: {
    filename: "[name].bundle.js",
    path: path.resolve(__dirname, "dist"),
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
      {
        test: /\.(png|jpg|gif|svg|ico)$/,
        type: "asset/resource",
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"],
      inject: "body",
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands.html",
      chunks: ["commands"],
      inject: "body",
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "assets"),
      publicPath: "/assets",
    },
    port: 3000,
    https: {
      key: path.resolve(os.homedir(), ".office-addin-dev-certs", "localhost.key"),
      cert: path.resolve(os.homedir(), ".office-addin-dev-certs", "localhost.crt"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
};
