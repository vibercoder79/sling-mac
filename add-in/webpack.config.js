/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

async function buildConfig(env, options) {
  const devCerts = require("office-addin-dev-certs");
  const httpsOptions = await devCerts.getHttpsServerOptions();

  return {
    entry: {
      commands: "./src/commands/commands.ts",
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
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
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/picker.html", to: "picker.html" },
        ],
      }),
    ],
    devServer: {
      hot: true,
      host: "localhost",
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
}

module.exports = buildConfig;
