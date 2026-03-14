const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const isDevelopment = process.env.NODE_ENV !== "production";

const homeDir = process.env.USERPROFILE || process.env.HOME || "";
const certDir = path.join(homeDir, ".office-addin-dev-certs");
const certFile = path.join(certDir, "localhost.crt");
const keyFile = path.join(certDir, "localhost.key");
const caFile = path.join(certDir, "ca.crt");
const hasDevCerts =
  fs.existsSync(certFile) && fs.existsSync(keyFile) && fs.existsSync(caFile);

module.exports = {
  entry: {
    taskpane: "./src/taskpane.tsx",
    functions: "./src/functions/functions.ts"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    publicPath: "/"
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"]
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        exclude: /node_modules/,
        use: "ts-loader"
      },
      {
        test: /\.css$/i,
        use: [MiniCssExtractPlugin.loader, "css-loader"]
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"],
      minify: false
    }),
    new HtmlWebpackPlugin({
      filename: "functions.html",
      template: "./src/functions.html",
      chunks: ["functions"],
      minify: false
    }),
    new CopyPlugin({
      patterns: [
        { from: "src/assets", to: "assets", noErrorOnMissing: true },
        { from: "src/functions/metadata.json", to: "functions.json" }
      ]
    }),
    new MiniCssExtractPlugin(),
    new webpack.DefinePlugin({
      "process.env.API_BASE_URL": JSON.stringify(process.env.API_BASE_URL || "https://localhost:8000"),
      "process.env.DEFAULT_PROVIDER": JSON.stringify(process.env.DEFAULT_PROVIDER || "mock")
    })
  ],
  devServer: {
    port: 3000,
    hot: true,
    server: hasDevCerts
      ? {
          type: "https",
          options: {
            cert: fs.readFileSync(certFile),
            key: fs.readFileSync(keyFile),
            ca: fs.readFileSync(caFile),
          },
        }
      : "https",
    allowedHosts: "all",
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
      "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization"
    },
    client: {
      overlay: {
        errors: true,
        warnings: false,
        runtimeErrors: (error) => {
          // Suppress cross-origin "Script error." from office.js CDN
          if (error?.message === "Script error.") return false;
          return true;
        }
      }
    }
  },
  devtool: isDevelopment ? "inline-source-map" : false,
  mode: isDevelopment ? "development" : "production"
};

