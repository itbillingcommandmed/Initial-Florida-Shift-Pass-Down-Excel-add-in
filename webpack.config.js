const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const isDevelopment = argv.mode === "development";

  const config = {
    entry: "./src/taskpane.js",

    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "taskpane.js",
      clean: true,
    },

    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane.html",
        filename: "taskpane.html",
      }),

      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest.xml",
            to: "manifest.xml",
          },
        ],
      }),
    ],
  };

  // Only configure the localhost HTTPS server during development.
  // GitHub/Azure production builds do not need a local certificate.
  if (isDevelopment) {
    const httpsOptions = await devCerts.getHttpsServerOptions();

    config.devServer = {
      server: {
        type: "https",
        options: {
          ca: httpsOptions.ca,
          key: httpsOptions.key,
          cert: httpsOptions.cert,
        },
      },
      port: 3000,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    };
  }

  return config;
};
