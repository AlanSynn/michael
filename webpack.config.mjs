// Webpack build config for the Michael Outlook add-in. ESM (package.json
// "type": "module"). Produces two entry chunks — taskpane + commands — and
// rewrites the dev manifest into the prod manifest (origin + AppDomain strip).

import devCerts from "office-addin-dev-certs";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import path from "node:path";
import webpack from "webpack";

const urlDev = "https://localhost:3000/";
const urlProd = "https://alansynn.com/michael/";

const addinName = "Michael";
const addinDevName = "Michael [Local]";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

export default async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/index.js",
      commands: "./src/commands/commands.js",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // The Z.AI API key is entered at runtime (Settings -> General) and stored
      // in Office.context.roamingSettings. The build-time ZAI_API_KEY env var is
      // honored ONLY for dev builds (local convenience); production builds always
      // embed "" so a key can never leak into the published bundle.
      new webpack.DefinePlugin({
        __ZAI_API_KEY__: JSON.stringify(dev ? process.env.ZAI_API_KEY || "" : ""),
        __ZAI_CODING_BASE_URL__: JSON.stringify(
          process.env.ZAI_CODING_BASE_URL || "https://api.z.ai/api/coding/paas/v4"
        ),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            filter: (resourcePath) => {
              return resourcePath.endsWith(".png") || resourcePath.endsWith(".svg");
            },
          },
          {
            from: "manifest*.xml",
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                const contentStr = content.toString();
                const escapedAddinDevName = addinDevName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                const escapedUrlDev = urlDev.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

                // Rewrite dev URLs to the production origin and drop the dev-only
                // localhost AppDomain entry (the alansynn.com AppDomain already covers prod).
                return contentStr
                  .replace(new RegExp(escapedAddinDevName, "g"), addinName)
                  .replace(new RegExp(escapedUrlDev, "g"), urlProd)
                  .replace(/[ \t]*<AppDomain>https:\/\/localhost:3000<\/AppDomain>[ \t]*\n/g, "");
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      historyApiFallback: true,
    },
  };

  return config;
};
