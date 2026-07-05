// Webpack build config for the Michael Outlook add-in. ESM (package.json
// "type": "module"). Produces two entry chunks — taskpane + commands — and
// rewrites the dev manifest into the prod manifest (origin + AppDomain strip).

import devCerts from "office-addin-dev-certs";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
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
      // No polyfill entry. babel.config.json is a bare @babel/preset-env with
      // no useBuiltIns, so Babel transpiles SYNTAX only (async/await etc. are
      // left native — the package.json browserslist "last 2 versions" target
      // is modern enough). Importing core-js/stable + regenerator-runtime
      // previously shipped a 232KB polyfill.js the Office webview never needs;
      // both deps have been removed.
      taskpane: "./src/taskpane/index.js",
      commands: "./src/commands/commands.js",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    optimization: {
      // Keep each entry chunk self-contained. Webpack's default splitChunks
      // extracts shared modules (zai.js etc.) into a separate chunk, which for
      // this tiny 2-entry graph produced an orphan chunk + a broken <script src>
      // reference in commands.html (a hashed filename that was never emitted,
      // so the Quick Translate function-file would 404). Disabling splitting
      // inlines shared code into taskpane.js + commands.js — predictable + small.
      splitChunks: false,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          // Bare @babel/preset-env config lives in babel.config.json so it
          // applies uniformly; no inline override here.
          use: "babel-loader",
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
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
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
            // Self-hosted Inter woff2 (latin, weights 400/500/600/700). Copied
            // verbatim to dist/fonts/ and referenced by @font-face in
            // taskpane.css. Self-hosted => covered by CSP font-src 'self', no
            // external font CDN. Inter is latin-only; Korean glyphs fall back
            // to the system Korean font in the body font-family stack.
            from: "node_modules/@fontsource/inter/files/inter-latin-{400,500,600,700}-normal.woff2",
            to: "fonts/[name][ext]",
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
