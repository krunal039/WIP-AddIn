const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const webpack = require("webpack");
const fs = require("fs");
const dotenv = require("dotenv");

// --- 1. Ensure NODE_ENV is defined ---
if (!process.env.NODE_ENV) {
  process.env.NODE_ENV = "development";
  console.log("[Webpack] NODE_ENV not set, defaulting to 'development'");
}

// --- 2. Load local .env file for development ---
if (process.env.NODE_ENV === "development") {
  const envPath = path.resolve(__dirname, ".env.local");
  try {
    dotenv.config({ path: envPath });
    console.log(`[Webpack] Loaded local env file: ${envPath}`);
  } catch (err) {
    console.warn(`[Webpack] No local env file found at: ${envPath}`);
  }
}

// --- 3. Function to dynamically get client environment ---
function getClientEnvironment() {
  const clientEnv = {};
  Object.keys(process.env).forEach((key) => {
    // Expose only variables with REACT_APP_ prefix + NODE_ENV
    if (key.startsWith("REACT_APP_") || key === "NODE_ENV") {
      clientEnv[key] = process.env[key];
    }
  });

  // Debug: show which variables are being injected into client
  console.log("[Webpack] Environment variables injected into client:", clientEnv);

  return clientEnv;
}

// --- 4. Helper to log all pipeline variables for debugging ---
function logAllPipelineVariables() {
  // Only log in pipeline builds, not in production client bundle
  if (process.env.BUILD_REASON) { // BUILD_REASON is set automatically in Azure DevOps
    console.log("[Webpack] Azure DevOps pipeline variables:");
    Object.keys(process.env)
      .sort()
      .forEach((key) => {
        // Mask secrets if they contain common secret keywords
        const masked = /(KEY|SECRET|PASSWORD|TOKEN)/i.test(key)
          ? "*****"
          : process.env[key];
        console.log(`  ${key} = ${masked}`);
      });
  }
}

// Call the pipeline debug helper
logAllPipelineVariables();

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
      process: require.resolve("process/browser"),
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
      ],
    }),
    new webpack.ProvidePlugin({
      process: "process/browser",
    }),
    new webpack.DefinePlugin({
      "process.env": JSON.stringify(getClientEnvironment()),
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
  mode: process.env.NODE_ENV === "production" ? "production" : "development",
  devtool: "source-map",
};
