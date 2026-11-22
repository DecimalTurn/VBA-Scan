export default {
  input: {
    app: "./app.js",
    "analysis-worker": "./analysis-worker.ts",
  },
  output: {
    dir: "site",
    format: "es",
    entryFileNames: "[name].js",
    chunkFileNames: "chunks/[name]-[hash].js",
  },
};
