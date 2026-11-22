export default {
  input: {
    app: "./app.ts",
    "analysis-worker": "./analysis-worker.ts",
  },
  output: {
    dir: "site",
    format: "es",
    entryFileNames: "[name].js",
    chunkFileNames: "chunks/[name]-[hash].js",
  },
};
