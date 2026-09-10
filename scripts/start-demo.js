// Local-only interactive harness. Uses the esbuild version installed by tsx;
// it does not install dependencies or contact a Dataverse environment.
const path = require("node:path");
const fs = require("node:fs");
const esbuild = require(require.resolve("esbuild", {
    paths: [path.dirname(require.resolve("tsx/package.json"))],
}));
const root = path.resolve(__dirname, "..");
const output = path.join(root, "out", "demo");

async function start() {
    fs.mkdirSync(output, { recursive: true });
    fs.copyFileSync(path.join(root, "assets", "harness-demo.html"), path.join(output, "index.html"));
    const context = await esbuild.context({
        absWorkingDir: root,
        entryPoints: ["scripts/harness-demo.tsx"],
        bundle: true,
        outfile: path.join(output, "bundle.js"),
        platform: "browser",
        target: "es2020",
        define: { "process.env.NODE_ENV": '"development"' },
        logLevel: "info",
    });
    await context.watch();
    const server = await context.serve({ host: "127.0.0.1", port: 8183, servedir: output });
    console.log(`Interactive PCF demo: http://127.0.0.1:${server.port}/ (sample data only)`);
    let stopping = false;
    const stop = async () => {
        if (stopping) return;
        stopping = true;
        await context.dispose();
    };
    process.on("SIGINT", () => { void stop(); });
    process.on("SIGTERM", () => { void stop(); });
}
start().catch((error) => { console.error(error); process.exitCode = 1; });
