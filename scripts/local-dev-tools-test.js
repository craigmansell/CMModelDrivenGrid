/* Standalone tests: no real processes are started and no files are changed. */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { EventEmitter } = require("node:events");

const source = fs.readFileSync(path.join(__dirname, "dev-loop-watch.js"), "utf8");
const scripts = JSON.parse(fs.readFileSync(path.join(__dirname, "../package.json"), "utf8")).scripts;
assert.equal(scripts["dev:loop"], "npm run test:local && npm run build");
assert.equal(scripts["dev:watch"], "node scripts/dev-loop-watch.js");

for (const platform of ["win32", "linux"]) {
    const fakeProcess = new EventEmitter();
    const kills = [];
    Object.assign(fakeProcess, {
        env: { npm_execpath: "/tools/npm-cli.js" },
        execPath: "/tools/node",
        platform,
        stdout: { write: () => undefined },
        stderr: { write: () => undefined },
        exit: () => assert.fail("Watcher unexpectedly exited"),
        kill: (pid, signal) => kills.push([pid, signal]),
    });
    let watchCallback;
    let closed = false;
    const timers = new Map();
    let timerId = 0;
    const calls = [];
    vm.runInNewContext(source, {
        __dirname,
        process: fakeProcess,
        require: (name) => {
            if (name === "path") return path;
            if (name === "fs") return {
                watch: (_root, _options, callback) => {
                    watchCallback = callback;
                    return { close: () => { closed = true; } };
                },
            };
            if (name === "child_process") return {
                spawn: (command, args, options) => {
                    const child = new EventEmitter();
                    child.pid = 500 + calls.length;
                    child.kill = () => kills.push([child.pid]);
                    calls.push({ command, args: Array.from(args), options, child });
                    return child;
                },
            };
            throw new Error(`Unexpected dependency: ${name}`);
        },
        setTimeout: (callback) => { timers.set(++timerId, callback); return timerId; },
        clearTimeout: (id) => timers.delete(id),
    });
    const flush = () => {
        const pending = Array.from(timers.values());
        timers.clear();
        pending.forEach((callback) => callback());
    };
    assert.equal(calls.length, 0, "Starting the watcher must not deploy or build");
    watchCallback("change", "CMModelDrivenGrid/generated/ManifestTypes.d.ts");
    flush();
    assert.equal(calls.length, 0, "Generated files must not trigger a rebuild loop");
    watchCallback("change", "CMModelDrivenGrid/Grid.tsx");
    watchCallback("change", "CMModelDrivenGrid/index.ts");
    flush();
    assert.equal(calls.length, 1, "Burst changes should produce one build");
    assert.equal(calls[0].command, "/tools/node");
    assert.deepEqual(calls[0].args, ["/tools/npm-cli.js", "run", "build"]);
    assert.equal(calls[0].options.windowsHide, true);
    watchCallback("change", "CMModelDrivenGrid/Grid.tsx");
    watchCallback("change", "CMModelDrivenGrid/Grid.tsx");
    flush();
    assert.equal(calls.length, 1, "Only one build may run at a time");
    calls[0].child.emit("exit", 0);
    flush();
    assert.equal(calls.length, 2, "Changes during a build should queue one rerun");
    watchCallback("change", "CMModelDrivenGrid/Grid.tsx");
    fakeProcess.emit("SIGINT");
    assert.equal(closed, true);
    if (platform === "win32") {
        assert.equal(calls[2].command, "taskkill");
        assert.deepEqual(calls[2].args, ["/PID", "501", "/T", "/F"]);
        assert.equal(calls[2].options.windowsHide, true);
    } else {
        assert.deepEqual(kills, [[-501, "SIGTERM"]]);
    }
    const afterShutdown = calls.length;
    calls[1].child.emit("exit", 1);
    flush();
    assert.equal(calls.length, afterShutdown, "Shutdown must not start a queued build");
}
console.log("Development tooling regression checks passed (Windows and Unix).");
