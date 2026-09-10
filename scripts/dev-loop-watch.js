const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const projectRoot = path.resolve(__dirname, "..");
// npm supplies its actual JavaScript entry point. Invoking it with Node works
// on Windows without trying to execute an npm.cmd file directly.
const npmCli = process.env.npm_execpath;
if (!npmCli) {
	throw new Error("Start this watcher with npm run dev:watch.");
}
const debounceMs = 500;

const watchedTopLevelFiles = new Set([
	"package.json",
	"pcfconfig.json",
	"tsconfig.json",
	"eslint.config.mjs",
	"version-bump.js",
]);

const watchedRoots = [
	"CMModelDrivenGrid",
	"CMModelDrivenGridSolution/src/Other",
	"scripts",
	"assets",
];

const ignoredPrefixes = [
	".git/",
	"node_modules/",
	"obj/",
	"out/",
	"CMModelDrivenGrid/generated/",
];

function toRelativePath(filePath) {
	return filePath.split(path.sep).join("/").replace(/^\.\//, "");
}

function isWatchedPath(relativePath) {
	if (!relativePath) {
		return false;
	}

	if (ignoredPrefixes.some((prefix) => relativePath === prefix.slice(0, -1) || relativePath.startsWith(prefix))) {
		return false;
	}

	if (watchedTopLevelFiles.has(relativePath)) {
		return true;
	}

	return watchedRoots.some((root) => relativePath === root || relativePath.startsWith(`${root}/`));
}

function timestamp() {
	return new Date().toLocaleTimeString("en-US", { hour12: false });
}

function log(message) {
	process.stdout.write(`[dev:watch ${timestamp()}] ${message}\n`);
}

let debounceTimer;
let activeChild = null;
let stopping = false;
let rerunQueued = false;
let rerunLogged = false;
const changedPaths = new Set();

function runDevLoop() {
	if (stopping) return;
	const triggerSummary = Array.from(changedPaths).slice(0, 4);
	const extraCount = Math.max(0, changedPaths.size - triggerSummary.length);
	changedPaths.clear();
	rerunQueued = false;
	rerunLogged = false;

	const reason =
		triggerSummary.length > 0
			? `${triggerSummary.join(", ")}${extraCount > 0 ? ` (+${extraCount} more)` : ""}`
			: "manual trigger";

	log(`running local build for ${reason}`);
	activeChild = spawn(process.execPath, [npmCli, "run", "build"], {
		cwd: projectRoot,
		stdio: "inherit",
		windowsHide: true,
		detached: process.platform !== "win32",
	});

	activeChild.on("exit", (code, signal) => {
		activeChild = null;

		if (signal) {
			log(`local build stopped by signal ${signal}`);
		} else if (code === 0) {
			log("local build completed");
		} else {
			log(`local build failed with exit code ${code}`);
		}

		if (!stopping && (rerunQueued || changedPaths.size > 0)) {
			scheduleRun();
		}
	});

	activeChild.on("error", (error) => {
		activeChild = null;
		log(`failed to start local build: ${error.message}`);
	});
}

function scheduleRun(changedPath) {
	if (stopping) return;
	if (changedPath) {
		changedPaths.add(changedPath);
		log(`change detected: ${changedPath}`);
	}

	if (activeChild) {
		rerunQueued = true;
		if (!rerunLogged) {
			rerunLogged = true;
			log("local build is still running; queued one rerun");
		}
		return;
	}

	if (debounceTimer) {
		clearTimeout(debounceTimer);
	}

	debounceTimer = setTimeout(() => {
		debounceTimer = undefined;
		runDevLoop();
	}, debounceMs);
}

let watcher;

try {
	watcher = fs.watch(projectRoot, { recursive: true }, (_eventType, filename) => {
		if (typeof filename !== "string" || filename.length === 0) {
			return;
		}

		const relativePath = toRelativePath(filename);
		if (!isWatchedPath(relativePath)) {
			return;
		}

		scheduleRun(relativePath);
	});
} catch (error) {
	const message = error instanceof Error ? error.message : String(error);
	process.stderr.write(`Unable to start dev watcher: ${message}\n`);
	process.exit(1);
}

function shutdown() {
	if (stopping) return;
	stopping = true;
	if (debounceTimer) {
		clearTimeout(debounceTimer);
	}

	if (watcher) {
		watcher.close();
	}

	if (activeChild?.pid) {
		if (process.platform === "win32") {
			// Stop only this watcher's build process and its descendants.
			const killer = spawn("taskkill", ["/PID", String(activeChild.pid), "/T", "/F"], {
				windowsHide: true,
				stdio: "ignore",
			});
			killer.on("error", () => activeChild?.kill());
		} else {
			try { process.kill(-activeChild.pid, "SIGTERM"); } catch { activeChild.kill(); }
		}
	}
	log("watcher stopped");
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

log("watching for changes and will build locally after edits (no Dataverse deployment)");
