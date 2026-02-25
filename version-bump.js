const fs = require("fs");

// Bump control manifest version (major.minor.patch)
const manifestPath = "./CMModelDrivenGrid/ControlManifest.Input.xml";
let manifest = fs.readFileSync(manifestPath, "utf8");
manifest = manifest.replace(/version="(\d+)\.(\d+)\.(\d+)"/, (_, major, minor, patch) => {
	const next = `version="${major}.${minor}.${parseInt(patch) + 1}"`;
	console.log(`Control version: ${major}.${minor}.${patch} → ${major}.${minor}.${parseInt(patch) + 1}`);
	return next;
});
fs.writeFileSync(manifestPath, manifest, "utf8");

// Bump solution version (major.minor.build.revision)
const solutionPath = "./CMModelDrivenGridSolution/src/Other/Solution.xml";
let solution = fs.readFileSync(solutionPath, "utf8");
solution = solution.replace(/<Version>(\d+)\.(\d+)\.(\d+)\.(\d+)<\/Version>/, (_, a, b, c, d) => {
	const next = `<Version>${a}.${b}.${parseInt(c) + 1}.${d}</Version>`;
	console.log(`Solution version: ${a}.${b}.${c}.${d} → ${a}.${b}.${parseInt(c) + 1}.${d}`);
	return next;
});
fs.writeFileSync(solutionPath, solution, "utf8");
