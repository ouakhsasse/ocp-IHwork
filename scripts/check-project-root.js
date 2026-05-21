import { existsSync } from "node:fs";
import { join, resolve } from "node:path";
import process from "node:process";

const currentDirectory = process.cwd();
const nestedProjectDirectory = join(currentDirectory, "HK_energies-main");
const currentPackageJson = join(currentDirectory, "package.json");
const nestedPackageJson = join(nestedProjectDirectory, "package.json");

if (existsSync(currentPackageJson)) {
  console.log(`Project root detected: ${currentDirectory}`);
  process.exit(0);
}

console.error("You are not in the project root. Please run:");
console.error("cd HK_energies-main");
console.error("npm run dev");

if (existsSync(nestedPackageJson)) {
  console.error("");
  console.error(`Detected project root: ${resolve(nestedProjectDirectory)}`);
}

process.exit(1);

