#!/usr/bin/env node
const { spawn } = require("child_process");

function run(cmd, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { stdio: "inherit", shell: true, ...options });
    child.on("close", (code) => {
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`${cmd} failed with code ${code}`));
      }
    });
  });
}

async function main() {
  await run("npx", ["office-addin-dev-certs", "install"]);
  await run("node", ["scripts/dev-server.js"]);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
