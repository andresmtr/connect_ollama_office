#!/usr/bin/env node
const fs = require("fs");
const os = require("os");
const path = require("path");

const paths = [
  "Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef",
  "Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef",
  "Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
].map((p) => path.join(os.homedir(), p));

let removed = 0;
for (const target of paths) {
  try {
    fs.rmSync(target, { recursive: true, force: true });
    removed += 1;
    console.log(`Removed: ${target}`);
  } catch (err) {
    console.error(`Failed to remove ${target}: ${err.message}`);
  }
}

if (removed === 0) {
  console.log("No cache folders removed.");
}
