#!/usr/bin/env node
// Entry point for the `outlook-mcp` binary. With no arguments it runs the MCP
// server over stdio, which is what every MCP client launches; `setup` and
// `login` are for people at a terminal. Each module runs itself on import.
const command = process.argv[2];

const modules: Record<string, string> = {
  setup: './setup.js',
  login: './login.js',
  serve: './index.js',
};

const target = command === undefined ? './index.js' : modules[command];
if (!target) {
  console.error(`Unknown command "${command}". Usage: outlook-mcp [setup | login]`);
  process.exit(1);
}
import(target).catch((error) => {
  console.error(error);
  process.exit(1);
});
