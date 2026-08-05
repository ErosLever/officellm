#!/usr/bin/env node
/**
 * Generates a random API token and writes it to:
 *   .env                              (read by webpack DefinePlugin at JS build time)
 *   src/mcp-server/GeneratedToken.cs  (compiled into the C# server binary)
 *
 * Both files are gitignored — never committed.
 * Run once before building: `node scripts/generate-token.mjs`
 * Re-run to rotate the token (then rebuild both server and add-in).
 */

import { randomBytes } from "crypto";
import { writeFileSync, readFileSync, existsSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const envPath = resolve(root, ".env");
const csPath = resolve(root, "src/mcp-server/GeneratedToken.cs");

// If .env already exists and has a token, reuse it (idempotent).
let token = null;
if (existsSync(envPath)) {
  const existing = readFileSync(envPath, "utf8");
  const match = existing.match(/^MCP_API_TOKEN=(.+)$/m);
  if (match) token = match[1].trim();
}

if (!token) {
  token = randomBytes(32).toString("base64url");
  console.log("Generated new token.");
} else {
  console.log("Reusing existing token from .env.");
}

// Write .env (preserves any other lines already present)
let envContent = existsSync(envPath) ? readFileSync(envPath, "utf8") : "";
if (/^MCP_API_TOKEN=/m.test(envContent)) {
  envContent = envContent.replace(/^MCP_API_TOKEN=.*$/m, `MCP_API_TOKEN=${token}`);
} else {
  envContent += (envContent.endsWith("\n") || envContent === "" ? "" : "\n") +
    `MCP_API_TOKEN=${token}\n`;
}
writeFileSync(envPath, envContent, "utf8");

// Write GeneratedToken.cs
writeFileSync(csPath, `// AUTO-GENERATED — do not commit. Re-run scripts/generate-token.mjs to rotate.
namespace OfficeMcpServer;

internal static class BuildToken
{
    internal const string Value = "${token}";
}
`, "utf8");

console.log(`Token written to:`);
console.log(`  ${envPath}`);
console.log(`  ${csPath}`);
console.log(`\nShare the token with external MCP clients (e.g. Claude Desktop) by reading .env.`);
