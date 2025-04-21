import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { exec } from "child_process";
import path from "node:path";

async function main() {
  // 1. Connect to Python MCP server
  const transport = new StdioClientTransport({
    command: "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3",
    args: ["mcp_server.py"]
  });
  const client = new Client(
    { name: "ppt-automator-client", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );
  await client.connect(transport);

  // 2. Ask user for the text
  const rl = readline.createInterface({ input, output });
  const text = await rl.question("Enter text to put in the rectangle: ");
  rl.close();

  // 3. Call update_slide
  const result = (await client.callTool({
    name: "update_slide",
    arguments: { text, ppt_path: "rectangle_demo.pptx", slide_index: 1 }
  })) as CallToolResult;
  console.log(result.content[0].text);

  // 4. Open the PPTX with the default app on macOS
  const pptPath = path.resolve(process.cwd(), "rectangle_demo.pptx");
  exec(`open "${pptPath}"`);

  // 5. Clean up
  await client.disconnect();
}

main().catch(console.error);
