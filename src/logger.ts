import * as fs from "fs/promises";
import * as path from "path";
import { getConfigDir } from "./config.js";

const LOG_FILE = "action-log.jsonl";

export interface ActionLogEntry {
  timestamp: string; tool: string; account: string; summary: string; ids?: string[];
}

export async function logAction(entry: ActionLogEntry): Promise<void> {
  const logPath = path.join(getConfigDir(), LOG_FILE);
  const line = JSON.stringify(entry) + "\n";
  await fs.appendFile(logPath, line, "utf-8");
}
