import * as fs from "fs/promises";
import * as path from "path";
import * as os from "os";
import * as crypto from "crypto";

const CONFIG_DIR = path.join(
  process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config"),
  "mariana-outlook-mcp"
);
const CONFIG_PATH = path.join(CONFIG_DIR, "config.json");
const TOKENS_DIR = path.join(CONFIG_DIR, "tokens");

export interface AccountEntry { label: string; emailHash: string; }
export interface AppConfig {
  defaultAccount: string;
  accounts: Record<string, AccountEntry>;
}

export function getConfigDir(): string { return CONFIG_DIR; }
export function getTokensDir(): string { return TOKENS_DIR; }
export function hashEmail(email: string): string {
  return crypto.createHash("sha256").update(email.toLowerCase()).digest("hex");
}
export function getTokenPath(emailHash: string): string {
  return path.join(TOKENS_DIR, `${emailHash}.json`);
}
export async function ensureDirs(): Promise<void> {
  await fs.mkdir(CONFIG_DIR, { recursive: true });
  await fs.mkdir(TOKENS_DIR, { recursive: true, mode: 0o700 });
  // mkdir's mode only applies on creation; enforce on every start.
  await fs.chmod(TOKENS_DIR, 0o700);
}
export async function loadConfig(): Promise<AppConfig> {
  try {
    const raw = await fs.readFile(CONFIG_PATH, "utf-8");
    return JSON.parse(raw) as AppConfig;
  } catch { return { defaultAccount: "primary", accounts: {} }; }
}
export async function saveConfig(config: AppConfig): Promise<void> {
  await ensureDirs();
  await fs.writeFile(CONFIG_PATH, JSON.stringify(config, null, 2), "utf-8");
}
export async function addAccount(name: string, email: string): Promise<AppConfig> {
  const config = await loadConfig();
  config.accounts[name] = { label: name, emailHash: hashEmail(email) };
  if (Object.keys(config.accounts).length === 1) config.defaultAccount = name;
  await saveConfig(config);
  return config;
}
export async function resolveAccount(accountParam?: string): Promise<{ name: string; emailHash: string } | null> {
  const config = await loadConfig();
  const name = accountParam || config.defaultAccount;
  const entry = config.accounts[name];
  if (!entry) return null;
  return { name, emailHash: entry.emailHash };
}
