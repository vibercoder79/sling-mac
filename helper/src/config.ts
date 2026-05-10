import * as fs from "fs";
import * as path from "path";
import * as os from "os";

export interface AccountConfig {
  vaultPath: string;
  defaultFolder: string;
}

export interface SlingMacConfig {
  accounts: Record<string, AccountConfig>;
}

export const CONFIG_PATH = path.join(os.homedir(), ".sling-mac.json");

export function loadConfig(): SlingMacConfig {
  if (!fs.existsSync(CONFIG_PATH)) {
    throw new Error(
      `Config nicht gefunden: ${CONFIG_PATH}\nBitte im helper-Verzeichnis 'npm run setup' ausfuehren.`
    );
  }
  return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8")) as SlingMacConfig;
}

export function saveConfig(config: SlingMacConfig): void {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2), "utf-8");
}

export function getAccountConfig(config: SlingMacConfig, email: string): AccountConfig {
  const account = config.accounts[email];
  if (account) return account;

  const first = Object.values(config.accounts)[0];
  if (first) return first;

  throw new Error(
    `Kein Account-Config fuer ${email}. Bitte 'npm run setup' ausfuehren.`
  );
}
