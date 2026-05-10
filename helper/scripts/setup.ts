import * as readline from "readline";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { CONFIG_PATH, SlingMacConfig, saveConfig } from "../src/config";

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

function ask(question: string): Promise<string> {
  return new Promise((resolve) => rl.question(question, (a) => resolve(a.trim())));
}

async function main() {
  console.log("\nSling-Mac Setup\n");

  let config: SlingMacConfig = { accounts: {} };
  if (fs.existsSync(CONFIG_PATH)) {
    config = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
    console.log(`Bestehende Config: ${CONFIG_PATH}`);
  }

  const email = await ask("Outlook E-Mail-Adresse: ");

  const defaultVault = config.accounts[email]?.vaultPath ?? path.join(os.homedir(), "Obsidian");
  const vaultInput = await ask(`Vault-Pfad [${defaultVault}]: `);
  const vaultPath = vaultInput || defaultVault;

  if (!fs.existsSync(vaultPath)) {
    console.error(`\nPfad existiert nicht: ${vaultPath}`);
    rl.close();
    process.exit(1);
  }

  const defaultFolder = config.accounts[email]?.defaultFolder ?? "01 Inbox";
  const folderInput = await ask(`Standard-Sling-Ordner (relativ zum Vault) [${defaultFolder}]: `);
  const folder = folderInput || defaultFolder;

  const fullFolderPath = path.join(vaultPath, folder);
  if (!fs.existsSync(fullFolderPath)) {
    console.log(`Hinweis: Ordner existiert noch nicht — wird beim ersten Sling angelegt.`);
  }

  config.accounts[email] = { vaultPath, defaultFolder: folder };
  saveConfig(config);

  console.log(`\nConfig gespeichert: ${CONFIG_PATH}`);
  console.log(`  Account: ${email}`);
  console.log(`  Vault:   ${vaultPath}`);
  console.log(`  Inbox:   ${folder}\n`);

  rl.close();
}

main().catch((err) => {
  console.error(err.message);
  rl.close();
  process.exit(1);
});
