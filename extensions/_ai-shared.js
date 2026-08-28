// Shared helpers for AI-powered extensions (ai-summary.js, ai-narrative.js).
import fs   from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname2   = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname2, '..', 'ui-config.json');

export function readAiSettings() {
  try {
    if (fs.existsSync(UI_CONFIG_PATH)) {
      const cfg = JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'));
      return cfg.settings?.aiSettings || {};
    }
  } catch { /* ignore */ }
  return {};
}

export async function fetchNotes(aiSettings) {
  // ui-config contextFile takes precedence over SUMMARY_NOTES_FILE env var
  const configFile = aiSettings.contextFile;
  const envFile    = process.env.SUMMARY_NOTES_FILE;
  const url        = process.env.SUMMARY_NOTES_URL;

  const file = configFile || envFile;

  if (file) {
    const resolved = path.resolve(file);
    if (!fs.existsSync(resolved)) {
      console.warn(`  ⚠  Context file not found: ${resolved}`);
      // Fall through to URL if available
    } else {
      console.log(`  Notes loaded from file: ${resolved}`);
      return fs.readFileSync(resolved, 'utf8');
    }
  }

  if (url) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      console.log(`  Notes fetched from URL: ${url}`);
      return await res.text();
    } catch (e) {
      console.warn(`  ⚠  Could not fetch notes from SUMMARY_NOTES_URL — ${e.message}`);
      return null;
    }
  }

  return null;
}
