// Loads .env before any other module reads process.env.
// This must be the first import in generate-report.js because ES module imports
// are evaluated before the importing module's own code runs — so dotenv.config()
// called inside generate-report.js would be too late for config.js to see the values.
import dotenv from 'dotenv';
import path   from 'path';
import { fileURLToPath } from 'url';

dotenv.config({ path: path.join(path.dirname(fileURLToPath(import.meta.url)), '.env') });
