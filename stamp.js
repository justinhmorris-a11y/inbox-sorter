// Rewrites ?v=<short git hash> onto the local script/style links in docs/taskpane.html so a
// republished pane is picked up on the next open instead of after GitHub Pages' 10-minute cache.
// Run by `npm run stamp` (and by `npm test`, so the stamp is always fresh before a push).
const fs = require('fs'), path = require('path'), cp = require('child_process');
const file = path.join(__dirname, 'docs', 'taskpane.html');
let v;
try { v = cp.execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim(); } catch (e) { v = String(Date.now()); }
const before = fs.readFileSync(file, 'utf8');
const after = before.replace(/((?:href|src)=")((?:taskpane\.css|config\.js|mock\.js|msal-browser\.min\.js|engine\.js|graph\.js|app\.js))(?:\?v=[^"]*)?"/g, '$1$2?v=' + v + '"');
if (after !== before) { fs.writeFileSync(file, after); console.log('taskpane.html stamped v=' + v); } else console.log('taskpane.html already stamped v=' + v);
