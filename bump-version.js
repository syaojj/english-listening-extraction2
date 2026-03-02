/**
 * 버전 0.1 올리기: config.json + js/app.js 의 APP_VERSION 동시 갱신
 * 사용: node bump-version.js
 */
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname);
const configPath = path.join(root, 'config.json');
const appPath = path.join(root, 'js', 'app.js');

const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const current = config.version;
const parts = current.split('.').map(Number);
if (parts.length >= 2) {
  parts[1] = (parts[1] || 0) + 1;
  if (parts[1] >= 10) {
    parts[0] = (parts[0] || 0) + 1;
    parts[1] = 0;
  }
} else {
  parts[0] = (parts[0] || 0) + 1;
  parts.push(0);
}
const next = parts.join('.');

config.version = next;
fs.writeFileSync(configPath, JSON.stringify(config, null, 2) + '\n', 'utf8');

let appCode = fs.readFileSync(appPath, 'utf8');
appCode = appCode.replace(/var APP_VERSION = '[^']+';/, "var APP_VERSION = '" + next + "';");
fs.writeFileSync(appPath, appCode, 'utf8');

console.log('버전 업데이트: ' + current + ' → ' + next + ' (config.json, js/app.js)');
