// Builds both deliverables and their release zips:
//   dist/plugin/     → dist/outlook-web-calendar-plugin-v<version>.zip
//                      (upload in Super Productivity: Settings → Plugins)
//   dist/extension/  → dist/outlook-web-calendar-chrome-extension-v<version>.zip
//                      (unzip, then chrome://extensions → Load unpacked)
// The version in package.json is written into both manifests, so a release tag,
// the plugin and the extension always agree.
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const esbuild = require('esbuild');

const root = path.resolve(__dirname, '..');
const dist = path.join(root, 'dist');
const { version } = require(path.join(root, 'package.json'));

const bundle = (entry, outfile, minify) =>
  esbuild.build({
    entryPoints: [path.join(root, entry)],
    outfile,
    bundle: true,
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    minify,
    logLevel: 'warning',
  });

const writeManifest = (src, outDir) => {
  const manifest = JSON.parse(fs.readFileSync(path.join(root, src), 'utf8'));
  manifest.version = version;
  fs.writeFileSync(path.join(outDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
};

// Flat archive: manifest.json must sit at the zip root for both consumers.
const zipDir = (dir, zipFile) => {
  if (process.platform === 'win32') {
    execFileSync('powershell.exe', [
      '-NoProfile',
      '-Command',
      `Compress-Archive -Path '${path.join(dir, '*')}' -DestinationPath '${zipFile}'`,
    ]);
  } else {
    execFileSync('zip', ['-qr', zipFile, '.'], { cwd: dir });
  }
};

(async () => {
  fs.rmSync(dist, { recursive: true, force: true });
  const pluginDir = path.join(dist, 'plugin');
  const extDir = path.join(dist, 'extension');
  fs.mkdirSync(pluginDir, { recursive: true });
  fs.mkdirSync(extDir, { recursive: true });

  await bundle('src/plugin.ts', path.join(pluginDir, 'plugin.js'), true);
  writeManifest('src/manifest.json', pluginDir);
  fs.copyFileSync(path.join(root, 'src/config-schema.json'), path.join(pluginDir, 'config-schema.json'));
  fs.copyFileSync(path.join(root, 'icon.svg'), path.join(pluginDir, 'icon.svg'));

  // Unminified: people sideloading an extension should be able to read what it does.
  await bundle('extension/content.ts', path.join(extDir, 'content.js'), false);
  await bundle('extension/background.ts', path.join(extDir, 'background.js'), false);
  writeManifest('extension/manifest.json', extDir);

  const pluginZip = path.join(dist, `outlook-web-calendar-plugin-v${version}.zip`);
  const extZip = path.join(dist, `outlook-web-calendar-chrome-extension-v${version}.zip`);
  zipDir(pluginDir, pluginZip);
  zipDir(extDir, extZip);
  console.log(`Built v${version}:\n  ${pluginZip}\n  ${extZip}`);
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
