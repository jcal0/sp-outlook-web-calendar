// Shared by the plugin's listener and the Chrome extension. The port is also in
// extension/manifest.json host_permissions.
export const SYNC_PORT = 39217;
export const SYNC_PATH = '/outlook-week';
/** Required request header: forces a CORS preflight, which the listener never
 * grants, so web pages cannot post here; the extension's host permission skips it. */
export const SYNC_HEADER = 'x-sp-outlook-sync';
