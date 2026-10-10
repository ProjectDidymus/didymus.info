// Hinode only purges CSS when this file exists, and it runs PostCSS against this folder.
// The configuration itself is Hinode's own, loaded from the vendored module, so the purge
// safelist follows the Hinode version pinned in go.mod. _vendor is refreshed by
// `npm run mod:vendor`, which `npm start` and `npm run build` run first.
module.exports = require('../_vendor/github.com/gethinode/hinode/v3/config/postcss.config.js')
