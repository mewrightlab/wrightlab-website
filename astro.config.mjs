// @ts-check
import { defineConfig } from 'astro/config';

// https://astro.build/config
export default defineConfig({
  site: 'https://mewrightlab.org',
  trailingSlash: 'always',
  redirects: {
    '/research/proximity-interaction-networks/': '/research/proximal-interaction-networks/',
  },
  build: {
    format: 'directory',
  },
});
