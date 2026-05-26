// @ts-check
import { defineConfig } from 'astro/config';

// https://astro.build/config
export default defineConfig({
  site: 'https://mewrightlab.org',
  trailingSlash: 'always',
  build: {
    format: 'directory',
  },
});
