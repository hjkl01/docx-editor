import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';

const monorepoRoot = path.resolve(__dirname, '../..');

export default defineConfig({
  plugins: [react()],
  root: __dirname,
  resolve: {
    alias: [
      {
        find: '@eigenpal/docx-editor-react',
        replacement: path.join(monorepoRoot, 'packages/react/src/index.ts'),
      },
      {
        find: /^@eigenpal\/docx-editor-core\/(.+)/,
        replacement: path.join(monorepoRoot, 'packages/core/src/$1'),
      },
      {
        find: /^@eigenpal\/docx-editor-core$/,
        replacement: path.join(monorepoRoot, 'packages/core/src/core.ts'),
      },
    ],
  },
  server: {
    port: 5180,
    open: true,
  },
  build: {
    outDir: 'dist',
  },
});
