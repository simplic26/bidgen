export default {
  optimizeDeps: {
    noDiscovery: true,
    include: [],
  },
  test: {
    environment: 'jsdom',
    setupFiles: './vitest.setup.ts',
    globals: true,
  },
};
