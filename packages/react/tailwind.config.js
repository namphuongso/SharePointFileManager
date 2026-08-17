/** @type {import('tailwindcss').Config} */
export default {
  prefix: "spm-",
  content: ["./src/**/*.{ts,tsx}"],
  corePlugins: {
    preflight: false,
  },
  theme: {
    extend: {
      colors: {
        sp: {
          bg: "var(--spm-bg, #ffffff)",
          surface: "var(--spm-surface, #f8fafc)",
          border: "var(--spm-border, #e2e8f0)",
          text: "var(--spm-text, #0f172a)",
          muted: "var(--spm-muted, #64748b)",
          primary: "var(--spm-primary, #2563eb)",
          danger: "var(--spm-danger, #dc2626)",
        },
      },
    },
  },
};
