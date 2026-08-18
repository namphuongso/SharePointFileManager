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
          surface: "var(--spm-surface, #faf9f8)",
          border: "var(--spm-border, #edebe9)",
          text: "var(--spm-text, #323130)",
          muted: "var(--spm-muted, #8a8886)",
          primary: "var(--spm-primary, #0078d4)",
          danger: "var(--spm-danger, #d13438)",
          hover: "var(--spm-hover, #f3f2f1)",
          header: "var(--spm-header, #faf9f8)",
          link: "var(--spm-link, #0078d4)",
        },
      },
    },
  },
};
