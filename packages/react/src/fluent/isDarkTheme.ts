/** Theme Fluent: light / dark / theo hệ thống. */
export function isDarkTheme(theme: "light" | "dark" | "system"): boolean {
  if (theme === "dark") return true;
  if (theme === "system" && typeof matchMedia !== "undefined") {
    return matchMedia("(prefers-color-scheme: dark)").matches;
  }
  return false;
}
