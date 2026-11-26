(() => {
  const root = document.documentElement;
  const storageKey = "vba-scan-theme";
  const systemPreference = window.matchMedia("(prefers-color-scheme: dark)");
  const themeColor = document.querySelector("#theme-color");
  let toggle;

  try {
    const savedTheme = window.localStorage.getItem(storageKey);
    if (savedTheme === "light" || savedTheme === "dark") {
      root.dataset.theme = savedTheme;
    }
  } catch (error) {
    console.warn("[VBA-Scan] Could not read the saved theme preference.", error);
  }

  function getCurrentTheme() {
    return root.dataset.theme || (systemPreference.matches ? "dark" : "light");
  }

  function updateThemeControls() {
    if (!toggle) return;

    const currentTheme = getCurrentTheme();
    const nextTheme = currentTheme === "dark" ? "light" : "dark";
    toggle.dataset.action = nextTheme;
    toggle.setAttribute("aria-label", `Switch to ${nextTheme} theme`);
    toggle.setAttribute("title", `Switch to ${nextTheme} theme`);
    toggle.setAttribute("aria-pressed", String(currentTheme === "dark"));
    themeColor.content = currentTheme === "dark" ? "#171a18" : "#f8faf8";
  }

  document.addEventListener("DOMContentLoaded", () => {
    toggle = document.querySelector("#theme-toggle");
    updateThemeControls();

    toggle.addEventListener("click", () => {
      const nextTheme = getCurrentTheme() === "dark" ? "light" : "dark";
      root.dataset.theme = nextTheme;
      try {
        window.localStorage.setItem(storageKey, nextTheme);
      } catch (error) {
        console.warn("[VBA-Scan] Could not save the theme preference.", error);
      }
      updateThemeControls();
    });
  });

  systemPreference.addEventListener("change", () => {
    if (!root.dataset.theme) updateThemeControls();
  });
})();
