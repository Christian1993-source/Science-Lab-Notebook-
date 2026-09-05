/* Check for an updated document only before the student starts writing. */
(async function () {
  const version = document.querySelector('meta[name="notebook-release"]')?.content || "1";
  if (location.protocol !== "file:") {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 4000);
    try {
      const url = new URL("./index.html", location.href);
      url.searchParams.set("release-check", String(Date.now()));
      const response = await fetch(url, { cache: "no-store", signal: controller.signal });
      if (response.ok) {
        const documentText = new DOMParser().parseFromString(await response.text(), "text/html");
        const latest = documentText.querySelector('meta[name="notebook-release"]')?.content;
        if (latest && latest !== version) {
          const destination = new URL(location.href);
          if (destination.searchParams.get("notebook-update") !== latest) {
            destination.searchParams.set("notebook-update", latest);
            location.replace(destination.href);
            return;
          }
        }
      }
    } catch (_error) {
      // Offline copies can still open; never interrupt an active report to update.
    } finally {
      clearTimeout(timeout);
    }
  }
  try {
    for (const file of ["figures.js", "app.js"]) {
      await new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = `${file}?v=${encodeURIComponent(version)}`;
        script.onload = resolve;
        script.onerror = reject;
        document.body.append(script);
      });
    }
  } catch (_error) {
    document.getElementById("saveState").textContent = "The notebook could not finish loading. Check your connection and reopen the page before starting your report.";
  }
})();
