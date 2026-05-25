// Register the service worker for PWA offline support

export function register() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      const swUrl = `${process.env.PUBLIC_URL}/service-worker.js`;
      navigator.serviceWorker
        .register(swUrl)
        .then(reg => console.log("SW registered:", reg.scope))
        .catch(err => console.log("SW registration failed:", err));
    });
  }
}

export function unregister() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.ready
      .then(reg => reg.unregister())
      .catch(err => console.error(err.message));
  }
}
