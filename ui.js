/* ui.js
 * UI event handlers and status management
 */

// ===== UI Elements =====
const $ = (s) => document.querySelector(s);
const fileInput = $("#file");
const goBtn = $("#go");
const statusEl = $("#status");

let loadedFile = null;

// ===== Status Management =====
function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

// ===== Event Handlers =====
function initializeUI() {
  // File input change handler
  fileInput?.addEventListener("change", () => {
    loadedFile = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
    if (!loadedFile) {
      setStatus("Esperando archivo…");
      if (goBtn) goBtn.disabled = true;
      return;
    }
    setStatus(`Archivo listo: ${loadedFile.name}`);
    if (goBtn) goBtn.disabled = false;
  });

  // Process button click handler
  goBtn?.addEventListener("click", async () => {
    if (!loadedFile) return;
    goBtn.disabled = true;
    setStatus("Procesando…");
    try {
      const outName = await Processor.processFile(loadedFile, Processor.CONFIG);
      setStatus(`Listo ✅ Descargado: ${outName}`);
    } catch (err) {
      console.error(err);
      setStatus("Error: " + (err?.message || err));
    } finally {
      goBtn.disabled = false;
    }
  });
}

// Initialize UI when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeUI);
} else {
  initializeUI();
}

// Export for use in other modules
window.UI = {
  setStatus,
  initializeUI
};
