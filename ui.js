/* ui.js
 * UI event handlers and status management
 */

// ===== UI Elements =====
const $ = (s) => document.querySelector(s);
let fileInput, goBtn, statusEl;
let loadedFile = null;

// ===== Status Management =====
function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

// ===== Event Handlers =====
function initializeUI() {
  console.log("initializeUI");
  
  // Get UI elements when DOM is ready
  fileInput = $("#file");
  goBtn = $("#go");
  statusEl = $("#status");
  
  // File input change handler
  fileInput?.addEventListener("change", () => {
    console.log("fileInput change");
    loadedFile = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
    if (!loadedFile) {
      setStatus("Esperando archivo…");
      if (goBtn) goBtn.disabled = true;
      return;
    }
    
    // Track file upload event
    if (typeof umami !== 'undefined') {
      umami.track('file_upload', {
        file_name: loadedFile.name,
        file_size: loadedFile.size,
        file_type: loadedFile.type
      });
    }
    
    setStatus(`Archivo listo: ${loadedFile.name}`);
    if (goBtn) goBtn.disabled = false;
  });

  // Process button click handler
  goBtn?.addEventListener("click", async () => {
    if (!loadedFile) return;
    goBtn.disabled = true;
    setStatus("Procesando…");
    
    // Track file processing start event
    if (typeof umami !== 'undefined') {
      umami.track('file_process_start', {
        file_name: loadedFile.name,
        file_size: loadedFile.size
      });
    }
    
    try {
      const outName = await Processor.processFile(loadedFile, Processor.CONFIG);
      
      // Track successful file processing
      if (typeof umami !== 'undefined') {
        umami.track('file_process_success', {
          file_name: loadedFile.name,
          output_name: outName
        });
      }
      
      setStatus(`Listo ✅ Descargado: ${outName}`);
    } catch (err) {
      console.error(err);
      
      // Track file processing error
      if (typeof umami !== 'undefined') {
        umami.track('file_process_error', {
          file_name: loadedFile.name,
          error_message: err?.message || err
        });
      }
      
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
