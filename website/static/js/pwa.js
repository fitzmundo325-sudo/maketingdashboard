(function () {
  if (window.__pwa_initialized) return;
  window.__pwa_initialized = true;

  let deferredInstallPrompt = null;

  function isStandalone() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  }

  function isInstalled() {
    if (isStandalone()) return true;
    try {
      return window.localStorage.getItem('pwa_installed') === 'true';
    } catch (e) {
      return false;
    }
  }

  function getButtons() {
    return Array.from(document.querySelectorAll('.pwa-install-btn'));
  }

  function setButtonState() {
    const installed = isInstalled();
    getButtons().forEach((button) => {
      button.classList.toggle('hidden', installed);
      button.disabled = installed;
      button.setAttribute('aria-disabled', installed ? 'true' : 'false');
    });
  }

  function showInstallButton() {
    getButtons().forEach((button) => {
      button.classList.remove('hidden');
      button.disabled = false;
      button.setAttribute('aria-disabled', 'false');
    });
  }

  function showInstallHelp() {
    const isApple = /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');
    const isSafari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent || '');
    const message = isApple && isSafari
      ? 'To install on Mac Safari, open File > Add to Dock. On iPhone or iPad, use Share > Add to Home Screen.'
      : 'Use your browser menu and choose Install HYG Marketing or Add to desktop.';

    if (window.toast && typeof window.toast.show === 'function') {
      window.toast.show(message, 'info', 7000);
      return;
    }

    window.alert(message);
  }

  async function installApp() {
    if (!deferredInstallPrompt) {
      showInstallHelp();
      return;
    }

    deferredInstallPrompt.prompt();
    await deferredInstallPrompt.userChoice;
    deferredInstallPrompt = null;
    setButtonState();
  }

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredInstallPrompt = event;
    showInstallButton();
  });

  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    try {
      window.localStorage.setItem('pwa_installed', 'true');
    } catch (e) {
      /* ignore */
    }
    setButtonState();
    if (window.toast && typeof window.toast.show === 'function') {
      window.toast.show('HYG Marketing installed successfully.', 'success');
    }
  });

  // Since DOMContentLoaded might have already fired, let's run setup immediately
  function initPwaButtons() {
    getButtons().forEach((button) => button.classList.add('hidden'));
    getButtons().forEach((button) => {
      // Avoid duplicate click listeners
      button.removeEventListener('click', installApp);
      button.addEventListener('click', installApp);
    });

    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js', { scope: '/' }).catch(() => {
        if (window.toast && typeof window.toast.show === 'function') {
          window.toast.show('Install support is not available in this browser session.', 'warning');
        }
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initPwaButtons);
  } else {
    initPwaButtons();
  }
})();
