/* global Office */
Office.onReady(() => {
  // Outlook Desktop chargera ce script depuis GitHub Pages (HTTPS)
});

function onTestButtonClick(event) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ok",
      {
        type: "informationalMessage",
        message: "Tout est OK ✅ (DEV GitHub Pages)",
        icon: "icon16",
        persistent: false
      },
      () => event.completed()
    );
  } catch (e) {
    // Toujours terminer l'action
    event.completed();
  }
}

// Exposer la fonction pour Add-in Commands
if (typeof window !== "undefined") {
  window.onTestButtonClick = onTestButtonClick;
}
