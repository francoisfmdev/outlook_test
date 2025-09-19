/* global Office */
Office.onReady(() => {
  console.log("Add-in DEV prêt ✅");
});

function onTestButtonClick(event) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ok",
      {
        type: "informationalMessage",
        message: "Tout est OK ✅ (DEV)",
        persistent: false
      },
      () => event.completed()
    );
  } catch (e) {
    console.error("Erreur add-in:", e);
    event.completed();
  }
}

// Expose la fonction pour Outlook
if (typeof window !== "undefined") {
  window.onTestButtonClick = onTestButtonClick;
}
