/* global Office */
Office.onReady(() => {});

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
    event.completed();
  }
}

if (typeof window !== "undefined") {
  window.onTestButtonClick = onTestButtonClick;
}
