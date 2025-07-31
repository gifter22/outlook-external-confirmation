Office.onReady(() => {
  const internalDomain = "@pmsc.com.ph"; // Change this to match your domain

  Office.context.mailbox.item.to.getAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = result.value;
      let externalFound = false;

      for (const r of recipients) {
        if (!r.emailAddress.toLowerCase().endsWith(internalDomain)) {
          externalFound = true;
          break;
        }
      }

      if (externalFound) {
        const proceed = confirm("⚠️ You are sending to external recipients. Do you want to continue?");
        document.getElementById("status").textContent = proceed
          ? "Confirmed to send."
          : "You should review recipients.";
      } else {
        document.getElementById("status").textContent = "All recipients are internal.";
      }
    } else {
      document.getElementById("status").textContent = "Failed to load recipients.";
    }
  });
});
