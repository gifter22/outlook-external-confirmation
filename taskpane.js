Office.initialize = function (reason) {
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      const internalDomain = "@yourcompany.com";
      let externalFound = false;

      for (let i = 0; i < recipients.length; i++) {
        const email = recipients[i].emailAddress.toLowerCase();
        if (!email.endsWith(internalDomain)) {
          externalFound = true;
          break;
        }
      }

      if (externalFound) {
        const confirmSend = confirm("You are sending this email to an external recipient. Do you want to proceed?");
        if (!confirmSend) {
          Office.context.mailbox.item.notificationMessages.addAsync("externalWarning", {
            type: "errorMessage",
            message: "Email sending cancelled by user."
          });
          Office.context.ui.messageParent("cancel");
        } else {
          Office.context.ui.messageParent("ok");
        }
      }
    }
  });
};
