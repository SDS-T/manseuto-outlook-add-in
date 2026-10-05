/*
 * Mansueto Ventures - Outlook Marketing / RFP Intake
 *
 * Event: OnMessageRecipientsChanged
 * Trigger: marketing@mansueto.com is present in the To field
 *
 * Kept to ECMAScript 2016-compatible syntax for older classic Outlook builds.
 */

/* global Office */

var TARGET_EMAIL = "marketing@mansueto.com";

function onMessageRecipientsChangedHandler(event) {
  var item = Office.context.mailbox.item;

  if (!item || !item.to) {
    event.completed();
    return;
  }

  item.to.getAsync(function (recipientResult) {
    if (recipientResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Mansueto RFP Intake: unable to read To recipients.");
      event.completed();
      return;
    }

    var recipients = recipientResult.value || [];
    var foundTarget = false;
    var i;
    var address;

    for (i = 0; i < recipients.length; i += 1) {
      address = recipients[i] && recipients[i].emailAddress;

      if (
        address &&
        address.toLowerCase().trim() === TARGET_EMAIL.toLowerCase()
      ) {
        foundTarget = true;
        break;
      }
    }

    if (!foundTarget) {
      event.completed();
      return;
    }

    checkForExistingTemplate(event);
  });
}

function checkForExistingTemplate(event) {
  var item = Office.context.mailbox.item;

  item.body.getTypeAsync(function (typeResult) {
    if (typeResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Mansueto RFP Intake: unable to determine body type.");
      event.completed();
      return;
    }

    var bodyType = typeResult.value;

    item.body.getAsync(bodyType, function (bodyResult) {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Mansueto RFP Intake: unable to read message body.");
        event.completed();
        return;
      }

      var currentBody = bodyResult.value || "";

      if (currentBody.toUpperCase().indexOf("RFP INTAKE") !== -1) {
        event.completed();
        return;
      }

      insertIntakeTemplate(event, bodyType);
    });
  });
}

function insertIntakeTemplate(event, bodyType) {
  var item = Office.context.mailbox.item;
  var template = getTextTemplate();

  if (bodyType === Office.CoercionType.Html) {
    template = getHtmlTemplate();
  }

  item.body.prependAsync(
    template,
    { coercionType: bodyType },
    function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(
          "Mansueto RFP Intake: unable to insert template. " +
            result.error.message
        );
      }

      event.completed();
    }
  );
}

function getHtmlTemplate() {
  return (
    "<p><b><u>RFP INTAKE</u></b></p>" +
    "<p><b>RFP (Y/N):</b>&nbsp;</p>" +
    "<p><b>Account:</b>&nbsp;</p>" +
    "<p><b>Deliverable Needed (formats):</b>&nbsp;</p>" +
    "<p><b>Brand being activated (Inc. / FC / Both):</b>&nbsp;</p>" +
    "<p><b>Due Date:</b>&nbsp;</p>" +
    "<p><b>Goals &amp; Objectives:</b>&nbsp;</p>" +
    "<p><b>Budget:</b>&nbsp;</p>" +
    "<p><b>General ideas or thoughts on tactics that are appropriate:</b>&nbsp;</p>" +
    "<p><b>Splits should be included:</b>&nbsp;</p>" +
    "<p><b>Your planner in RFP requests:</b>&nbsp;</p>" +
    "<p>&nbsp;</p>"
  );
}

function getTextTemplate() {
  return (
    "RFP INTAKE\\n\\n" +
    "RFP (Y/N):\\n\\n" +
    "Account:\\n\\n" +
    "Deliverable Needed (formats):\\n\\n" +
    "Brand being activated (Inc. / FC / Both):\\n\\n" +
    "Due Date:\\n\\n" +
    "Goals & Objectives:\\n\\n" +
    "Budget:\\n\\n" +
    "General ideas or thoughts on tactics that are appropriate:\\n\\n" +
    "Splits should be included:\\n\\n" +
    "Your planner in RFP requests:\\n\\n"
  );
}

Office.actions.associate(
  "onMessageRecipientsChangedHandler",
  onMessageRecipientsChangedHandler
);
