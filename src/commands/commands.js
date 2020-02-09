import { validateSubjectAndCC } from "../helper/validate_classification.js";

var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function(reason) {};

// Add any ui-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "github-error",
    {
      type: "errorMessage",
      message: error
    },
    function(result) {}
  );
}

//On mail send
function validateClassfication() {
  //Get Default
  config = getConfig();
  //Check if  Default available
  if (config.default_classfication) {
  } else {
  }
}

//Set all four type of commands function
function setInternal(event) {
  setClassfication("Internal");
}
function setScrete(event) {
  setClassfication("Screte");
}
function setConfidential(event) {
  setClassfication("Confidential");
}
function setPublic(event) {
  setClassfication("Public");
}

function setClassfication(type) {
  // Get the default gist content and insert.
  try {
    Office.context.mailbox.item.body.setSelectedDataAsync(type, { coercionType: Office.CoercionType.Html }, function(
      result
    ) {
      event.completed();
    });
  } catch (err) {
    showError(err);
    event.completed();
  }
}
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

var g = getGlobal();

// The add-in command functions need to be available in global scope.
g.setInternal = setInternal;
