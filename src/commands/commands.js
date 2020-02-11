var config = {};
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function(reason) {
  mailboxItem = Office.context.mailbox.item;
};

// Add any ui-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "Addin-error",
    {
      type: "errorMessage",
      message: error
    },
    function(result) {}
  );
}

//On mail send
async function validateClassfication(event) {
  //Get config
  config.default_classfication = Office.context.roamingSettings.get("default_classfication");
  //Check if classfied
  var is_classified = Office.context.roamingSettings.get("quick_classfication");
  //Check if  Default available
  if (config.default_classfication) {
    try {
      await setClassfication(config.default_classfication);
      //Allow
      setTimeout(function() {
        event.completed({ allowEvent: true });
      }, 2000);
    } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      showNotification(error, "error", "validateClassfication");
    }
  } else if (is_classified) {
    //Mark classfication set as true
    remove_quick_classfication();
      //Allow
      setTimeout(function() {
        event.completed({ allowEvent: true });
      }, 2000);
  } else {
    //Show error
    showNotification(
      " Please set a classification for this email.[validateClassfication]",
      "error",
      "validateClassfication"
    );
    event.completed({ allowEvent: false });
  }
}

//Set all four type of commands function
async function setInternal(event) {
  try {
    await setClassfication("Internal");
    //Mark classfication set as true
    set_quick_classfication();
    //Colplete event
    event.completed();
  } catch (error) {
    //Mark classfication set as true
    remove_quick_classfication();
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    showNotification(error, "error", "setClassfication");
    event.completed();
  }
}
async function setScrete(event) {
  try {
    await setClassfication("Screte");
    //Mark classfication set as true
    //Mark classfication set as true
    set_quick_classfication();
    //Colplete event
    event.completed();
  } catch (error) {
    //Mark classfication set as true
    remove_quick_classfication();
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    showNotification(error, "error", "setClassfication");
    event.completed();
  }
}
async function setConfidential(event) {
  try {
    await setClassfication("Confidential");
    //Mark classfication set as true
    //Mark classfication set as true
    set_quick_classfication();
    //Colplete event
    event.completed();
  } catch (error) {
    //Mark classfication set as true
    remove_quick_classfication();
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    showNotification(error, "error", "setClassfication");
    event.completed();
  }
}
async function setPublic(event) {
  try {
    await setClassfication("Public");
    //Mark classfication set as true
    set_quick_classfication();
    //Colplete event
    event.completed();
  } catch (error) {
    event.completed();
    //Mark classfication set as true
    remove_quick_classfication();
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    showNotification(error, "error", "setClassfication");
  }
}

async function setClassfication(type) {
  var user_email = Office.context.mailbox.userProfile.emailAddress;
  var timezone = Office.context.mailbox.userProfile.timeZone;
  var body =
    "This email is classified as " +
    type +
    ". By " +
    user_email +
    " at " +
    new Date().toLocaleDateString() +
    " " +
    timezone;
  var subject = "This email is classified [" + type + "]";
  var Footer = "This email is classified [" + type + "]";
  // Get the default gist content and insert.
  try {
    //Set Category
    var masterCategoriesToAdd = [
      {
        displayName: type,
        color: Office.MailboxEnums.CategoryColor.Preset0
      }
    ];
    Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        Office.context.mailbox.item.categories.addAsync([type], function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            showNotification("Successfully added categories", "success", "category");
          } else {
            showNotification(
              "categories.addAsync call failed with error: " + asyncResult.error.message,
              "error",
              "category"
            );
          }
        });
      } else {
        Office.context.mailbox.item.categories.addAsync([type], function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            showNotification("Successfully added categories", "success", "category");
          } else {
            showNotification(
              "categories.addAsync call failed with error: " + asyncResult.error.message,
              "error",
              "category"
            );
          }
        });
        showNotification("Unable to set the category MasterCategories.addAsync"+ asyncResult.error.message, "error", "Createcategory");
      }
    });

    //Set Subject
    Office.context.mailbox.item.subject.setAsync(subject, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        showNotification("Successfully added subject", "success", "subject");
      } else {
        showNotification("Unable to set the subject: " + asyncResult.error.message, "error", "subject");
      }
    });

    //set body
    //Set Body (top)
    Office.context.mailbox.item.body.prependAsync(body, { coercionType: "html" }, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        showNotification("Successfully added body", "success", "body");
        //Footer nd body
        Office.context.mailbox.item.body.getAsync("html", function callback(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            var new_body = asyncResult.value + "</br></br></br><h4>" + Footer + "</h4>";
            Office.context.mailbox.item.body.setAsync(new_body, { coercionType: "html" }, function callback(asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                showNotification("Successfully added footer", "success", "footer");
              } else {
                showNotification("body.setAsync call failed with error: " + asyncResult.error.message, "error", "footer");
              }
            });
          } else {
            showNotification("Unable to set the footer: " + asyncResult.error.message, "error", "footer");
          }
        });
      } else {
        showNotification("Unable to set the body: " + asyncResult.error.message, "error", "body");
      }
    });

  } catch (err) {
    //Show error on body (only dev)
    Office.context.mailbox.item.body.setAsync("<b>" + err + "</b>", { coercionType: "html" }, function callback(
      result
    ) {});
    throw TypeError("SetClassification Error");
  }
}

// Display notifications in message banner at the top of the task pane.
function showNotification(content, type, name) {
  if (type && type == "error") {
    Office.context.mailbox.item.notificationMessages.addAsync(name, {
      type: "errorMessage",
      message: content
    });
  } else {
    Office.context.mailbox.item.notificationMessages.addAsync(name, {
      type: "informationalMessage",
      message: content,
      icon: "iconid",
      persistent: false
    });
  }
}
function set_quick_classfication() {
  //Get value
  Office.context.roamingSettings.set("quick_classfication", true);
  Office.context.roamingSettings.saveAsync();
}
function remove_quick_classfication(value) {
  //Remove value
  Office.context.roamingSettings.remove("quick_classfication");
  Office.context.roamingSettings.saveAsync();
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
g.setConfidential = setConfidential;
g.setPublic = setPublic;
g.setScrete = setScrete;
g.validateClassfication = validateClassfication;
