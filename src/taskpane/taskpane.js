export function getConfig() {
  var config = {};
  config.default_classfication = Office.context.roamingSettings.get("default_classfication");
  return config;
}

export async function setConfig() {
  //Get value
  let value = $("input[type='radio']:checked").val();
  Office.context.roamingSettings.set("default_classfication", value);

  Office.context.roamingSettings.saveAsync();
}

Office.onReady(info => {
  //Get Option and data
  $.getJSON("/data.json", function(data) {
    //Set ttile
    $(".header_text").text(data.title);
    //Set inputs
    data.inputs.forEach(function(input) {
      //check input type
      if (input.type == "radio") {
        //Check all option
        input.values.forEach(function(radio) {
          //Append html
          var option =
            '<p><input type="radio" value=' + radio.value + ' name="radio-group"><label>' + radio.label + "</p>";
          //Add to main screen
          $("#app-body").append(option);
        });
      }
    });
    //Set properties
    $("input[type='radio']").click(() => tryCatch(setConfig));
    //Initiate
    tryCatch(setConfig);
  });
});

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    errorHandler(error);
  }
}

function errorHandler(error) {
  showNotification(error);
}

// Display notifications in message banner at the top of the task pane.
function showNotification(content) {
  $(".error").text(content);
}
