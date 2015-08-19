// Add any initialization logic to this function.
var customProperties;
var customPropertiesAreLoaded = false;

Office.initialize = function (reason) {
    initApp();
};

// Initializes the mail app for Outlook.
function initApp() {
    Office.context.mailbox.item.loadCustomPropertiesAsync(loadCustomPropertiesComplete);
    $("#footer").hide();
};

// Function called after the local custom properties are loaded from the Exchange server.
function loadCustomPropertiesComplete(asyncResult) {
    customProperties = asyncResult.value;
    customPropertiesAreLoaded = true;
};

// Gets the value of the custom property with the specified key.
function getCustomProperty(key) {
    var result = null;
    if (customPropertiesAreLoaded) {
        result = customProperties.get(key);
    } else {
        showToast("Load error", "Custom properties are not loaded.");
    }
    return result;
}

// Sets the value of a custom property in local storage.
function setCustomProperty(key, value) {
    if (customPropertiesAreLoaded) {
        customProperties.set(key, value);
    } else {
        showToast("Load error", "Custom properties are not loaded.");
    }
}

// Gets a property from local storage.
function GetProperty() {
    try {
        var propertyName = document.getElementById("storedProperty");
        var propertyValue = getCustomProperty(propertyName.value);

        if (propertyValue == null) {
            showToast("Error", "No property with that name.");
        }
        else {
            showToast("Info Retrieved", "Key: "+propertyName.value+", value: " + propertyValue);
        }
    }
    catch (err) {
        showToast(err.name, err.message);
    }
}

// Saves local changes to properties to the Exchange server.
function PersistProperties() {
    if (customPropertiesAreLoaded) {
        customProperties.saveAsync(savePropertiesCallback);
    }
}

// Callback function for saving properties to the Exchange server.
function savePropertiesCallback(asyncResult) {
    showToast("Saved", "Properties saved to Exchange server.");
}

// Removes a property in local storage.
function RemoveProperty() {
    try {
        var manageName = document.getElementById("storedProperty");
        var manageValue = getCustomProperty(manageName.value);

        if (manageValue == null) {
            showToast("Error", "No property with that name.");
        }
        else {
            if (customPropertiesAreLoaded) {
                customProperties.remove(manageName.value);
                showToast("Removed", "Property " + manageName.value + " removed.");
            }
        }
    }
    catch (err) {
        showToast(err.name, err.message);
    }
}

// Saves a property to local storage.
function SaveProperty() {
    try {
        var newProperty = document.getElementById("newProperty");
        var newValue = document.getElementById("newValue");

        setCustomProperty(newProperty.value, newValue.value);
    }
    catch (err) {
        showToast(err.name, err.message);
    }
};

// Displays the toast for 10 seconds.
function showToast(title, message) {

    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;

    $("#footer").show("slow");

    window.setTimeout(function () { $("#footer").hide("slow") }, 10000);
};

// Switches the page that is visible in the interface.
function switchPage(node) {
    var savePage = document.getElementById("saveui");
    var getPage = document.getElementById("getui");
    var managePage = document.getElementById("manageui");
    var pageTitle = document.getElementById("page");

    switch (node) {
        case "Save":
            getPage.setAttribute("class", "hiddenPage");
            managePage.setAttribute("class", "hiddenPage");
            savePage.setAttribute("class", "displayedPage");
            break;
        case "Get":
            savePage.setAttribute("class", "hiddenPage");
            managePage.setAttribute("class", "hiddenPage");
            getPage.setAttribute("class", "displayedPage");
            break;
        default:
            savePage.setAttribute("class", "hiddenPage");
            getPage.setAttribute("class", "hiddenPage");
            managePage.setAttribute("class", "displayedPage");
            break;
    }
};