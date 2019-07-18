---
topic: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/19/2015 11:13:31 AM
---
# Outlook-Add-in-Store-Custom-Properties-On-Exchange-Server
This sample shows how to set a property on an email message and then store that property on your Exchange server so that you can retrieve it the next time the item is returned. For example, if your mail add-in for Outlook adds contacts to an external contacts database, you can set a property on an item to show that a contact was added so that you are not prompted to add the same contact a second time.

The  [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) method on the item object returns a  [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) object that contains and manages the custom properties that you've stored for an item. After you loaded the custom properties, you can do the following:

* Use the [get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) method and [set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) method to read and write custom properties. 
* Use the [remove](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) method to delete custom properties that you've created. 
* Use the [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) method to persist any changes that you've made back to the Exchange server. 

You must call the [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) method to store the properties on the Exchange server; otherwise, all the changes that you made are discarded when the current item is changed.

The sample UI has three pages: one to set the key and value of a custom property, one to retrieve the value of a custom property, and one to remove custom properties or to persist the changes that you make to the Exchange server.

The JavaScript file contains click handlers for buttons in the UI to get, set, remove, and save custom properties by using the corresponding methods on the [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) object. A local Boolean variable, customPropertiesAreLoaded, is set in the callback function for the  loadCustomPropertiesAsync method to show that the custom properties object is loaded. The handlers check this value to make sure that the [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) object is available before calling functions on the object. 

*Prerequisites*

This sample requires that you have the following:

* Visual Studio 2012, with the apps for Office project templates. 
* A computer running Exchange 2013 with at least one email account, or an Office 365 developer account. You can [join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup).
* Familiarity with JavaScript programming and web services. 
* Internet Explorer 9 or Internet Explorer 10 Preview. 

*Key components of the sample*

The sample solution contains the following files:

* CustomProperties project 
  * CustomProperties.xml – The manifest file for the mail add-in for Outlook. 
* CustomPropertiesWeb project
  * Home.html – The HTML user interface for the mail add-in for Outlook. 
  * Home.js – The JavaScript file that handles requesting and using the Exchange Web Services (EWS) request. 
  * Scripts\Lib – The mail add-in for Outlook and Outlook Web App API. 


*Configure the sample*

The mail add-in will be activated on any email message in the user's Inbox. You can make it easier to test the add-in by sending one or more email messages to your test account before you run the sample.

*Build the sample*

Press F5 to build and deploy the sample application. Complete the following tasks to deploy the application:

1. Connect to an Exchange account by providing the email address and password for an Exchange 2013 server. 
2. Allow the server to configure the email account. 

*Run and test the sample*

You run and test the sample in the web browser that is started by Visual Studio when you build and deploy the sample.

If you are running the sample on an Exchange server that is using the default self-signed certificate, you will receive a certificate error when the web browser starts. After you verify that the web browser is opening the correct URL by looking at the web address, you can click Continue to this Web site to start Outlook Web App.

Follow these steps to run the sample:

1. Log on to the email account by entering the account name and password. 
2. Select a message in the Inbox. 
3. Wait for the App bar to appear over the message. 
4. In the App bar, click Custom Properties. 
5. When the Custom Properties mail add-in appears, type a property name and value into the text boxes and then click the Save button to save the property value. 
6. Click Get, type a property name, and then click the Get button to retrieve a property. 
7. Click Manage, and either click Persist to save the stored properties to the Exchange server, or type a property name and click Remove to delete the property from storage. 

*Troubleshooting*

The following are common errors that can occur when you use Outlook Web App to test a mail add-in for Outlook:

* The App bar does not appear when a message is selected. If this occurs, restart the application by selecting Debug – Stop Debugging in the Visual Studio window, then press F5 to rebuild and deploy the app. 
* Changes to the JavaScript code may not be picked up when you deploy and run the app. If the changes are not picked up, clear the cache on the web browser by selecting Tools – Internet options and clicking the Delete… button. Delete the temporary Internet files and then restart the app. 

*Additional resources*

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [CustomProperties object](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [loadCustomPropertiesAsync method](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [get method](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [set method](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [remove method](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [saveAsync method](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
