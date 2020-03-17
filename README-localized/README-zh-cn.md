---
page_type: sample
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
本示例介绍如何设置电子邮件上的属性并将该属性存储在 Exchange 服务器上，以便在下次返回该项时可以检索它。例如，如果 Outlook 的邮件加载项会将联系人添加到外部联系人数据库中，则可在某项上设置一个属性来显示已添加某个联系人，以免被提示再次添加相同的联系人。

该项对象上的 [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) 方法将返回 [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) 对象，后者会包含并管理你为某个项存储的自定义属性。加载自定义属性后，可执行下列操作：

* 使用 [get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) 方法和 [set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) 方法来读取和写入自定义属性。 
* 使用 [remove](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) 方法来删除已创建的自定义属性。 
* 使用 [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) 方法将所做的任何更改保存回 Exchange 服务器。 

你必须调用 [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) 方法来将属性存储在 Exchange 服务器上；否则，在当前项发生更改时，将丢弃你所做的所有更改。

示例 UI 有三个页面：一个用于设置自定义属性的键和值，一个用于检索自定义属性的值，还有一个用于删除自定义属性或将所做的更改保存到 Exchange 服务器。

JavaScript 文件中包含针对 UI 中的各个按钮的单击处理程序，通过在 [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) 对象上使用相应方法来获取、设置、删除和保存自定义属性。loadCustomPropertiesAsync 方法的回调函数中设置了局部 Boolean 变量 customPropertiesAreLoaded，以显示已加载自定义属性对象。这些处理程序将检查此值，从而在系统在 [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) 对象上调用函数前，确保该对象是可用的。 

*先决条件*

此示例要求如下：

* Visual Studio 2012，包含 Office 相关应用的项目模板。 
* 运行至少具有一个电子邮件帐户或 Office 365 开发人员帐户的 Exchange 2013 的计算机。你可以[参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
* 熟悉 JavaScript 编程和 Web 服务。 
* Internet Explorer 9 或 Internet Explorer 10 预览版。 

*示例的主要组件*

本示例解决方案包含以下文件：

* CustomProperties 项目 
  * CustomProperties.xml – Outlook 邮件加载项的清单文件。 
* CustomPropertiesWeb 项目
  * Home.html – Outlook 邮件加载项的 HTML 用户界面。 
  * Home.js – 处理请求和使用 Exchange Web 服务 (EWS) 请求的 JavaScript 文件。 
  * Scripts\Lib – Outlook 邮件加载项和 Outlook Web App API。 


*配置示例*

用户收件箱中的任何电子邮件均会激活邮件加载项。可以在运行本示例之前向测试帐户发送一封或多封电子邮件，以便更轻松地测试该加载项。

*生成示例*

按 F5 生成并部署示例应用程序。完成以下任务以部署应用程序：

1. 通过为 Exchange 2013 服务器提供电子邮件地址和密码连接至 Exchange 帐户。 
2. 允许服务器配置该电子邮件帐户。 

*运行并测试示例*

你将在生成和部署示例时由 Visual Studio 启动的 Web 浏览器中运行并测试示例。

如果你在使用默认自签名证书的 Exchange 服务器上运行本示例，则在 Web 浏览器启动时，将会收到一条证书错误消息。通过查看 Web 地址确认浏览器打开的是正确的 URL 之后，可单击“继续转到此网站”以启动 Outlook Web App。

请按照以下步骤运行该示例：

1. 通过输入帐户名称和密码登录电子邮件帐户。 
2. 选择收件箱中的一封邮件。 
3. 等待应用栏出现在邮件上方。 
4. 在应用栏中，单击“自定义属性”。 
5. 出现“自定义属性”邮件加载项后，在文本框中键入属性名称和值，然后单击“保存”按钮来保存属性值。 
6. 单击“获取”，键入属性名称，然后单击“获取”按钮来检索属性。 
7. 单击“管理”，然后单击“保持”来将存储的属性保存到 Exchange 服务器，或者键入某个属性名称并单击“删除”来从存储内容中删除相应属性。 

*疑难解答*

以下是当你使用 Outlook Web App 测试 Outlook 的邮件加载项时可能发生的常见错误：

* 选中邮件后，没有出现应用栏。如果发生此情况，请通过在 Visual Studio 窗口中依次选择“调试”、“停止调试”来重启应用程序，然后按 F5 重新生成并部署应用。 
* 部署和运行应用时，可能不会取用对 JavaScript 代码所做的更改。如果更改未被取用，请清除 Web 浏览器上的缓存，方法是依次选择“工具”、“Internet 选项”，然后单击“删除…”按钮。删除临时 Internet 文件，然后重启应用。 

*其他资源*

* [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [CustomProperties 对象](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [loadCustomPropertiesAsync 方法](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [get 方法](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [set 方法](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [remove 方法](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [saveAsync 方法](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
