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
このサンプルでは、電子メール メッセージのプロパティを設定する方法と、そのプロパティを Exchange サーバーに保存して、次にアイテムが返されたときにプロパティを取得できるようにする方法を示します。たとえば、Outlook 用のメール アドインが外部連絡先データベースに連絡先を追加した場合、アイテムにプロパティを設定して、同じ連絡先をもう一度追加するように求められないように、連絡先が追加されたことを示すことができます。

アイテム オブジェクトの [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) メソッドは、アイテムに対して保存したプロパティを格納および管理する [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) オブジェクトを返します。カスタム プロパティを読み込むと、次の操作を行うことができます。

* [get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) メソッドと [set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) メソッドを使用して、カスタム プロパティを読み取り/書き込みます。 
* [remove](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) メソッドを使用して、作成したカスタム プロパティを削除します。 
* [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) メソッドを使用して、Exchange サーバーに行った変更を永続化します。 

Exchange サーバーにプロパティを保存するには、[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) メソッドを呼び出す必要があります。そうしないと、現在のアイテムが変更されたときに、行ったすべての変更が破棄されます。

サンプル UI には、カスタム プロパティのキーおよび値を設定する、カスタム プロパティの値を取得する、カスタム プロパティを削除するか、Exchange サーバーに行った変更を永続化するための 3 つのページが含まれます。

JavaScript ファイルには、UI のボタン用のクリック ハンドラーが含まれており、[CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) オブジェクトの対応するメソッドを使用して、カスタム プロパティの取得、設定、削除、保存を行うことができます。ローカル Boolean 変数 customPropertiesAreLoaded は、loadCustomPropertiesAsync メソッドのコールバック関数で設定され、カスタム プロパティが読み込まれたことを示します。ハンドラーは、オブジェクトの関数を呼び出す前に、この値をチェックして、[CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) オブジェクトが使用可能であることを確認します。 

*前提条件*

このサンプルを実行するには次のものが必要です。

* Office プロジェクトテンプレート用のアプリを含む Visual Studio 2012。 
* 少なくとも 1 つのメール アカウントまたは Office 365 開発者アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer プログラムに参加すると、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
* JavaScript プログラミングと Web サービスに関する知識。 
* Internet Explorer 9 または Internet Explorer 10 Preview。 

*サンプルの主な構成要素*

このサンプル ソリューションに含まれるファイルは次のとおりです。

* CustomProperties プロジェクト 
  * CustomProperties.xml – Outlook 用メール アドインのマニフェスト ファイル。 
* CustomPropertiesWeb プロジェクト
  * Home.html – Outlook 用メール アドインの HTML ユーザー インターフェイス。 
  * Home.js – Exchange Web サービス (EWS) 要求を処理および使用する JavaScript ファイル。 
  * Scripts\\Lib – Outlook および Outlook Web App API 用メール アドイン。 


*サンプルを構成する*

メール アドインは、ユーザーの受信トレイのすべてのメール メッセージで有効になります。サンプルを実行する前に、1 つまたは複数のメール メッセージをテスト アカウントに送信しておくと、アドインを簡単にテストできます。

*サンプルをビルドする*

F5 キーを押して、サンプル アプリケーションをビルドおよび展開します。次のタスクを完了して、アプリケーションを展開します。

1. Exchange 2013 Server 用のメール アドレスとパスワードを入力して Exchange アカウントに接続します。 
2. サーバーがメール アカウントを構成できるようにします。 

*サンプルを実行してテストする*

サンプルをビルドして展開するとき Visual Studio によって開始された Web ブラウザーでサンプルを実行してテストします。

既定の自己署名証明書を使用している Exchange サーバーでサンプルを実行している場合、Web ブラウザーが起動するとき、証明書エラーが発生します。Web ブラウザーが正しい URL を開いていることを Web アドレスを見て確認したら、[この Web サイトに進む] をクリックして、Outlook Web App を起動します。

次の手順に従ってサンプルを実行します。

1. アカウント名とパスワードを入力して、メール アカウントにログオンします。 
2. 受信トレイのメッセージを選択します。 
3. メッセージにアプリ バーが表示されるまで待ちます。 
4. アプリ バーで、[カスタム プロパティ] をクリックします。 
5. [カスタム プロパティ] メール アドインが表示されたら、テキスト ボックスにプロパティの名前と値を入力し、[保存] ボタンをクリックしてプロパティ値を保存します。 
6. [取得] をクリックし、プロパティ名を入力して、[取得] ボタンをクリックしてプロパティを取得します。 
7. [管理] をクリックし、[保持] をクリックして、Exchange サーバーに保存したプロパティを保存するか、プロパティ名を入力して [削除] をクリックし、ストレージからプロパティを削除します。 

*トラブルシューティング*

Outlook Web App を使用して Outlook のメール アドインをテストするときに発生する可能性がある一般的なエラーは次のとおりです。

* メッセージが選択されているときに、アプリ バーが表示されない。この問題が発生した場合は、Visual Studio ウィンドウで、[デバッグ]、[デバッグの停止] の順に選択してアプリケーションを再起動し、次に F5 キーを押して、アプリをリビルドして展開します。 
* アプリの展開と実行時に JavaScript コードの変更が認識されない場合がある。変更が認識されない場合は、[ツール]、[インターネット オプション] の順に選択し、[削除…] ボタンをクリックして Web ブラウザーのキャッシュをクリアします。インターネット一時ファイルを削除してからアプリを再起動します。 

*その他のリソース*

* [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [CustomProperties オブジェクト](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [loadCustomPropertiesAsync メソッド](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [get メソッド](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [set メソッド](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [remove メソッド](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [saveAsync メソッド](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
