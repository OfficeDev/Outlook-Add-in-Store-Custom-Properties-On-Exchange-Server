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
Este exemplo mostra como definir uma propriedade em uma mensagem de email e armazenar a propriedade no servidor Exchange para que você possa recuperá-la na próxima vez em que o item for retornado. Por exemplo, se seu suplemento de e-mail do Outlook adicionar contatos a um banco de dados de contatos externos, você poderá definir uma propriedade em um item para mostrar que um contato foi adicionado para que você não seja solicitado a adicionar o mesmo contato uma segunda vez.

O método [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) no objeto Item retorna um objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) que contém e gerencia as propriedades personalizadas que você armazenou para um item. Depois de carregar as propriedades personalizadas, você poderá fazer o seguinte:

* Use o método [Get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) e o método[Set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) para ler e gravar propriedades personalizadas. 
* Use o método [Remover](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) para excluir as propriedades personalizadas que você criou. 
* Use o método [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) para salvar todas as alterações que você fez no Exchange Server. 

Você deve chamar o método [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) para armazenar as propriedades no Exchange Server. Caso contrário, todas as alterações feitas serão descartadas quando o item atual for alterado.

A interface de usuário de exemplo tem três páginas: uma para definir a chave e o valor de uma propriedade personalizada, uma para recuperar o valor de uma propriedade personalizada e outra para remover propriedades personalizadas ou para salvar as alterações feitas no servidor do Exchange.

O arquivo JavaScript contém manipuladores de clique para botões na interface do usuário para obter, definir, remover e salvar propriedades personalizadas usando os métodos correspondentes no objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3). Uma variável Boolean local, customPropertiesAreLoaded, é definida na função callback do método loadCustomPropertiesAsync para mostrar que o objeto de propriedades personalizadas está carregado. Os manipuladores verificam esse valor para verificar se o objeto [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) está disponível antes de chamar funções no objeto. 

*Pré-requisitos*

Este exemplo exige que você tenha o seguinte:

* O Visual Studio 2012, com os aplicativos do Office para modelos de projeto. 
* Um computador executando o Exchange 2013 com pelo menos uma conta de e-mail ou uma conta de desenvolvedor do Office 365. Você pode [participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de um ano do Office 365](https://aka.ms/devprogramsignup).
* Familiaridade com programação em JavaScript e serviços Web. 
* Internet Explorer 9 ou Internet Explorer 10 Preview. 

*Componentes principais do exemplo*

A solução do exemplo contém os seguintes arquivos:

* Projeto CustomProperties 
  * CustomProperties.xml - O arquivo de manifesto do suplemento de e-mail do Outlook. 
* Projeto CustomPropertiesWeb
  * Home.html - A interface do usuário HTML do suplemento do Outlook. 
  * Home.js – o arquivo JavaScript que manipula a solicitação e o uso da solicitação de serviços Web do Exchange (EWS). 
  * Scripts\Lib – o suplemento de e-mail do Outlook e da API do Outlook Web App. 


*Configurar o exemplo*

O suplemento e-mail será ativado em qualquer mensagem de e-mail na caixa de entrada do usuário. Você pode facilitar o teste do suplemento enviando uma ou mais mensagens de e-mail para a sua conta de teste antes de executar o exemplo.

*Criar o exemplo*

Pressione F5 para criar e implementar o aplicativo do exemplo. Conclua as seguintes tarefas para implantar o aplicativo:

1. Conecte-se a uma conta do Exchange fornecendo o endereço de e-mail e a senha de um servidor do Exchange 2013. 
2. Permitir que o servidor configure a conta de e-mail. 

*Executar e testar o exemplo*

Você executará e testará o exemplo no navegador da Web que é iniciado pelo Visual Studio ao criar e implantar o exemplo.

Se você estiver executando o exemplo em um servidor Exchange que usa o certificado autoassinado padrão, receberá um erro de certificado quando o navegador da Web for iniciado. Depois de verificar se o navegador da web está abrindo a URL correta, verificando o endereço da Web, selecione Continuar neste site para iniciar o Outlook Web App.

Siga estas etapas para executar o exemplo:

1. No navegador, faça logon com a conta de e-mail digitando o nome e a senha da conta. 
2. Salvar uma mensagem na Caixa de entrada. 
3. Aguarde até que a barra do aplicativo seja exibida sobre a mensagem. 
4. Na barra do aplicativo, clique em Propriedades Personalizadas. 
5. Quando o suplemento de propriedades personalizadas for exibido, digite o nome e o valor de uma propriedade nas caixas de texto e, em seguida, clique no botão Salvar para salvar o valor da propriedade. 
6. Clique em Obter, digite um nome de propriedade e, em seguida, clique no botão obter para recuperar uma propriedade. 
7. Clique em Gerenciar, e clique em persistir para salvar as propriedades armazenadas no servidor do Exchange ou digite o nome da propriedade e clique em Remover para excluir a propriedade do armazenamento. 

*Solução de problemas*

Estes são erros comuns que podem ocorrer quando você usa o Outlook Web App para testar um suplemento de e-mail do Outlook:

* A barra do aplicativo não aparece quando uma mensagem é selecionada. Se isso ocorrer, reinicie o suplemento selecionando Depurar - Parar a depuração na janela do Visual Studio e, em seguida, pressione F5 para recriar e implementar o aplicativo. 
* Pode ser que as alterações no código JavaScript não sejam selecionadas quando você implementar e executar o aplicativo. Se as alterações não forem selecionadas, limpe o cache do navegador da Web selecionando Ferramentas - Opções da Internet e selecionando o botão Excluir. Exclua os arquivos temporários da Internet e reinicie o aplicativo. 

*Recursos adicionais*

* [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Objeto CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [método loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [método get](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [método set](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [método remover](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [método saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
