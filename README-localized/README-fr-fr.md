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
Cet exemple présente comment définir une propriété sur un courrier puis la stocker sur votre serveur Exchange, afin de pouvoir la récupérer lorsque l’élément est renvoyé. Par exemple, si votre complément courrier pour Outlook ajoute des contacts externes à une base de données, vous pouvez définir une propriété à l'élément pour indiquer qu’un contact a été ajouté pour ne pas être invité à ajouter le même contact une deuxième fois.

La méthode [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261) sur l’objet élément renvoie un objet [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) qui contient et gère les propriétés personnalisées stockées d'un élément. Une fois les propriétés personnalisées chargées, vous pouvez effectuer ce qui suit :

* Utilisez la méthode [obtenir](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b) et la méthode [définir](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d) pour lire et écrire des propriétés personnalisées. 
* Utilisez la méthode [supprimer](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3) pour effacer des propriétés personnalisées que vous avez créées. 
* Utilisez la méthode [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) pour maintenir des modifications que vous avez apportées au serveur Exchange. 

Vous devez appeler la méthode [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56) pour stocker les propriétés sur le serveur Exchange. Dans le cas contraire, toutes les modifications effectuées sont annulées lors de la modification de l’élément actif.

L’exemple d’interface utilisateur contient trois pages : une pour définir la clé et la valeur d’une propriété personnalisée, une pour récupérer la valeur d’une propriété personnalisée, une pour supprimer des propriétés personnalisées ou pour conserver des modifications apportées au serveur Exchange.

Le fichier JavaScript contient des gestionnaires de clic pour les boutons de l’interface utilisateur permettant d’obtenir, définir, supprimer et enregistrer des propriétés personnalisées à l’aide des méthodes correspondantes sur l’objet [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3). Une variable Boolean locale, customPropertiesAreLoaded, est définie dans la fonction de rappel pour la méthode loadCustomPropertiesAsync afin d’indiquer que l’objet des propriétés personnalisées est chargé. Les gestionnaires contrôlent cette valeur pour s’assurer que l’objet [CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3) est disponible avant d’appeler des fonctions sur l’objet. 

*Conditions préalables*

Vous devez disposer des éléments suivants pour cet exemple :

* Visual Studio 2012 et les applications pour les modèles de projets Office. 
* Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte de développeur Office 365. Vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
* Être familiarisé avec les services web et la programmation JavaScript. 
* Internet Explorer 9 ou Internet Explorer 10 en préversion. 

*Composants clés de l’exemple*

La solution d’exemple contient les fichiers suivants :

* projet CustomProperties 
  * CustomProperties.xml : le fichier manifeste pour le complément de courrier pour Outlook. 
* Projet CustomPropertiesWeb
  * Home.html : l'interface utilisateur HTML pour le complément de messagerie pour Outlook. 
  * Home.js : le fichier JavaScript qui gère les demandes et l’utilisation de la demande Exchange Web Services (EWS). 
  * Scripts\Lib : le complément courrier pour Outlook et l’API d'Outlook Web App. 


*Configurer l’exemple*

Le complément messagerie sera activé sur tout courrier électronique figurant dans la boîte de réception de l’utilisateur. Vous pouvez simplifier le test du complément en envoyant un ou plusieurs e-mails à votre compte de test avant d’exécuter l’exemple.

*Créer l’exemple*

Appuyez sur F5 pour créer et déployer l'application de l’exemple. Achevez les tâches suivantes pour déployer l’application :

1. Connectez-vous à un compte Exchange en fournissant l’adresse de messagerie et le mot de passe d’un serveur Exchange 2013. 
2. Autorisez le serveur à configurer le compte de courrier. 

*Exécuter et tester le complément*

Vous exécutez et testez l’exemple dans le navigateur web démarré par Visual Studio lorsque vous créez et déployez l’exemple.

Si vous exécutez l’exemple sur un serveur Exchange qui utilise le certificat auto-signé par défaut, vous recevrez une erreur de certificat lorsque le navigateur web démarre. Après avoir vérifié que le navigateur ouvre l’URL correcte en contrôlant l’adresse web, sélectionnez Continuer sur ce site web pour démarrer Outlook Web App.

Procédez comme suit pour exécuter l’exemple :

1. Connectez-vous au compte de messagerie en entrant le nom du compte et le mot de passe. 
2. Sélectionnez un message dans la boîte de réception. 
3. Patientez jusqu’à ce que la barre App s’affiche au-dessus du message. 
4. Dans la barre App, cliquez sur Propriétés personnalisées. 
5. Lorsque les Propriétés personnalisées de courrier apparaissent, tapez un nom et une valeur de propriété dans les zones de texte, puis cliquez sur le bouton Enregistrer pour sauvegarder la valeur de propriété. 
6. Cliquez sur Obtenir, tapez un nom de propriété, puis cliquez sur le bouton Obtenir pour récupérer une propriété. 
7. Cliquez sur Gérer, puis cliquez sur Conserver pour enregistrer les propriétés stockées sur le serveur Exchange, ou tapez un nom de propriété, puis cliquez sur Supprimer pour effacer la propriété de l’espace de stockage. 

*Résolution de problèmes*

Voici quelques erreurs courantes qui peuvent se produire lors de l’utilisation d’Outlook Web App pour tester un complément de messagerie pour Outlook :

* La barre App n'apparaît pas lorsque le message est sélectionné. Si c’est le cas, redémarrez l’application en sélectionnant Déboguer – Arrêter le débogage dans la fenêtre Visual Studio, puis appuyez sur F5 pour regénérer et déployer l'application. 
* Les modifications apportées au code JavaScript peuvent ne pas être prises en compte lorsque vous déployez et exécutez l'application. Si les modifications ne sont pas prises en compte, effacez le cache du navigateur web en sélectionnant Outils – Options Internet, puis en cliquant sur le bouton Supprimer... Supprimez les fichiers Internet temporaires et redémarrez ensuite l'application. 

*Ressources supplémentaires*

* [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Objet CustomProperties](http://msdn.microsoft.com/library/%2095a69bd6-c4dc-429a-8b27-e2b68f74f3e3)
* [Méthode loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261)
* [Obtenir la méthode](http://msdn.microsoft.com/library/3ab90551-138a-482d-9d93-4cdb20db193b)
* [Définir la méthode](http://msdn.microsoft.com/library/03a8b253-b681-4a09-b828-80d9cf46ca9d)
* [Supprimer la méthode](http://msdn.microsoft.com/library/01983beb-766f-4308-9e23-e840e950f7e3)
* [méthode saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56)



Ce projet a adopté le [Code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
