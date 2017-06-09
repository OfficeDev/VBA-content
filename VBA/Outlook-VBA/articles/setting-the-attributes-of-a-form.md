---
title: Setting the Attributes of a Form
ms.prod: outlook
ms.assetid: 7c170a7b-fe1e-32be-3841-535e8f29dae4
ms.date: 06/08/2017
---


# Setting the Attributes of a Form

To set the attributes of a custom form that has form pages, use the  **Properties** and **Actions** pages in design mode. The **Properties** page includes information such as Description, Category, and Form Number that helps users to find and identify your form in the **Choose Form** dialog box. The **Actions** page includes all of the various ways that users can reply to a form, including Reply, Reply All, Forward, and so on.

When you customize a form with form regions, you set the attributes for a form region inside the form region manifest file. The values that you specify on the properties and actions pages in the forms designer are not saved with a form region and are not used when a form region is loaded. For more information, see  [How to: Create a Form Region](create-a-form-region.md).

 One important attribute of your form to consider is whether to send the form definition with the form. The form definition includes all the fields and the code that you add to the form. As a general rule, publish the form definition to a forms library instead of sending the form definition with the item. If you cannot publish your form to a forms library, you can select the **Send form definition with item** check box on the **Properties** page so that other users can see the form pages when they receive items that are composed by using the form.

Forms that you only intend to use once and not publish are referred to as "one-off forms." Because of security concerns with one-off forms, users might not see the form correctly when they open items sent to them with a one-off form. In this case, sending the form definition with the one-off form provides the necessary information required to display the form correctly for the users. 
If you plan to publish your form to a forms library that other users can access, such as an  **Organizational Forms Library** or a Public Folder library, you do not need to send the form definition; it is stored in the library. The form definition can add considerable size to items that are saved by using your form.
To change how users reply to your form, click the  **Actions** page. The **Actions** page lists the default Reply forms that are available. You can also add your own custom Reply forms. For example, forms based on a new e-mail message have built-in Reply, Reply to All, Forward, and Reply to Folder forms. When users receive your form, the form contains buttons and menu commands so that users can respond to the form. You can disable some or all of these default forms and set attributes that define how these Reply forms appear.

