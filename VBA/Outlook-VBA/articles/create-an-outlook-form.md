---
title: Create an Outlook Form
ms.prod: outlook
ms.assetid: c2674dd0-f033-ecea-3262-8b591acab784
ms.date: 06/08/2017
---


# Create an Outlook Form

You must base all custom forms in Microsoft Outlook on standard forms. Some of the default pages of these forms can be customized. If you do not want to use the functionality in a default form that Outlook provides, you can customize the form by hiding, adding, or replacing portions of pages or entire pages, or by replacing entire standard forms.

There are a few things to consider when you select the standard Outlook form for your custom form solution:

- Routing versus folder-based solutions
    
    If you plan to distribute the custom form by e-mail, choose the standard mail message form. If you plan to post the custom form in a public folder, choose the standard post form.
    
- Built-in fields on the standard form
    
    Select the form for the type of item that has fields best suited to the needs of your custom solution. Each type of Outlook item has a set of fields built into it. For example, to see all of the fields that are available in an e-mail message, click  **All Mail Fields** in the **Field Chooser**. For more information, see  [Using the Field Chooser](using-the-field-chooser.md).
    
- Extent of customization
    
    When you select a standard form, consider the extent to which you want to customize the form. Most standard forms have more than one page on the form, but only the pages on the mail and post forms, and the  **General** page on the contact form are customizable. To change many of the existing standard form pages, you can:
    
      - Hide the existing page on the form, create a new page, and add fields or controls to that page. 
    
  - Use additive form regions to extend the user interface on the default form or to add an extra page to a standard form. 
    
  - Use replacement form regions to replace a default page or an entire standard form.
    


## To design an Outlook form


1. On the  **Developer** tab, in the **Custom Forms** group, click **Design a Form**, and then select the standard form on which to base your custom form. 
    
2. Add the fields, controls, and code that you want to your new form. For more information, see  [Using Fields with Controls](using-fields-with-controls.md),  [Using Visual Basic with Outlook](using-visual-basic-with-outlook.md), and  [How to: Create a Form Region](create-a-form-region.md).
    
3. Set form attributes for the custom form. 
    
4. Publish the form. (For more information, see  [How to: Publish a Form](publish-a-form.md).)
     
	|**Note**|
	|:-----|  
	|<ul><li>To make the custom form available so that you or other users can create new items in a folder, you must publish the form to the folder. If you want the form to be available to other users, publish the form to a public folder so that it is available to users who have permissions to that folder. If you want the form to be available only to you, publish it in a personal folder.</li><li>Form regions cannot be published to the server; you must deploy them by using an add-in. For more information, see [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md ).</li></ul>|




