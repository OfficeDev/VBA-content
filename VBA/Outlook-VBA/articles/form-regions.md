---
title: Form Regions
ms.prod: outlook
ms.assetid: 66e80f83-60db-e3b1-47e9-097f855f6512
ms.date: 06/08/2017
---


# Form Regions

Form regions are custom pieces of user interface that can be used to customize a standard form. There are four types of form regions:


- adjoining form region
    
- separate form region
    
- replacement form region
    
- replace-all form region
    



 An adjoining form region is a form region added to the bottom of the default page of a standard Outlook form, and a separate form region is an individual page added to a standard Outlook form. Adjoining form regions and separate form regions are additive form regions. You can have up to 50 adjoining form regions and up to 30 separate form regions in a form.
A replacement form region is a page that replaces the default page of a standard form, and a replace-all form region replaces all pages in a standard Outlook form.
Although you can create a custom form that contains a form region without using an add-in, add-ins allow you to deploy and further extend a standard Outlook form. Add-ins customize form regions of a form in a way similar to other custom pages of a form, by adding fields from the Field Chooser and controls from the control toolbox. However, while you add code behind a custom form page using the Script Editor and Visual Basic Scripting Edition (VBScript), you add code behind a form region using an add-in. For more information on using an add-in to extend a form region, see  [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md).
Form regions allow greater flexibility in customizing and extending a standard Outlook form in the following ways:


- You can add new user interface as adjoining form regions to the default page of any standard Outlook form.
    
- If you are using adjoining or separate form regions to add user interface to a standard form, you can choose to specify the message class of the form regions as the same message class of the standard form (for example,  **IPM.Contact**), or as a custom message class derived from the standard message class (for example,  **IPM.Contact.PersonalContacts**).
    
- You can use a replacement form region to replace the default page of a standard form, or a replace-all form region to replace the entire standard form. In this case, you must specify a derived message class for the form region and register the form region for that message class.
    
- Multiple add-ins can add form regions to the same form.
    
- Administrators can more conveniently deploy a form customized with form regions by installing and running an add-in. For more information, see the Office Resource Kit Web site.
    


