---
title: Changing the Form Used by Existing Items in a Folder
keywords: olfm10.chm3077111
f1_keywords:
- olfm10.chm3077111
ms.prod: outlook
ms.assetid: cfe2d03e-68ff-e16a-5daf-5eb8bdef17e2
ms.date: 06/08/2017
---


# Changing the Form Used by Existing Items in a Folder

In some cases you may need to change the form associated with items that are already in a folder. This is often necessary after importing items, or if you create a custom form after you have already created items based on a standard Outlook form.

The Message Class field cannot be directly changed using the Outlook user interface, but you can use VBScript, Visual Basic, or Visual Basic for Applications to change the Message Class field.

The following Automation code can be used as a basis for developing your own solution. This code assumes that the name of the new form is MyForm. It will change all contacts in your default contacts folder so that they will use MyForm.




```vb
Sub ChangeMessageClass() 
Set olNS = Application.GetNameSpace("MAPI") 
Set ContactsFolder = _ 
 olNS.GetDefaultFolder(olFolderContacts) 
Set ContactItems = ContactsFolder.Items 
 
For Each Itm in ContactItems 
 If Itm.MessageClass <> "IPM.Contact.MyForm" Then 
 Itm.MessageClass = "IPM.Contact.MyForm" 
 Itm.Save 
 End If 
Next 
End Sub
```


 **Note**  If you want to use a folder other than a default folder, use the  **[Folders](folders-object-outlook.md)** collection object to refer to any folder that is available in your Folder List.


