---
title: ContactItem.UserProperties Property (Outlook)
keywords: vbaol11.chm955
f1_keywords:
- vbaol11.chm955
ms.prod: outlook
api_name:
- Outlook.ContactItem.UserProperties
ms.assetid: f52b8fb8-945b-a406-b3cb-1c9dcc150184
ms.date: 06/08/2017
---


# ContactItem.UserProperties Property (Outlook)

Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.


## Syntax

 _expression_ . **UserProperties**

 _expression_ A variable that represents a **ContactItem** object.


## Example

This Visual Basic for Applications (VBA) example finds a custom property named  `LastDateContacted` for the contact 'Jeff Smith' and displays it to the user. To run this example, you need to replace 'Jeff Smith' with a valid contact name and create a user-defined property called `LastDateContacted` for the contact.


```vb
Sub FindContact() 
 
 'Finds and displays last contacted info for a contact 
 
 
 
 Dim objContact As Outlook.ContactItem 
 
 Dim objContacts As Outlook.Folder 
 
 Dim objNameSpace As Outlook.NameSpace 
 
 Dim objProperty As Outlook.UserProperty 
 
 
 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 Set objContacts = objNameSpace.GetDefaultFolder(olFolderContacts) 
 
 Set objContact = objContacts.Items.Find( _ 
 
 "[FileAs] = ""Smith, Jeff"" and [FirstName] = ""Jeff""") 
 
 If Not TypeName(objContact) = "Nothing" Then 
 
 Set objProperty = _ 
 
 objContact.UserProperties.Find("LastDateContacted") 
 
 If TypeName(objProperty) <> "Nothing" Then 
 
 MsgBox "Last Date Contacted: " &; objProperty.Value 
 
 End If 
 
 Else 
 
 MsgBox "The contact was not found." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

