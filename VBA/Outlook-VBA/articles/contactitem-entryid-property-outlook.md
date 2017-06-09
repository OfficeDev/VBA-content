---
title: ContactItem.EntryID Property (Outlook)
keywords: vbaol11.chm939
f1_keywords:
- vbaol11.chm939
ms.prod: outlook
api_name:
- Outlook.ContactItem.EntryID
ms.assetid: 04f4bd28-5edf-4e69-5b7c-d3bec749fc4f
ms.date: 06/08/2017
---


# ContactItem.EntryID Property (Outlook)

Returns a  **String** representing the unique Entry ID of the object. Read-only.


## Syntax

 _expression_ . **EntryID**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagEntryId** .

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term **EntryID**s, search http://msdn.microsoft.com for  **PidTagEntryId** .


## Example

This Visual Basic for Applications (VBA) example uses the  **EntryID** property to compare the Entry ID of one contact to the Entry ID of a contact returned by a search operation to determine whether the objects represent the same contact. Replace the name with a valid contact name in your Contacts folder before running this example.


```vb
Sub UseEntryID() 
 Dim myNamespace As Outlook.NameSpace 
 Dim myContacts As Outlook.Folder 
 Dim myItem1 As Outlook.ContactItem 
 Dim myItem2 As Outlook.ContactItem 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myContacts = myNameSpace.GetDefaultFolder(olFolderContacts) 
 Set myItem1 = myContacts.Items.Find("[FirstName] = ""Dan""") 
 Set myitem2 = myContacts.Items.Find("[FileAs] = ""Wil"" and [FirstName] = ""Dan""") 
 If Not TypeName(myitem2) = "Nothing" Then 
 If myItem1.EntryID = myitem2.EntryID Then 
 MsgBox "These two contact items refer to the same contact." 
 End If 
 Else 
 MsgBox "The contact items were not found." 
 End If 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

