---
title: Items.Add Method (Outlook)
keywords: vbaol11.chm61
f1_keywords:
- vbaol11.chm61
ms.prod: outlook
api_name:
- Outlook.Items.Add
ms.assetid: 0ee68068-1452-0f29-b85a-88b801ac0448
ms.date: 06/08/2017
---


# Items.Add Method (Outlook)

Creates a new Outlook item in the  **[Items](items-object-outlook.md)** collection for the folder.


## Syntax

 _expression_ . **Add** **_Type_**

 _expression_ A variable that represents an **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|The Outlook item type for the new item. Specifies a  **[MessageClass](mailitem-messageclass-property-outlook.md)** to create custom forms. Can be one of the following **OlItemType** constants: **olAppointmentItem** , **olContactItem** , **olJournalItem** , **olMailItem** , **olNoteItem** , **olPostItem** , or **olTaskItem,** , or any valid message class.|

### Return Value

An  **Object** value that represents the new Outlook item.


## Remarks

If not specified, the  **Type** property of the Outlook item defaults to the type of the folder or to **[MailItem](mailitem-object-outlook.md)** if the parent folder is not typed.


## Example

This VBA example gets the current Contacts folder and adds a new ContactItem object to it and sets some initial values in the fields based on another contact. To run this example without any error, replace 'Dan Wilson' with a valid contact name that exists in your Contacts folder.


```vb
Sub AddContact() 
 Dim myNamespace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myItem As Outlook.ContactItem 
 Dim myOtherItem As Outlook.ContactItem 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts) 
 Set myOtherItem = myFolder.Items("Dan Wilson") 
 Set myItem = myFolder.Items.Add 
 myItem.CompanyName = myOtherItem.CompanyName 
 myItem.BusinessAddress = myOtherItem.BusinessAddress 
 myItem.BusinessTelephoneNumber = myOtherItem.BusinessTelephoneNumber 
 myItem.Display 
End Sub
```

This VBA example adds a custom form to the default Tasks folder.




```vb
Sub AddForm() 
 Dim myNamespace As outlook.NameSpace 
 Dim myItems As outlook.Items 
 Dim myFolder As outlook.Folder 
 Dim myItem As outlook.TaskItem 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 Set myFolder = _ 
 myNamespace.GetDefaultFolder(olFolderTasks) 
 Set myItems = myFolder.Items 
 Set myItem = myItems.Add("IPM.Task.myTask") 
End Sub
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

