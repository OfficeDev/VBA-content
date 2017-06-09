---
title: ContactItem.FullName Property (Outlook)
keywords: vbaol11.chm1006
f1_keywords:
- vbaol11.chm1006
ms.prod: outlook
api_name:
- Outlook.ContactItem.FullName
ms.assetid: 3036dc57-31fb-45ad-f51e-49336206581d
ms.date: 06/08/2017
---


# ContactItem.FullName Property (Outlook)

Returns or sets a  **String** specifying the whole, unparsed full name for the contact. Read/write.


## Syntax

 _expression_ . **FullName**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed into the  **[FirstName](contactitem-firstname-property-outlook.md)** , **[MiddleName](contactitem-middlename-property-outlook.md)** , **[LastName](contactitem-lastname-property-outlook.md)** , and **[Suffix](contactitem-suffix-property-outlook.md)** properties, which may be changed or typed independently if they are parsed incorrectly. Any changes or entries to the **FirstName** , **LastName** , **MiddleName** , or **Suffix** properties will be overwritten by any subsequent changes or entries to **FullName** .


## Example

This Visual Basic for Applications (VBA) example uses the  **[Restrict](items-restrict-method-outlook.md)** method to apply a filter to the contact items based on the item's **[LastModificationTime](mailitem-lastmodificationtime-property-outlook.md)** property, and then it displays the full name of the contacts returned by the filter.


```vb
Public Sub ContactDateCheck() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myContacts As Outlook.Items 
 
 Dim myItems As Outlook.Items 
 
 Dim myItem As Object 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items 
 
 Set myItems = myContacts.Restrict("[LastModificationTime] > '01/1/2003'") 
 
 For Each myItem In myItems 
 
 If (myItem.Class = olContact) Then 
 
 MsgBox myItem.FullName &; ": " &; myItem.LastModificationTime 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

