---
title: NameSpace.Dial Method (Outlook)
keywords: vbaol11.chm774
f1_keywords:
- vbaol11.chm774
ms.prod: outlook
api_name:
- Outlook.NameSpace.Dial
ms.assetid: 1fd29ed8-e983-c668-c48f-f642c56bfcd2
ms.date: 06/08/2017
---


# NameSpace.Dial Method (Outlook)

Displays the  **New Call** dialog box that allows users to dial the primary phone number of a specified contact.


## Syntax

 _expression_ . **Dial**( **_ContactItem_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContactItem_|Optional| **Variant**|The  **[ContactItem](contactitem-object-outlook.md)** object of the contact you want to dial.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example opens the  **New Call** dialog box.


```vb
Sub DialContact() 
 
 'Opens the New Call dialog 
 
 Application.GetNamespace("MAPI").Dial 
 
End Sub
```

The following VBA example opens the  **New Call** dialog box with the contact's information. To run this example, replace 'Jeff Smith' with a valid contact name.




```vb
Sub DialContact() 
 
 'Opens the New Call dialog with the contact info 
 
 Dim objContact As Outlook.ContactItem 
 
 
 
 Set objContact = Application.GetNamespace("MAPI"). _ 
 
 GetDefaultFolder(olFolderContacts).Items("Jeff Smith") 
 
 Application.GetNamespace("MAPI").Dial objContact 
 
 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

