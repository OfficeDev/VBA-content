---
title: Folder.MoveTo Method (Outlook)
keywords: vbaol11.chm1998
f1_keywords:
- vbaol11.chm1998
ms.prod: outlook
api_name:
- Outlook.Folder.MoveTo
ms.assetid: 5e8ece38-aaba-4971-643e-969956c2a196
ms.date: 06/08/2017
---


# Folder.MoveTo Method (Outlook)

Moves a folder to the specified destination folder.


## Syntax

 _expression_ . **MoveTo**( **_DestinationFolder_** )

 _expression_ A variable that represents a **[Folder](folder-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DestinationFolder_|Required| **Folder**|The destination  **Folder** for the **Folder** that is being moved.|

## Remarks

Setting the REG_MULTI_SZ value,  `DisableCrossAccountCopy`, in  `HKCU\Software\Microsoft\Office\14.0\Outlook` in the Windows registry has the side effect of disabling this method.


## Example

This Visual Basic for Applications (VBA) example uses the  **MoveTo** method to move the "My Test Contacts" folder in the default Contacts folder to the Inbox folder.


```vb
Sub MoveFolder() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myNewFolder As Outlook.Folder 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderContacts) 
 Set myNewFolder = myFolder.Folders.Add("My Test Contacts") 
 myNewFolder.MoveTo myNameSpace.GetDefaultFolder _ 
 (olFolderInbox) 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

