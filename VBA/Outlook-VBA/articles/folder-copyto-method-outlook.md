---
title: Folder.CopyTo Method (Outlook)
keywords: vbaol11.chm1994
f1_keywords:
- vbaol11.chm1994
ms.prod: outlook
api_name:
- Outlook.Folder.CopyTo
ms.assetid: ddd010e2-54af-f291-b9a9-92cc55a83cca
ms.date: 06/08/2017
---


# Folder.CopyTo Method (Outlook)

Copies the current folder in its entirety to the destination folder. 


## Syntax

 _expression_ . **CopyTo**( **_DestinationFolder_** )

 _expression_ A variable that represents a **Folder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DestinationFolder_|Required| **[Folder](folder-object-outlook.md)**|Required  **Folder** object that represents the destination folder.|

### Return Value

A  **Folder** object that represents the new copy of the current folder.


## Remarks

Setting the REG_MULTI_SZ value,  `DisableCrossAccountCopy`, in  `HKCU\Software\Microsoft\Office\14.0\Outlook` in the Windows registry has the side effect of disabling this method.


## Example

This Visual Basic for Applications (VBA) example uses the  **CopyTo** method to copy the default Contacts folder to the default Inbox folder.


```vb
Sub CopyFolder() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myInboxFolder As Outlook.Folder 
 Dim myContactsFolder As Outlook.Folder 
 Dim myNewFolder As Outlook.Folder 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myInboxFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 Set myContactsFolder = myNameSpace.GetDefaultFolder(olFolderContacts) 
 Set myNewFolder = myContactsFolder.CopyTo(myInboxFolder) 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

