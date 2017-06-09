---
title: NameSpace.AddStoreEx Method (Outlook)
keywords: vbaol11.chm777
f1_keywords:
- vbaol11.chm777
ms.prod: outlook
api_name:
- Outlook.NameSpace.AddStoreEx
ms.assetid: 15b8948d-cbe4-a499-ec03-b1bbf56ead82
ms.date: 06/08/2017
---


# NameSpace.AddStoreEx Method (Outlook)

Adds a Personal Folders file (.pst) in the specified format to the current profile.


## Syntax

 _expression_ . **AddStoreEx**( **_Store_** , **_Type_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **Variant**|The path of the .pst file to be added to the profile. If the .pst file does not exist, Microsoft Outlook creates it.|
| _Type_|Required| **[OlStoreType](olstoretype-enumeration-outlook.md)**|The format in which the data file should be created.|

## Remarks

Use the  **olStoreUnicode** constant to add a new .pst file that has greater storage capacity for items and folders and supports multilingual Unicode data, to the user's profile. The **olStoreANSI** constant allows you to create .pst files that do not provide full support for multilingual Unicode data, but are compatible with earlier versions of Outlook. The **olStoreDefault** constant helps you create a .pst file in the default format that is compatible with the mailbox mode in which Outlook runs on the Microsoft Exchange Server.


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a new Personal Folders (.pst) file that has greater storage capacity for items and folders and supports Unicode to the user?s profile.


```vb
Sub CreateUnicodePST() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 myNameSpace.AddStoreEx "c:\" &; myNameSpace.CurrentUser &; "\.pst",olStoreUnicode 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

