---
title: NameSpace.AddStore Method (Outlook)
keywords: vbaol11.chm771
f1_keywords:
- vbaol11.chm771
ms.prod: outlook
api_name:
- Outlook.NameSpace.AddStore
ms.assetid: c9390982-2408-fda5-a14d-de6f0daaadf1
ms.date: 06/08/2017
---


# NameSpace.AddStore Method (Outlook)

Adds a Personal Folders (.pst) file to the current profile.


## Syntax

 _expression_ . **AddStore**( **_Store_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **Variant**|The path of the .pst file to be added to the profile. If the .pst file does not exist, Microsoft Outlook creates it.|

## Remarks

Use the  **RemoveStore** method to remove a .pst that is already added to a profile.


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a new Personal Folders (.pst) file to the user?s profile.


```vb
Sub CreatePST() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 myNameSpace.AddStore "c:\" &; myNameSpace.CurrentUser &; "\.pst" 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

