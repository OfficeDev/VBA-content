---
title: SpellingOptions.IgnoreFileNames Property (Excel)
keywords: vbaxl10.chm717078
f1_keywords:
- vbaxl10.chm717078
ms.prod: excel
api_name:
- Excel.SpellingOptions.IgnoreFileNames
ms.assetid: 346b454b-b501-9836-4d45-dbe551f4c2cb
ms.date: 06/08/2017
---


# SpellingOptions.IgnoreFileNames Property (Excel)

 **False** instructs Microsoft Excel to check for Internet and file addresses, **True** instructs Excel to ignore Internet and file addresses when using the spell checker. Read/write **Boolean** .


## Syntax

 _expression_ . **IgnoreFileNames**

 _expression_ A variable that represents a **SpellingOptions** object.


## Example

In this example, Microsoft Excel determines what the setting is for checking spelling of Internet and file addresses and notifies the user.


```vb
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreFileNames = True Then 
 MsgBox "Spelling options for checking Internet and file addresses is disabled." 
 Else 
 MsgBox "Spelling options for checking Internet and file addresses is enabled." 
 End If 
 
End Sub
```


## See also


#### Concepts


[SpellingOptions Object](spellingoptions-object-excel.md)

