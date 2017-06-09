---
title: Workbook.FullNameURLEncoded Property (Excel)
keywords: vbaxl10.chm199203
f1_keywords:
- vbaxl10.chm199203
ms.prod: excel
api_name:
- Excel.Workbook.FullNameURLEncoded
ms.assetid: 589d98f7-e6fa-bc28-2c8f-7cb72009737a
ms.date: 06/08/2017
---


# Workbook.FullNameURLEncoded Property (Excel)

Returns a  **String** indicating the name of the object, including its path on disk, as a string. Read-only.


## Syntax

 _expression_ . **FullNameURLEncoded**

 _expression_ A variable that represents a **Workbook** object.


## Example

In this example, Microsoft Excel displays the path and file name of the active workbook to the user.


```vb
Sub UseCanonical() 
 
 ' Display the full path to user. 
 MsgBox ActiveWorkbook.FullNameURLEncoded 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

