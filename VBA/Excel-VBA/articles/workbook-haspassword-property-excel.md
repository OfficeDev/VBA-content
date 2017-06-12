---
title: Workbook.HasPassword Property (Excel)
keywords: vbaxl10.chm199104
f1_keywords:
- vbaxl10.chm199104
ms.prod: excel
api_name:
- Excel.Workbook.HasPassword
ms.assetid: e3cfdc90-1e82-5556-0064-e8269ba92539
ms.date: 06/08/2017
---


# Workbook.HasPassword Property (Excel)

 **True** if the workbook has a protection password. Read-only **Boolean** .


## Syntax

 _expression_ . **HasPassword**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

You can assign a protection password to a workbook by using the  **[SaveAs](workbook-saveas-method-excel.md)** method.


## Example

This example displays a message if the active workbook has a protection password.


```vb
If ActiveWorkbook.HasPassword = True Then 
 MsgBox "Remember to obtain the workbook password" &; Chr(13) &; _ 
 " from the Network Administrator." 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

