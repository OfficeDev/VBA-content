---
title: Workbook.ReadOnlyRecommended Property (Excel)
keywords: vbaxl10.chm199216
f1_keywords:
- vbaxl10.chm199216
ms.prod: excel
api_name:
- Excel.Workbook.ReadOnlyRecommended
ms.assetid: 3cae84e4-d5f0-f01c-64d9-ec586ffdf79c
ms.date: 06/08/2017
---


# Workbook.ReadOnlyRecommended Property (Excel)

 **True** if the workbook was saved as read-only recommended. Read-only **Boolean** .


## Syntax

 _expression_ . **ReadOnlyRecommended**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

When you open a workbook that was saved as read-only recommended, Microsoft Excel displays a message recommending that you open the workbook as read-only.

Use the  **[SaveAs](workbook-saveas-method-excel.md)** method to change this property.


## Example

This example displays a message if the active workbook is saved as read-only recommended.


```vb
If ActiveWorkbook.ReadOnlyRecommended = True Then 
 MsgBox "This workbook is saved as read-only recommended" 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

