---
title: Workbook.Saved Property (Excel)
keywords: vbaxl10.chm199147
f1_keywords:
- vbaxl10.chm199147
ms.prod: excel
api_name:
- Excel.Workbook.Saved
ms.assetid: 37eb8e08-2bfa-8065-2520-a71e291ab50c
ms.date: 06/08/2017
---


# Workbook.Saved Property (Excel)

 **True** if no changes have been made to the specified workbook since it was last saved. Read/write **Boolean** .


## Syntax

 _expression_ . **Saved**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

If a workbook has never been saved, its  **[Path](workbook-path-property-excel.md)** property returns an empty string ("").

You can set this property to  **True** if you want to close a modified workbook without either saving it or being prompted to save it.


## Example

This example displays a message if the active workbook contains unsaved changes.


```vb
If Not ActiveWorkbook.Saved Then 
 MsgBox "This workbook contains unsaved changes." 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

