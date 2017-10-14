---
title: Worksheet.EnableSelection Property (Excel)
keywords: vbaxl10.chm175095
f1_keywords:
- vbaxl10.chm175095
ms.prod: excel
api_name:
- Excel.Worksheet.EnableSelection
ms.assetid: e1647c07-3863-9268-864c-1c62b2eebbb1
ms.date: 06/08/2017
---


# Worksheet.EnableSelection Property (Excel)

Returns or sets what can be selected on the sheet. Read/write  **[XlEnableSelection](xlenableselection-enumeration-excel.md)** .


## Syntax

 _expression_ . **EnableSelection**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

This property takes effect only when the worksheet is protected:  **xlNoSelection** prevents any selection on the sheet, **xlUnlockedCells** allows only those cells whose **Locked** property is **False** to be selected, and **xlNoRestrictions** allows any cell to be selected.


## Example

This example sets worksheet one so that nothing on it can be selected.


```vb
With Worksheets(1) 
 .EnableSelection = xlNoSelection 
 .Protect Contents:=True, UserInterfaceOnly:=True 
End With
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

