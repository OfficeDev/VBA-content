---
title: Workbook.IconSets Property (Excel)
keywords: vbaxl10.chm199261
f1_keywords:
- vbaxl10.chm199261
ms.prod: excel
api_name:
- Excel.Workbook.IconSets
ms.assetid: c837d2a8-d21d-7432-a409-f49426368556
ms.date: 06/08/2017
---


# Workbook.IconSets Property (Excel)

This property is used to filter data in a workbook based on a cell icon from the  **IconSet** collection. Read-only.


## Syntax

 _expression_ . **IconSets**

 _expression_ A variable that represents a **Workbook** object.


## Example

In the following example, data is filtered by a cell icon.


```vb
Selection.AutoFilter Field:=1, Criteria1:=ActiveWorkbook.IconSets(xl3Arrows).Item(1), Operator:=xlFilterIcon
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

