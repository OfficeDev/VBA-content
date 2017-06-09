---
title: Workbook.CustomViews Property (Excel)
keywords: vbaxl10.chm199164
f1_keywords:
- vbaxl10.chm199164
ms.prod: excel
api_name:
- Excel.Workbook.CustomViews
ms.assetid: 286f6d5a-fb91-a339-8e74-9014ab7f4835
ms.date: 06/08/2017
---


# Workbook.CustomViews Property (Excel)

Returns a  **[CustomViews](customviews-object-excel.md)** collection that represents all the custom views for the workbook.


## Syntax

 _expression_ . **CustomViews**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example creates a new custom view named "Summary" in the active workbook.


```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

