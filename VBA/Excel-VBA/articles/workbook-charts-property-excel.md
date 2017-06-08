---
title: Workbook.Charts Property (Excel)
keywords: vbaxl10.chm199084
f1_keywords:
- vbaxl10.chm199084
ms.prod: excel
api_name:
- Excel.Workbook.Charts
ms.assetid: 582d9a78-d86f-ab69-0c22-85f8a59412d9
ms.date: 06/08/2017
---


# Workbook.Charts Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the chart sheets in the specified workbook.


## Syntax

 _expression_ . **Charts**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example deletes every chart sheet in the active workbook.




 **Note**  For the following sample code to work you must have a chart sheet in the active workbook.




```vb
ActiveWorkbook.Charts.Delete
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

