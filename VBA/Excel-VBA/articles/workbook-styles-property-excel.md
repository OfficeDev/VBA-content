---
title: Workbook.Styles Property (Excel)
keywords: vbaxl10.chm199154
f1_keywords:
- vbaxl10.chm199154
ms.prod: excel
api_name:
- Excel.Workbook.Styles
ms.assetid: c9a70be9-cab5-ea5f-2e3f-949b1acf43d9
ms.date: 06/08/2017
---


# Workbook.Styles Property (Excel)

Returns a  **[Styles](styles-object-excel.md)** collection that represents all the styles in the specified workbook. Read-only.


## Syntax

 _expression_ . **Styles**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example deletes the user-defined style "Stock Quote Style" from the active workbook.


```vb
ActiveWorkbook.Styles("Stock Quote Style").Delete
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

