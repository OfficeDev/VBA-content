---
title: Worksheet.Outline Property (Excel)
keywords: vbaxl10.chm175113
f1_keywords:
- vbaxl10.chm175113
ms.prod: excel
api_name:
- Excel.Worksheet.Outline
ms.assetid: e53d8038-f20b-9d55-1ee0-c5f6b4a099d4
ms.date: 06/08/2017
---


# Worksheet.Outline Property (Excel)

Returns an  **[Outline](outline-object-excel.md)** object that represents the outline for the specified worksheet. Read-only.


## Syntax

 _expression_ . **Outline**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example sets the outline on Sheet1 to use automatic styles.


```vb
Worksheets("Sheet1").Outline.AutomaticStyles = True
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

