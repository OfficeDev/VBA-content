---
title: Columns Property
keywords: vbagr10.chm65777
f1_keywords:
- vbagr10.chm65777
ms.prod: excel
api_name:
- Excel.Columns
ms.assetid: 7c5bd414-aa86-49e6-c853-0fa0c56d11a7
ms.date: 06/08/2017
---


# Columns Property

Returns a Range object that represents the columns in the specified range or all the columns on the datasheet. Read-only Range object.

 _expression_. **Range**

 _expression_ Required. An expression that returns an object in the Applies To List.

For information about returning a single member of a collection, see  [Returning an Object from a Collection (Excel)](returning-an-object-from-a-collection-excel.md).

## Example

This example clears column A of the datasheet.


```
myChart.Application.DataSheet.Columns(2).ClearContents
```


