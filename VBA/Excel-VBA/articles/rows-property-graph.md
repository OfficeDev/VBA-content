---
title: Rows Property (Graph)
keywords: vbagr10.chm5207942
f1_keywords:
- vbagr10.chm5207942
ms.prod: excel
ms.assetid: 045405b7-3f7c-bcf6-7757-f116ed8d7e37
ms.date: 06/08/2017
---


# Rows Property (Graph)

Returns a  **Range** object that represents the rows in the specified **Range** or **DataSheet** object. Read-only.

For information about returning a single member of a collection, see  [Returning an Object from a Collection](returning-an-object-from-a-collection-excel.md).

## Example

This example deletes row three on the datasheet.


```
myChart.Application.DataSheet.Rows(3).Delete
```


