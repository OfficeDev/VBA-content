---
title: ListObject.DataBodyRange Property (Excel)
keywords: vbaxl10.chm734082
f1_keywords:
- vbaxl10.chm734082
ms.prod: excel
api_name:
- Excel.ListObject.DataBodyRange
ms.assetid: fe906555-d006-8220-d9f8-59636cca68d5
ms.date: 06/08/2017
---


# ListObject.DataBodyRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range of values, excluding the header row, in a table. Read-only.


## Syntax

 _expression_ . **DataBodyRange**

 _expression_ A variable that represents a **ListObject** object.


## Example

This example selects the active data range in the list.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.ListObjects.Item(1).DataBodyRange.Select
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

