---
title: OLEObject.ZOrder Property (Excel)
keywords: vbaxl10.chm415096
f1_keywords:
- vbaxl10.chm415096
ms.prod: excel
api_name:
- Excel.OLEObject.ZOrder
ms.assetid: dd7c2c81-6582-5de9-d254-66061d4345ef
ms.date: 06/08/2017
---


# OLEObject.ZOrder Property (Excel)

Returns the z-order position of the object. Read-only  **Long** .


## Syntax

 _expression_ . **ZOrder**

 _expression_ A variable that represents an **OLEObject** object.


## Remarks

In any collection of objects, the object at the back of the z-order is  _collection_(1), and the object at the front of the z-order is  _collection_( _collection_. **Count** ). For example, if there are embedded charts on the active sheet, the chart at the back of the z-order is `ActiveSheet.ChartObjects(1)`, and the chart at the front of the z-order is  `ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count)`.


## Example

This example displays the z-order position of embedded chart one on Sheet1.


```vb
MsgBox "The chart's z-order position is " &; _ 
 Worksheets("Sheet1").ChartObjects(1).ZOrder
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

