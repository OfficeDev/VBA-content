---
title: Insert Method
keywords: vbagr10.chm3077620
f1_keywords:
- vbagr10.chm3077620
ms.prod: excel
api_name:
- Excel.Insert
ms.assetid: 5f6a5961-9278-a2fa-6f08-4360646a7566
ms.date: 06/08/2017
---


# Insert Method

Inserts a cell or a range of cells into the datasheet and shifts other cells away to make space.

 _expression_. **Insert**( **_Shift_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Shift** Optional
 **XlInsertShiftDirection**
. Specifies which way to shift the cells.


|XlInsertShiftDirection can be one of these XlInsertShiftDirection constants.|
| **xlShiftToRight**|
| **xlShiftDown**If this argument is omitted, Microsoft Graph decides based on the shape of the range.|

## Example

This example inserts a new row before row four on the datasheet.


```
myChart.Application.DataSheet.Rows(4).Insert
```

This example inserts new cells at the range A1:C5 on the datasheet and shifts cells downward.




```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:C5").Insert Shift:=xlShiftDown
```


