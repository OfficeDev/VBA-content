---
title: Delete Method (Graph)
keywords: vbagr10.chm3077617
f1_keywords:
- vbagr10.chm3077617
ms.prod: excel
ms.assetid: f5bc861f-67e4-05e9-765f-d9ed34e0e936
ms.date: 06/08/2017
---


# Delete Method (Graph)

Delete method as it applies to all objects in the Applies To list except the  **Range** object.

Deletes the specified object.

 _expression_. **Delete**

 _expression_ Required. An expression that returns one of the above objects.
Delete method as it applies to the  **Range** object.
Deletes the specified object.
 _expression_. **Delete**( **_Shift_**)
 _expression_ Required. An expression that returns one of the above objects.
 **Shift** Optional
 **XlDeleteShiftDirection**
 . Used only with **Range** objects. Specifies how to shift cells to replace deleted cells.


|XlDeleteShiftDirection can be one of these XlDeleteShiftDirection constants.|
| **xlShiftToLeft**|
| **xlShiftUp** If this argument is omitted, Microsoft Graph decides how to shift cells based on the shape of the specified range.|

## Remarks

Deleting a  **Point** or **LegendKey** object deletes the entire series.


## Example

This example deletes cells A1:D10 on the datasheet and shifts the remaining cells to the left.


```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D10").Delete Shift:=xlShiftToLeft
```

This example deletes the chart title.




```
myChart.ChartTitle.Delete
```


