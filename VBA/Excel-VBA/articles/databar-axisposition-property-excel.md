---
title: Databar.AxisPosition Property (Excel)
keywords: vbaxl10.chm810092
f1_keywords:
- vbaxl10.chm810092
ms.prod: excel
api_name:
- Excel.Databar.AxisPosition
ms.assetid: 0e239fd1-8bdf-2355-10ae-b7766b9befaf
ms.date: 06/08/2017
---


# Databar.AxisPosition Property (Excel)

Returns or sets the position of the axis of the data bars specified by a conditional formatting rule. Read/write


## Syntax

 _expression_ . **AxisPosition**

 _expression_ A variable that represents a **[Databar](databar-object-excel.md)** object.


### Return Value

 **[XlDataBarAxisPosition](xldatabaraxisposition-enumeration-excel.md)**


## Remarks

The axis for data bars is displayed only when the  **AxisPosition** property is either **xlDataBarAxisAutomatic** or **xlDataBarAxisMidpoint** , and when there are negative values in the range of values specified with a data bar conditional formatting rule. If the conditional formatting rule is created programmatically, the default value for the **AxisPosition** property is **xlDataBarAxisNone** . If the conditional formatting rule is created using the user interface, the default value for the **AxisPosition** property is **xlDataBarAxisAutomatic** .


## Example

The following code example selects a range of cells, adds data bar formatting, and then sets the axis position to display in the middle of the cells when negative values are present.


```vb
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
myDataBar.AxisPosition = xlDataBarAxisMidpoint
```


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

