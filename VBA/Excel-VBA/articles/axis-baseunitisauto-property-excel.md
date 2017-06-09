---
title: Axis.BaseUnitIsAuto Property (Excel)
keywords: vbaxl10.chm561105
f1_keywords:
- vbaxl10.chm561105
ms.prod: excel
api_name:
- Excel.Axis.BaseUnitIsAuto
ms.assetid: e6f72a37-cfa7-4888-2688-f236fa61d259
ms.date: 06/08/2017
---


# Axis.BaseUnitIsAuto Property (Excel)

 **True** if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **BaseUnitIsAuto**

 _expression_ A variable that represents an **Axis** object.


## Remarks

You cannot set this property for a value axis.


## Example

This example sets the category axis in embedded chart one on worksheet one to use a time scale with automatic base units.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnitIsAuto = True 
 End With 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

