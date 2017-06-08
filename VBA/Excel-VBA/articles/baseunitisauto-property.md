---
title: BaseUnitIsAuto Property
keywords: vbagr10.chm67184
f1_keywords:
- vbagr10.chm67184
ms.prod: excel
api_name:
- Excel.BaseUnitIsAuto
ms.assetid: 5c3257a8-2348-ff5c-53ce-86a7529d2dbb
ms.date: 06/08/2017
---


# BaseUnitIsAuto Property

True if Microsoft Graph chooses appropriate base units for the specified category axis. The default value is True. Read/write Boolean.

 _expression_. **BaseUnitIsAuto**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

You cannot set this property for a value axis.


## Example

This example sets the category axis on the chart to use a time scale with automatic base units.


```vb
With myChart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnitIsAuto = True 
 End With 
End With
```


