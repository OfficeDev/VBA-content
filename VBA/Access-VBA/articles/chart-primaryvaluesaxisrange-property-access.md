---
title: Chart.PrimaryValuesAxisRange Property (Access)
keywords: vbaac10.chm6125
f1_keywords:
- vbaac10.chm6125
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisRange
ms.date: 05/02/2018
---


# Chart.PrimaryValuesAxisRange Property (Access)

Returns or sets the behavior for representing minimum and maximum values on the primary values axis. Read/write **[AcAxisRange](acaxisrange-enumeration-access.md)** .


## Syntax

 _expression_ . **PrimaryValuesAxisRange**

 _expression_ A variable that represents a **Chart** object.


## Remarks

**PrimaryValuesAxisMinimum** and **PrimaryValuesAxisMaximum** are enforced when the **PrimaryValuesAxisRange** 
property is set to **Fixed**. Otherwise, the **Auto** setting will determine the range based on the lowest and 
highest values in the set.


## See also


#### Concepts


[AcAxisRange Enumeration](acaxisrange-enumeration-access.md)

[PrimaryValuesAxisMinimum Property](chart-primaryvaluesaxisminimum-property-access.md)

[PrimaryValuesAxisMaximum Property](chart-primaryvaluesaxismaximum-property-access.md)

[Chart Object](chart-object-access.md)