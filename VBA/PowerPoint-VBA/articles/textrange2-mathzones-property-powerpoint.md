---
title: TextRange2.MathZones Property (PowerPoint)
ms.assetid: 77e13bb5-e1c2-4438-a9eb-a475fd5f372c
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.MathZones Property (PowerPoint)

Sets the starting point and length of a math zone within a text range. Read-only


## Syntax

 _expression_. **MathZones**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Integer**|The starting point for the math zone.|
| _Length_|Optional|**Integer**|The length of the math zone.|

## Remarks

A math zone is a text range within which math typography rules apply and outside of which math typography rules do not apply. In addition to containing special mathematical symbols, math zones can also contain text such as in the equation rate = distance/time where text appears with math symbols.


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


