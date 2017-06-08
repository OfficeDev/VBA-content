---
title: TextRange2.MathZones Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.MathZones
ms.assetid: 277aa819-d717-e2f5-5bc7-607abfce20a4
ms.date: 06/08/2017
---


# TextRange2.MathZones Property (Office)

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


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

