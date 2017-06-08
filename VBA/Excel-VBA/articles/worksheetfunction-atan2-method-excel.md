---
title: WorksheetFunction.Atan2 Method (Excel)
keywords: vbaxl10.chm137118
f1_keywords:
- vbaxl10.chm137118
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Atan2
ms.assetid: d6a6597d-9d46-fdad-3bf1-05cee4cf9e20
ms.date: 06/08/2017
---


# WorksheetFunction.Atan2 Method (Excel)

Returns the arctangent, or inverse tangent, of the specified x- and y-coordinates. The arctangent is the angle from the x-axis to a line containing the origin (0, 0) and a point with coordinates (x_num, y_num). The angle is given in radians between -pi and pi, excluding -pi.


## Syntax

 _expression_ . **Atan2**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The x-coordinate of the point.|
| _Arg2_|Required| **Double**|The y-coordinate of the point.|

### Return Value

Double


## Remarks




- A positive result represents a counterclockwise angle from the x-axis; a negative result represents a clockwise angle.
    
- The following conditions apply:
    
      - Where x > 0 ATAN2(x,y) = ATAN(y/x)
    
  - Where y >= 0, x < 0 ATAN2(x,y) = ATAN(y/x)+PI()
    
  - Where y < 0, x < 0 ATAN2(x,y) = ATAN(y/x) - PI()
    
  - Where y > 0, x = 0 ATAN2(x,y) = PI()/2
    
  - Where y < 0, x = 0 ATAN2(x,y) = -PI()/2
    
  - If both x and y are 0, Atan2 returns an error value.
    
- To express the arctangent in degrees, multiply the result by 180/PI( ) or use the [Degrees](worksheetfunction-degrees-method-excel.md) method.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

