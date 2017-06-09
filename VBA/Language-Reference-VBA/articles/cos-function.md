---
title: Cos Function
keywords: vblr6.chm1008879
f1_keywords:
- vblr6.chm1008879
ms.prod: office
ms.assetid: a08a4706-223b-3d94-838a-4ac92b04744a
ms.date: 06/08/2017
---


# Cos Function



Returns a  **Double** specifying the cosine of an angle.
 **Syntax**
 **Cos(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md) that expresses an angle in radians.
 **Remarks**
The  **Cos** function takes an angle and returns the ratio of two sides of a right triangle. The ratio is the length of the side adjacent to the angle divided by the length of the hypotenuse.
The result lies in the range -1 to 1.
To convert degrees to radians, multiply degrees by [pi](vbe-glossary.md)/180. To convert radians to degrees, multiply radians by 180/pi.

## Example

This example uses the  **Cos** function to return the cosine of an angle.


```vb
Dim MyAngle, MySecant
MyAngle = 1.3    ' Define angle in radians.
MySecant = 1 / Cos(MyAngle)    ' Calculate secant.


```


