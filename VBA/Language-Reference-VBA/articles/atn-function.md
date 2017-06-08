---
title: Atn Function
keywords: vblr6.chm1008860
f1_keywords:
- vblr6.chm1008860
ms.prod: office
ms.assetid: ab5272cf-b372-8665-28c6-ee0318aa9bac
ms.date: 06/08/2017
---


# Atn Function



Returns a  **Double** specifying the arctangent of a number.
 **Syntax**
 **Atn(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md).
 **Remarks**
The  **Atn** function takes the ratio of two sides of a right triangle ( _number_ ) and returns the corresponding angle in radians. The ratio is the length of the side opposite the angle divided by the length of the side adjacent to the angle.
The range of the result is  **-**[pi](vbe-glossary.md)/2 to pi/2 radians.
To convert degrees to radians, multiply degrees by pi/180. To convert radians to degrees, multiply radians by 180/pi.

 **Note**   **Atn** is the inverse trigonometric function of **Tan**, which takes an angle as its argument and returns the ratio of two sides of a right triangle. Do not confuse **Atn** with the cotangent, which is the simple inverse of a tangent (1/tangent).


## Example

This example uses the  **Atn** function to calculate the value of pi.


```vb
Dim IntVar, StrVar, DateVar, MyCheck
' Initialize variables.
IntVar = 459: StrVar = "Hello World": DateVar = #2/12/69# 
MyCheck = VarType(IntVar)    ' Returns 2.
MyCheck = VarType(DateVar)    ' Returns 7.
MyCheck = VarType(StrVar)    ' Returns 8.


```


