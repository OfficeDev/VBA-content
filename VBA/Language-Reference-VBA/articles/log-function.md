---
title: Log Function
keywords: vblr6.chm1008966
f1_keywords:
- vblr6.chm1008966
ms.prod: office
ms.assetid: 09ff0a30-0138-cfad-6cb8-2172c8ff37f2
ms.date: 06/08/2017
---


# Log Function



Returns a  **Double** specifying the natural logarithm of a number.
 **Syntax**
 **Log(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md) greater than zero.
 **Remarks**
The natural logarithm is the logarithm to the base  _e_. The[constant ](vbe-glossary.md) _e_ is approximately 2.718282.
You can calculate base- _n_ logarithms for any number _x_ by dividing the natural logarithm of _x_ by the natural logarithm of _n_ as follows:
Log _n(x)_ = **Log** ( _x_ ) / **Log** ( _n_ )
The following example illustrates a custom  **Function** that calculates base-10 logarithms:



```vb
Static Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function
```


## Example

This example uses the  **Log** function to return the natural logarithm of a number.


```vb
Dim MyAngle, MyLog
' Define angle in radians.
MyAngle = 1.3
' Calculate inverse hyperbolic sine.
MyLog = Log(MyAngle + Sqr(MyAngle * MyAngle + 1))


```


