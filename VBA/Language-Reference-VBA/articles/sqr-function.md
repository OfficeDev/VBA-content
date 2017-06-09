---
title: Sqr Function
keywords: vblr6.chm1009029
f1_keywords:
- vblr6.chm1009029
ms.prod: office
ms.assetid: ce2add56-f943-9470-0caa-befda14d124a
ms.date: 06/08/2017
---


# Sqr Function



Returns a  **Double** specifying the square root of a number.
 **Syntax**
 **Sqr(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md) greater than or equal to zero.

## Example

This example uses the  **Sqr** function to calculate the square root of a number.


```vb
Dim MySqr
MySqr = Sqr(4)    ' Returns 2.
MySqr = Sqr(23)    ' Returns 4.79583152331272.
MySqr = Sqr(0)    ' Returns 0.
MySqr = Sqr(-4)    ' Generates a run-time error.


```


