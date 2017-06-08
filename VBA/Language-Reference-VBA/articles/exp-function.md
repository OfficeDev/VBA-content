---
title: Exp Function
keywords: vblr6.chm1008917
f1_keywords:
- vblr6.chm1008917
ms.prod: office
ms.assetid: cd9d5f30-63b5-2025-1b23-4fbed4aeef1e
ms.date: 06/08/2017
---


# Exp Function



Returns a  **Double** specifying _e_ (the base of natural logarithms) raised to a power.
 **Syntax**
 **Exp(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md).
 **Remarks**
If the value of  _number_ exceeds 709.782712893, an error occurs. The[constant ](vbe-glossary.md) _e_ is approximately 2.718282.

 **Note**  The  **Exp** function complements the action of the **Log** function and is sometimes referred to as the antilogarithm.


## Example

This example uses the  **Exp** function to return _e_ raised to a power.


```vb
Dim MyAngle, MyHSin
' Define angle in radians.
MyAngle = 1.3    
' Calculate hyperbolic sine.
MyHSin = (Exp(MyAngle) - Exp(-1 * MyAngle)) / 2  

```


