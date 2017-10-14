---
title: Int, Fix Functions
keywords: vblr6.chm1008795
f1_keywords:
- vblr6.chm1008795
ms.prod: office
ms.assetid: 32ce40ac-fdf8-bd6d-e7f9-154c480a9602
ms.date: 06/08/2017
---


# Int, Fix Functions



Returns the integer portion of a number.
 **Syntax**
 **Int(**_number_**)**
 **Fix(**_number_**)**
The required  _number_[argument](vbe-glossary.md) is a[Double](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md). If  _number_ contains[Null](vbe-glossary.md),  **Null** is returned.
 **Remarks**
Both  **Int** and **Fix** remove the fractional part of _number_ and return the resulting integer value.
The difference between  **Int** and **Fix** is that if _number_ is negative, **Int** returns the first negative integer less than or equal to _number,_ whereas **Fix** returns the first negative integer greater than or equal to _number._ For example, **Int** converts -8.4 to -9, and **Fix** converts -8.4 to -8.
 **Fix(**_number_**)** is equivalent to:



```vb
Sgn(number) * Int(Abs(number))

```


## Example

This example illustrates how the  **Int** and **Fix** functions return integer portions of numbers. In the case of a negative number argument, the **Int** function returns the first negative integer less than or equal to the number; the **Fix** function returns the first negative integer greater than or equal to the number.


```vb
Dim MyNumber
MyNumber = Int(99.8)    ' Returns 99.
MyNumber = Fix(99.2)    ' Returns 99.

MyNumber = Int(-99.8)    ' Returns -100.
MyNumber = Fix(-99.8)    ' Returns -99.

MyNumber = Int(-99.2)    ' Returns -100.
MyNumber = Fix(-99.2)    ' Returns -99.


```


