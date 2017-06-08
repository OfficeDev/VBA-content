---
title: Sgn Function
keywords: vblr6.chm1009021
f1_keywords:
- vblr6.chm1009021
ms.prod: office
ms.assetid: 9da078d4-8c97-ea76-c095-46a4e46518ac
ms.date: 06/08/2017
---


# Sgn Function



Returns a  **Variant** ( **Integer** ) indicating the sign of a number.
 **Syntax**
 **Sgn(**_number_**)**
The required  _number_[argument](vbe-glossary.md) can be any valid[numeric expression](vbe-glossary.md).
 **Return Values**


|**If  _number_ is**|**Sgn returns**|
|:-----|:-----|
|Greater than zero|1|
|Equal to zero|0|
|Less than zero|-1|
 **Remarks**
The sign of the  _number_ argument determines the return value of the **Sgn** function.

## Example

This example uses the  **Sgn** function to determine the sign of a number.


```vb
Dim MyVar1, MyVar2, MyVar3, MySign
MyVar1 = 12: MyVar2 = -2.4: MyVar3 = 0
MySign = Sgn(MyVar1)    ' Returns 1.
MySign = Sgn(MyVar2)    ' Returns -1.
MySign = Sgn(MyVar3)    ' Returns 0.
```


