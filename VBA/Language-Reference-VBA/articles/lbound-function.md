---
title: LBound Function
keywords: vblr6.chm1008956
f1_keywords:
- vblr6.chm1008956
ms.prod: office
ms.assetid: 49520e9d-305b-4f5b-3ae6-df92f875d1eb
ms.date: 06/08/2017
---


# LBound Function



Returns a [Long](vbe-glossary.md) containing the smallest available subscript for the indicated dimension of an[array](vbe-glossary.md).
 **Syntax**
 **LBound(**_arrayname_ [ **,**_dimension_ ] **)**
The  **LBound** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _arrayname_|Required. Name of the array [variable](vbe-glossary.md); follows standard variable naming conventions.|
| _dimension_|Optional;  **Variant** ( **Long** ). Whole number indicating which dimension's lower bound is returned. Use 1 for the first dimension, 2 for the second, and so on. If _dimension_ is omitted, 1 is assumed.|
 **Remarks**
The  **LBound** function is used with the **UBound** function to determine the size of an array. Use the **UBound** function to find the upper limit of an array dimension.
 **LBound** returns the values in the following table for an array with the following dimensions:


|**Statement**|**Return Value**|
|:-----|:-----|
|LBound(A, 1)|1|
| `LBound(A, 2)`|0|
| `LBound(A, 3)`|-3|

The default lower bound for any dimension is either 0 or 1, depending on the setting of the  **Option** **Base** statement. The base of an array created with the **Array** function is zero; it is unaffected by **Option Base**.
Arrays for which dimensions are set using the  **To** clause in a **Dim**, **Private**, **Public**, **ReDim**, or **Static** statement can have any integer value as a lower bound.

## Example

This example uses the  **LBound** function to determine the smallest available subscript for the indicated dimension of an array. Use the **Option Base** statement to override the default base array subscript value of 0.


```vb
Dim Lower
Dim MyArray(1 To 10, 5 To 15, 10 To 20)     ' Declare array variables.
Dim AnyArray(10)
Lower = Lbound(MyArray, 1)     ' Returns 1.
Lower = Lbound(MyArray, 3)    ' Returns 10.
Lower = Lbound(AnyArray)    ' Returns 0 or 1, depending on
    ' setting of Option Base.


```


