---
title: IIf Function
keywords: vblr6.chm1012957
f1_keywords:
- vblr6.chm1012957
ms.prod: office
ms.assetid: a31d9f49-1f5a-324b-77a2-276eb573552a
ms.date: 06/08/2017
---


# IIf Function



Returns one of two parts, depending on the evaluation of an [expression](vbe-glossary.md).
 **Syntax**
 **IIf( _expr_,** **_truepart_,** **_falsepart_ )**
The  **IIf** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>               | <strong>Description</strong>                                                                         |
|:------------------------------------|:-----------------------------------------------------------------------------------------------------|
| <strong><em>expr</em></strong>      | Required. Expression you want to evaluate.                                                           |
| <strong><em>truepart</em></strong>  | Required. Value or expression returned if  <strong><em>expr</em></strong> is <strong>True</strong>.  |
| <strong><em>falsepart</em></strong> | Required. Value or expression returned if  <strong><em>expr</em></strong> is <strong>False</strong>. |

 **Remarks**
 **IIf** always evaluates both **_truepart_** and **_falsepart_**, even though it returns only one of them. Because of this, you should watch for undesirable side effects. For example, if evaluating **_falsepart_** results in a division by zero error, an error occurs even if **_expr_** is **True**.

## Example

This example uses the  **IIf** function to evaluate the `TestMe` parameter of the `CheckIt` procedure and returns the word "Large" if the amount is greater than 1000; otherwise, it returns the word "Small".


```vb
Function CheckIt (TestMe As Integer)
    CheckIt = IIf(TestMe > 1000, "Large", "Small")
End Function
```


