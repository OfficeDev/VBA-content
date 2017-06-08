---
title: Choose Function
keywords: vblr6.chm1010091
f1_keywords:
- vblr6.chm1010091
ms.prod: office
ms.assetid: ccf3fe4c-9507-5ff3-b834-9a16e2a19ae2
ms.date: 06/08/2017
---


# Choose Function



Selects and returns a value from a list of [arguments](vbe-glossary.md).
 **Syntax**
 **Choose(**_index_, _choice-1_ [, _choice-2_, _..._ [, _choice-n_ ]] **)**
The  **Choose** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _index_|Required. [Numeric expression](vbe-glossary.md) or field that results in a value between 1 and the number of available choices.|
| _choice_|Required. [Variant expression](vbe-glossary.md) containing one of the possible choices.|
 **Remarks**
 **Choose** returns a value from the list of choices based on the value of _index_. If _index_ is 1, **Choose** returns the first choice in the list; if _index_ is 2, it returns the second choice, and so on.
You can use  **Choose** to look up a value in a list of possibilities. For example, if _index_ evaluates to 3 and _choice-1_ = "one", _choice-2_ = "two", and _choice-3_ = "three", **Choose** returns "three". This capability is particularly useful if _index_ represents the value in an option group.
 **Choose** evaluates every choice in the list, even though it returns only one. For this reason, you should watch for undesirable side effects. For example, if you use the **MsgBox** function as part of an[expression](vbe-glossary.md) in all the choices, a message box will be displayed for each choice as it is evaluated, even though **Choose** returns the value of only one of them.
The  **Choose** function returns a[Null](vbe-glossary.md) if _index_ is less than 1 or greater than the number of choices listed.
If  _index_ is not a whole number, it is rounded to the nearest whole number before being evaluated.

## Example

This example uses the  **Choose** function to display a name in response to an index passed into the procedure in the `Ind` parameter.


```vb
Function GetChoice(Ind As Integer)
    GetChoice = Choose(Ind, "Speedy", "United", "Federal")
End Function
```


