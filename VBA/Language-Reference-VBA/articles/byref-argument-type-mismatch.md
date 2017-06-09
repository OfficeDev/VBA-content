---
title: ByRef argument type mismatch
keywords: vblr6.chm1011308
f1_keywords:
- vblr6.chm1011308
ms.prod: office
ms.assetid: 6adca657-8620-e3f1-3587-e317f988979c
ms.date: 06/08/2017
---


# ByRef argument type mismatch

An [argument](vbe-glossary.md) passed **ByRef** ([by reference](vbe-glossary.md)), the default, must have the precise [data type](vbe-glossary.md) expected in the[procedure](vbe-glossary.md). This error has the following cause and solution:



- You passed an argument of one type that could not be coerced to the type expected. 
    
    For example, this error occurs if you try to pass an  **Integer** variable when a **Long** is expected. If you want coercion to occur, even if it causes information to be lost, you can pass the argument in its own set of parentheses. For example, to pass the **Variant** argument `MyVar` to a procedure that expects an **Integer** argument, you can write the call as follows:
    


```vb
Dim MyVar 
MyVar = 3.1415 
Call SomeSub((MyVar)) 
 
Sub SomeSub (MyNum As Integer) 
MyNum = MyNum + MyNum 
End Sub
```


    Placing the argument in its own set of parentheses forces evaluation of it as an [expression](vbe-glossary.md). During this evaluation, the fractional portion of the number is rounded (not truncated) to make it conform to the expected argument type. The result of the evaluation is placed in a temporary location, and a reference to the temporary location is received by the procedure. Thus, the original  `MyVar` retains its value.
    
     **Note**  If you don't specify a type for a [variable](vbe-glossary.md), the variable receives the default type,  **Variant**. This isn't always obvious. For example, the following code declares two variables, the first, `MyVar`, is a  **Variant**; the second, `AnotherVar`, is an  **Integer**.



```vb
Dim MyVar, AnotherVar As Integer 

```

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

