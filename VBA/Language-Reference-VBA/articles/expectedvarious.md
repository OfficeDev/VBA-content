---
title: "Expected: <various>"
keywords: vblr6.chm1011312
f1_keywords:
- vblr6.chm1011312
ms.prod: office
ms.assetid: 05deb22f-01c5-ff47-9f02-b31217740b95
ms.date: 06/08/2017
---


# Expected: <various>

An expected part of the syntax was not found. The error is usually located to the left of the selected item, but isn't always obvious. For example, you can invoke a  **Sub** procedure with or without the **Call** keyword. However, if you use the **Call** keyword, you must enclose the argument list in parentheses. This error has the following causes and solutions:



-  **Expected: End of Statement**. Improper use of parentheses in a[procedure](vbe-glossary.md) invocation:
    
```vb
X = Workbook.Add F:= 5    ' Error due to no parentheses. 
Call MySub 5                ' Error due to no parentheses. 
```


    Use parentheses in a function call that specifies [argument](vbe-glossary.md)s or with a  **Sub** procedure invocation that uses the **Call** keyword.
    
-  **Expected: )**. Incorrect syntax for a procedure call. For example, a function call can't stand by itself, and **Sub** procedure calls sometimes require the **Call** keyword, depending on how you specify their arguments.
    
  ```
  Workbook.Add (X:=5, Y:=7)    ' Function call without expression. 
YourSub(5, 7)                ' Sub invocation without Call. 

  ```


     Always use function calls in[expressions](vbe-glossary.md). If you have multiple arguments enclosed in parentheses in a  **Sub** procedure call, you must use the **Call** keyword.
    
-  **Expected: Expression**. For example, when pasting code from the[Object Browser](vbe-glossary.md), you may have forgotten to specify a value for a [named argument](vbe-glossary.md).
    
  ```
  Workbook.Add (X:= )    ' Error because no value assigned to 
' named argument. 

  ```


    Either add a value for the argument, or delete the argument if it's optional.
    
-  **Expected: Variable**. For example, you may have used restricted[keywords](vbe-glossary.md) for variable names. In the following example, the **Input #** statement expects a variable as the second argument. Since **Type** is a restricted keyword, it can't be used as a variable name.
    
```vb
Input # 1, Type    ' Type keyword invalidly used as 
' variable name. 
```


    Rename the variable so it doesn't conflict with restricted keywords.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

