---
title: Writing a Function Procedure
keywords: vbcn6.chm1076690
f1_keywords:
- vbcn6.chm1076690
ms.prod: office
ms.assetid: 80e2ad00-a12f-2f40-3cb8-9878a595dde3
ms.date: 06/08/2017
---


# Writing a Function Procedure

A  **Function** procedure is a series of Visual Basic[statements](vbe-glossary.md) enclosed by the **Function** and **End Function** statements. A **Function** procedure is similar to a **Sub** procedure, but a function can also return a value. A **Function** procedure can take[arguments](vbe-glossary.md), such as [constants](vbe-glossary.md), [variables](vbe-glossary.md), or [expressions](vbe-glossary.md) that are passed to it by a calling procedure. If a **Function** procedure has no arguments, its **Function** statement must include an empty set of parentheses. A function returns a value by assigning a value to its name in one or more statements of the procedure.

In the following example, the  **Celsius** function calculates degrees Celsius from degrees Fahrenheit. When the function is called from the **Main** procedure, a variable containing the argument value is passed to the function. The result of the calculation is returned to the calling procedure and displayed in a message box.



```vb
Sub Main() 
 temp = Application.InputBox(Prompt:= _ 
 "Please enter the temperature in degrees F.", Type:=1) 
 MsgBox "The temperature is " &; Celsius(temp) &; " degrees C." 
End Sub 
 
Function Celsius(fDegrees) 
 Celsius = (fDegrees - 32) * 5 / 9 
End Function
```


