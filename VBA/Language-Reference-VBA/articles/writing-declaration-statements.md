---
title: Writing Declaration Statements
keywords: vbcn6.chm1076693
f1_keywords:
- vbcn6.chm1076693
ms.prod: office
ms.assetid: 9aaee08c-09d3-b70b-0d8f-9ca949fbd04a
ms.date: 06/08/2017
---


# Writing Declaration Statements

You use declaration statements to name and define [procedures](vbe-glossary.md), [variables](vbe-glossary.md), [arrays](vbe-glossary.md), and [constants](vbe-glossary.md). When you declare a procedure, variable, or constant, you also define its [scope](vbe-glossary.md), depending on where you place the declaration and what [keywords](vbe-glossary.md) you use to declare it.

The following example contains three declarations.



```vb
Sub ApplyFormat() 
    Const limit As Integer = 33 
    Dim myCell As Range 
    ' More statements 
End Sub
```

The  **Sub** statement (with matching **End Sub** statement) declares a procedure named `ApplyFormat`. All the statements enclosed by the  **Sub** and **End Sub** statements are executed whenever the `ApplyFormat` procedure is called or run.
The  **Const** statement declares the constant `limit` specifying the **Integer** data type and a value of 33.
The  **Dim** statement declares the `myCell` variable. The data type is an object, in this case, a Microsoft Excel **Range** object. You can declare a variable to be any object that is exposed in the application you are using. **Dim** statements are one type of statement used to declare variables. Other keywords used in declarations are **ReDim**, **Static**, **Public**, **Private**, and **Const**.

## See also


#### Concepts


[Writing a Sub Procedure](writing-a-sub-procedure.md)
[Declaring Constants](declaring-constants.md)
[Declaring Variables](declaring-variables.md)

