---
title: Variable required. Can't assign to this expression
keywords: vblr6.chm1011215
f1_keywords:
- vblr6.chm1011215
ms.prod: office
ms.assetid: bf661c7f-3cda-7e89-8952-e565397a8987
ms.date: 06/08/2017
---


# Variable required. Can't assign to this expression

This error typically occurs when you attempt to assign a value to something that can't accept the assignment. This error has the following causes and solutions:



- You attempted to use a [numeric expression](vbe-glossary.md) as an[argument](vbe-glossary.md) to the **Len** function.
    
    The  **Len** function doesn't accept a numeric expression, a numeric literal, or a binary numeric expression, but it does accept either a string or numeric variable, a[string expression](vbe-glossary.md), or a [variable](vbe-glossary.md) of[user-defined type](vbe-glossary.md).
    
- You used a function call or an [expression](vbe-glossary.md) as an argument to **Input #**, **Let**, **Get**, or **Put**. For example, you may have used an argument that appears to be a valid reference to an[array](vbe-glossary.md) variable, but instead is a call to a function of the same name.
    
     **Input #**, **Let**, **Get**, and **Put** don't accept function calls as arguments.
    
- You attempted to assign a value to an [identifier](vbe-glossary.md) previously declared as a[constant](vbe-glossary.md).
    
    Choose another name for the identifier.
    
- You tried to use a nonvariable as a loop counter in a  **For...Next** construction. Use a variable as the counter.
    
- You tried to assign a value to a read-only [property](vbe-glossary.md) or to an expression that consists of more than one variable (such as X + Y). An assignment places a value at a memory location. The specified expression must represent a single, writable location.
    
    Rewrite the assignment to a single variable name that can accept the data.
    
- You tried to use an undeclared variable that is defined as a constant in a [type library](vbe-glossary.md).
    
    Either use a different name for the variable, or declare it explicitly.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

