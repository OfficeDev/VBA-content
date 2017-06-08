---
title: Duplicate declaration in current scope
keywords: vblr6.chm1011221
f1_keywords:
- vblr6.chm1011221
ms.prod: office
ms.assetid: 55b4c056-c787-b30a-bc8c-30e552a3cdbb
ms.date: 06/08/2017
---


# Duplicate declaration in current scope

The specified name is already used at this level of [scope](vbe-glossary.md). For example, two [variables](vbe-glossary.md) can have the same name if they are defined in different[procedures](vbe-glossary.md), but not if they are defined within the same procedure. This error has the following causes and solutions:



- A new variable or procedure has the same name as an existing variable or procedure. For example:
    
```vb
Sub MySub() 
Dim A As Integer 
Dim A As Variant 
. . .        ' Other declarations or procedure code here. 
End Sub
```


     Check the current procedure,[module](vbe-glossary.md), or [project](vbe-glossary.md) and remove any duplicate declarations.
    
- A  **Const** statement uses the same name as an existing variable or procedure. Remove or rename the[constant](vbe-glossary.md) in question.
    
- You declared a fixed [array](vbe-glossary.md) more than once.
    
    Remove or rename one of the arrays.
    

Search for the duplicate name. When specifying the name to search for, omit any [type-declaration character](vbe-glossary.md) because a conflict occurs if the names are the same and the type-declaration characters are different.
Note that a [module-level](vbe-glossary.md) variable can have the same name as a variable declared in a procedure, but when you want to refer to the module-level variable within the procedure, you must qualify it with the module name. Module names and the names of[referenced projects](vbe-glossary.md) can be reused as variable names within procedures and can also be qualified.
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

