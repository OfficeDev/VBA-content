---
title: Argument not optional (Error 449)
keywords: vblr6.chm1011248
f1_keywords:
- vblr6.chm1011248
ms.prod: office
ms.assetid: 04d08e66-7084-8c94-52b1-b471423846ca
ms.date: 06/08/2017
---


# Argument not optional (Error 449)

The number and types of [arguments](vbe-glossary.md) must match those expected. This error has the following causes and solutions:



- Incorrect number of arguments. Supply all necessary arguments. For example, the  **Left** function requires two arguments; the first representing the character string being operated on, and the second representing the number of characters to return from the left side of the string. Because neither argument is optional, both must be supplied.
    
- Omitted argument isn't optional. An argument can only be omitted from a call to a user-defined [procedure](vbe-glossary.md) if it was declared **Optional** in the procedure declaration. Either supply the argument in the call or declare the[parameter ](vbe-glossary.md) **Optional** in the definition.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

