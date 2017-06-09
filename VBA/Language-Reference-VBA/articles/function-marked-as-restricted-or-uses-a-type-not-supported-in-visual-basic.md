---
title: Function marked as restricted or uses a type not supported in Visual Basic
keywords: vblr6.chm1035026
f1_keywords:
- vblr6.chm1035026
ms.prod: office
ms.assetid: b013d6ca-2e99-f2c9-d64b-87ef0990493d
ms.date: 06/08/2017
---


# Function marked as restricted or uses a type not supported in Visual Basic

Not every [procedure](vbe-glossary.md) that appears in a[type library](vbe-glossary.md) or[object library](vbe-glossary.md) can be accessed by every programming language. The creator of a type or object library can designate some functions as restricted to prevent their use by macro languages. This error has the following causes and solutions:



- You tried to use a function with a restricted specification. You can't use the function in your program. If you have documentation for the object represented by the library, check to see if a [method](vbe-glossary.md) is provided that gives equivalent functionality.
    
- You tried to use a function that requires a [parameter](vbe-glossary.md) type or has a return type that isn't available in Visual Basic.
    
    Sometimes you can simulate return types with Visual Basic equivalents. Check the subtypes of the [Variant data type](vbe-glossary.md) . This may also work for non-Basic parameter types that are expected as references. However, you can't pass a **Variant** data type[by value](vbe-glossary.md) in an effort to simulate a non-Basic type.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

