---
title: Visual Basic Naming Rules
keywords: vbcn6.chm1076688
f1_keywords:
- vbcn6.chm1076688
ms.prod: office
ms.assetid: d3e2b5d5-ac45-a1e0-9935-3630fd033a7d
ms.date: 06/08/2017
---


# Visual Basic Naming Rules

Use the following rules when you name [procedures](vbe-glossary.md), [constants](vbe-glossary.md), [variables](vbe-glossary.md), and [arguments](vbe-glossary.md) in a Visual Basic[module](vbe-glossary.md):



- You must use a letter as the first character.
    
- You can't use a space, period ( **.** ), exclamation mark ( **!** ), or the characters **@**, **&;**, **$**, **#** in the name.
    
- Name can't exceed 255 characters in length.
    
- Generally, you shouldn't use any names that are the same as the [functions](vbe-glossary.md), [statements](vbe-glossary.md), and [methods](vbe-glossary.md) in Visual Basic. You end up shadowing the same[keywords](vbe-glossary.md) in the language. To use an intrinsic language function, statement, or method that conflicts with an assigned name, you must explicitly identify it. Precede the intrinsic function, statement, or method name with the name of the associated[type library](vbe-glossary.md). For example, if you have a variable called  `Left`, you can only invoke the  **Left** function using `VBA.Left`.
    
- You can't repeat names within the same level of [scope](vbe-glossary.md). For example, you can't declare two variables named  `age` within the same procedure. However, you can declare a private variable named `age` and a[procedure-level](vbe-glossary.md) variable named `age` within the same module.
    
     **Note**  Visual Basic isn't case-sensitive, but it preserves the capitalization in the statement where the name is declared.


