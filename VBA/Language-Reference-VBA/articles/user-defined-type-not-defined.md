---
title: User-defined type not defined
keywords: vblr6.chm1011292
f1_keywords:
- vblr6.chm1011292
ms.prod: office
ms.assetid: 60e0da5e-c498-7a2f-46c6-c09d59fc607a
ms.date: 06/08/2017
---


# User-defined type not defined

You can create your own [data types](vbe-glossary.md) in Visual Basic, but they must be defined first in a **Type...End Type** statement or in a properly registered[object library](vbe-glossary.md) or[type library](vbe-glossary.md). This error has the following causes and solutions:



- You tried to declare a [variable](vbe-glossary.md) or[argument](vbe-glossary.md) with an undefined data type or you specified an unknown[class](vbe-glossary.md) or object.
    
    Use the  **Type** statement in a[module](vbe-glossary.md) to define a new data type. If you are trying to create a reference to a class, the class must be visible to the[project](vbe-glossary.md). If you are referring to a class in your program, you must have a [class module](vbe-glossary.md) of the specified name in your project. Check the spelling of the type name or name of the object.
    
- The type you want to declare is in another module but has been declared  **Private**. Move the definition of the type to a[standard module](vbe-glossary.md) where it can be **Public**.
    
- The type is a valid type, but the object library or type library in which it is defined isn't registered in Visual Basic. Display the  **References** dialog box, and then select the appropriate object library or type library. For example, if you don't check the **Data Access Object** in the **References** dialog box, types like Database, Recordset, and TableDef aren't recognized and references to them in code cause this error.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

