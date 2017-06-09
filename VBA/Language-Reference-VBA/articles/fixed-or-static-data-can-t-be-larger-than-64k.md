---
title: Fixed or static data can't be larger than 64K
keywords: vblr6.chm1011320
f1_keywords:
- vblr6.chm1011320
ms.prod: office
ms.assetid: e41c3342-6ea7-38d9-be17-f058858ec006
ms.date: 06/08/2017
---


# Fixed or static data can't be larger than 64K

Fixed and static data include nonautomatic [variables](vbe-glossary.md), fixed-length strings, and fixed [arrays](vbe-glossary.md). This error has the following causes and solutions:



- You attempted to allocate more than 64K of [module-level](vbe-glossary.md) data.
    
    Reduce the amount of declared data. Note that although the size limit for module-level data is 64K, module-level variable-length strings and arrays can exceed this limit.
    
- You attempted to allocate more than 64K of static [procedure-level](vbe-glossary.md) data in the[module](vbe-glossary.md).
    
    Reduce the amount of this type of data declared. Static data from all [procedures](vbe-glossary.md) in a module is limited to a total of 64K (not 64K per procedure). Note that static variable-length strings and arrays can exceed this limit.
    
- The size of a [user-defined type](vbe-glossary.md) exceeds 64K.
    
    Reduce the size of the user-defined type. Generally the size of a user-defined type equals the sum of the sizes specified for its elements. On some platforms there may be padding between the elements to keep them aligned on word boundaries. If you nest one user-defined type in another, the size of the nested type must be included in the size of the new type.
    
- In a procedure, you tried to declare a variable of user-defined type that requires more than 32K. Although the size limit of a variable of user-defined type is 64K at module level, variables of user-defined type in procedures can't exceed 32K. Reduce the size required for the user-defined type, or use a module-level variable.
    
- The size of a fixed-length string declared within a procedure exceeds 65,464. Reduce the length of the fixed-length string. Note that variable-length strings can exceed this limit.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

