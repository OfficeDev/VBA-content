---
title: Invalid type-declaration character
keywords: vblr6.chm1011193
f1_keywords:
- vblr6.chm1011193
ms.prod: office
ms.assetid: 6c6411c0-6ed1-3cdb-061b-563ed3b91766
ms.date: 06/08/2017
---


# Invalid type-declaration character

[Type-declaration characters](vbe-glossary.md) are valid, but don't exist for all[data types](vbe-glossary.md); they aren't permitted in some situations. This error has the following causes and solutions:



- A type-declaration character is appended to a [variable](vbe-glossary.md) declared in a **Private**, **Public**, or **Static** statement with an **As** clause.
    
    Remove the type-declaration character.
    
- A type-declaration character is appended to an inconsistent literal. For example, since the ampersand ( **&;** ) is the type-declaration character for a **Long** integer, appending it to a literal of a different type causes this error:
    
  ```
  10.253&; 

  ```


     Remove the type-declaration character or replace it with the correct one.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

