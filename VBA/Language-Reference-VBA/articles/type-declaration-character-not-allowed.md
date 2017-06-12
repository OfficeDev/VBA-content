---
title: Type-declaration character not allowed
keywords: vblr6.chm1040101
f1_keywords:
- vblr6.chm1040101
ms.prod: office
ms.assetid: 83b717bb-e16c-f205-9c94-c8dda735a8a1
ms.date: 06/08/2017
---


# Type-declaration character not allowed

While using [type-declaration characters](vbe-glossary.md) is valid in Visual Basic, some[data types](vbe-glossary.md) (including **Byte**, **Boolean**, **Date**, **Object**, and **Variant** ) have no associated type-declaration characters. This error has the following causes and solutions:



- You tried to use a type-declaration character in the declaration of a [variable](vbe-glossary.md) that uses the **As** clause, for example, with **Dim**, **Static**, **Public**, and so on.
    
    Either remove the type-declaration character or remove the  **As** clause.
    
- You tried to use a [type-declaration character](vbe-glossary.md) in reference to a variable that was implicitly declared without a type-declaration character:
    
  ```
  MyVar = 20    ' Implicit declaration. 
MyVar% = 25    ' Generates an error. 

  ```


```vb
Dim MyStr$  
MyStr = "Because it was explicitly declared, the $ is optional." 

  ```


    Either remove the type-declaration character or redeclare the original variable.
    
     **Note**  If an explicit variable declaration contains a type-declaration character, inclusion of the character is optional in later references. For example:

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

