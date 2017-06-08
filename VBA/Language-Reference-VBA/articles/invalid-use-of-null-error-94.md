---
title: Invalid use of Null (Error 94)
keywords: vblr6.chm1000094
f1_keywords:
- vblr6.chm1000094
ms.prod: office
ms.assetid: c1c987fb-8b4c-bbc2-a69b-c5e9047bb94a
ms.date: 06/08/2017
---


# Invalid use of Null (Error 94)

[Null](vbe-glossary.md) is a **Variant** subtype used to indicate that a data item contains no valid data. This error has the following cause and solution:



- You are trying to obtain the value of a  **Variant** variable or an[expression](vbe-glossary.md) that is **Null**. For example:
    
  ```
  MyVar = Null 
For Count = 1 To MyVar 
. . . 
Next Count 

  ```


    Make sure the [variable](vbe-glossary.md) contains a valid value.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

