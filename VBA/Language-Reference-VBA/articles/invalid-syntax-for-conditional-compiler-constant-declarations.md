---
title: Invalid syntax for conditional compiler constant declarations
keywords: vblr6.chm1040354
f1_keywords:
- vblr6.chm1040354
ms.prod: office
ms.assetid: 815e6833-9813-5341-838d-55e0b4a4aae5
ms.date: 06/08/2017
---


# Invalid syntax for conditional compiler constant declarations

Entering [conditional compiler constants](vbe-glossary.md) in an **Options** dialog box differs from declaring[constants](vbe-glossary.md) in code. This error has the following cause and solution:



- You used improper syntax when entering a constant declaration in the in an  **Options** dialog box. The only valid syntax is a simple assignment of an integer value to the[identifier](vbe-glossary.md). Make sure the syntax for the entry is as follows, with each constant separated by a colon ( **:** ):
    
  ```
  constantname= [{+ | - }]integervalue: [{+ | - }]constantname=integervalue  [...] 

  ```


    
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

