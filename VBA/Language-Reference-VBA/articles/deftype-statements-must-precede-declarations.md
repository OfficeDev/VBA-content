---
title: Deftype statements must precede declarations
keywords: vblr6.chm1040057
f1_keywords:
- vblr6.chm1040057
ms.prod: office
ms.assetid: 1cbcf2e1-cd74-7d92-2d7a-2b6c3086e89a
ms.date: 06/08/2017
---


# Deftype statements must precede declarations

 **Def**_type_ statements include **DefInt**, **DefDbl**, **DefCur**, and so on. This error has the following causes and solutions:



- A [variable](vbe-glossary.md)[declaration](vbe-glossary.md) precedes a **Def**_type_ statement at[module level](vbe-glossary.md).
    
    Move the  **Def**_type_ statement to precede all variable declarations.
    
- A  **Def**_type_ statement appears in a[procedure](vbe-glossary.md).
    
    Move the  **Def**_type_ statement to module level, preceding all variable declarations.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

