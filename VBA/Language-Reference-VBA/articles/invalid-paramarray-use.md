---
title: Invalid ParamArray use
keywords: vblr6.chm1040132
f1_keywords:
- vblr6.chm1040132
ms.prod: office
ms.assetid: 791f4e2b-c37e-6e68-e5f6-5ef258d4fab0
ms.date: 06/08/2017
---


# Invalid ParamArray use

The [parameter](vbe-glossary.md) defined as **ParamArray** is used incorrectly in the[procedure](vbe-glossary.md). This error has the following causes and solutions:



- You attempted to pass  **ParamArray** as an[argument](vbe-glossary.md) to another procedure that expects an[array](vbe-glossary.md) or a **ByRef Variant**.
    
    Assign the  **ParamArray** parameter to a **Variant**, and then pass the variant.
    
- You attempted to use an  **Erase** or **ReDim** statement with a **ParamArray** parameter within its procedure. Remove the **Erase** or **ReDim**. These operations can't be performed on the **ParamArray** parameter.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

