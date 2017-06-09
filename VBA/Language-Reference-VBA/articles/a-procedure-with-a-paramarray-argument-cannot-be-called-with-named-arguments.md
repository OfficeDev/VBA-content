---
title: A procedure with a ParamArray argument cannot be called with named arguments
keywords: vblr6.chm1040130
f1_keywords:
- vblr6.chm1040130
ms.prod: office
ms.assetid: 59cbcba9-b3bf-5e5d-1002-5529fa6226ad
ms.date: 06/08/2017
---


# A procedure with a ParamArray argument cannot be called with named arguments

All [arguments](vbe-glossary.md) in a call to a[procedure](vbe-glossary.md) defined with a **ParamArray** must be positional. This error has the following cause and solution:


- [Named-argument](vbe-glossary.md) syntax appears in a procedure call.
    
    The named-argument calling syntax can't be used to call a procedure that includes a  **ParamArray** parameter. To supply only some elements of the **ParamArray**, use commas as placeholders for those elements you want to omit. For example, in the following call, if the **ParamArray** arguments begin after `Arg2`, values are being passed only for the first, third, and sixth values in the  **ParamArray**:
    


  ```
  MySub Arg1, Arg2, 7,, 44,,,3 
  ```


     **Note**  The  **ParamArray** always represents the last items in the argument list.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


