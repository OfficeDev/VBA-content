---
title: Named argument already specified
keywords: vblr6.chm1011224
f1_keywords:
- vblr6.chm1011224
ms.prod: office
ms.assetid: 8fa1e0f1-2484-8344-038c-438ab21d2b71
ms.date: 06/08/2017
---


# Named argument already specified

You can use a [named argument](vbe-glossary.md) only once in the[argument](vbe-glossary.md) list of each[procedure](vbe-glossary.md) invocation. This error has the following causes and solutions:



- You specified the same named argument more than once in a single call. For example, if the procedure  `MySub` expects the named arguments `Arg1` and `Arg2`, the following call would generate this error:
    
  ```
  Call MySub(Arg1 := 3, Arg1 := 5) 

  ```


     Remove one of the duplicate specifications.
    
- You specified the same [argument](vbe-glossary.md) both by position and with a named argument, for example:
    
  ```
  Call MySub(1, Arg1 := 3) 

  ```


    Remove one of the duplicate specifications.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

