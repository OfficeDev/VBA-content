---
title: Qualified name disallowed in module scope
keywords: vblr6.chm1011256
f1_keywords:
- vblr6.chm1011256
ms.prod: office
ms.assetid: 463cafc7-1af6-95b3-ee63-1681a82fb4ac
ms.date: 06/08/2017
---


# Qualified name disallowed in module scope

Under some circumstances, some [host applications](vbe-glossary.md) don't permit[procedure](vbe-glossary.md) calls that include qualified names. This error has the following cause and solution:



- You specified a [module](vbe-glossary.md) name in a procedure call using dot notation ( _qualifier_. _item_ ).
    
    If you are receiving this error it is probably because the host application already knows the specified qualifier and doesn't need that information in the procedure call. In such a case, you can simply omit the qualifier altogether and the host application will make the procedure call correctly. Check the host application's documentation to find the reason for any other restrictions on qualified names.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

