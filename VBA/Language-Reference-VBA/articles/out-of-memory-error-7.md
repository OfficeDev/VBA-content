---
title: Out of memory (Error 7)
keywords: vblr6.chm1011242
f1_keywords:
- vblr6.chm1011242
ms.prod: office
ms.assetid: b04a1604-738c-2425-1d4b-a5c595cd798d
ms.date: 06/08/2017
---


# Out of memory (Error 7)

More memory was required than is available, or a 64K segment boundary was encountered. This error has the following causes and solutions:



- You have too many applications, documents, or source files open. Close any unnecessary applications, documents, or source files that are open.
    
- You have a [module](vbe-glossary.md) or [procedure](vbe-glossary.md) that's too large. Break large modules or procedures into smaller ones. This doesn't save memory, but it can prevent hitting 64K segment boundaries.
    
- You are running Microsoft Windows in standard mode. Restart Microsoft Windows in enhanced mode.
    
- You are running Microsoft Windows in enhanced mode, but have run out of virtual memory. Increase virtual memory by freeing some disk space, or at least ensure that some space is available.
    
- You have terminate-and-stay-resident programs running. Eliminate terminate-and-stay-resident programs.
    
- You have many device drivers loaded. Eliminate unnecessary device drivers.
    
- You have run out of space for **Public** [variables](vbe-glossary.md). Reduce the number of  **Public** variables.
    
- You have attempted to update a property that is read-only. Do not assign values to read-only properties.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

