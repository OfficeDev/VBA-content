---
title: Destination label too far away; loop, Select Case, or block If too large
keywords: vblr6.chm1011341
f1_keywords:
- vblr6.chm1011341
ms.prod: office
ms.assetid: 56644b8d-3a38-874d-1e5e-0091bcd86f0b
ms.date: 06/08/2017
---


# Destination label too far away; loop, Select Case, or block If too large

[Procedures](vbe-glossary.md) can be as large as 64K from beginning to end, but because branching can occur either forward or backward within a procedure, such branching is limited to 32,767 bytes in either direction. This error has the following causes and solutions:



- You have a branching statement ( **GoTo**, **GoSub** ) whose destination label is farther away than 32,767 bytes from the source branching statement. Move the label closer, or make the procedure smaller.
    
- You have a very large loop structure that occupies more than 32K of memory from beginning to end. Make the loop smaller.
    
- You have a very large block  **If** structure that contains a **Then** or **Else** clause that occupies more than 32K of memory from beginning to end. Reduce the size of the offending portion of the structure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

