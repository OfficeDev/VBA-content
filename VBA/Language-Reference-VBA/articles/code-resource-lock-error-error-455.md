---
title: Code resource lock error (Error 455)
keywords: vblr6.chm1011111
f1_keywords:
- vblr6.chm1011111
ms.prod: office
ms.assetid: 17e5089a-2578-f40e-7147-c87fedfa50d8
ms.date: 06/08/2017
---


# Code resource lock error (Error 455)

This error can only occur on the Macintosh. When you access a code resource, you must lock it. This error has the following cause and solution:



- A call was made to a [procedure](vbe-glossary.md) in a code resource. The code resource was found, but an error occurred when an attempt was made to lock the resource.
    
    Check for an error returned by Hlock, for example, " `Illegal on empty handle`" or " `Illegal on free block`".
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

