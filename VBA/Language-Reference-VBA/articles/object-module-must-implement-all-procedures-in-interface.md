---
title: Object module must implement all procedures in interface
keywords: vblr6.chm1109544
f1_keywords:
- vblr6.chm1109544
ms.prod: office
ms.assetid: 9b8ccb3a-92e3-20d8-1263-0425c53286a5
ms.date: 06/08/2017
---


# Object module must implement all procedures in interface

An interface is a collection of unimplemented [procedure](vbe-glossary.md) prototypes. This error has the following cause and solution:



- You specified an interface in an  **Implements** statement, but you didn't add code for all the procedures in the interface. You must write code for each of the procedures specified in the interface. An empty procedure is adequate; the procedure should implement the required behavior.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

