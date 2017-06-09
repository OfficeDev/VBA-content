---
title: Cyclic reference of projects not allowed
keywords: vblr6.chm1011121
f1_keywords:
- vblr6.chm1011121
ms.prod: office
ms.assetid: 40b1af10-726c-6a66-a2c9-12cf380ac8e9
ms.date: 06/08/2017
---


# Cyclic reference of projects not allowed

Circular references to [projects](vbe-glossary.md) aren't allowed. For example, if `MyProj` references `YourProj`, then  `YourProj` (or a project that references `YourProj`) can't reference references  `YourProj`, then  `YourProj` (or a project that references `YourProj`) can't reference  `MyProj`. This error has the following cause and solution:



- You tried to add a reference to a project that is already part of the project. Remove the circular reference or references.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

