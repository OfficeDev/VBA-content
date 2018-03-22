---
title: Object Data Type
keywords: vblr6.chm1008829
f1_keywords:
- vblr6.chm1008829
ms.prod: office
ms.assetid: cffe448d-29dd-52aa-4a5c-2155c07b5bf3
ms.date: 06/08/2017
---


# Object Data Type

[Object variables](vbe-glossary.md) are stored as 32-bit (4-byte) addresses that refer to objects. Using the **Set** statement, a variable declared as an **Object** can have any object reference assigned to it.


 **Note** Although a variable declared with **Object** type is flexible enough to contain a reference to any object, binding to the object referenced by that variable is always late ([run-time](vbe-glossary.md) binding). To force early binding ([compile-time](vbe-glossary.md) binding), assign the object reference to a variable declared with a specific [class](vbe-glossary.md) name.


