---
title: Invalid attribute in Sub, Function, or Property
keywords: vblr6.chm1011196
f1_keywords:
- vblr6.chm1011196
ms.prod: office
ms.assetid: 86a5ff38-4f00-060f-5087-453758f27e68
ms.date: 06/08/2017
---


# Invalid attribute in Sub, Function, or Property

Some attributes are invalid within [procedures](vbe-glossary.md). This error has the following cause and solution:



- A  **Public** or **Private** attribute appears within the body of a procedure definition. Remove the attribute from the procedure. To give the[variable](vbe-glossary.md) wider[scope](vbe-glossary.md), move the declaration to [module level](vbe-glossary.md). Variables declared within procedures are always  **Private**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

