---
title: Method not valid without suitable object
keywords: vblr6.chm1057032
f1_keywords:
- vblr6.chm1057032
ms.prod: office
ms.assetid: a7dc857a-e803-35d1-d7df-d2b9a3c79652
ms.date: 06/08/2017
---


# Method not valid without suitable object

Not all [methods](vbe-glossary.md) can be performed by all objects. This error has the following cause and solution:



- You called a method without specifying an object, and the method isn't valid for the implicit object. For example, you can't use the  **Line** method in a[standard module](vbe-glossary.md) without a valid object qualifier because a standard module can't display the output of the **Line** method.
    
    Explicitly qualify the method call with an object that can accept the method. For example, you can specify a form or picture box with the  **Line** method.
    
     **Note**  Other methods that need an explicit object qualifier when used in a standard module include  **Circle**, **Print**, and **PSet**.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

