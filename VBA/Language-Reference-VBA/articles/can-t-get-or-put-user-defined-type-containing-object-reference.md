---
title: Can't Get or Put user-defined type containing object reference
keywords: vblr6.chm1011399
f1_keywords:
- vblr6.chm1011399
ms.prod: office
ms.assetid: 138dcd9a-75a5-3ded-a6ed-9d2fae2c9b14
ms.date: 06/08/2017
---


# Can't Get or Put user-defined type containing object reference

An object reference is temporary and can easily become invalid between closing and opening a file. This error has the following cause and solution:



- The [variable](vbe-glossary.md) in your **Get** or **Put** statement contains, or is declared to contain, a reference to an object.
    
    If the variable is an object reference you can't use it with  **Get** and **Put** statements. To place the value of some or all of the object's[properties](vbe-glossary.md) in the file, each property must be individually specified.
    
- The [user-defined type](vbe-glossary.md) variable in your **Get** or **Put** statement contains an element that is an object reference.
    
    If the variable's  **Type** statement contains an element representing an object (for example, it is defined in a[class module](vbe-glossary.md), has [Object data type](vbe-glossary.md), is a form or a control, and so on), remove it from the definition, or define a new type for use with the  **Get** and **Put** statements that has no **Object** type element in its definition.
    
    If you have elements in the user-defined type with  **Variant** type, make sure no object reference is assigned to that element. A **Variant** can accept such an assignment, but will cause this error if its user-defined type is used in a **Get** or **Put**.
    
    Note that you can use  **Input #**, **Line Input #**, **Print #**, or **Write #** to write the default property of an object to disk.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

