---
title: Function call on left-hand side of assignment must return Variant or Object
keywords: vblr6.chm1011177
f1_keywords:
- vblr6.chm1011177
ms.prod: office
ms.assetid: 5c0b6c52-ab00-1c1b-96f8-7dfb3fcb749e
ms.date: 06/08/2017
---


# Function call on left-hand side of assignment must return Variant or Object

A function call can appear on the left side of an assignment, but only if the return value of the function is an  **Object** or **Variant**. This error has the following cause and solution:



- The return type of the function on the left side of the assignment isn't a  **Variant** or **Object**. Change the return type. Note that if the return value is an object or a **Variant** that contains an object, the assignment is to the default[property](vbe-glossary.md) of the object. If the **Variant** returned isn't an object, the assignment has no effect.
    
- Everything in the call is correct, however, it can't be completed. For example, you may be trying to set a property that can only be set at design time. Enter design mode and set the property in the  **Property** window. Remove the code that tried to set the property programmatically.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

