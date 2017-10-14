---
title: Definitions of property procedures for the same property are inconsistent
keywords: vblr6.chm1019369
f1_keywords:
- vblr6.chm1019369
ms.prod: office
ms.assetid: 0dbd5698-d475-fdc5-ce9a-803835530afa
ms.date: 06/08/2017
---


# Definitions of property procedures for the same property are inconsistent

The [parameters](vbe-glossary.md) for **Property Get**, **Property Let**, and **Property Set**[procedures](vbe-glossary.md) for the same[property](vbe-glossary.md) must match exactly, except that the **Property Let** has one extra parameter, whose type must match the return type of the corresponding **Property Get**, and the **Property Set** has one more parameter than the corresponding **Property Get**, whose type is either **Variant**, **Object**, a[class](vbe-glossary.md) name, or an object library type specified in an[object library](vbe-glossary.md). This error has the following causes and solutions:



- The number of parameters for the  **Property Get** procedure isn't one less than the number of parameters for the matching **Property Let** or **Property Set** procedure. Add a parameter to **Property Let** or **Property Set** or remove a parameter from **Property Get**, as appropriate.
    
- The parameter types of  **Property Get** must exactly match the corresponding parameters of **Property Let** or **Property Set**, except for the extra **Property Set** parameter. Modify the parameter declarations in the corresponding procedure definitions so they are appropriately matched.
    
- The parameter type of the extra parameter of the  **Property Let** must match the return type of the corresponding **Property Get** procedure. Modify either the extra parameter declaration in the **Property Let** or the return type of the corresponding **Property Get** so they are appropriately matched.
    
- The parameter type of the extra parameter of the  **Property Set** can differ from the return type of the corresponding **Property Get**, but it must be either a **Variant**, **Object**,[class](vbe-glossary.md) name, or a valid[object library](vbe-glossary.md) type.
    
    Make sure the extra parameter of the  **Property Set** procedure is either a **Variant**, **Object**, class name, or object library type.
    
- You defined a  **Property** procedure with an **Optional** or a **ParamArray** parameter. **ParamArray** and **Optional** parameters aren't permitted in **Property** procedures. Redefine the procedures without using these[keywords](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

