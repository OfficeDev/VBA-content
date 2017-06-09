---
title: Syntax error
keywords: vblr6.chm1011279
f1_keywords:
- vblr6.chm1011279
ms.prod: office
ms.assetid: ca84aa92-e41a-1167-ab66-032ab9626005
ms.date: 06/08/2017
---


# Syntax error

Visual Basic can't determine what action to take. This error has the following causes and solutions:



- A [keyword](vbe-glossary.md) or[argument](vbe-glossary.md) is misspelled.
    
    Keywords and the names of [named arguments](vbe-glossary.md) must exactly match those specified in their syntax specifications. Check online Help, and then correct the spelling.
    
- Punctuation is incorrect. For example, when you omit optional arguments positionally, you must substitute a comma ( **,** ) as a placeholder for the omitted argument.
    
- A [procedure](vbe-glossary.md) isn't defined.
    
    Check the spelling of the procedure name.
    
- You tried to specify both  **Optional** and **ParamArray** in the same procedure declaration. A **ParamArray** argument can't be **Optional**. Choose one and delete the other.
    
- You tried to define an event procedure with an  **Optional** or **ParamArray** parameter. Remove the **Optional** or **ParamArray** keyword from the parameter specification.
    
- You tried to use a named argument in a  **RaiseEvent** statement. Events do not support named arguments.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

