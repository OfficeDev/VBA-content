---
title: Invalid use of base class name
ms.prod: office
ms.assetid: 67c2bc6c-5717-450c-b80a-3217fee436ed
ms.date: 06/08/2017
---


# Invalid use of base class name
You cannot use the name of a base class by itself. This error has the following causes and solutions:


- You tried to use the name of a base class by itself without making clear that you were trying to access the base class' default member. Place the base-class name within parentheses to indicate you want to access the default member.
    
- You used the base-class name in an expression but the member you were trying to access was ambiguously specified. Use a disambiguator ( for example, an exclamation point) between the base-class name and the member you are interested in.
    
- You used the base-class name in a  **Set** statement as though it contained a reference to the class. Use the base-class name to retrieve a reference for example, using GetObject.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

