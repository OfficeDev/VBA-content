---
title: Expected Function or variable
keywords: vblr6.chm1040076
f1_keywords:
- vblr6.chm1040076
ms.prod: office
ms.assetid: c4f8e6fb-43b7-3dcd-c93a-7f9b2e542817
ms.date: 06/08/2017
---


# Expected Function or variable

The syntax of your statement indicates a [variable](vbe-glossary.md) or function call. This error has the following cause and solution:


- The name isn't that of a known variable or  **Function** procedure.
    
    Check the spelling of the name. Make sure that any variable or function with that name is visible in the portion of the program from which you are referencing it. For example, if a function is defined as  **Private** or a variable isn't defined as **Public**, it's only visible within its own[module](vbe-glossary.md).
    
- You are trying to inappropriately assign a value to a [procedure](vbe-glossary.md) name.
    
    For example if  `MySub` is a **Sub** procedure, the following code generates this error:
    


  ```
  MySub = 237    ' Causes Expected Function or variable error
  ```


    Although you can use assignment syntax with a  **Property Let** procedure or with a **Function** that returns an object or a **Variant** containing an object, you can't use assignment syntax with a **Sub**, **Property Get**, or **Property Set** procedure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


