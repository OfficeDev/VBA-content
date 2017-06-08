---
title: Invalid use of AddressOf operator
keywords: vblr6.chm1107785
f1_keywords:
- vblr6.chm1107785
ms.prod: office
ms.assetid: 96ce20a6-175e-a006-f0fe-98080d630c7f
ms.date: 06/08/2017
---


# Invalid use of AddressOf operator

The  **AddressOf** operator modifies an[argument](vbe-glossary.md) to pass the address of a function rather than passing the result of the function call. This error has the following cause and solution:



- You tried to use  **AddressOf** with the name of a class method. Only the names of Visual Basic procedures in a .bas module can be modified with **AddressOf**. You can't specify a class method.
    
- The procedure name modified by  **AddressOf** is defined in a[module](vbe-glossary.md) in a different[project](vbe-glossary.md).
    
- You tried to modify the name a DLL function or a function defined in a [type library](vbe-glossary.md) with **AddressOf**.
    
- DLL and type library functions can't be modified with  **AddressOf**. The procedure definition must be in a module in the current project. Move the definition to a module in this project or include its current module in the project.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

