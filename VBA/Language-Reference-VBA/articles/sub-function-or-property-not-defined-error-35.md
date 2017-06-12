---
title: Sub, Function, or Property not defined (Error 35)
keywords: vblr6.chm1011291
f1_keywords:
- vblr6.chm1011291
ms.prod: office
ms.assetid: 3f770754-8929-b15e-5bcc-d07fb2c353f4
ms.date: 06/08/2017
---


# Sub, Function, or Property not defined (Error 35)

A  **Sub**, **Function**, or **Property** procedure must be defined to be called. This error has the following causes and solutions:



- You misspelled the name of your [procedure](vbe-glossary.md).
    
    Check the spelling and correct it.
    
- You tried to call a procedure from another [project](vbe-glossary.md) without explicitly adding a reference to that project in the **References** dialog box.
    
     **To add a reference**
    
    
    
      1. Display the  **References** dialog box.
    
  2. Find the name of the project containing the procedure you want to call. If the project name doesn't appear in the  **References** dialog box, click the **Browse** button to search for it.
    
  3. Click the check box to the left of the project name.
    
  4. Click  **OK**.
    

    
    
- The specified procedure isn't visible to the calling procedure. Procedures declared  **Private** in one[module](vbe-glossary.md) can't be called from procedures outside the module. If **Option Private Module** is in effect, procedures in the module aren't available to other projects. Search to locate the procedure.
    
- You declared a Windows [dynamic-link library (DLL)](vbe-glossary.md) routine or Macintosh code-resource routine, but the routine isn't in the specified library or code resource.
    
- Check the ordinal (if you used one) or the name of the routine. Make sure your version of the DLL or Macintosh code-resource is the correct one. The routine may only exist in later versions of the DLL or Macintosh code-resource. If the directory containing the wrong version precedes the directory containing the correct one in your path, the wrong DLL or Macintosh code-resource is accessed. You gave the right DLL name or Macintosh code-resource, but it isn't the version that contains the specified function.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

