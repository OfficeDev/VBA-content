---
title: "Compile error in hidden module: <module name>"
keywords: vblr6.chm1011113
f1_keywords:
- vblr6.chm1011113
ms.prod: office
ms.assetid: 14deea0e-46ae-bcbd-b1c4-2363c90365f9
ms.date: 06/08/2017
---


# Compile error in hidden module: <module name>

A protected [module](vbe-glossary.md) contains a compilation error. Because the error is in a protected module it cannot be displayed.

This error commonly occurs when code is incompatible with the version or architecture of this application (for example, code in a document targets 32-bit Microsoft Office applications but it is attempting to run on 64-bit Office).

 This error has the following cause and solution:

Cause of the error:


- The error is raised when a compilation error exists in the VBA code inside a protected (hidden) module. The specific compilation error is not exposed because the module is protected.
    

Possible solutions:

- If you have access to the VBA code in the document or project, unprotect the module, and then run the code again to view the specific error.
    
- If you do not have access to the VBA code in the document, then contact the document author to have the code in the hidden module updated.
    

