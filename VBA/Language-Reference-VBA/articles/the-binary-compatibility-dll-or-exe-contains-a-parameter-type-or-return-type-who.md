---
title: The binary compatibility DLL or EXE contains a parameter type or return type whose definition cannot be found
keywords: vblr6.chm1040374
f1_keywords:
- vblr6.chm1040374
ms.prod: office
ms.assetid: 4c13f8e8-76ba-9b7e-c6a2-8501e7bfbfd2
ms.date: 06/08/2017
---


# The binary compatibility DLL or EXE contains a parameter type or return type whose definition cannot be found

If you have a Binary Compatible server which contains a parameter or return type that is contained in another DLL, you must be careful when recompiling it. This warning has the following cause and solution:



- When you set Binary Compatibility on a project and then recompile the project, Project Compatibility is set automatically, changing the interface's internal GUID. Since this is not a visible change, this can be an unexpected error. Basically, this error occurs when a project's binary compatible DLL or EXE has a typelib with a broken reference. Broken references can occur when a referenced typelib is overwritten by another file (such as a re-compiled DLL/EXE), when you delete the typelib file, or when you move a referencing typelib over to a machine, but either don't move the referenced typelib or don't register the referenced typelib. One possible fix is to obtain a copy of the referenced typelib onto your machine and register it. You won't be able to use the old one because it was overwritten on recompile. Failing this, all that can be done is to stop using the DLL/EXE as your binary compatible version.
    


